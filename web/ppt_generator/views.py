"""
Views for PPT Generator application.
"""

import sys
import json
import traceback
import uuid
from pathlib import Path
from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.http import require_http_methods
from django.contrib.auth.decorators import login_required, permission_required
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.conf import settings
from django.core.files.base import ContentFile
from django.utils import timezone

from .models import GlobalLLMConfig, PPTGeneration, TemplateEditSession
from .forms import PPTGenerationForm

# Add parent directory to path to import S2S modules
sys.path.insert(0, str(settings.BASE_DIR.parent))
from scripts.docx_to_config import generate_config_data
from scripts.generate_slides import render_slides


def _guess_template_json(template_path: Path) -> Path:
    """根据模板 PPT 路径自动推断对应的 template.json 配置文件路径。

    优先级：
    1. 与 PPT 同名的 JSON（同文件夹）：<folder>/<stem>.json
    2. 同文件夹下的 template.json
    3. 全局模板目录下的 template.json（兼容老行为）
    """
    base_dir = settings.S2S_TEMPLATE_DIR

    candidates = []

    # 1) 同名 JSON：template1/template.pptx -> template1/template.json
    candidates.append(template_path.with_suffix(".json"))

    # 2) 同目录下的 template.json
    candidates.append(template_path.parent / "template.json")

    # 3) 全局默认 template.json
    candidates.append(base_dir / "template.json")

    for candidate in candidates:
        if candidate.exists():
            return candidate

    # 如果都不存在，最后仍然返回全局默认路径，让后续报出清晰错误
    return base_dir / "template.json"


@login_required
def index(request):
    """Main page with upload form and history."""
    # Check if user is developer
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    if request.method == "POST":
        form = PPTGenerationForm(request.POST, request.FILES)
        if form.is_valid():
            generation = form.save(commit=False)

            # Set the user
            generation.user = request.user

            # For config template: only developers can explicitly set it
            if is_developer:
                config_choice = form.cleaned_data.get("config_template_choice", "auto")
                if config_choice == "select":
                    # Use dropdown selection
                    generation.config_template = (
                        form.cleaned_data.get("config_template") or None
                    )
                    generation.config_template_file = None
                elif config_choice == "upload":
                    # Use uploaded file (handled by ModelForm)
                    generation.config_template = None
                    # config_template_file is handled by ModelForm automatically
                else:
                    # Auto-match
                    generation.config_template = None
                    generation.config_template_file = None
            else:
                generation.config_template = None
                generation.config_template_file = None

            # Set template name based on choice
            template_choice = form.cleaned_data.get("template_choice")
            if template_choice == "preset":
                # Store the relative path to the selected preset template
                generation.template_name = form.cleaned_data.get("preset_template_path")
            else:
                generation.template_name = "custom"

            # Handle LLM configuration: populate llm_provider/llm_model for display in admin
            if generation.use_llm:
                llm_config_choice = form.cleaned_data.get("llm_config_choice", "preset")
                if llm_config_choice == "preset" and generation.llm_preset_config:
                    # 使用预设配置时，将预设的 provider 和 model 复制到记录字段中
                    # 这样 Admin 界面可以正确显示
                    preset = generation.llm_preset_config
                    generation.llm_provider = preset.llm_provider
                    generation.llm_model = preset.get_model_for_provider()
                    # API Key 和 Base URL 不复制，保持安全性
                # 如果是自定义配置，字段已经由表单填充，无需额外处理

            generation.save()
            return redirect("generation_detail", pk=generation.pk)
    else:
        form = PPTGenerationForm()

    # Get recent generations for current user
    recent_generations = PPTGeneration.objects.filter(user=request.user).order_by(
        "-created_at"
    )[:10]

    # Get available templates from template directory
    template_dir = settings.S2S_TEMPLATE_DIR
    available_templates = []
    available_config_templates = []
    if template_dir.exists():
        # Scan for template.pptx in subdirectories (e.g. template/template1/template.pptx)
        # Also include template.pptx in root for backward compatibility

        # 1. Root template.pptx
        if (template_dir / "template.pptx").exists():
            available_templates.append(
                {"name": "默认模板 (template.pptx)", "path": "template.pptx"}
            )

        # 2. Subdirectories
        for subdir in template_dir.iterdir():
            if subdir.is_dir() and not subdir.name.startswith("."):
                ppt_path = subdir / "template.pptx"
                if ppt_path.exists():
                    available_templates.append(
                        {
                            "name": f"{subdir.name} (template.pptx)",
                            "path": str(subdir.name) + "/template.pptx",
                        }
                    )

        for json_file in template_dir.rglob("*.json"):
            # 使用相对路径方便前端展示和回填
            rel_path = json_file.relative_to(template_dir)
            available_config_templates.append(str(rel_path))

    context = {
        "form": form,
        "recent_generations": recent_generations,
        "available_templates": available_templates,
        "available_config_templates": available_config_templates,
        "is_developer": is_developer,
    }
    return render(request, "ppt_generator/index.html", context)


@login_required
def generation_detail(request, pk):
    """Detail page for a specific generation."""
    generation = get_object_or_404(PPTGeneration, pk=pk)

    # Check if user is developer
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    # Check if preprocessed script exists
    has_preprocessed_script = False
    if is_developer:
        run_dir = settings.S2S_TEMP_DIR / f"web-{generation.id}"
        script_path = run_dir / "preprocessed_script.md"
        has_preprocessed_script = script_path.exists()

    context = {
        "generation": generation,
        "is_developer": is_developer,
        "has_preprocessed_script": has_preprocessed_script,
    }
    return render(request, "ppt_generator/detail.html", context)


@login_required
@require_http_methods(["POST"])
def start_generation(request, pk):
    """Start PPT generation process (AJAX endpoint)."""
    generation = get_object_or_404(PPTGeneration, pk=pk)

    if generation.status != "pending":
        return JsonResponse(
            {"success": False, "error": "该任务已经开始处理或已完成"}, status=400
        )

    try:
        generation.mark_processing()

        # Determine template path
        if generation.template_file:
            template_path = Path(generation.template_file.path)
        else:
            template_path = settings.S2S_TEMPLATE_DIR / generation.template_name

        if not template_path.exists():
            raise FileNotFoundError(f"模板文件不存在: {template_path}")

        # Prepare paths
        docx_path = Path(generation.docx_file.path)

        # Determine template.json config: uploaded file > dropdown selection > auto-guess
        if generation.config_template_file:
            # Priority 1: Uploaded JSON file
            template_json = Path(generation.config_template_file.path)
        elif generation.config_template:
            # Priority 2: Dropdown selection
            template_json = settings.S2S_TEMPLATE_DIR / generation.config_template
        else:
            # Priority 3: Auto-guess based on PPTX template
            template_json = _guess_template_json(template_path)

        if not template_json.exists():
            raise FileNotFoundError(f"配置模板不存在: {template_json}")

        template_list = settings.S2S_TEMPLATE_DIR / "template.txt"

        # Create run directory
        run_dir = settings.S2S_TEMP_DIR / f"web-{generation.id}"
        run_dir.mkdir(parents=True, exist_ok=True)

        # Prepare metadata overrides
        metadata_overrides = {
            "course": generation.course_name,
            "college": generation.college_name,
            "lecturer": generation.lecturer_name,
        }

        # Prepare LLM configuration
        if generation.use_llm:
            # 判断使用预设配置还是自定义配置
            if (
                generation.llm_config_choice == "preset"
                and generation.llm_preset_config
            ):
                # 使用预设配置
                preset = generation.llm_preset_config
                llm_provider = preset.llm_provider
                llm_model = preset.get_model_for_provider()
                llm_api_key = preset.llm_api_key
                llm_base_url = preset.llm_base_url
                user_prompt = generation.user_prompt or preset.default_prompt
            else:
                # 使用自定义配置
                llm_provider = generation.llm_provider
                llm_model = generation.llm_model
                llm_api_key = generation.llm_api_key
                llm_base_url = generation.llm_base_url
                user_prompt = generation.user_prompt

                # 如果自定义配置不完整，回退到全局默认配置
                if not llm_provider or not llm_api_key:
                    global_config = GlobalLLMConfig.get_config()
                    llm_provider = llm_provider or global_config.llm_provider
                    llm_model = llm_model or global_config.get_model_for_provider()
                    llm_api_key = llm_api_key or global_config.llm_api_key
                    llm_base_url = llm_base_url or global_config.llm_base_url
                    user_prompt = user_prompt or global_config.default_prompt

        else:
            # LLM not enabled
            llm_provider = None
            llm_model = None
            llm_base_url = None
            llm_api_key = None
            user_prompt = None

        # Set API key in environment if provided
        import os

        # 清除所有 LLM 相关的环境变量，避免使用旧的 API Key
        for key in [
            "DEEPSEEK_API_KEY",
            "TAICHU_API_KEY",
            "GLM_API_KEY",
            "LOCAL_LLM_API_KEY",
        ]:
            os.environ.pop(key, None)

        # 设置当前使用的 API Key
        if llm_api_key:
            if llm_provider == "deepseek":
                os.environ["DEEPSEEK_API_KEY"] = llm_api_key
            elif llm_provider == "taichu":
                os.environ["TAICHU_API_KEY"] = llm_api_key
            elif llm_provider == "glm" or llm_provider == "zhipu":
                os.environ["GLM_API_KEY"] = llm_api_key
            elif llm_provider == "local":
                os.environ["LOCAL_LLM_API_KEY"] = llm_api_key

        # Step 1: Generate config JSON
        config_data = generate_config_data(
            docx_path=str(docx_path),
            template_json=str(template_json),
            template_list=str(template_list),
            use_llm=generation.use_llm,
            llm_provider=llm_provider,
            llm_model=llm_model,
            llm_base_url=llm_base_url,
            metadata_overrides=metadata_overrides,
            run_dir=run_dir,
            user_prompt=user_prompt,
        )

        # Save config JSON
        config_path = run_dir / "config.json"
        config_path.write_text(
            json.dumps(config_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        # Step 2: Render slides
        result = render_slides(
            template_path=template_path,
            config=config_data,
            output_name="slides.pptx",
            run_dir=run_dir,
        )

        output_path = result["output_path"]

        # Save output files to model
        with open(output_path, "rb") as f:
            generation.output_ppt.save(
                f"generation_{generation.id}.pptx", ContentFile(f.read()), save=False
            )

        with open(config_path, "rb") as f:
            generation.config_json.save(
                f"config_{generation.id}.json", ContentFile(f.read()), save=False
            )

        generation.mark_completed(
            output_path=generation.output_ppt.name,
            config_path=generation.config_json.name,
            run_dir=run_dir,
        )

        return JsonResponse(
            {
                "success": True,
                "message": "PPT生成成功！",
                "download_url": generation.output_ppt.url,
                "generation_id": generation.id,
            }
        )

    except Exception as e:
        error_msg = f"{str(e)}\n\n{traceback.format_exc()}"
        generation.mark_failed(error_msg)

        return JsonResponse(
            {
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
            status=500,
        )


@login_required
@require_http_methods(["GET"])
def check_status(request, pk):
    """Check generation status (AJAX endpoint)."""
    generation = get_object_or_404(PPTGeneration, pk=pk)

    # Check if user is developer
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    response_data = {
        "status": generation.status,
        "status_display": generation.get_status_display(),
    }

    if generation.status == "completed":
        response_data["download_url"] = generation.output_ppt.url
        # Only show config URL to developers
        if is_developer and generation.config_json:
            response_data["config_url"] = generation.config_json.url
    elif generation.status == "failed":
        response_data["error"] = generation.error_message

    return JsonResponse(response_data)


@login_required
def download_ppt(request, pk):
    """Download generated PPT file."""
    generation = get_object_or_404(PPTGeneration, pk=pk)

    # 优先从 temp 目录下载
    run_dir = settings.S2S_TEMP_DIR / f"web-{generation.id}"
    file_path = run_dir / "slides.pptx"

    # 如果 temp 不存在，尝试 media 目录（兼容旧记录）
    if not file_path.exists() and generation.output_ppt:
        file_path = Path(generation.output_ppt.path)

    if not file_path.exists():
        raise Http404("PPT文件不存在")

    response = FileResponse(
        open(file_path, "rb"),
        content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    response["Content-Disposition"] = (
        f'attachment; filename="generated_{generation.id}.pptx"'
    )
    return response


@login_required
@require_http_methods(["GET"])
def download_config_json(request, pk):
    """
    下载配置 JSON（仅管理员/开发者可用）
    """
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )
    if not is_developer:
        return JsonResponse({"error": "权限不足"}, status=403)

    generation = get_object_or_404(PPTGeneration, pk=pk)

    # 优先从 temp 目录下载
    run_dir = settings.S2S_TEMP_DIR / f"web-{generation.id}"
    file_path = run_dir / "config.json"

    # 如果 temp 不存在，尝试 media 目录（兼容旧记录）
    if not file_path.exists() and generation.config_json:
        file_path = Path(generation.config_json.path)

    if not file_path.exists():
        raise Http404("配置文件不存在")

    response = FileResponse(
        open(file_path, "rb"),
        content_type="application/json",
    )
    response["Content-Disposition"] = (
        f'attachment; filename="config_{generation.id}.json"'
    )
    return response


@login_required
@require_http_methods(["GET"])
def download_preprocessed_script(request, pk):
    """
    下载预分页讲稿（仅管理员/开发者可用）

    Args:
        pk: PPT生成记录的主键

    Returns:
        Markdown 文件下载响应
    """
    # 检查权限：仅管理员和开发者可以下载
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )
    if not is_developer:
        return JsonResponse({"error": "权限不足，仅管理员/开发者可下载"}, status=403)

    generation = get_object_or_404(PPTGeneration, pk=pk)

    # 构建预分页讲稿路径 - 使用 run_dir 而不是 config_json 路径
    run_dir = settings.S2S_TEMP_DIR / f"web-{generation.id}"
    script_path = run_dir / "preprocessed_script.md"

    if not script_path.exists():
        raise Http404("预分页讲稿不存在（可能该生成使用了带标记的讲稿）")

    response = FileResponse(
        open(script_path, "rb"),
        content_type="text/markdown; charset=utf-8",
    )
    response["Content-Disposition"] = (
        f'attachment; filename="preprocessed_script_{generation.id}.md"'
    )
    return response


@login_required
def history(request):
    """View generation history (filtered by user)."""
    # Each user can only see their own generation history
    generations = PPTGeneration.objects.filter(user=request.user).order_by(
        "-created_at"
    )

    # Check if user is developer
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    context = {
        "generations": generations,
        "is_developer": is_developer,
    }
    return render(request, "ppt_generator/history.html", context)


def user_login(request):
    """User login view."""
    if request.user.is_authenticated:
        return redirect("index")

    if request.method == "POST":
        username = request.POST.get("username")
        password = request.POST.get("password")
        user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)
            next_url = request.GET.get("next", "index")
            return redirect(next_url)
        else:
            messages.error(request, "用户名或密码错误")

    return render(request, "ppt_generator/login.html")


@login_required
def user_logout(request):
    """User logout view."""
    logout(request)
    messages.success(request, "已成功登出")
    return redirect("login")


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def developer_tools(request):
    """Developer tools for managing LLM config templates."""
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    # 获取已发布的模板列表
    published_templates = []
    template_dir = settings.S2S_TEMPLATE_DIR
    if template_dir.exists():
        for item in sorted(template_dir.iterdir()):
            if item.is_dir():
                pptx_file = item / "template.pptx"
                json_file = item / "template.json"
                if pptx_file.exists():
                    published_templates.append(
                        {
                            "name": item.name,
                            "path": str(item),
                            "has_json": json_file.exists(),
                        }
                    )

    context = {
        "is_developer": is_developer,
        "published_templates": published_templates,
    }
    return render(request, "ppt_generator/developer_tools.html", context)


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def config_generator_page(request):
    """Config generator independent page."""
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    context = {
        "is_developer": is_developer,
    }
    return render(request, "ppt_generator/config_generator.html", context)


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def config_editor_page(request):
    """Config editor independent page."""
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    # 检查是否是嵌入模式
    embedded = request.GET.get("embedded") == "1"

    # 检查是否有初始配置数据（来自向导的配置生成）
    init_config = request.GET.get("init_config")

    context = {
        "is_developer": is_developer,
        "embedded": embedded,
        "init_config": init_config,
    }

    response = render(request, "ppt_generator/config_editor.html", context)

    # 嵌入模式允许同源 iframe 加载
    if embedded:
        response["X-Frame-Options"] = "SAMEORIGIN"

    return response


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def template_editor_page(request):
    """Template editor independent page."""
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    # 检查是否是嵌入模式（从向导页面的 iframe 加载）
    embedded = request.GET.get("embedded") == "1"

    context = {
        "is_developer": is_developer,
        "embedded": embedded,
    }
    response = render(request, "ppt_generator/template_editor.html", context)

    # 如果是嵌入模式，允许在同源 iframe 中加载
    if embedded:
        response["X-Frame-Options"] = "SAMEORIGIN"

    return response


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def generate_config_template(request):
    """Generate config template from PPTX (AJAX endpoint)."""
    if request.method == "POST":
        template_file = request.FILES.get("template_file")
        mode = request.POST.get("mode", "semantic")

        if not template_file:
            return JsonResponse({"error": "请上传模板文件"}, status=400)

        try:
            # Save template temporarily
            import tempfile
            import os
            from pathlib import Path

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                for chunk in template_file.chunks():
                    tmp.write(chunk)
                tmp_path = tmp.name

            # Import template analysis function
            from scripts.export_template_structure import export_template_structure

            # Analyze template
            template_data = export_template_structure(
                template_path=Path(tmp_path),
                mode=mode,
                include_pages=None,  # Export all pages
            )

            # Clean up
            os.unlink(tmp_path)

            return JsonResponse(template_data, safe=False)

        except Exception as e:
            import traceback

            return JsonResponse(
                {"error": str(e), "traceback": traceback.format_exc()}, status=500
            )

    return JsonResponse({"error": "仅支持 POST 请求"}, status=405)


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def ai_enrich_template_view(request):
    """AI enrich template configuration (AJAX endpoint)."""
    if request.method == "POST":
        try:
            import json as json_module

            # Get template data from request
            template_data = json_module.loads(request.body)

            # Get LLM configuration from global config
            global_config = GlobalLLMConfig.get_config()
            llm_provider = global_config.llm_provider
            llm_model = global_config.llm_model
            llm_base_url = global_config.llm_base_url
            llm_api_key = global_config.llm_api_key

            # Set API key in environment if provided
            if llm_api_key:
                import os

                if llm_provider == "deepseek":
                    os.environ["DEEPSEEK_API_KEY"] = llm_api_key
                elif llm_provider == "local":
                    os.environ["LOCAL_LLM_API_KEY"] = llm_api_key

            # Import AI enrich function
            from scripts.export_template_structure import ai_enrich_template

            # Enrich template
            enriched_data = ai_enrich_template(
                template_data=template_data,
                llm_provider=llm_provider,
                llm_model=llm_model,
                llm_base_url=llm_base_url,
            )

            return JsonResponse(enriched_data, safe=False)

        except Exception as e:
            import traceback

            return JsonResponse(
                {"error": str(e), "traceback": traceback.format_exc()}, status=500
            )

    return JsonResponse({"error": "仅支持 POST 请求"}, status=405)


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def parse_ppt_template(request):
    """
    解析 PPT 模板，提取元素信息并生成标注截图

    Request:
        - ppt_file: PPT 模板文件

    Response:
        {
            "template_id": "uuid",
            "pages": [
                {
                    "page_num": 1,
                    "image_url": "/media/temp/page_1_annotated.png",
                    "shapes": [...]
                }
            ]
        }
    """
    try:
        import uuid
        from .utils import (
            extract_shapes_info,
            convert_ppt_to_pdf,
            convert_pdf_to_images,
            annotate_screenshot,
        )

        # 获取上传的文件
        ppt_file = request.FILES.get("ppt_file")

        if not ppt_file:
            return JsonResponse({"error": "请上传 PPT 文件"}, status=400)

        # 创建临时目录
        template_id = str(uuid.uuid4())
        temp_dir = settings.MEDIA_ROOT / "template_editor" / template_id
        temp_dir.mkdir(parents=True, exist_ok=True)

        # 保存 PPT 文件（统一命名为 template.pptx，便于后续发布）
        ppt_path = temp_dir / "template.pptx"
        with open(ppt_path, "wb+") as f:
            for chunk in ppt_file.chunks():
                f.write(chunk)

        # 提取元素信息
        shapes_data = extract_shapes_info(ppt_path)

        # 使用 LibreOffice 将 PPT 转换为 PDF
        pdf_path = convert_ppt_to_pdf(ppt_path, temp_dir)

        # 转换 PDF 为图片（使用 150 DPI 以减小文件大小，加快加载速度）
        images_dir = temp_dir / "images"
        image_paths = convert_pdf_to_images(pdf_path, images_dir, dpi=150)

        # 获取幻灯片尺寸
        slide_width = shapes_data.get("slide_width", 12192000)
        slide_height = shapes_data.get("slide_height", 6858000)

        # 为每个页面生成标注图片
        pages = []
        for page_data in shapes_data["pages"]:
            page_num = page_data["page_num"]

            if page_num <= len(image_paths):
                image_path = image_paths[page_num - 1]

                # 生成标注图片（使用幻灯片尺寸计算坐标）
                annotated_path = annotate_screenshot(
                    image_path,
                    page_data["shapes"],
                    slide_width=slide_width,
                    slide_height=slide_height,
                )

                # 生成相对 URL
                relative_path = annotated_path.relative_to(settings.MEDIA_ROOT)
                image_url = f"/media/{relative_path}"

                # 根据元素类型推断页面类型
                shapes = page_data["shapes"]
                text_count = sum(1 for s in shapes if s.get("type") == "text")
                image_count = sum(1 for s in shapes if s.get("type") == "image")

                if text_count == 0 and image_count > 0:
                    page_type = "纯图页"
                elif text_count <= 3 and image_count == 0:
                    page_type = "标题页"
                elif image_count > 0:
                    page_type = "图文页"
                elif text_count > 0:
                    page_type = "文字页"
                else:
                    page_type = f"第{page_num}页"

                pages.append(
                    {
                        "page_num": page_num,
                        "page_type": page_type,
                        "image_url": image_url,
                        "shapes": page_data["shapes"],
                    }
                )

        return JsonResponse(
            {
                "template_id": template_id,
                "ppt_path": str(ppt_path.relative_to(settings.MEDIA_ROOT)),
                "slide_width": shapes_data.get("slide_width", 12192000),
                "slide_height": shapes_data.get("slide_height", 6858000),
                "pages": pages,
            }
        )

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def update_shape_name_api(request):
    """
    更新元素名称

    Request:
        {
            "template_id": "uuid",
            "page_num": 1,
            "shape_index": 3,  # 元素在 slide.shapes 中的索引
            "new_name": "标题区"
        }

    Response:
        {"success": true}
    """
    try:
        import json
        from .utils import update_shape_name

        data = json.loads(request.body)
        template_id = data.get("template_id")
        page_num = data.get("page_num")
        # 使用 shape_id 定位元素（支持 GROUP 内的元素）
        shape_id = data.get("shape_id")
        new_name = data.get("new_name")

        if not all([template_id, page_num is not None, shape_id is not None, new_name]):
            return JsonResponse({"error": "缺少必要参数"}, status=400)

        # 获取 PPT 文件路径
        ppt_path = settings.MEDIA_ROOT / "template_editor" / template_id
        ppt_files = list(ppt_path.glob("*.pptx"))

        if not ppt_files:
            return JsonResponse({"error": "找不到 PPT 文件"}, status=404)

        # 更新元素名称（使用 shape_id 支持 GROUP 内元素）
        print(
            f"[update_shape_name] 更新形状名称: 文件={ppt_files[0]}, 页码={page_num}, shape_id={shape_id}, 新名称={new_name}"
        )
        update_shape_name(ppt_files[0], page_num, shape_id, new_name)
        print(f"[update_shape_name] 保存成功")

        return JsonResponse({"success": True})

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


def _estimate_max_chars(shape: dict) -> int:
    """
    根据文本框尺寸估算最大字符数

    PPT 文字宜精简，采用保守估算策略：
    - 使用当前内容长度作为参考（不额外增加）
    - 如果没有内容，根据面积粗略估算但限制上限

    Args:
        shape: 包含 width, height, char_count 等属性的字典

    Returns:
        估算的最大字符数（建议手动校准）
    """
    char_count = shape.get("char_count", 0)

    # 如果有现有内容，直接使用当前字数（PPT 模板通常已经是合适长度）
    if char_count > 0:
        return char_count

    # 没有现有内容时，根据面积粗略估算
    # 假设平均每个中文字符占约 40x40 pt 的区域（更保守的估算）
    EMU_PER_POINT = 914400 / 72
    width = shape.get("width", 0)
    height = shape.get("height", 0)

    if width and height:
        width_pt = width / EMU_PER_POINT
        height_pt = height / EMU_PER_POINT
        # 假设字符占用面积约 1600 平方点（40x40）- 更保守
        area = width_pt * height_pt
        estimated = int(area / 1600)
        # 限制最大值为 150（PPT 文字不宜过长）
        return min(max(estimated, 10), 150)

    return 20  # 默认值


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def generate_template_config(request):
    """
    生成配置 JSON（符合 S2S 标准格式）

    Request:
        {
            "template_id": "uuid"
        }

    Response:
        {
            "config": {
                "manifest": [...],
                "ppt_pages": [...]
            }
        }
    """
    try:
        import json
        from .utils import extract_shapes_info

        data = json.loads(request.body)
        template_id = data.get("template_id")

        if not template_id:
            return JsonResponse({"error": "缺少 template_id"}, status=400)

        # 获取 PPT 文件路径
        ppt_path = settings.MEDIA_ROOT / "template_editor" / template_id
        ppt_files = list(ppt_path.glob("*.pptx"))

        if not ppt_files:
            return JsonResponse({"error": "找不到 PPT 文件"}, status=404)

        # 提取元素信息
        shapes_data = extract_shapes_info(ppt_files[0])

        # 生成符合 S2S 标准的配置 JSON
        manifest = []
        ppt_pages = []

        for page_data in shapes_data["pages"]:
            page_num = page_data["page_num"]
            text_slots = 0
            image_slots = 0
            content = {}

            for shape in page_data["shapes"]:
                # 只包含已命名的元素（非隐藏）
                if shape.get("is_named") and not shape.get("is_hidden"):
                    name = shape["name"]
                    is_text = shape["type"] == "text"

                    if is_text:
                        text_slots += 1
                        # 智能估算 max_chars
                        max_chars = _estimate_max_chars(shape)
                        content[name] = {
                            "type": "text",
                            "hint": f"填写{name}的内容",
                            "required": True,
                            "value": "",
                            "max_chars": max_chars,
                        }
                    else:
                        image_slots += 1
                        content[name] = {
                            "type": "image",
                            "hint": "插入与本页主题相关的图片路径",
                            "required": True,
                            "value": "",
                            "preferred_format": "png/jpg",
                        }

            # 只添加有内容的页面
            if content:
                # 根据内容推断页面类型
                page_type = f"第{page_num}页"
                if text_slots == 0 and image_slots > 0:
                    page_type = "纯图页"
                elif text_slots <= 3 and image_slots == 0:
                    page_type = "标题页"
                elif image_slots > 0:
                    page_type = "图文页"
                else:
                    page_type = "文字页"

                manifest.append(
                    {
                        "template_page_num": page_num,
                        "page_type": page_type,
                        "text_slots": text_slots,
                        "image_slots": image_slots,
                    }
                )

                ppt_pages.append(
                    {
                        "page_type": page_type,
                        "template_page_num": page_num,
                        "content": content,
                        "meta": {
                            "layout": page_type,
                            "scene": [],
                            "style": "",
                            "text_slots": text_slots,
                            "image_slots": image_slots,
                            "notes": "请根据实际需要填写内容",
                        },
                    }
                )

        config = {"manifest": manifest, "ppt_pages": ppt_pages}

        return JsonResponse({"config": config})

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["GET"])
def download_template_ppt(request, template_id):
    """
    下载编辑后的 PPT 模板

    Args:
        template_id: 模板 ID

    Response:
        PPT 文件下载
    """
    try:
        from django.http import FileResponse

        # 获取 PPT 文件路径
        ppt_path = settings.MEDIA_ROOT / "template_editor" / template_id
        ppt_files = list(ppt_path.glob("*.pptx"))

        if not ppt_files:
            return JsonResponse({"error": "找不到 PPT 文件"}, status=404)

        # 返回文件下载
        response = FileResponse(
            open(ppt_files[0], "rb"),
            content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response["Content-Disposition"] = f'attachment; filename="{ppt_files[0].name}"'

        return response

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def toggle_shape_visibility(request):
    """
    切换元素的隐藏/显示状态

    Request:
        {
            "template_id": "uuid",
            "page_num": 1,
            "shape_id": 123,  # 元素的 shape_id（支持 GROUP 内元素）
            "is_hidden": true
        }

    Response:
        {"success": true}

    Note: 这个功能只是在前端标记，不修改 PPT 文件
    """
    try:
        import json

        data = json.loads(request.body)
        template_id = data.get("template_id")
        page_num = data.get("page_num")
        shape_id = data.get("shape_id")
        is_hidden = data.get("is_hidden")

        if not all(
            [
                template_id,
                page_num is not None,
                shape_id is not None,
                is_hidden is not None,
            ]
        ):
            return JsonResponse({"error": "缺少必要参数"}, status=400)

        # 这个功能只在前端维护状态，不需要修改 PPT 文件
        # 前端会在生成配置 JSON 时自动过滤隐藏的元素

        return JsonResponse({"success": True})

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def refresh_page_preview(request):
    """
    刷新页面预览图（用于隐藏/显示元素后更新标注）

    Request:
        {
            "template_id": "uuid",
            "page_num": 1,
            "shapes": [...]  // 包含 is_hidden 状态的元素列表
        }

    Response:
        {"success": true, "image_url": "/media/..."}
    """
    try:
        import json
        import logging
        from .utils import annotate_screenshot

        logger = logging.getLogger(__name__)

        data = json.loads(request.body)
        template_id = data.get("template_id")
        page_num = data.get("page_num")
        shapes = data.get("shapes", [])

        # 调试日志
        hidden_shapes = [s for s in shapes if s.get("is_hidden")]
        logger.info(
            f"[refresh_preview] page={page_num}, total={len(shapes)}, hidden={len(hidden_shapes)}"
        )
        for s in hidden_shapes:
            logger.info(f"  隐藏元素: {s.get('name')} (id={s.get('shape_id')})")

        if not all([template_id, page_num is not None]):
            return JsonResponse({"error": "缺少必要参数"}, status=400)

        # 获取原始截图路径（图片在 images 子目录中）
        template_path = settings.MEDIA_ROOT / "template_editor" / template_id
        images_dir = template_path / "images"
        original_image = images_dir / f"page_{page_num}.png"

        logger.info(f"[refresh_preview] 查找图片: {original_image}")

        if not original_image.exists():
            return JsonResponse(
                {"error": f"找不到页面 {page_num} 的截图: {original_image}"}, status=404
            )

        # 获取 PPT 文件以获取幻灯片尺寸
        ppt_files = list(template_path.glob("*.pptx"))
        if not ppt_files:
            return JsonResponse({"error": "找不到 PPT 文件"}, status=404)

        from pptx import Presentation

        prs = Presentation(str(ppt_files[0]))
        slide_width = prs.slide_width
        slide_height = prs.slide_height

        # 重新生成标注图片
        annotated_path = annotate_screenshot(
            original_image,
            shapes,
            slide_width=slide_width,
            slide_height=slide_height,
        )

        # 生成相对 URL
        relative_path = annotated_path.relative_to(settings.MEDIA_ROOT)
        image_url = f"/media/{relative_path}"

        return JsonResponse({"success": True, "image_url": image_url})

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def ai_auto_name_shapes(request):
    """
    使用多模态 AI 自动为页面元素命名

    Request:
        {
            "template_id": "uuid",
            "page_num": 1,
            "image_url": "/media/...",
            "shapes": [...],  # 当前页面的元素列表
            "llm_provider": "glm",  # 可选，默认使用全局配置
            "llm_model": "glm-4v-plus"  # 可选
        }

    Response:
        {
            "success": true,
            "named_shapes": [
                {"shape_index": 3, "suggested_name": "章节标题"},
                ...
            ]
        }
    """
    import base64
    import json
    import mimetypes
    import os
    import sys
    from pathlib import Path

    try:
        data = json.loads(request.body)
        template_id = data.get("template_id")
        page_num = data.get("page_num")
        image_url = data.get("image_url")
        shapes = data.get("shapes", [])
        existing_config = data.get("existing_config", {})  # 现有配置（查漏补缺模式）
        wizard_mode = data.get("wizard_mode", False)  # 是否向导模式

        if not all([template_id, page_num, image_url, shapes]):
            return JsonResponse({"error": "缺少必要参数"}, status=400)

        # 获取 LLM 配置
        llm_provider = data.get("llm_provider")
        llm_model = data.get("llm_model")

        # 如果没有指定，使用多模态默认配置
        if not llm_provider:
            from .models import GlobalLLMConfig

            # 优先使用多模态默认配置
            multimodal_config = GlobalLLMConfig.get_multimodal_config()
            if multimodal_config:
                llm_provider = multimodal_config.llm_provider
                # 使用 get_model_for_provider() 获取正确的模型名
                llm_model = llm_model or multimodal_config.get_model_for_provider()
                # 设置 API Key
                if multimodal_config.llm_api_key:
                    if llm_provider == "glm":
                        os.environ["GLM_API_KEY"] = multimodal_config.llm_api_key
                    elif llm_provider == "taichu":
                        os.environ["TAICHU_API_KEY"] = multimodal_config.llm_api_key

                print(
                    f"[AI命名] 使用多模态配置: provider={llm_provider}, model={llm_model}"
                )
            else:
                return JsonResponse(
                    {
                        "error": "未配置多模态 LLM，请先在管理后台配置 glm 或 taichu 模型并设为多模态默认"
                    },
                    status=400,
                )

        # 注意：多模态支持由用户在配置中标记，这里不再硬编码检查
        # get_multimodal_config 已确保返回的配置支持多模态

        # 加载图片（去除可能的时间戳参数）
        clean_image_url = image_url.split("?")[0]  # 移除 ?t=xxx 等参数
        image_path = settings.MEDIA_ROOT / clean_image_url.lstrip("/media/")
        if not image_path.exists():
            return JsonResponse({"error": f"图片不存在: {clean_image_url}"}, status=404)

        # 读取并编码图片
        with open(image_path, "rb") as f:
            image_data = f.read()
        base64_image = base64.b64encode(image_data).decode("utf-8")
        mime_type, _ = mimetypes.guess_type(str(image_path))
        if not mime_type:
            mime_type = "image/png"

        # 过滤出可见元素（隐藏的元素不显示在图片上也不需要命名）
        visible_shapes = [s for s in shapes if not s.get("is_hidden")]

        # 构建元素描述（编号与图片上的标注一致，从1开始）
        shapes_desc = []
        existing_elements = existing_config.get("elements", [])
        for i, shape in enumerate(visible_shapes, 1):
            desc = f"#{i}: 类型={shape.get('type', '未知')}"
            if shape.get("text_sample"):
                desc += f', 文本预览="{shape["text_sample"][:30]}..."'
            # 添加现有配置信息
            if i <= len(existing_elements):
                elem = existing_elements[i - 1]
                if elem.get("name"):
                    desc += f', 已命名="{elem["name"]}"'
                if elem.get("hint"):
                    desc += f', 已有提示="{elem["hint"]}"'
            shapes_desc.append(desc)

        # 检查是否有现有配置（查漏补缺模式）
        existing_page_type = existing_config.get("page_type", "")
        existing_page_note = existing_config.get("page_note", "")
        has_existing = bool(existing_page_type) or any(
            e.get("name") or e.get("hint") for e in existing_elements
        )

        # 构建 Prompt
        if has_existing:
            # 查漏补缺模式
            prompt = f"""你是一个 PPT 模板分析专家。请分析这张 PPT 幻灯片截图，**补全缺失的配置信息**。

**重要说明**
这是一个通用模板，会用于多种不同主题的演讲/课程/汇报。请使用**通用、抽象**的命名。

**现有配置**（保留已有内容，补全缺失部分）
- 页面类型: {existing_page_type or "（未填写，请补全）"}
- 页面备注: {existing_page_note or "（未填写，请补全）"}

**页面元素**（{len(visible_shapes)} 个）：
{chr(10).join(shapes_desc)}

**补全规则**
1. 已有的 name 和 hint 请**保持不变**，只补全空缺的字段
2. 使用通用名称，如：主标题、副标题、正文内容、配图、日期等
3. hint 提示要通用，如"填写本页主题"而非具体内容
4. max_chars: 尽量往少估算，PPT文字宜精简。标题类10-20，正文类30-80，长文本最多150，图片填null
5. required: 重要元素为true，装饰性元素为false

请以 JSON 格式返回完整配置（包含已有和补全的内容）：
```json
{{
  "page_type": "封面页",
  "page_note": "展示演讲主题和演讲者信息",
  "elements": [
    {{"index": 1, "name": "主标题", "hint": "填写演讲或课程主题", "max_chars": 15, "required": true}},
    {{"index": 2, "name": "副标题", "hint": "补充说明或副主题", "max_chars": 30, "required": false}}
  ]
}}
```

只返回 JSON，不要其他解释。确保 elements 数量与可见元素数量 ({len(visible_shapes)}) 一致。"""
        else:
            # 全新配置模式
            prompt = f"""你是一个 PPT 模板分析专家。请分析这张 PPT 幻灯片截图，为这个**通用模板**配置元素信息。

**重要说明**
这是一个通用模板，会用于多种不同主题的演讲/课程/汇报。请使用**通用、抽象**的命名，不要根据图片中的具体内容命名。

**页面信息**
图片上标注了 {len(visible_shapes)} 个可编辑元素（黄色/蓝色编号圈）：

{chr(10).join(shapes_desc)}

**命名原则**
1. 根据元素的**位置和布局作用**命名，而非具体内容
2. 使用通用名称，如：
   - 标题类：主标题、副标题、页面标题、小标题
   - 内容类：正文内容、说明文字、描述文本、要点列表
   - 图片类：主图、配图、插图、背景图
   - 信息类：日期、作者、单位名称、页码
3. hint 提示也要通用，如"填写本页主题"而非"填写活动名称"

**配置规则**
- page_type: 页面类型，如"封面页"、"目录页"、"内容页"、"图文页"、"结束页"，不超过6字
- page_note: 简要说明这类页面的通用用途
- name: 元素名称，反映布局位置/功能，不超过10字
- hint: 通用的内容提示，不涉及具体主题
- max_chars: 建议最大字数，尽量往少估算（PPT文字宜精简）。标题类10-20，正文类30-80，长文本最多150，图片填null
- required: 是否必填

请以 JSON 格式返回：
```json
{{
  "page_type": "封面页",
  "page_note": "展示演讲主题和演讲者信息",
  "elements": [
    {{"index": 1, "name": "主标题", "hint": "填写演讲或课程主题", "max_chars": 15, "required": true}},
    {{"index": 2, "name": "副标题", "hint": "补充说明或副主题", "max_chars": 30, "required": false}},
    {{"index": 3, "name": "日期", "hint": "填写日期信息", "max_chars": 12, "required": false}}
  ]
}}
```

只返回 JSON，不要其他解释。确保 elements 数量与可见元素数量 ({len(visible_shapes)}) 一致。"""

        # 初始化 LLM（使用管理后台配置的多模态模型）
        sys.path.insert(0, str(Path(__file__).parent.parent.parent / "scripts"))
        from llm_client import GLMLLM, TaichuLLM

        # llm_model 已从 multimodal_config.llm_model 获取，不再使用硬编码默认值
        if not llm_model:
            return JsonResponse(
                {"error": "未配置多模态模型名称，请在管理后台设置"},
                status=400,
            )

        if llm_provider == "glm":
            llm = GLMLLM(model=llm_model)
        else:
            llm = TaichuLLM(model=llm_model)

        # 构建多模态消息
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime_type};base64,{base64_image}"},
                    },
                ],
            }
        ]

        # 调用 LLM（不重试，避免加重限流问题）
        try:
            print(f"[AI命名] 开始调用 {llm_provider} 模型 {llm_model}...")
            response = llm.generate(messages)
            print(f"[AI命名] 调用成功，响应长度: {len(response)}")
        except Exception as e:
            error_str = str(e)
            print(f"[AI命名] 调用失败: {error_str}")

            # 检查是否是限流错误，给出明确提示
            if "429" in error_str or "1302" in error_str or "并发" in error_str:
                raise Exception(
                    "GLM API 限流，请等待 30 秒后重试。如频繁遇到此问题，可在管理后台切换到太初(taichu)模型。"
                ) from e
            raise

        # 解析响应
        print(f"[AI命名] 原始响应:\n{response[:500]}...")  # 打印前500字符用于调试

        # 提取 JSON 部分
        json_match = response.strip()

        # 尝试多种方式提取 JSON
        if "```json" in json_match:
            json_match = json_match.split("```json")[1].split("```")[0].strip()
        elif "```" in json_match:
            # 可能是 ```\n{...}\n```
            parts = json_match.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("{") or part.startswith("["):
                    json_match = part
                    break

        # 如果还是没找到 JSON，尝试直接找 { 或 [ 开头的内容
        if not json_match.startswith("{") and not json_match.startswith("["):
            # 尝试找到第一个 { 或 [
            start_brace = json_match.find("{")
            start_bracket = json_match.find("[")
            if start_brace >= 0 and (start_bracket < 0 or start_brace < start_bracket):
                # 找到最后一个匹配的 }
                end_brace = json_match.rfind("}")
                if end_brace > start_brace:
                    json_match = json_match[start_brace : end_brace + 1]
            elif start_bracket >= 0:
                end_bracket = json_match.rfind("]")
                if end_bracket > start_bracket:
                    json_match = json_match[start_bracket : end_bracket + 1]

        print(f"[AI命名] 提取的 JSON:\n{json_match[:300]}...")

        if not json_match:
            raise json.JSONDecodeError("AI 返回内容为空", response, 0)

        parsed_response = json.loads(json_match)

        # 支持新格式（包含 page_type 和 elements）和旧格式（仅元素列表）
        if isinstance(parsed_response, dict):
            page_type = parsed_response.get("page_type", "")
            page_note = parsed_response.get("page_note", "")
            suggested_names = parsed_response.get("elements", [])
        else:
            # 兼容旧格式：直接是元素列表
            page_type = ""
            page_note = ""
            suggested_names = parsed_response

        # 映射到 shape_index（visible_shapes 已在上面定义）
        named_shapes = []

        for suggestion in suggested_names:
            idx = suggestion.get("index", 0) - 1  # 转为0-based
            if 0 <= idx < len(visible_shapes):
                shape = visible_shapes[idx]
                named_shapes.append(
                    {
                        "shape_index": shape.get("shape_index"),
                        "shape_id": shape.get("shape_id"),
                        "suggested_name": suggestion.get("name", ""),
                        # 新增：额外配置信息（向导模式下使用）
                        "hint": suggestion.get("hint", ""),
                        "max_chars": suggestion.get("max_chars"),
                        "required": suggestion.get("required", False),
                    }
                )

        return JsonResponse(
            {
                "success": True,
                "named_shapes": named_shapes,
                "page_type": page_type,
                "page_note": page_note,  # 新增：页面备注
            }
        )

    except json.JSONDecodeError as e:
        return JsonResponse(
            {"error": f"AI 返回格式错误: {str(e)}", "raw_response": response},
            status=500,
        )
    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


# ============ 编辑记录管理 API ============


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["GET"])
def list_edit_sessions(request):
    """
    获取当前用户的编辑记录列表

    Query Params:
        editor_type: 可选，'ppt' 或 'config'，不传则返回所有类型

    Response:
        {
            "sessions": [
                {
                    "id": 1,
                    "session_id": "uuid",
                    "editor_type": "ppt",
                    "editor_type_display": "PPT 模板编辑器",
                    "template_name": "template1.pptx",
                    "progress_summary": "10/20 已命名 (50%)",
                    "thumbnail_url": "/media/...",
                    "created_at": "2024-01-01 12:00:00",
                    "updated_at": "2024-01-01 13:00:00"
                }
            ]
        }
    """
    editor_type = request.GET.get("editor_type")

    sessions = TemplateEditSession.objects.filter(user=request.user)
    if editor_type:
        sessions = sessions.filter(editor_type=editor_type)

    sessions = sessions.order_by("-updated_at")[:20]  # 最多返回20条

    result = []
    for session in sessions:
        result.append(
            {
                "id": session.id,
                "session_id": session.session_id,
                "editor_type": session.editor_type,
                "editor_type_display": session.get_editor_type_display(),
                "template_name": session.template_name,
                "progress_summary": session.progress_summary,
                "progress_data": session.progress_data,
                "thumbnail_url": session.thumbnail_url,
                "created_at": session.created_at.strftime("%Y-%m-%d %H:%M"),
                "updated_at": session.updated_at.strftime("%Y-%m-%d %H:%M"),
            }
        )

    return JsonResponse({"sessions": result})


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["POST"])
def save_edit_session(request):
    """
    保存/更新编辑会话记录

    Request:
        {
            "session_id": "uuid",
            "editor_type": "ppt" | "config",
            "template_name": "template.pptx",
            "progress_data": {...},
            "thumbnail_url": "/media/..."  // 可选
        }

    Response:
        {"success": true, "id": 1}
    """
    try:
        data = json.loads(request.body)
        session_id = data.get("session_id")
        editor_type = data.get("editor_type", "ppt")
        template_name = data.get("template_name", "未命名模板")
        progress_data = data.get("progress_data", {})
        thumbnail_url = data.get("thumbnail_url")

        if not session_id:
            return JsonResponse({"error": "缺少 session_id"}, status=400)

        # 创建或更新
        session, created = TemplateEditSession.objects.update_or_create(
            user=request.user,
            session_id=session_id,
            defaults={
                "editor_type": editor_type,
                "template_name": template_name,
                "progress_data": progress_data,
                "thumbnail_url": thumbnail_url,
            },
        )

        return JsonResponse({"success": True, "id": session.id, "created": created})

    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["DELETE", "POST"])
def delete_edit_session(request, session_id):
    """
    删除编辑会话记录

    Response:
        {"success": true}
    """
    try:
        session = TemplateEditSession.objects.get(
            user=request.user, session_id=session_id
        )

        # 如果是 PPT 编辑器，同时删除临时文件
        if session.editor_type == "ppt":
            import shutil

            temp_dir = settings.MEDIA_ROOT / "template_editor" / session_id
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

        session.delete()
        return JsonResponse({"success": True})

    except TemplateEditSession.DoesNotExist:
        return JsonResponse({"error": "记录不存在"}, status=404)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["GET"])
def get_edit_session(request, session_id):
    """
    获取单个编辑会话的详细信息

    Response:
        {
            "session": {...},
            "exists": true  // 对于 PPT 编辑器，检查临时文件是否还存在
        }
    """
    try:
        session = TemplateEditSession.objects.get(
            user=request.user, session_id=session_id
        )

        # 检查文件是否还存在
        exists = True
        if session.editor_type == "ppt":
            temp_dir = settings.MEDIA_ROOT / "template_editor" / session_id
            exists = temp_dir.exists()

        return JsonResponse(
            {
                "session": {
                    "id": session.id,
                    "session_id": session.session_id,
                    "editor_type": session.editor_type,
                    "template_name": session.template_name,
                    "progress_data": session.progress_data,
                    "thumbnail_url": session.thumbnail_url,
                    "created_at": session.created_at.strftime("%Y-%m-%d %H:%M"),
                    "updated_at": session.updated_at.strftime("%Y-%m-%d %H:%M"),
                },
                "exists": exists,
            }
        )

    except TemplateEditSession.DoesNotExist:
        return JsonResponse({"error": "记录不存在"}, status=404)


@login_required
@permission_required("ppt_generator.is_developer", raise_exception=True)
@require_http_methods(["GET"])
def restore_edit_session(request, session_id):
    """
    恢复 PPT 编辑会话 - 重新加载已上传的模板数据

    Response:
        与 parse_ppt_template 相同的结构
    """
    try:
        from .utils import extract_shapes_info, annotate_screenshot

        # 尝试查找会话（ppt 类型或从向导传入的直接加载）
        session = TemplateEditSession.objects.filter(
            user=request.user, session_id=session_id, editor_type="ppt"
        ).first()

        # 获取 PPT 文件路径
        template_path = settings.MEDIA_ROOT / "template_editor" / session_id
        if not template_path.exists():
            return JsonResponse({"error": "编辑会话文件已过期"}, status=404)

        ppt_files = list(template_path.glob("*.pptx"))
        if not ppt_files:
            return JsonResponse({"error": "找不到 PPT 文件"}, status=404)

        ppt_path = ppt_files[0]

        # 提取元素信息
        shapes_data = extract_shapes_info(ppt_path)

        # 获取幻灯片尺寸
        slide_width = shapes_data.get("slide_width", 12192000)
        slide_height = shapes_data.get("slide_height", 6858000)

        # 构建页面数据（使用已有的标注图片）
        images_dir = template_path / "images"
        pages = []

        for page_data in shapes_data["pages"]:
            page_num = page_data["page_num"]

            # 查找已有的标注图片
            annotated_image = images_dir / f"page_{page_num}_annotated.png"
            if annotated_image.exists():
                image_url = f"/media/template_editor/{session_id}/images/page_{page_num}_annotated.png"
            else:
                # 如果没有标注图片，使用原始图片重新标注
                original_image = images_dir / f"page_{page_num}.png"
                if original_image.exists():
                    annotated_path = annotate_screenshot(
                        original_image,
                        page_data["shapes"],
                        slide_width,
                        slide_height,
                    )
                    image_url = (
                        f"/media/template_editor/{session_id}/images/"
                        + annotated_path.name
                    )
                else:
                    image_url = None

            pages.append(
                {
                    "page_num": page_num,
                    "image_url": image_url,
                    "shapes": page_data["shapes"],
                }
            )

        # 获取模板名称（从 session 或 ppt 文件名）
        template_name = session.template_name if session else ppt_path.stem

        return JsonResponse(
            {
                "template_id": session_id,
                "template_name": template_name,
                "ppt_path": str(ppt_path.relative_to(settings.MEDIA_ROOT)),
                "slide_width": slide_width,
                "slide_height": slide_height,
                "pages": pages,
            }
        )
    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
def template_wizard_page(request):
    """模板制作向导页面"""
    import shutil
    from .utils import (
        extract_shapes_info,
        convert_ppt_to_pdf,
        convert_pdf_to_images,
        annotate_screenshot,
    )

    # 检查是否是编辑已发布模板的请求
    edit_template = request.GET.get("edit")
    edit_mode_data = None

    if edit_template:
        # 编辑已发布模板模式
        template_dir = settings.S2S_TEMPLATE_DIR / edit_template
        pptx_file = template_dir / "template.pptx"
        json_file = template_dir / "template.json"

        if pptx_file.exists():
            # 创建临时会话目录，复制 PPT 文件
            ppt_session_id = str(uuid.uuid4())
            session_dir = settings.MEDIA_ROOT / "template_editor" / ppt_session_id
            session_dir.mkdir(parents=True, exist_ok=True)

            # 复制 PPT 文件
            ppt_path = session_dir / "template.pptx"
            shutil.copy2(pptx_file, ppt_path)

            # 生成预览图片（和 parse_ppt_template 相同的逻辑）
            try:
                # 提取元素信息
                shapes_data = extract_shapes_info(ppt_path)

                # 使用 LibreOffice 将 PPT 转换为 PDF
                pdf_path = convert_ppt_to_pdf(ppt_path, session_dir)

                # 转换 PDF 为图片
                images_dir = session_dir / "images"
                convert_pdf_to_images(pdf_path, images_dir, dpi=150)

                # 获取幻灯片尺寸
                slide_width = shapes_data.get("slide_width", 12192000)
                slide_height = shapes_data.get("slide_height", 6858000)

                # 为每个页面生成标注图片
                for page_data in shapes_data["pages"]:
                    page_num = page_data["page_num"]
                    image_path = images_dir / f"page_{page_num}.png"
                    if image_path.exists():
                        annotate_screenshot(
                            image_path,
                            page_data["shapes"],
                            slide_width=slide_width,
                            slide_height=slide_height,
                        )
            except Exception as e:
                print(f"[template_wizard_page] 生成预览图失败: {e}")

            # 加载 JSON 配置（如果存在）
            config_data = None
            if json_file.exists():
                try:
                    config_data = json.loads(json_file.read_text(encoding="utf-8"))
                except Exception:
                    pass

            edit_mode_data = json.dumps(
                {
                    "template_name": edit_template,
                    "ppt_session_id": ppt_session_id,
                    "config_data": config_data,
                    "is_edit_mode": True,
                }
            )

    # 只有通过 URL 参数指定 session 时才恢复会话
    # 直接点击卡片进入时不传 session 参数，开始新的向导
    session_id = request.GET.get("session")

    existing_session = None
    existing_session_json = None
    if session_id and not edit_template:
        existing_session = TemplateEditSession.objects.filter(
            user=request.user, session_id=session_id, editor_type="wizard"
        ).first()
        # 如果已发布，不恢复
        if existing_session and existing_session.progress_data.get("published"):
            existing_session = None

        # 转换为 JSON 友好的格式
        if existing_session:
            existing_session_json = json.dumps(
                {
                    "session_id": existing_session.session_id,
                    "template_name": existing_session.template_name,
                    "updated_at": existing_session.updated_at.strftime(
                        "%Y-%m-%d %H:%M"
                    ),
                    "progress_data": existing_session.progress_data or {},
                }
            )

    context = {
        "existing_session": existing_session,
        "existing_session_json": existing_session_json,
        "edit_mode_data": edit_mode_data,
    }
    return render(request, "ppt_generator/template_wizard.html", context)


@login_required
@require_http_methods(["POST"])
def publish_template(request):
    """
    发布模板到 templates 目录

    请求体:
    {
        "template_name": "课程介绍模板",
        "ppt_session_id": "xxx",
        "config_data": {...}
    }
    """
    try:
        data = json.loads(request.body)
        template_name = data.get("template_name", "").strip()
        ppt_session_id = data.get("ppt_session_id")
        config_data = data.get("config_data")
        is_edit_mode = data.get("is_edit_mode", False)
        original_template_name = data.get("original_template_name")

        if not template_name:
            return JsonResponse({"error": "模板名称不能为空"}, status=400)
        if not ppt_session_id:
            return JsonResponse({"error": "缺少 PPT 会话 ID"}, status=400)
        if not config_data:
            return JsonResponse({"error": "缺少配置数据"}, status=400)

        # 验证模板名称（只允许中文、英文、数字、下划线、横线）
        import re

        if not re.match(r"^[\u4e00-\u9fa5a-zA-Z0-9_-]+$", template_name):
            return JsonResponse(
                {"error": "模板名称只能包含中文、英文、数字、下划线和横线"}, status=400
            )

        # 目标目录
        target_dir = settings.S2S_TEMPLATE_DIR / template_name

        # 编辑模式：允许覆盖原模板
        if is_edit_mode and original_template_name:
            # 如果名称改变了，需要检查新名称是否已存在
            if template_name != original_template_name and target_dir.exists():
                return JsonResponse(
                    {"error": f"模板 '{template_name}' 已存在，请使用其他名称"},
                    status=400,
                )
            # 如果名称改变，需要删除原目录
            if template_name != original_template_name:
                old_dir = settings.S2S_TEMPLATE_DIR / original_template_name
                if old_dir.exists():
                    import shutil as shutil_old

                    shutil_old.rmtree(old_dir)
        else:
            # 新建模式：不允许覆盖
            if target_dir.exists():
                return JsonResponse(
                    {"error": f"模板 '{template_name}' 已存在，请使用其他名称"},
                    status=400,
                )

        # 获取 PPT 文件路径
        ppt_source_dir = settings.MEDIA_ROOT / "template_editor" / ppt_session_id
        ppt_source_path = ppt_source_dir / "template.pptx"

        # 如果 template.pptx 不存在，尝试查找任意 .pptx 文件（兼容旧会话）
        if not ppt_source_path.exists():
            ppt_files = list(ppt_source_dir.glob("*.pptx"))
            if ppt_files:
                ppt_source_path = ppt_files[0]
            else:
                return JsonResponse({"error": "PPT 文件不存在，请重新上传"}, status=400)

        # 创建目标目录
        target_dir.mkdir(parents=True, exist_ok=True)

        # 复制 PPT 文件
        import shutil

        pptx_target = target_dir / "template.pptx"
        shutil.copy2(ppt_source_path, pptx_target)

        # 保存 JSON 配置
        json_target = target_dir / "template.json"
        json_target.write_text(
            json.dumps(config_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        # 发布成功后清理相关会话记录
        import shutil as shutil_cleanup

        # 删除 wizard 会话
        wizard_sessions = TemplateEditSession.objects.filter(
            user=request.user, editor_type="wizard"
        )
        wizard_sessions.delete()

        # 删除 PPT 编辑器会话
        ppt_session = TemplateEditSession.objects.filter(
            user=request.user, session_id=ppt_session_id
        ).first()
        if ppt_session:
            ppt_session.delete()

        # 删除临时文件目录
        if ppt_source_dir.exists():
            shutil_cleanup.rmtree(ppt_source_dir)

        return JsonResponse(
            {
                "success": True,
                "message": f"模板 '{template_name}' 发布成功！",
                "template_path": str(target_dir),
                "pptx_path": str(pptx_target),
                "json_path": str(json_target),
            }
        )

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )
