"""
Views for PPT Generator application.
"""

import sys
import json
import traceback
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

from .models import GlobalLLMConfig, PPTGeneration
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

    context = {
        "generation": generation,
        "is_developer": is_developer,
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

    if not generation.output_ppt:
        raise Http404("PPT文件不存在")

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

    context = {
        "is_developer": is_developer,
    }
    return render(request, "ppt_generator/developer_tools.html", context)


@login_required
@permission_required("ppt_generator.can_export_template_json", raise_exception=True)
def template_editor_page(request):
    """Template editor independent page."""
    is_developer = (
        request.user.groups.filter(name="开发者").exists() or request.user.is_superuser
    )

    context = {
        "is_developer": is_developer,
    }
    return render(request, "ppt_generator/template_editor.html", context)


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
@permission_required("ppt_generator.can_use_developer_tools", raise_exception=True)
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

        # 保存 PPT 文件
        ppt_path = temp_dir / ppt_file.name
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

        # 为每个页面生成标注图片
        pages = []
        for page_data in shapes_data["pages"]:
            page_num = page_data["page_num"]

            if page_num <= len(image_paths):
                image_path = image_paths[page_num - 1]

                # 生成标注图片（使用 150 DPI）
                annotated_path = annotate_screenshot(
                    image_path, page_data["shapes"], dpi=150
                )

                # 生成相对 URL
                relative_path = annotated_path.relative_to(settings.MEDIA_ROOT)
                image_url = f"/media/{relative_path}"

                pages.append(
                    {
                        "page_num": page_num,
                        "image_url": image_url,
                        "shapes": page_data["shapes"],
                    }
                )

        return JsonResponse(
            {
                "template_id": template_id,
                "ppt_path": str(ppt_path.relative_to(settings.MEDIA_ROOT)),
                "pages": pages,
            }
        )

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.can_use_developer_tools", raise_exception=True)
@require_http_methods(["POST"])
def update_shape_name_api(request):
    """
    更新元素名称

    Request:
        {
            "template_id": "uuid",
            "page_num": 1,
            "shape_id": 0,
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
        shape_id = data.get("shape_id")
        new_name = data.get("new_name")

        if not all([template_id, page_num is not None, shape_id is not None, new_name]):
            return JsonResponse({"error": "缺少必要参数"}, status=400)

        # 获取 PPT 文件路径
        ppt_path = settings.MEDIA_ROOT / "template_editor" / template_id
        ppt_files = list(ppt_path.glob("*.pptx"))

        if not ppt_files:
            return JsonResponse({"error": "找不到 PPT 文件"}, status=404)

        # 更新元素名称
        update_shape_name(ppt_files[0], page_num, shape_id, new_name)

        return JsonResponse({"success": True})

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.can_use_developer_tools", raise_exception=True)
@require_http_methods(["POST"])
def generate_template_config(request):
    """
    生成配置 JSON

    Request:
        {
            "template_id": "uuid"
        }

    Response:
        {
            "config": {...}
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

        # 生成配置 JSON
        config = {"template_name": ppt_files[0].stem, "ppt_pages": []}

        for page_data in shapes_data["pages"]:
            page_config = {
                "page_num": page_data["page_num"],
                "page_type": "content",  # 默认类型
                "fields": [],
            }

            for shape in page_data["shapes"]:
                # 只包含已命名的元素
                if shape.get("is_named") and not shape.get("is_hidden"):
                    field = {
                        "name": shape["name"],
                        "type": "text" if shape["type"] == "text" else "image",
                        "required": True,
                    }

                    if shape["type"] == "text" and shape.get("char_count"):
                        field["max_chars"] = shape["char_count"] * 2  # 预留空间

                    page_config["fields"].append(field)

            config["ppt_pages"].append(page_config)

        return JsonResponse({"config": config})

    except Exception as e:
        import traceback

        return JsonResponse(
            {"error": str(e), "traceback": traceback.format_exc()}, status=500
        )


@login_required
@permission_required("ppt_generator.can_use_developer_tools", raise_exception=True)
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
@permission_required("ppt_generator.can_use_developer_tools", raise_exception=True)
@require_http_methods(["POST"])
def toggle_shape_visibility(request):
    """
    切换元素的隐藏/显示状态

    Request:
        {
            "template_id": "uuid",
            "page_num": 1,
            "shape_id": 0,
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
