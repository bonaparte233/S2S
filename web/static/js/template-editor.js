/**
 * PPT 模板编辑器 - Canvas 交互式预览
 * 
 * 功能：
 * 1. Canvas 渲染 PPT 截图 + 编号标注
 * 2. 点击、悬停交互
 * 3. 双向联动（Canvas ↔ 列表）
 * 4. 元素命名编辑
 */

class CanvasPreview {
    constructor(canvasId, screenshotUrl, shapesData) {
        this.canvas = document.getElementById(canvasId);
        if (!this.canvas) {
            console.error('Canvas element not found:', canvasId);
            return;
        }

        this.ctx = this.canvas.getContext('2d');
        this.screenshotUrl = screenshotUrl;
        this.shapes = shapesData || [];

        // 状态
        this.selectedShapeId = null;
        this.hoveredShapeId = null;
        this.scale = 1.0;
        this.offsetX = 0;
        this.offsetY = 0;
        this.isPanning = false;
        this.lastMouseX = 0;
        this.lastMouseY = 0;

        // 常量
        this.EMU_PER_INCH = 914400;
        this.DPI = 300;

        this.init();
    }

    init() {
        // 加载截图
        const img = new Image();
        img.onload = () => {
            this.screenshot = img;
            this.canvas.width = img.width;
            this.canvas.height = img.height;
            this.render();
        };
        img.onerror = () => {
            console.error('Failed to load screenshot:', this.screenshotUrl);
        };
        img.src = this.screenshotUrl;

        // 绑定事件
        this.canvas.addEventListener('click', (e) => this.handleClick(e));
        this.canvas.addEventListener('mousemove', (e) => this.handleMouseMove(e));
        this.canvas.addEventListener('mousedown', (e) => this.handleMouseDown(e));
        this.canvas.addEventListener('mouseup', (e) => this.handleMouseUp(e));
        this.canvas.addEventListener('wheel', (e) => this.handleWheel(e));
        this.canvas.addEventListener('mouseleave', () => {
            this.isPanning = false;
            this.canvas.style.cursor = 'crosshair';
        });
    }

    render() {
        if (!this.screenshot) return;

        // 清空画布
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);

        // 应用缩放和平移
        this.ctx.save();
        this.ctx.translate(this.offsetX, this.offsetY);
        this.ctx.scale(this.scale, this.scale);

        // 1. 绘制截图背景
        this.ctx.drawImage(this.screenshot, 0, 0);

        // 2. 绘制元素边框和编号
        this.shapes.forEach((shape, index) => {
            if (shape.is_hidden) return;
            this.drawShape(shape, index + 1);
        });

        this.ctx.restore();
    }

    drawShape(shape, number) {
        const x = shape.left / this.EMU_PER_INCH * this.DPI;
        const y = shape.top / this.EMU_PER_INCH * this.DPI;
        const w = shape.width / this.EMU_PER_INCH * this.DPI;
        const h = shape.height / this.EMU_PER_INCH * this.DPI;

        // 确定状态
        const isSelected = this.selectedShapeId === shape.shape_id;
        const isHovered = this.hoveredShapeId === shape.shape_id;
        const isNamed = shape.is_named;

        // 绘制边框和填充
        if (isSelected || isHovered) {
            // 边框
            this.ctx.strokeStyle = isSelected ? '#007BFF' : (isNamed ? '#28A745' : '#FFC107');
            this.ctx.lineWidth = isSelected ? 3 : 2;
            this.ctx.setLineDash(isNamed ? [] : [5, 5]);
            this.ctx.strokeRect(x, y, w, h);

            // 半透明填充
            this.ctx.fillStyle = isSelected
                ? 'rgba(0, 123, 255, 0.15)'
                : (isNamed ? 'rgba(40, 167, 69, 0.05)' : 'rgba(255, 193, 7, 0.05)');
            this.ctx.fillRect(x, y, w, h);

            this.ctx.setLineDash([]);
        }

        // 绘制编号圆圈
        const circleX = x + 20;
        const circleY = y + 20;
        const radius = isHovered ? 24 : 22;

        // 圆圈颜色
        let circleColor;
        if (isSelected) {
            circleColor = '#28A745';  // 绿色
        } else if (isNamed) {
            circleColor = '#007BFF';  // 蓝色
        } else {
            circleColor = '#FFC107';  // 橙色
        }

        // 绘制阴影
        this.ctx.shadowColor = 'rgba(0, 0, 0, 0.2)';
        this.ctx.shadowBlur = 4;
        this.ctx.shadowOffsetX = 0;
        this.ctx.shadowOffsetY = 2;

        // 绘制圆圈
        this.ctx.fillStyle = circleColor;
        this.ctx.beginPath();
        this.ctx.arc(circleX, circleY, radius, 0, 2 * Math.PI);
        this.ctx.fill();

        // 绘制白色边框
        this.ctx.strokeStyle = '#FFFFFF';
        this.ctx.lineWidth = 3;
        this.ctx.stroke();

        // 重置阴影
        this.ctx.shadowColor = 'transparent';

        // 绘制编号
        this.ctx.fillStyle = '#FFFFFF';
        this.ctx.font = 'bold 24px Arial';
        this.ctx.textAlign = 'center';
        this.ctx.textBaseline = 'middle';
        this.ctx.fillText(number, circleX, circleY);
    }

    handleClick(e) {
        if (this.isPanning) return;

        const rect = this.canvas.getBoundingClientRect();
        const x = (e.clientX - rect.left - this.offsetX) / this.scale;
        const y = (e.clientY - rect.top - this.offsetY) / this.scale;

        // 查找点击的元素
        const clickedShape = this.shapes.find(shape => {
            if (shape.is_hidden) return false;

            const sx = shape.left / this.EMU_PER_INCH * this.DPI;
            const sy = shape.top / this.EMU_PER_INCH * this.DPI;
            const sw = shape.width / this.EMU_PER_INCH * this.DPI;
            const sh = shape.height / this.EMU_PER_INCH * this.DPI;

            return x >= sx && x <= sx + sw && y >= sy && y <= sy + sh;
        });

        if (clickedShape) {
            this.selectShape(clickedShape);
        }
    }

    handleMouseMove(e) {
        if (this.isPanning) {
            const dx = e.clientX - this.lastMouseX;
            const dy = e.clientY - this.lastMouseY;
            this.offsetX += dx;
            this.offsetY += dy;
            this.lastMouseX = e.clientX;
            this.lastMouseY = e.clientY;
            this.render();
            return;
        }

        const rect = this.canvas.getBoundingClientRect();
        const x = (e.clientX - rect.left - this.offsetX) / this.scale;
        const y = (e.clientY - rect.top - this.offsetY) / this.scale;

        // 查找悬停的元素
        const hoveredShape = this.shapes.find(shape => {
            if (shape.is_hidden) return false;

            const sx = shape.left / this.EMU_PER_INCH * this.DPI;
            const sy = shape.top / this.EMU_PER_INCH * this.DPI;
            const sw = shape.width / this.EMU_PER_INCH * this.DPI;
            const sh = shape.height / this.EMU_PER_INCH * this.DPI;

            return x >= sx && x <= sx + sw && y >= sy && y <= sy + sh;
        });

        if (hoveredShape) {
            this.canvas.style.cursor = 'pointer';
            this.hoveredShapeId = hoveredShape.shape_id;
        } else {
            this.canvas.style.cursor = this.isPanning ? 'grabbing' : 'crosshair';
            this.hoveredShapeId = null;
        }

        this.render();
    }

    handleMouseDown(e) {
        if (e.button === 0 && e.shiftKey) {  // Shift + 左键 = 平移
            this.isPanning = true;
            this.lastMouseX = e.clientX;
            this.lastMouseY = e.clientY;
            this.canvas.style.cursor = 'grabbing';
        }
    }

    handleMouseUp(e) {
        if (e.button === 0) {
            this.isPanning = false;
            this.canvas.style.cursor = 'crosshair';
        }
    }

    handleWheel(e) {
        e.preventDefault();

        // 缩放
        const delta = e.deltaY > 0 ? 0.9 : 1.1;
        this.scale *= delta;
        this.scale = Math.max(0.1, Math.min(5, this.scale));

        // 更新缩放显示
        const zoomLevel = document.getElementById('zoomLevel');
        if (zoomLevel) {
            zoomLevel.textContent = Math.round(this.scale * 100) + '%';
        }

        this.render();
    }

    selectShape(shape) {
        this.selectedShapeId = shape.shape_id;
        this.render();

        // 触发自定义事件，通知外部组件
        const event = new CustomEvent('shapeSelected', { detail: shape });
        this.canvas.dispatchEvent(event);
    }

    zoomIn() {
        this.scale *= 1.2;
        this.scale = Math.min(5, this.scale);
        document.getElementById('zoomLevel').textContent = Math.round(this.scale * 100) + '%';
        this.render();
    }

    zoomOut() {
        this.scale *= 0.8;
        this.scale = Math.max(0.1, this.scale);
        document.getElementById('zoomLevel').textContent = Math.round(this.scale * 100) + '%';
        this.render();
    }

    fitToWindow() {
        if (!this.screenshot) return;

        const container = this.canvas.parentElement;
        const containerWidth = container.clientWidth;
        const containerHeight = container.clientHeight;

        const scaleX = containerWidth / this.screenshot.width;
        const scaleY = containerHeight / this.screenshot.height;

        this.scale = Math.min(scaleX, scaleY, 1);
        this.offsetX = 0;
        this.offsetY = 0;

        document.getElementById('zoomLevel').textContent = Math.round(this.scale * 100) + '%';
        this.render();
    }

    updateShapes(shapesData) {
        this.shapes = shapesData || [];
        this.render();
    }
}

