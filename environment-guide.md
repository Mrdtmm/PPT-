# PPT制作环境配置与关键技术分析

本文档分析当前环境制作高质量PPT的关键因素，便于在其他环境中复现。

## 当前环境核心工具

### 1. Python 包依赖

```bash
pip install python-pptx pillow python-docx openpyxl
```

| 包名 | 版本 | 用途 |
|------|------|------|
| python-pptx | 1.0.2 | PPT生成与编辑 |
| pillow | 12.0.0 | 图片处理（保持宽高比） |
| python-docx | 1.2.0 | 读取Word文档 |
| openpyxl | 3.1.5 | Excel处理（如需要） |

### 2. MCP Server（可选）

```
pptx-mcp-server 0.1.0  # PPT生成MCP服务
mcp 1.27.1             # MCP协议支持
fastmcp 1.0            # FastMCP框架
```

## 高质量PPT的关键技术

### 1. 图片宽高比保持（核心）

```python
from PIL import Image as PILImage

def add_image(slide, img_path, left, top, width, height):
    """添加图片（保持宽高比，自动缩放居中）"""
    if os.path.exists(img_path):
        with PILImage.open(img_path) as img:
            img_width, img_height = img.size
        aspect = img_width / img_height

        if width / height > aspect:
            new_height = height
            new_width = height * aspect
        else:
            new_width = width
            new_height = width / aspect

        offset_x = left + (width - new_width) / 2
        offset_y = top + (height - new_height) / 2

        slide.shapes.add_picture(
            str(img_path),
            Inches(offset_x), Inches(offset_y),
            Inches(new_width), Inches(new_height)
        )
```

**关键点：** 使用 Pillow 读取图片原始尺寸，计算缩放比例，避免拉伸变形。

### 2. 幻灯片尺寸配置

```python
prs = Presentation()
prs.slide_width = Inches(13.333)   # 16:9 宽屏比例
prs.slide_height = Inches(7.5)
```

### 3. 颜色配置（学术风格）

```python
SCHOOL_BLUE = RGBColor(0x1a, 0x3a, 0x7a)  # 主色调蓝色
ACCENT = RGBColor(0x1a, 0xaf, 0x6c)       # 绿色强调
BAD = RGBColor(0xc1, 0x3a, 0x3a)           # 红色警示
DARK_TEXT = RGBColor(0x0c, 0x0d, 0x10)     # 深色文字
GRAY_TEXT = RGBColor(0x55, 0x59, 0x6a)     # 灰色文字
LIGHT_GRAY = RGBColor(0xf5, 0xf5, 0xf6)     # 浅灰背景
```

### 4. 卡片式布局

```python
def add_card(slide, left, top, width, height, items, title=None):
    """添加卡片内容"""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_GRAY
    card.line.color.rgb = RGBColor(0xe8, 0xe8, 0xe8)
    # ... 添加标题和列表项
```

### 5. 页面装饰元素

```python
def add_slide_header(slide):
    """顶部蓝色装饰条"""
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
        Inches(13.333), Inches(0.1)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = SCHOOL_BLUE
    header.line.fill.background()

def add_slide_footer(slide, section_name, page_num):
    """底部页脚（区域名 + 页码）"""
    footer_left = slide.shapes.add_textbox(
        Inches(0.5), Inches(7.1), Inches(4), Inches(0.3)
    )
    footer_left.text_frame.paragraphs[0].text = section_name
    # ... 页码右侧显示
```

## 图片尺寸优化

### 大图展示区域
```python
# 右侧大图：6.7 x 5.0 英寸
add_image(slide, img_path, 6.3, 1.3, 6.7, 5.0)

# 三图并排：4.2 x 5.0 英寸
add_image(slide9, img1, 0.3, 1.5, 4.2, 5)
add_image(slide9, img2, 4.55, 1.5, 4.2, 5)
add_image(slide9, img3, 8.8, 1.5, 4.2, 5)
```

## 环境复现检查清单

- [ ] Python >= 3.8
- [ ] python-pptx >= 1.0.0
- [ ] pillow >= 10.0.0
- [ ] python-docx >= 1.0.0
- [ ] 使用 16:9 幻灯片比例
- [ ] 使用 Pillow 保持图片宽高比
- [ ] 配置学术风格颜色（蓝色主调）

## 相关文件

- PPT生成脚本：E:\CODE\zhukefu\sunxiaoyu\create_ppt_v3.py
- 参考HTML：E:\CODE\zhukefu\sunxiaoyu\结题汇报_v3.html
- html-ppt-skill：E:\CODE\zhukefu\html-ppt-skill\
