# DOCX to PPT with Images - Skill

## 描述
从Word文档(.docx)中提取图片素材，并根据内容重命名，最后生成图文并茂的PPT演示文稿。

## 适用场景
- 学术论文结题汇报PPT制作
- 从docx演讲稿自动生成PPT
- 需要提取docx中内嵌图片并用于PPT

## 依赖工具

### Python库
```bash
pip install python-docx pillow python-pptx
```

| 库 | 用途 |
|----|------|
| `python-docx` | 读取docx文档结构 |
| `pillow` | 读取图片尺寸，保持宽高比 |
| `python-pptx` | 生成和编辑PPT文件 |

## 工作流程

### 步骤1：提取docx中的图片

```python
import zipfile
import os

docx_path = 'path/to/你的文档.docx'
output_dir = 'path/to/图片素材目录'

os.makedirs(output_dir, exist_ok=True)

with zipfile.ZipFile(docx_path, 'r') as zf:
    media_files = [n for n in zf.namelist() 
                   if n.startswith('word/media/') and not n.endswith('/')]
    
    for i, name in enumerate(media_files, 1):
        basename = os.path.basename(name)
        ext = os.path.splitext(basename)[1]
        out_name = f'image_{i:02d}{ext}'
        out_path = os.path.join(output_dir, out_name)
        
        with zf.open(name) as src, open(out_path, 'wb') as dst:
            dst.write(src.read())
        print(f'Extracted: {out_name}')
```

### 步骤2：重命名图片为中文标识

根据HTML参考文件或内容分析，将图片重命名为有意义的中文名称：
```bash
mv image_01.png "01_研究概念图.png"
mv image_02.png "02_技术路线图.png"
# ... 以此类推
```

### 步骤3：创建PPT脚本（关键模式）

```python
#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image as PILImage

# ========== 配置 ==========
BASE_DIR = r"E:\path\to\output"
IMG_DIR = os.path.join(BASE_DIR, "图片素材")
OUTPUT_FILE = os.path.join(BASE_DIR, "最终PPT.pptx")

# 颜色配置
SCHOOL_BLUE = RGBColor(0x1a, 0x3a, 0x7a)
ACCENT = RGBColor(0x1a, 0xaf, 0x6c)  # 绿色
BAD = RGBColor(0xc1, 0x3a, 0x3a)    # 红色
DARK_TEXT = RGBColor(0x0c, 0x0d, 0x10)
GRAY_TEXT = RGBColor(0x55, 0x59, 0x6a)
LIGHT_GRAY = RGBColor(0xf5, 0xf5, 0xf6)

# 幻灯片尺寸
prs = Presentation()
prs.slide_width = Inches(13.333)  # 16:9 比例
prs.slide_height = Inches(7.5)

# ========== 辅助函数 ==========

def add_slide_header(slide):
    """顶部蓝色装饰条"""
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), 
        Inches(13.333), Inches(0.1)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = SCHOOL_BLUE
    header.line.fill.background()

def add_slide_footer(slide, section_name, page_num, total=13):
    """底部页脚"""
    footer_left = slide.shapes.add_textbox(
        Inches(0.5), Inches(7.1), Inches(4), Inches(0.3)
    )
    footer_left.text_frame.paragraphs[0].text = section_name
    footer_left.text_frame.paragraphs[0].font.size = Pt(10)
    footer_left.text_frame.paragraphs[0].font.color.rgb = GRAY_TEXT
    
    footer_right = slide.shapes.add_textbox(
        Inches(11.5), Inches(7.1), Inches(1.5), Inches(0.3)
    )
    footer_right.text_frame.paragraphs[0].text = f"{page_num} / {total}"
    footer_right.text_frame.paragraphs[0].font.size = Pt(12)
    footer_right.text_frame.paragraphs[0].font.color.rgb = GRAY_TEXT
    footer_right.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

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

def add_card(slide, left, top, width, height, items, title=None):
    """添加卡片内容"""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_GRAY
    card.line.color.rgb = RGBColor(0xe8, 0xe8, 0xe8)
    
    content_top = top + 0.2
    if title:
        title_box = slide.shapes.add_textbox(
            Inches(left + 0.2), Inches(content_top),
            Inches(width - 0.4), Inches(0.4)
        )
        title_box.text_frame.paragraphs[0].text = title
        title_box.text_frame.paragraphs[0].font.size = Pt(14)
        title_box.text_frame.paragraphs[0].font.bold = True
        title_box.text_frame.paragraphs[0].font.color.rgb = DARK_TEXT
        content_top += 0.5
    
    for i, item in enumerate(items):
        item_box = slide.shapes.add_textbox(
            Inches(left + 0.3), Inches(content_top + i*0.45),
            Inches(width - 0.5), Inches(0.4)
        )
        tf = item_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "• " + item
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_TEXT

# ========== 创建幻灯片 ==========

# 示例：创建包含图片的幻灯片
slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局
add_slide_header(slide)

# 左侧文字卡片
add_card(slide, 0.5, 1.4, 5.5, 2.5, [
    "要点1",
    "要点2",
    "要点3"
], "标题")

# 右侧图片（保持宽高比）
add_image(slide, os.path.join(IMG_DIR, "图片.png"), 
          6.3, 1.3, 6.7, 5.0)

add_slide_footer(slide, "section", 1)

# ========== 保存 ==========
prs.save(OUTPUT_FILE)
print(f"PPT已保存至: {OUTPUT_FILE}")
```

## 关键配置说明

### 幻灯片尺寸
- 16:9 比例：`Inches(13.333)` x `Inches(7.5)`
- 4:3 比例：`Inches(10)` x `Inches(7.5)`

### 图片保持宽高比的核心代码
```python
def add_image(slide, img_path, left, top, width, height):
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

### 颜色配置
| 颜色 | RGB值 | 用途 |
|------|-------|------|
| 学校蓝 | (0x1a, 0x3a, 0x7a) | 主色调、标题 |
| 强调绿 | (0x1a, 0xaf, 0x6c) | 正面标记、成功 |
| 警示红 | (0xc1, 0x3a, 0x3a) | 问题标记、强调 |
| 深色文字 | (0x0c, 0x0d, 0x10) | 正文 |
| 灰色文字 | (0x55, 0x59, 0x6a) | 次要文字 |

## 文件命名建议

从docx提取的图片按内容重命名：
```
01_研究概念图.png
02_技术路线图.png
03_结果_western.png
04_结果_Cl检测.png
05_结果_VCAM1.png
06_结果_NLRP3.png
07_结果_对比图.png
08_Statistical_analysis.png
09_结果_蛋白表达.png
10_结论示意图.png
```

## 验证方式

1. 运行脚本生成PPT
2. 在PowerPoint中打开检查：
   - 图片是否清晰（无拉伸变形）
   - 图片是否居中显示
   - 文字与图片是否重叠
   - 页脚页码是否正确

## 相关文件
- PPT生成脚本：`E:\CODE\zhukefu\sunxiaoyu\create_ppt_v3.py`
- 图片素材：`E:\CODE\zhukefu\sunxiaoyu\图片素材\`
- 参考HTML：`E:\CODE\zhukefu\sunxiaoyu\结题汇报_v3.html`
