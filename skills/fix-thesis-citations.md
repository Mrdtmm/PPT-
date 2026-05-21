# 论文引用修复与格式标准化技能

## 描述
用于修复Word论文文档中的引用格式问题，包括上标引用添加、双括号清理、重复内容修复等。

## 来源
从 `E:\CODE\zhukefu` 项目中提取的经验总结（2026-05-18）

## 背景
处理中文论文文档时，遇到以下典型问题：
1. 引用文献数量不足，需要从30增加到49
2. 引用编号顺序错误（[49-31]而非[31-49]）
3. 正文中的引用需要改为上标格式
4. 各种格式错误：`[[`、`]]`、连续重复引用 `[num][num]`
5. 修复过程中产生的文字重复问题

## 核心工作流程

### 1. 从原始文件开始
```python
from docx import Document
doc = Document('原始论文.docx')
```

### 2. XML层面添加上标引用（避免双括号问题）
```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_superscript_xml(paragraph, keyword, citation):
    """通过XML操作添加上标引用"""
    text = paragraph.text
    if keyword not in text or citation in text:
        return False

    # 找到插入位置
    idx = text.find(keyword)
    keyword_end = idx + len(keyword)

    insert_pos = keyword_end
    for j in range(keyword_end, min(keyword_end + 5, len(text))):
        if text[j] in '。；，、':
            insert_pos = j
            break

    # 找到目标run
    pos = 0
    runs = paragraph._p.findall(qn('w:r'))
    target_run_elem = None
    split_offset = 0

    for run_elem in runs:
        run_text_elem = run_elem.find(qn('w:t'))
        if run_text_elem is None:
            pos += 1
            continue

        run_text = run_text_elem.text or ''
        run_start = pos
        run_end = pos + len(run_text)

        if run_start <= insert_pos <= run_end:
            target_run_elem = run_elem
            split_offset = insert_pos - run_start
            break
        pos = run_end

    if target_run_elem is None:
        return False

    t_elem = target_run_elem.find(qn('w:t'))
    run_text = t_elem.text or ''
    before_text = run_text[:split_offset]
    after_text = run_text[split_offset:]

    # 更新原run
    t_elem.text = before_text

    # 创建上标run（关键：citation已经是[31]格式，不要再包装）
    sup_r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    vertAlign = OxmlElement('w:vertAlign')
    vertAlign.set(qn('w:val'), 'superscript')
    rPr.append(vertAlign)
    sup_r.append(rPr)

    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = citation  # 例：[31]
    sup_r.append(t)

    target_run_elem.addnext(sup_r)

    if after_text:
        after_r = OxmlElement('w:r')
        after_t = OxmlElement('w:t')
        after_t.set(qn('xml:space'), 'preserve')
        after_t.text = after_text
        after_r.append(after_t)
        sup_r.addnext(after_r)

    return True
```

### 3. 验证格式问题
```python
import re

def check_format_issues(doc):
    """检查常见格式问题"""
    problems = {'[[': 0, ']]': 0, '连续引用': 0, '句末]': 0}

    for p in doc.paragraphs:
        text = p.text
        if '[[' in text:
            problems['[['] += 1
        if ']]' in text:
            problems[']]'] += 1
        if re.search(r'\[(\d+)\]\[(\d+)\]', text):
            problems['连续引用'] += 1
        if re.search(r'。\]', text):
            problems['句末]'] += 1

    return problems

def count_superscript(doc):
    """统计上标引用数量"""
    count = 0
    for p in doc.paragraphs:
        for run in p.runs:
            if 'superscript' in run._r.xml:
                count += 1
    return count
```

## 常见问题与解决方案

### 问题1：双括号 [[ 或 ]]
**原因**: 插入上标时文字分割不当
**解决**: 使用XML层面的插入，确保 citation 参数格式为 `[31]` 而非额外添加括号

### 问题2：连续重复引用 [num][num]
**原因**: 同一位置多次插入
**解决**: 每次插入前检查 `if citation in text: return False`

### 问题3：文字内容缺失
**原因**: 修复过程中误删或重复添加
**解决**: 从原始文件重新开始，避免重复修改

### 问题4：段落重复（如"利用利用"）
**原因**: 正则替换时未处理好边界
**解决**: 使用更精确的匹配模式，并从原文件重建

## 重要提示

1. **从原始文件开始**：每次修复都从原文件重建，避免误差累积
2. **使用 save() 而非 save(path) 时注意权限**：如果文件已打开会保存失败
3. **验证每次修改**：保存后立即检查格式问题
4. **分批处理**：避免一次性大量修改导致问题

## 相关文件
- 原始论文：`E:\CODE\zhukefu\论文\朱可夫+202165201135+论文.docx`
- 修复后：`E:\CODE\zhukefu\论文\朱可夫+202165201135+论文_新最终版.docx`
- 工作脚本目录：`E:\CODE\zhukefu\`

## 使用频率
- 适用场景：中文Word论文的引用格式修复
- 注意事项：每次修改后务必验证，避免产生新问题