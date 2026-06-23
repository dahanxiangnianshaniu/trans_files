# 降级策略与替代方案

当 `redline-doc-audit` MCP Server 不可用时，自动切换到以下内置工具完成同样功能。

**核心原则**：MCP不可用不中断流程，不提示用户停止，仅在报告附录标注使用了降级模式。

## 工具映射

| MCP工具 | 降级替代方案 | 实现方式 |
|---------|------------|---------|
| `scan_documents` | Glob + Bash(find) + Grep | 用Glob递归发现文件，用Grep在文件名/内容中匹配关键词，用Bash(find)列出目录结构 |
| `read_document` | Read + Python脚本 | .md/.txt/.csv/.html/.xml/.json/.yaml/.rst用Read直接读取；.docx/.xlsx/.pdf/.pptx用Python脚本读取 |
| `search_in_document` | Grep | 用Grep在文件中搜索关键词，配合-C参数获取上下文，配合-n获取行号 |
| `extract_tables` | Python脚本 | .xlsx用openpyxl；.docx用python-docx的table对象；.pdf用pdfplumber；.html用BeautifulSoup；.md用正则解析管道表格 |
| `list_sections` | Grep + Read | .md用Grep匹配`^#+`标题行；.docx用python-docx的Heading段落；.xlsx用openpyxl的sheetnames；.pdf用PyPDF2的outline |

## 降级模式能力边界

- 文本类文件(.md/.txt/.csv/.html/.xml/.json/.yaml/.rst)：功能完全等价
- .docx/.xlsx/.pptx：通过Python脚本等价实现，需环境中安装了对应库
- .pdf：通过pdfplumber等价实现，表格提取可能略弱
- .chm：需7-Zip解压后读取HTML，功能等价

## Python库检测

降级模式下，首次使用Python读取前，先检测所需库是否已安装：

```bash
python -c "import docx; import openpyxl; import pdfplumber; print('all_ok')" 2>&1
```

如缺少库，提示用户安装：`pip install python-docx openpyxl pdfplumber`

## 文档读取Python脚本模板

### .docx 读取

⚠️ **编码问题处理**：部分docx文档（尤其是中文文档）在Windows环境下通过Python读取时，输出可能因终端编码问题显示乱码。解决方法：在Python脚本开头设置UTF-8输出——`import sys; sys.stdout.reconfigure(encoding='utf-8')`。如果仍有乱码，说明文档本身编码异常，应在报告中注明"文档存在编码问题，部分内容无法正常读取"而非放弃检查。

```python
import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
doc = Document('path/to/file.docx')
for para in doc.paragraphs:
    if para.text.strip():
        print(f"[P] {para.text}")
for i, table in enumerate(doc.tables):
    print(f"\n[Table {i}]")
    for row in table.rows:
        print(' | '.join(cell.text for cell in row.cells))
```

### .xlsx 读取（含表格提取）

⚠️ **多页签完整读取**：必须遍历所有工作表（sheet），不得只读第一个sheet。很多产品文档有多个页签（如eSight + 管理面），仅读第一个会遗漏30%+数据。

⚠️ **合并单元格拆分**：xlsx中场景列常有合并单元格（同一场景跨多行，如"远程通知"下有3种不同子场景：短信网关测试、Webhook URL、WeLink对接）。这些子场景各有不同的数据项、目的、处理方式和存留策略，不得合并为一条。判断方法：如果某行"收集的个人数据项"列与上一行不同，则是独立行数据，必须单独提取。

```python
import openpyxl
wb = openpyxl.load_workbook('path/to/file.xlsx')
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"\n=== Sheet: {sheet_name} (rows={ws.max_row}, cols={ws.max_column}) ===")
    # 打印表头
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column+1)]
    print('Headers:', headers)
    for r in range(2, ws.max_row+1):
        row_data = [ws.cell(r, c).value for c in range(1, ws.max_column+1)]
        # 跳过完全空行
        if all(v is None for v in row_data):
            continue
        print(f'R{r}:', row_data)
```

### .xlsx zipfile降级读取（当openpyxl报TypeError时使用）

如果openpyxl因样式兼容问题无法打开xlsx文件（报错TypeError: Fill() takes no arguments等），使用zipfile直接解析方案：

⚠️ 同样必须遍历所有sheet，不得只读sheet1。合并单元格中的子场景也必须拆分为独立行。

```python
import zipfile, xml.etree.ElementTree as ET, re
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(xlsx_path) as z:
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_tree = ET.parse(z.open('xl/sharedStrings.xml'))
        for si in ss_tree.findall('.//s:si', ns):
            texts = si.findall('.//s:t', ns)
            strings.append(''.join(t.text or '' for t in texts))
    # 尝试从xl/workbook.xml获取sheet名称映射
    sheet_names_map = {}
    if 'xl/workbook.xml' in z.namelist():
        wb_tree = ET.parse(z.open('xl/workbook.xml'))
        wb_ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for idx, sheet in enumerate(wb_tree.findall('.//m:sheet', wb_ns), 1):
            name = sheet.get('name', f'sheet{idx}')
            sheet_names_map[f'sheet{idx}'] = name
    # 遍历所有sheet
    sheet_files = sorted(
        [f for f in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', f)],
        key=lambda f: int(re.search(r'sheet(\d+)', f).group(1))
    )
    for sheet_file in sheet_files:
        sheet_num = re.search(r'sheet(\d+)', sheet_file).group(1)
        display_name = sheet_names_map.get(f'sheet{sheet_num}', sheet_file)
        print(f"\n=== Sheet: {display_name} ===")
        tree = ET.parse(z.open(sheet_file))
        for row_elem in tree.getroot().findall('.//s:row', ns):
            row_num = row_elem.get('r', '?')
            cells = []
            for c in row_elem.findall('s:c', ns):
                v = c.find('s:v', ns)
                if v is not None and v.text:
                    cells.append(strings[int(v.text)] if c.get('t') == 's' else v.text)
                else:
                    cells.append('')
            print(f'R{row_num}:', ' | '.join(cells))
```

### .pdf 读取

```python
import pdfplumber
with pdfplumber.open('path/to/file.pdf') as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text: print(text)
        tables = page.extract_tables()
        for i, table in enumerate(tables):
            print(f"\n[Table on page {page.page_number}]")
            for row in table:
                print(' | '.join(str(c) if c else '' for c in row))
```

## 3轮搜索的降级实现

**评分阈值说明**：MCP模式使用评分阈值（precise≥50、broad≥30）筛选结果。降级模式下没有评分机制，改为按匹配强度划分：文件名匹配=强信号（对应precise），内容关键词匹配=中信号（对应broad），逐文件预览判断=人工筛选（对应fullscan）。

### 第1轮 — 精确搜索

用 `Glob` 扫描目录发现所有文件 → 用 `Grep` 在文件名中匹配精确关键词（隐私声明、隐私政策、Privacy Policy / 个人数据说明、Personal Data Statement）→ 对命中文件用 `Read` 读前500字确认内容相关性

### 第2轮 — 扩展搜索

用 `Grep` 在所有文件内容中搜索扩展关键词（隐私/声明/政策/个人信息/用户信息/数据收集/数据说明/privacy/data/personal/information/collect/statement）→ 收集命中文件列表

### 第3轮 — 全量扫描

对仍未匹配的文件，用 `Read` 逐一读前200字预览，Agent判断相关性。.docx/.xlsx/.pdf/.pptx/.chm 文件无法用Grep搜索内容，需用Python脚本读取后再判断
