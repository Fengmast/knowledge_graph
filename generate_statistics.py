from docx import Document
from collections import defaultdict

# 版本名称
version_names = [
    "新人教版",
    "沪科技版",
    "教科版",
    "粤教版",
    "鲁教版"
]

# 读取文档
doc = Document('知识图谱数据(1).docx')

# 提取所有表格数据
textbooks = []
for table in doc.tables:
    data = []
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        if any(row_data):
            data.append(row_data)
    if data:
        textbooks.append(data)

print("="*80)
print(" "*20 + "高中物理选择性必修2 - 五版本教材对比分析")
print("="*80)

# 统计每个版本
all_stats = []
for version_idx, textbook_data in enumerate(textbooks):
    version_name = version_names[version_idx]
    
    chapters = set()
    sections = set()
    points = []
    
    chapter_sections = defaultdict(set)
    section_points = defaultdict(list)
    
    current_chapter = None
    current_section = None
    
    for row in textbook_data:
        chapter = row[0] if len(row) > 0 else ''
        section = row[1] if len(row) > 1 else ''
        point = row[2] if len(row) > 2 else ''
        
        if chapter:
            current_chapter = chapter
            chapters.add(chapter)
        
        if section:
            current_section = section
            sections.add(section)
            if current_chapter:
                chapter_sections[current_chapter].add(section)
        
        if point:
            points.append(point)
            if current_section:
                section_points[current_section].append(point)
    
    stats = {
        'name': version_name,
        'chapters': len(chapters),
        'sections': len(sections),
        'points': len(points),
        'chapter_list': sorted(chapters),
        'avg_sections_per_chapter': len(sections) / len(chapters) if chapters else 0,
        'avg_points_per_section': len(points) / len(sections) if sections else 0
    }
    all_stats.append(stats)

# 打印总体统计
print("\n📊 整体统计对比\n")
print(f"{'版本':<12} {'章数':<8} {'节数':<8} {'知识点':<10} {'平均节/章':<12} {'平均点/节':<12}")
print("-"*80)

for stats in all_stats:
    print(f"{stats['name']:<12} {stats['chapters']:<8} {stats['sections']:<8} "
          f"{stats['points']:<10} {stats['avg_sections_per_chapter']:<12.2f} "
          f"{stats['avg_points_per_section']:<12.2f}")

# 详细章节结构
print("\n\n📚 各版本章节详细结构\n")
print("="*80)

for version_idx, textbook_data in enumerate(textbooks):
    version_name = version_names[version_idx]
    stats = all_stats[version_idx]
    
    print(f"\n【{version_name}选择性必修二】")
    print(f"共 {stats['chapters']} 章，{stats['sections']} 节，{stats['points']} 个知识点")
    print("-"*80)
    
    current_chapter = None
    chapter_num = 0
    section_num = 0
    
    for row in textbook_data:
        chapter = row[0] if len(row) > 0 else ''
        section = row[1] if len(row) > 1 else ''
        point = row[2] if len(row) > 2 else ''
        
        if chapter and chapter != current_chapter:
            current_chapter = chapter
            chapter_num += 1
            section_num = 0
            print(f"\n  {chapter}")
        
        if section:
            section_num += 1
            print(f"    {section}")

print("\n" + "="*80)
print("统计完成！")
print("="*80)

# 生成HTML报告
html_report = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>教材统计分析报告</title>
    <style>
        body {{
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            background: #f5f5f5;
            padding: 20px;
            margin: 0;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            text-align: center;
            border-bottom: 3px solid #667eea;
            padding-bottom: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 30px 0;
        }}
        th, td {{
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background: #667eea;
            color: white;
            font-weight: bold;
        }}
        tr:hover {{
            background: #f5f5f5;
        }}
        .version-section {{
            margin: 30px 0;
            padding: 20px;
            background: #f9f9f9;
            border-radius: 8px;
        }}
        .version-title {{
            font-size: 1.3em;
            color: #667eea;
            margin-bottom: 15px;
            font-weight: bold;
        }}
        .chapter {{
            margin: 15px 0;
            padding-left: 20px;
        }}
        .section {{
            margin: 5px 0;
            padding-left: 40px;
            color: #666;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 高中物理选择性必修2 - 五版本教材统计分析</h1>
        
        <h2>整体对比</h2>
        <table>
            <thead>
                <tr>
                    <th>版本</th>
                    <th>章数</th>
                    <th>节数</th>
                    <th>知识点数</th>
                    <th>平均节/章</th>
                    <th>平均知识点/节</th>
                </tr>
            </thead>
            <tbody>
"""

for stats in all_stats:
    html_report += f"""
                <tr>
                    <td><strong>{stats['name']}</strong></td>
                    <td>{stats['chapters']}</td>
                    <td>{stats['sections']}</td>
                    <td>{stats['points']}</td>
                    <td>{stats['avg_sections_per_chapter']:.2f}</td>
                    <td>{stats['avg_points_per_section']:.2f}</td>
                </tr>
"""

html_report += """
            </tbody>
        </table>
        
        <h2>各版本章节结构详情</h2>
"""

for version_idx, textbook_data in enumerate(textbooks):
    version_name = version_names[version_idx]
    stats = all_stats[version_idx]
    
    html_report += f"""
        <div class="version-section">
            <div class="version-title">{version_name}选择性必修二</div>
            <p>共 {stats['chapters']} 章，{stats['sections']} 节，{stats['points']} 个知识点</p>
"""
    
    current_chapter = None
    for row in textbook_data:
        chapter = row[0] if len(row) > 0 else ''
        section = row[1] if len(row) > 1 else ''
        
        if chapter and chapter != current_chapter:
            if current_chapter:
                html_report += "</div>"
            current_chapter = chapter
            html_report += f'<div class="chapter">📖 {chapter}'
        
        if section:
            html_report += f'<div class="section">└─ {section}</div>'
    
    html_report += "</div></div>"

html_report += """
    </div>
</body>
</html>
"""

# 保存HTML报告
with open('教材统计分析报告.html', 'w', encoding='utf-8') as f:
    f.write(html_report)

print("\n✓ HTML统计报告已生成: 教材统计分析报告.html")
