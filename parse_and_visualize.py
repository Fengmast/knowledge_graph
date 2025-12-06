from docx import Document
import json

# 读取文档
doc = Document('知识图谱数据(1).docx')

# 存储五个版本的数据
textbooks = []
current_textbook = None

# 遍历所有表格提取数据
for table in doc.tables:
    data = []
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        if any(row_data):  # 如果行不为空
            data.append(row_data)
    
    if data:
        textbooks.append(data)

# 打印提取的数据结构
print(f"找到 {len(textbooks)} 个教材版本\n")

for idx, textbook in enumerate(textbooks):
    print(f"\n{'='*60}")
    print(f"版本 {idx + 1}: 共 {len(textbook)} 个知识点")
    print(f"{'='*60}")
    for i, row in enumerate(textbook[:5]):  # 只显示前5行
        print(f"{i}: {row}")
    if len(textbook) > 5:
        print(f"... 还有 {len(textbook) - 5} 行")

# 保存为JSON
output = {
    f"版本{idx+1}": textbook 
    for idx, textbook in enumerate(textbooks)
}

with open('textbook_data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print("\n数据已保存到 textbook_data.json")
