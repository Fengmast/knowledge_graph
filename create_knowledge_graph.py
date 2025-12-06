from docx import Document
from pyvis.network import Network
import networkx as nx

# 版本名称
version_names = [
    "新人教版选择性必修二",
    "沪科技版选择性必修二",
    "教科版选择性必修二",
    "粤教版选择性必修二",
    "鲁教版选择性必修二"
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

print(f"开始创建 {len(textbooks)} 个知识图谱...")

# 为每个版本创建知识图谱
for version_idx, textbook_data in enumerate(textbooks):
    version_name = version_names[version_idx]
    print(f"\n正在处理: {version_name}")
    
    # 创建网络图
    net = Network(
        height='1000px',
        width='100%',
        bgcolor='#ffffff',
        font_color='#000000',
        directed=False
    )
    
    # 设置为静态布局，可手动拖动调整
    net.set_options('''
    {
        "physics": {
            "enabled": false
        },
        "interaction": {
            "dragNodes": true,
            "dragView": true,
            "zoomView": true
        },
        "nodes": {
            "font": {
                "size": 36,
                "face": "Microsoft YaHei",
                "bold": true
            },
            "margin": 25,
            "borderWidth": 3
        },
        "edges": {
            "smooth": {
                "type": "continuous",
                "roundness": 0.5
            }
        },
        "layout": {
            "randomSeed": 42,
            "improvedLayout": true,
            "hierarchical": {
                "enabled": false
            }
        }
    }
    ''')
    
    # 添加中心节点（教材名称）
    net.add_node(
        version_name,
        label=version_name,
        size=100,
        color='#FF6B6B',
        font={'size': 50, 'bold': True},
        shape='box'
    )
    
    current_chapter = None
    current_section = None
    
    # 遍历知识点数据
    for row in textbook_data:
        chapter = row[0] if len(row) > 0 else ''
        section = row[1] if len(row) > 1 else ''
        point = row[2] if len(row) > 2 else ''
        
        # 处理章节
        if chapter and chapter != current_chapter:
            current_chapter = chapter
            chapter_id = f"{version_name}_{chapter}"
            
            # 添加章节节点
            net.add_node(
                chapter_id,
                label=chapter,
                size=70,
                color='#4ECDC4',
                shape='ellipse',
                font={'size': 38, 'bold': True}
            )
            
            # 连接到教材中心节点
            net.add_edge(version_name, chapter_id, width=3, color='#888888')
        
        # 处理小节
        if section and section != current_section:
            current_section = section
            section_id = f"{version_name}_{chapter}_{section}"
            
            # 添加小节节点
            net.add_node(
                section_id,
                label=section,
                size=55,
                color='#95E1D3',
                shape='ellipse',
                font={'size': 32, 'bold': True}
            )
            
            # 连接到章节
            if current_chapter:
                net.add_edge(chapter_id, section_id, width=2, color='#AAAAAA')
        
        # 处理知识点
        if point:
            point_id = f"{version_name}_{chapter}_{section}_{point}"
            
            # 添加知识点节点
            net.add_node(
                point_id,
                label=point,
                size=40,
                color='#F38181',
                shape='dot',
                font={'size': 28, 'bold': True}
            )
            
            # 连接到小节
            if current_section:
                net.add_edge(section_id, point_id, width=1, color='#CCCCCC')
    
    # 保存HTML文件
    filename = f"知识图谱_{version_name}.html"
    net.save_graph(filename)
    
    # 读取生成的HTML并添加保存/恢复位置的功能
    with open(filename, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # 添加保存和恢复位置的JavaScript代码
    save_restore_script = '''
    <script type="text/javascript">
        // 保存节点位置到localStorage
        function savePositions() {
            const positions = network.getPositions();
            const storageKey = 'graph_positions_' + document.title;
            localStorage.setItem(storageKey, JSON.stringify(positions));
            
            // 显示保存成功提示
            const btn = document.getElementById('save-btn');
            const originalText = btn.innerHTML;
            btn.innerHTML = '✓ 已保存！';
            btn.style.background = '#4CAF50';
            setTimeout(() => {
                btn.innerHTML = originalText;
                btn.style.background = '#667eea';
            }, 2000);
        }
        
        // 从localStorage恢复节点位置
        function loadPositions() {
            const storageKey = 'graph_positions_' + document.title;
            const savedPositions = localStorage.getItem(storageKey);
            if (savedPositions) {
                const positions = JSON.parse(savedPositions);
                network.setPositions(positions);
                console.log('已恢复保存的布局');
            }
        }
        
        // 重置到初始布局
        function resetPositions() {
            const storageKey = 'graph_positions_' + document.title;
            localStorage.removeItem(storageKey);
            location.reload();
        }
        
        // 在network初始化后立即恢复位置
        network.on('stabilizationIterationsDone', function() {
            setTimeout(() => {
                loadPositions();
            }, 100);
        });
        
        // 页面加载完成后恢复位置
        window.addEventListener('load', function() {
            setTimeout(() => {
                loadPositions();
            }, 500);
        });
    </script>
    
    <style>
        .control-panel {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 1000;
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
        }
        
        .control-panel.minimized {
            padding: 0;
        }
        
        .control-panel.minimized .panel-content {
            display: none;
        }
        
        .panel-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 15px;
            cursor: pointer;
            user-select: none;
        }
        
        .panel-header h3 {
            margin: 0;
            font-size: 16px;
            color: #333;
        }
        
        .toggle-btn {
            background: none;
            border: none;
            font-size: 20px;
            cursor: pointer;
            padding: 0;
            width: 24px;
            height: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: transform 0.3s;
        }
        
        .toggle-btn:hover {
            transform: scale(1.2);
        }
        
        .panel-content {
            padding: 0 15px 15px 15px;
        }
        
        .control-btn {
            display: block;
            width: 140px;
            padding: 10px 15px;
            margin: 5px 0;
            border: none;
            border-radius: 5px;
            font-size: 14px;
            font-weight: bold;
            color: white;
            cursor: pointer;
            transition: all 0.3s;
            font-family: "Microsoft YaHei", Arial, sans-serif;
        }
        
        #save-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        
        #reset-btn {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }
        
        .control-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.3);
        }
        
        .tip {
            margin-top: 10px;
            padding: 8px;
            background: #f0f0f0;
            border-radius: 5px;
            font-size: 12px;
            color: #666;
            line-height: 1.4;
        }
    </style>
    
    <script type="text/javascript">
        function togglePanel() {
            const panel = document.querySelector('.control-panel');
            const toggleBtn = document.querySelector('.toggle-btn');
            panel.classList.toggle('minimized');
            
            if (panel.classList.contains('minimized')) {
                toggleBtn.innerHTML = '▼';
            } else {
                toggleBtn.innerHTML = '▲';
            }
        }
    </script>
    
    <div class="control-panel">
        <div class="panel-header" onclick="togglePanel()">
            <h3>🎨 布局控制</h3>
            <button class="toggle-btn">▲</button>
        </div>
        <div class="panel-content">
            <button id="save-btn" class="control-btn" onclick="savePositions()">💾 保存当前布局</button>
            <button id="reset-btn" class="control-btn" onclick="resetPositions()">🔄 恢复初始布局</button>
            <div class="tip">
                💡 拖动节点后点击"保存"<br>
                下次打开会自动恢复您的布局
            </div>
        </div>
    </div>
'''
    
    # 在</body>之前插入脚本
    html_content = html_content.replace('</body>', save_restore_script + '</body>')
    
    # 写回文件
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✓ 已生成: {filename}")

print(f"\n{'='*60}")
print("所有知识图谱已生成完成！")
print(f"{'='*60}")
print("\n生成的文件:")
for i, name in enumerate(version_names, 1):
    print(f"{i}. 知识图谱_{name}.html")
