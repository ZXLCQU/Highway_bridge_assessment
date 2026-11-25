import sys
import os
import subprocess
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QLabel, QPushButton, QHBoxLayout, QSpacerItem,
                               QSizePolicy, QMessageBox, QGroupBox, QTabWidget, QTextBrowser, QGridLayout)
from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QIcon, QFont

# 作者信息常量
AUTHOR_NAME = "Mingtao Zhang"
AUTHOR_EMAIL = "1045811673@qq.com"
# 桥梁类型分类配置
BRIDGE_CATEGORIES = {
    "梁桥": [
        {"name": "混凝土梁桥", "script": "Concrete_beam_bridge.py"},
        {"name": "钢结构梁桥", "script": "Steel_beam_bridge.py"}, ],
    "拱桥": [
        {"name": "圬工拱桥", "script": "Masonry_arch_bridge.py"},
        {"name": "钢筋混凝土拱桥", "script": "Reinforced_concrete_arch_bridge.py"},
        {"name": "双曲拱桥", "script": "Double_curved_arch_bridge.py"},
        {"name": "刚架拱桥", "script": "Rigid-framed_arch_bridge.py"},
        {"name": "桁架拱桥", "script": "Trussed_arch_bridge.py"},
        {"name": "钢-混组合拱桥", "script": "Steel-concrete_composite_arch_bridge.py"},
    ],
    "缆索承重桥": [
        {"name": "斜拉桥", "script": "Cable-stayed_bridge.py"},
        {"name": "悬索桥", "script": "Suspension_bridge.py"},
    ]
}


class HelpWindow(QMainWindow):
    """帮助文档窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("公路桥梁评估系统-帮助文档")
        self.setMinimumSize(900, 700)
        self.center_window()
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), "Basic_data/Highway_bridges.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.init_ui()

    def center_window(self):
        """将窗口居中显示"""
        screen_geometry = QApplication.primaryScreen().geometry()
        window_geometry = self.geometry()
        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2
        self.move(x, y)

    def init_ui(self):
        """初始化帮助文档界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 设置样式
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
            }
        """)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        # 标题
        title = QLabel("公路桥梁评估系统帮助文档")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(QFont("KaiTi", 18, QFont.Bold))
        title.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                padding: 15px;
                background-color: white;
                border-radius: 10px;
                margin-bottom: 20px;
            }
        """)
        layout.addWidget(title)
        # 创建选项卡
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #ecf0f1;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-family: KaiTi;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 2px solid #3498db;
            }
        """)

        # 添加各个帮助页面
        tab_widget.addTab(self.create_overview_tab(), "系统概述")
        tab_widget.addTab(self.create_usage_tab(), "使用指南")
        tab_widget.addTab(self.create_faq_tab(), "常见问题")
        tab_widget.addTab(self.create_contact_tab(), "联系我们")

        layout.addWidget(tab_widget)

    def create_overview_tab(self):
        """创建系统概述页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">系统概述</h2>
        <p style="font-size: 14px; line-height: 1.6;">公路桥梁评估系统是一款专业的桥梁技术状况评估软件，基于《公路桥梁技术状况评定标准》开发，为桥梁管理提供科学的技术支撑。</p>
        
        <h3 style="color: #34495e;">支持的桥梁类型</h3>
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 10px 0;">
            <h4 style="color: #2980b9;">梁桥类型</h4>
            <ul style="font-size: 14px; line-height: 1.8;">
                <li><strong>混凝土梁桥：</strong>包含桥面系、上部承重构件、上部一般构件、支座等评估模块</li>
                <li><strong>钢结构梁桥：</strong>专门针对钢结构特点的损伤评估体系</li>
            </ul>
            
            <h4 style="color: #2980b9;">拱桥类型</h4>
            <ul style="font-size: 14px; line-height: 1.8;">
                <li><strong>钢筋混凝土拱桥：</strong>钢筋混凝土拱桥技术状况评估</li>
                <li><strong>双曲拱桥：</strong>双曲拱桥专用评估模块</li>
                <li><strong>圬工拱桥：</strong>石拱桥等传统拱桥评估</li>
                <li><strong>桁架拱桥：</strong>桁架拱结构评估</li>
                <li><strong>钢-混组合拱桥：</strong>组合结构拱桥评估</li>
                <li><strong>刚架拱桥：</strong>刚架拱结构评估</li>
            </ul>
            
            <h4 style="color: #2980b9;">缆索桥类型</h4>
            <ul style="font-size: 14px; line-height: 1.8;">
                <li><strong>斜拉桥：</strong>包含斜拉索系统、主梁、索塔等专业评估模块</li>
                <li><strong>悬索桥：</strong>主缆、加劲梁、索塔等构件评估</li>
            </ul>
        </div>
        
        <h3 style="color: #34495e;">主要功能特点</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li><strong>三级评估体系：</strong>桥面系、上部结构、下部结构的全面技术状况评估</li>
            <li><strong>科学计算方法：</strong>采用高等代数方法代替传统计算方式</li>
            <li><strong>项目数据管理：</strong>支持项目创建、保存、加载、历史记录查询等完整的数据管理功能</li>
            <li><strong>专业报告生成：</strong>自动生成符合规范要求的Excel格式技术状况评估报告</li>
            <li><strong>权重自适应：</strong>根据桥梁实际构件情况自动调整权重分配</li>
        </ul>
        
        <h3 style="color: #34495e;">技术特色</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li><strong>标准化评估：</strong>严格按照行业技术标准进行评估计算</li>
            <li><strong>可视化界面：</strong>直观的病害分布界面，便于操作和结果查看</li>
            <li><strong>数据追溯：</strong>完整保存评估过程和计算依据，支持数据追溯</li>
        </ul>
        """)
        layout.addWidget(text_browser)
        return widget

    def create_usage_tab(self):
        """创建使用指南页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">使用指南</h2>
        
        <h3 style="color: #34495e;">1. 系统启动与桥梁类型选择</h3>
        <p style="font-size: 14px; line-height: 1.6;">启动系统后，在主界面选择待评估的桥梁类型。系统提供三大类别：</p>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li><strong>梁桥：</strong>混凝土梁桥、钢结构梁桥</li>
            <li><strong>拱桥：</strong>钢筋混凝土拱桥、双曲拱桥、圬工拱桥、桁架拱桥、钢-混拱桥、刚架拱桥</li>
            <li><strong>缆索桥：</strong>斜拉桥、悬索桥</li>
        </ul>
        
        <h3 style="color: #34495e;">2. 项目管理操作</h3>
        <ol style="font-size: 14px; line-height: 1.8;">
            <li><strong>加载项目：</strong>点击"加载项目"，从历史项目列表中选择已保存的评估项目</li>
            <li><strong>项目搜索：</strong>在项目管理界面可按桥梁名称、日期等条件搜索历史项目</li>
        </ol>
        
        <h3 style="color: #34495e;">3. 数据操作功能</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li><strong>导入数据：</strong>支持从Excel模板文件导入损伤数据</li>
            <li><strong>导出数据：</strong>将当前评估数据导出为Excel文件</li>
            <li><strong>保存项目：</strong>保存当前评估项目，便于后续查看和修改</li>
            <li><strong>生成报告：</strong>自动生成包含完整计算过程的技术状况评估报告</li>
        </ul>
        
        <h3 style="color: #34495e;">4. 操作注意事项</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li>评估前请仔细核对桥梁基本信息和构件配置</li>
            <li>损伤程度选择应基于现场检测结果，确保数据准确性</li>
            <li>对于不存在的构件，务必取消勾选，系统会自动调整权重</li>
            <li>建议定期保存评估数据，避免意外丢失</li>
        </ul>
        """)

        layout.addWidget(text_browser)
        return widget

    def create_faq_tab(self):
        """创建常见问题页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">常见问题</h2>

        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q1: 如何处理桥梁中不存在的构件？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 在评估界面中取消勾选相应构件的复选框，系统会自动调整权重分配，确保总权重为1.0。例如，某桥梁无人行道时，取消勾选人行道复选框，其10%的权重会按比例分配给其他构件。</p>
        </div>

        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q2: 项目数据保存在什么位置？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 项目数据保存在程序目录下的Saved_Projects文件夹中，按桥梁类型分类存储（如Concrete_beam_bridge、Cable-stayed_bridge等）。每个项目以JSON格式保存，包含完整的评估数据和计算结果。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q3: 如何进行数据备份和恢复？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 可以直接复制Saved_Projects文件夹进行整体备份，或者将重要项目导出为Excel文件单独保存。恢复时将备份的Saved_Projects文件夹覆盖即可。导出的Excel文件也可作为数据备份。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q4: 系统支持哪些桥梁损伤类型？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 系统支持混凝土桥梁的蜂窝麻面、剥落掉角、空洞孔洞、钢筋锈蚀、混凝土碳化、裂缝等损伤；钢结构桥梁的锈蚀、涂层劣化、焊缝开裂、构件变形等损伤；以及支座、伸缩缝、排水系统等附属设施的各类损伤。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q5: 如何使用数据导入功能？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 系统在Import_data文件夹中提供了各桥型的Excel模板文件。按照模板格式填写损伤数据后，在评估界面点击"导入数据"按钮选择相应的Excel文件即可批量导入损伤信息。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q6: 权重调整的原理是什么？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 当某个构件不存在时，系统会将该构件的权重按比例分配给同类别的其他构件，确保总权重始终为1.0。例如，桥面系中人行道权重为10%，若不存在人行道，这10%会按比例分配给桥面铺装、伸缩缝等其他构件。</p>
        </div>
        """)

        layout.addWidget(text_browser)
        return widget

    def create_contact_tab(self):
        """创建联系我们页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        text_browser = QTextBrowser()
        text_browser.setHtml(f"""
        <style>
        body {{ font-family: KaiTi, "Times New Roman", serif; }}
        h2, h3, h4 {{ font-family: KaiTi, "Times New Roman", serif; }}
        p, li {{ font-family: KaiTi, "Times New Roman", serif; }}
        </style>

        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #34495e;">开发者信息</h3>
            <p style="font-size: 16px; line-height: 1.8;">
                <strong>作者：</strong>{AUTHOR_NAME}<br>
                <strong>邮箱：</strong>{AUTHOR_EMAIL}<br>
                <strong>版本：</strong>v1.0<br>
                <strong>发布日期：</strong>2025年8月
            </p>
        </div>

        <div style="background-color: #e8f4fd; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #2980b9;">技术支持</h3>
            <p style="font-size: 14px; line-height: 1.6;">
                如果您在使用过程中遇到问题或有改进建议，欢迎通过邮箱联系作者。我们将及时为您提供技术支持和解答。
            </p>
            
            <h4 style="color: #2980b9;">常见技术问题</h4>
            <ul style="font-size: 14px; line-height: 1.6;">
                <li>软件安装和运行环境配置</li>
                <li>评估方法和计算公式解释</li>
                <li>数据导入导出格式问题</li>
                <li>特殊桥型评估方法咨询</li>
                <li>系统功能定制需求</li>
            </ul>
        </div>

        <div style="background-color: #d5f4e6; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #27ae60;">版权声明</h3>
            <p style="font-size: 14px; line-height: 1.6;">
                © 2025 公路桥梁评估系统。本软件基于《公路桥梁技术状况评定标准》开发，
                仅供学习和研究使用。如需商业使用，请联系开发者获得授权。
            </p>
        </div>
        
        <div style="background-color: #fff3cd; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #856404;">更新日志</h3>
            <p style="font-size: 14px; line-height: 1.6;">
                <strong>v1.0 (2025年8月)</strong><br>
                - 支持多种桥梁类型的技术状况评估<br>
                - 实现完整的评估计算体系<br>
                - 提供项目管理和数据导入导出功能<br>
                - 生成标准化的评估报告
            </p>
        </div>
        """)

        layout.addWidget(text_browser)
        return widget


class BridgeSelectionWindow(QMainWindow):
    """桥梁类型选择窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("公路桥梁评估系统")
        self.setMinimumSize(700, 500)
        # 设置窗口居中显示
        self.center_window()
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), "Basic_data/Highway_bridges.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        # 创建主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 设置窗口样式
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                border-radius: 10px;
            }
        """)
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(5)
        # 创建包含标题和帮助按钮的顶部布局
        self.create_header_layout(main_layout)
        # 添加副标题
        subtitle_label = QLabel("请选择要评估的桥梁类型")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_font = QFont("KaiTi")
        subtitle_font.setPointSize(14)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setStyleSheet("color: #34495e; margin-bottom: 15px; letter-spacing: 1px;")
        main_layout.addWidget(subtitle_label)
        # 创建分类按钮布局
        self.create_categorized_buttons(main_layout)
        # 作者信息框与按钮之间的布局
        main_layout.addSpacing(10)  # 作者信息框与按钮之间的布局10
        # 添加作者信息
        self.add_author_info(main_layout)

    def create_header_layout(self, main_layout):
        """创建包含标题的顶部布局"""
        # 添加标题
        title_label = QLabel("公路桥梁评估系统")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont("KaiTi")
        title_font.setPointSize(24)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 2px; letter-spacing: 2px;")
        main_layout.addWidget(title_label)

        # 添加帮助按钮（居中显示）
        help_center_layout = QHBoxLayout()
        help_center_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        help_main_button = QPushButton("📖查看帮助文档")
        help_main_button.setMinimumSize(150, 45)
        help_main_button.setFont(QFont("KaiTi", 12, QFont.Bold))
        help_main_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 22px;
                padding: 10px 20px;
                font-weight: bold;
                letter-spacing: 1px;
            }
            QPushButton:hover {
                background-color: #229954;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
        """)
        help_main_button.clicked.connect(self.show_help)
        help_center_layout.addWidget(help_main_button)
        help_center_layout.addItem(QSpacerItem(20, 10, QSizePolicy.Expanding, QSizePolicy.Minimum))
        main_layout.addLayout(help_center_layout)

    def show_help(self):
        """显示帮助文档窗口"""
        try:
            self.help_window = HelpWindow()
            self.help_window.show()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开帮助文档: {str(e)}")

    def center_window(self):
        """将窗口居中显示在屏幕上"""
        screen_geometry = QApplication.primaryScreen().geometry()
        window_geometry = self.geometry()
        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2
        self.move(x, y)

    def create_categorized_buttons(self, main_layout):
        """创建分类的桥梁类型按钮"""
        # 创建水平布局来放置三个分类组
        categories_layout = QHBoxLayout()
        categories_layout.setSpacing(10)
        main_layout.addLayout(categories_layout)
        # 为每个分类创建组框
        for category_name, bridges in BRIDGE_CATEGORIES.items():
            group_box = self.create_category_group(category_name, bridges)
            categories_layout.addWidget(group_box)

    def create_category_group(self, category_name, bridges):
        """创建桥梁分类组框"""
        group_box = QGroupBox(category_name)
        group_box.setFont(QFont("KaiTi", 12, QFont.Bold))
        group_box.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 8px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 8px 0 8px;
                color: #2c3e50;
                background-color: #f5f5f5;
                border-radius: 4px;
            }
        """)
        # 如果是拱桥类型，使用网格布局（2列）
        if category_name == "拱桥":
            group_layout = QGridLayout(group_box)
            group_layout.setSpacing(8)
            group_layout.setContentsMargins(10, 15, 10, 10)
            # 为每个桥梁类型创建按钮并按2列排列
            for i, bridge in enumerate(bridges):
                button = self.create_bridge_button(bridge["name"], bridge["script"])
                button.setMinimumSize(150, 50)  # 稍微缩小按钮适应网格
                row = i // 2  # 计算行号
                col = i % 2  # 计算列号
                group_layout.addWidget(button, row, col)
        else:
            # 其他分类保持原有的垂直布局
            group_layout = QVBoxLayout(group_box)
            group_layout.setSpacing(10)
            group_layout.setContentsMargins(10, 15, 10, 10)
            # 为每个桥梁类型创建按钮
            for bridge in bridges:
                button = self.create_bridge_button(bridge["name"], bridge["script"])
                group_layout.addWidget(button)
        return group_box

    def create_bridge_button(self, bridge_name, script_name):
        """创建桥梁类型按钮"""
        button = QPushButton(bridge_name)
        button.setMinimumSize(150, 50)
        button.setFont(QFont("KaiTi", 11, QFont.Bold))

        # 使用与作者信息框相同的渐变背景
        button.setStyleSheet(f"""
            QPushButton {{
                /* 渐变背景 */
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #3498db, stop:1 #2c3e50);
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }}
            QPushButton:hover {{
                /* 悬停时加深渐变 */
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #2980b9, stop:1 #232f3e);
                transform: translateY(-2px);
                transition: all 0.3s ease;
            }}
            QPushButton:pressed {{
                /* 按下时进一步加深渐变 */
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #1c6ea4, stop:1 #1a252f);
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }}
        """)
        button.clicked.connect(lambda: self.run_bridge_script(script_name))
        return button

    def run_bridge_script(self, script_name):
        """运行桥梁评估脚本"""
        try:
            script_path = os.path.join(os.path.dirname(__file__), script_name)
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "错误", f"找不到脚本文件: {script_path}")
                return
            # 添加淡出效果
            self.fade_out_animation = QPropertyAnimation(self, b"windowOpacity")
            self.fade_out_animation.setDuration(300)  # 300毫秒的动画时长
            self.fade_out_animation.setStartValue(1.0)
            self.fade_out_animation.setEndValue(0.0)
            self.fade_out_animation.setEasingCurve(QEasingCurve.InOutQuad)  # 添加缓动曲线使动画更自然
            # 当动画结束时再隐藏窗口并启动子进程
            self.fade_out_animation.finished.connect(lambda: self._run_script_after_fade(script_path))
            # 开始动画
            self.fade_out_animation.start()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"运行脚本时出错: {str(e)}")
            self.show()

    def _run_script_after_fade(self, script_path):
        """淡出动画结束后运行脚本"""
        # 隐藏当前窗口
        self.hide()
        # 使用Python解释器运行脚本
        python_exe = sys.executable
        process = subprocess.Popen([python_exe, script_path])

        # 等待脚本执行完毕后再显示主窗口
        def check_process():
            if process.poll() is not None:
                # 脚本结束，准备显示主窗口
                self.prepare_fade_in()
            else:
                # 每100毫秒检查一次进程状态
                QApplication.instance().processEvents()
                self.check_timer.singleShot(100, check_process)

        # 创建单次计时器
        self.check_timer = QTimer()
        self.check_timer.singleShot(100, check_process)

    def prepare_fade_in(self):
        """准备淡入动画"""
        # 设置窗口透明度为0
        self.setWindowOpacity(0.0)
        # 显示窗口但完全透明
        self.show()
        # 创建淡入动画
        self.fade_in_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_in_animation.setDuration(300)  # 300毫秒的动画时长
        self.fade_in_animation.setStartValue(0.0)
        self.fade_in_animation.setEndValue(1.0)
        self.fade_in_animation.setEasingCurve(QEasingCurve.InOutQuad)  # 添加缓动曲线
        # 开始动画
        self.fade_in_animation.start()

    def add_author_info(self, layout):
        """添加作者信息标签"""
        author_widget = QWidget()
        author_layout = QHBoxLayout(author_widget)
        author_layout.setContentsMargins(5, 5, 5, 5)
        # 创建一个带有渐变背景的容器
        info_container = QWidget()
        info_container.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                          stop:0 #3498db, stop:1 #2c3e50);
                border-radius: 15px;
                padding: 5px;
            }
        """)
        info_layout = QHBoxLayout(info_container)
        info_layout.setContentsMargins(10, 5, 10, 5)

        # 创建作者信息标签 - 使用混合字体
        author_label = QLabel()
        author_label.setText(
            f'<span style="font-family: KaiTi;">作者:</span> <span style="font-family: Times New Roman;">{AUTHOR_NAME}</span>')

        email_label = QLabel()
        email_label.setText(
            f'<span style="font-family: KaiTi;">邮箱:</span> <span style="font-family: Times New Roman;">{AUTHOR_EMAIL}</span>')

        # 设置基本样式
        for label in [author_label, email_label]:
            label.setStyleSheet("""
                QLabel {
                    color: white;
                    background-color: transparent;
                    font-size: 10pt;
                    font-weight: bold;
                }
            """)

        # 创建垂直布局放置作者和邮箱信息
        author_info_layout = QVBoxLayout()
        author_info_layout.setSpacing(2)
        author_info_layout.addWidget(author_label)
        author_info_layout.addWidget(email_label)
        info_layout.addLayout(author_info_layout)

        # 添加版权信息
        copyright_label = QLabel("© 2025 公路桥梁评估系统")
        copyright_label.setStyleSheet("""
            QLabel {
                color: #ecf0f1;
                background-color: transparent;
                font-size: 12px;
            }
        """)
        copyright_label.setFont(QFont("KaiTi"))
        info_layout.addWidget(copyright_label)

        # 将信息容器添加到主布局
        author_layout.addStretch()
        author_layout.addWidget(info_container)
        layout.addWidget(author_widget)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BridgeSelectionWindow()
    window.show()
    sys.exit(app.exec())