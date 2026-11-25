import json
import sys
import os
import pandas as pd
import numpy as np
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QWidget, QLabel, QTableWidget, QTableWidgetItem, QLineEdit,
    QHeaderView, QMessageBox, QCheckBox, QPushButton, QDialog,
    QTextEdit, QFormLayout, QComboBox, QTabWidget,
    QSizePolicy, QScrollArea, QInputDialog, QStyledItemDelegate) 
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon, QColor
import datetime
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ScoreTableDelegate(QStyledItemDelegate):
    """评分表格自定义委托，用于绘制背景色"""
    def __init__(self, parent=None):
        super().__init__(parent)
    def paint(self, painter, option, index):
        # 获取背景色数据
        background_color = index.data(Qt.BackgroundRole)
        if background_color and isinstance(background_color, QColor):
            # 绘制背景色
            painter.fillRect(option.rect, background_color)
        # 调用父类方法绘制文本和其他内容
        super().paint(painter, option, index)
    def sizeHint(self, option, index):
        # 返回默认大小
        return super().sizeHint(option, index)

class ProjectManagerWindow(QDialog):
    """项目管理窗口"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_system = parent
        self.projects_data = []
        self.filtered_projects = []
        self.init_ui()
        self.load_projects()
    
    def init_ui(self):
        self.setWindowTitle("项目管理-历史数据")
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        # 设置窗口大小和居中显示
        window_width = 1200  # 增加窗口宽度以容纳右侧面板
        window_height = 700   # 增加窗口高度
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        # 设置窗口图标
        try:
            self.setWindowIcon(QIcon("Basic_data/Highway_bridges.ico"))
        except:
            pass
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                font-family: "楷体", "Times New Roman", serif;
            }
            QLabel {
                font-family: "楷体", "Times New Roman", serif;
                color: #2c3e50;
                font-size: 13px;
            }
            QLineEdit {
                border: 2px solid #bdc3c7;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 13px;
                font-family: "楷体", "Times New Roman", serif;
                background-color: white;
            }
            QLineEdit:focus {
                border-color: #3498db;
                box-shadow: 0 0 5px rgba(52, 152, 219, 0.3);
            }
            QTableWidget {
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: white;
                gridline-color: #ecf0f1;
                font-family: "楷体", "Times New Roman", serif;
                font-size: 12px;
                selection-background-color: #3498db;
            }
            QTableWidget::item {
                padding: 10px 8px;
                border-bottom: 1px solid #ecf0f1;
                text-align: center;
            }
            QTableWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QTableWidget::item:hover {
                background-color: #ebf3fd;
            }
            QHeaderView::section {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                padding: 8px 5px;  /* 减少垂直padding，从15px改为8px */
                border: none;
                font-weight: bold;
                font-family: "楷体";  /* 使用单一字体族 */
                font-size: 12px;  /* 减小字体，从14px改为12px */
                text-align: center;
                min-height: 30px;  /* 确保最小高度 */
                max-height: 35px;  /* 限制最大高度 */
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 16px;
                font-size: 12px;
                font-weight: bold;
                font-family: "楷体", "Times New Roman", serif;
                min-width: 100px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5dade2, stop:1 #3498db);
                transform: translateY(-1px);
                box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2980b9, stop:1 #1f618d);
                transform: translateY(0px);
            }
            QPushButton:disabled {
                background: #bdc3c7;
                color: #7f8c8d;
            }
            /* 评分预览面板样式 */
            .score-preview-panel {
                background: white;
                border: 1px solid #bdc3c7;
                border-radius: 8px;
                padding: 15px;
            }
            .score-table {
                background: #f8f9fa;
                border: 1px solid #27ae60;
                border-radius: 6px;
            }
            .score-table QHeaderView::section {
                background: #27ae60;
                color: white;
                font-weight: bold;
                padding: 8px;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        # 标题
        title_label = QLabel("📂项目管理-历史数据")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 16, QFont.Bold))
        title_label.setStyleSheet("""
            color: #2c3e50;
            padding: 10px;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(52, 152, 219, 0.1), 
                stop:0.5 rgba(52, 152, 219, 0.05),
                stop:1 rgba(52, 152, 219, 0.1));
            border-radius: 10px;
            margin-bottom: 15px;
            font-family: "楷体", "Times New Roman", serif;
            border: 1px solid rgba(52, 152, 219, 0.2);
        """)
        layout.addWidget(title_label)
        # 搜索区域
        search_layout = QHBoxLayout()
        search_layout.setSpacing(5)  # 控制搜索标签和搜索框之间的间距
        search_label = QLabel("🔍搜索项目：")
        search_label.setFont(QFont("楷体", 12, QFont.Bold))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入项目名称进行实时搜索...")
        self.search_input.textChanged.connect(self.filter_projects)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)
        # **新增：主要内容区域 - 水平布局（左侧项目列表 + 右侧评分预览）**
        main_content_layout = QHBoxLayout()
        # **左侧：项目列表区域**
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        # 项目列表表格
        self.projects_table = QTableWidget()
        self.projects_table.setColumnCount(3)
        self.projects_table.setHorizontalHeaderLabels(["项目名称", "保存时间", "文件名"])
        # ===== 添加以下配置 =====
        vertical_header = self.projects_table.verticalHeader()
        vertical_header.setFont(QFont("Times New Roman", 10))       # 设置字体
        vertical_header.setDefaultAlignment(Qt.AlignCenter)        # 设置居中对齐
        vertical_header.setDefaultSectionSize(30)      
        # 设置列宽
        header = self.projects_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 项目名称
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # 保存时间
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # 文件名自适应剩余空间
        # 设置表格属性
        self.projects_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.projects_table.setAlternatingRowColors(True)
        self.projects_table.setSortingEnabled(True)
        # 双击事件
        self.projects_table.itemDoubleClicked.connect(self.load_selected_project)
        # 表格选择事件
        self.projects_table.itemSelectionChanged.connect(self.on_selection_changed)
        left_layout.addWidget(self.projects_table)
        # 状态标签
        self.status_label = QLabel("正在加载项目列表...")
        self.status_label.setStyleSheet("color: #7f8c8d; font-style: italic;")
        left_layout.addWidget(self.status_label)
        # **右侧：评分预览面板**
        right_panel = QWidget()
        right_panel.setFixedWidth(450)  # 固定宽度
        right_panel.setProperty("class", "score-preview-panel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(15, 15, 15, 15)
        # 评分预览标题
        score_title = QLabel("📊全桥技术状况评定")
        score_title.setFont(QFont("楷体", 14, QFont.Bold))
        score_title.setAlignment(Qt.AlignCenter)
        score_title.setStyleSheet("""
            color: #27ae60;
            padding: 10px;
            background: #e8f5e8;
            border-radius: 6px;
            border: 1px solid #27ae60;
        """)
        right_layout.addWidget(score_title)
        
        # 评分表格
        self.score_table = QTableWidget()
        self.score_table.setColumnCount(3)
        self.score_table.setHorizontalHeaderLabels(["结构部位", "技术状况评分", "技术状况等级"])
        self.score_table.setProperty("class", "score-table")
        
        # **新增：设置自定义委托**
        self.score_delegate = ScoreTableDelegate(self.score_table)
        self.score_table.setItemDelegate(self.score_delegate)
        
        # 设置评分表格样式（移除可能冲突的背景色设置）
        self.score_table.setStyleSheet("""
            QTableWidget {
                background: #f8f9fa;
                border: 1px solid #27ae60;
                border-radius: 6px;
                gridline-color: #d5e8d4;
            }
            QHeaderView::section {
                background: #27ae60;
                color: white;
                font-weight: bold;
                padding: 8px;
                border: none;
            }
            QTableWidget::item {
                padding: 8px;
                text-align: center;
                /* 移除背景色相关设置，让委托处理 */
            }
        """)
        # **修改：调整评分表格列宽设置**
        score_header = self.score_table.horizontalHeader()
        score_header.setSectionResizeMode(0, QHeaderView.Fixed)  # 结构部位固定宽度
        score_header.setSectionResizeMode(1, QHeaderView.Fixed)  # 评分固定宽度
        score_header.setSectionResizeMode(2, QHeaderView.Stretch)  # 等级自适应
        # 设置具体列宽
        self.score_table.setColumnWidth(0, 120)  # 结构部位
        self.score_table.setColumnWidth(1, 120)  # 技术状况评分
        
        # 禁用编辑和选择
        self.score_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.score_table.setSelectionMode(QTableWidget.NoSelection)
        
        right_layout.addWidget(self.score_table)
        
        # 初始化空的评分表格
        self.init_empty_score_table()
        
        # 添加左右面板到主布局
        main_content_layout.addWidget(left_panel, 2)  # 左侧占2/3
        main_content_layout.addWidget(right_panel, 1)  # 右侧占1/3
        layout.addLayout(main_content_layout)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        self.load_button = QPushButton("📥加载选中项目")
        self.load_button.clicked.connect(self.load_selected_project)
        self.load_button.setEnabled(False)
        
        self.refresh_button = QPushButton("🔄刷新列表")
        self.refresh_button.clicked.connect(self.load_projects)
        
        self.delete_button = QPushButton("🗑️删除项目")
        self.delete_button.clicked.connect(self.delete_selected_project)
        self.delete_button.setEnabled(False)
        self.delete_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #e74c3c, stop:1 #c0392b);
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ec7063, stop:1 #e74c3c);
            }
        """)
        close_button = QPushButton("❌关闭")
        close_button.clicked.connect(self.close)
        button_layout.addWidget(self.load_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch()
        button_layout.addWidget(self.refresh_button)
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)
    
    def load_projects(self):
        """加载所有保存的项目"""
        try:
            self.projects_data = []
            save_dir = "Saved_Projects/Steel-concrete_composite_arch_bridge"
            if not os.path.exists(save_dir):
                self.status_label.setText("📁保存目录不存在，暂无项目数据")
                self.update_table()
                return
            # 扫描保存目录中的所有JSON文件
            json_files = [f for f in os.listdir(save_dir) if f.endswith('.json')]
            if not json_files:
                self.status_label.setText("📄暂无保存的项目")
                self.update_table()
                return
            # 读取每个项目文件的信息
            for filename in json_files:
                filepath = os.path.join(save_dir, filename)
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)
                    project_info = project_data.get('project_info', {})
                    project_name = project_info.get('name', '未知项目')
                    save_time = project_info.get('save_time', '未知时间')
                    # 格式化保存时间
                    try:
                        if save_time != '未知时间':
                            dt = datetime.datetime.fromisoformat(save_time)
                            formatted_time = dt.strftime('%Y-%m-%d %H:%M:%S')
                        else:
                            formatted_time = save_time
                    except:
                        formatted_time = save_time
                    
                    self.projects_data.append({
                        'name': project_name,
                        'save_time': formatted_time,
                        'filename': filename,
                        'filepath': filepath,
                        'data': project_data
                    })
                except Exception as e:
                    QMessageBox.warning(self, "读取文件错误", f"读取项目文件{filename}时发生错误：\n{str(e)}")
                    continue
            # 按保存时间排序（最新的在前）
            self.projects_data.sort(key=lambda x: x['save_time'], reverse=True)
            self.filtered_projects = self.projects_data.copy()
            self.update_table()
            self.status_label.setText(f"📊共找到{len(self.projects_data)}个项目")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载项目列表时发生错误：\n{str(e)}")
            self.status_label.setText("❌加载项目列表失败")
    
    def filter_projects(self):
        """根据搜索关键词过滤项目"""
        search_text = self.search_input.text().strip().lower()
        if not search_text:
            self.filtered_projects = self.projects_data.copy()
        else:
            self.filtered_projects = [
                project for project in self.projects_data
                if search_text in project['name'].lower()
            ]
        self.update_table()
        if search_text:
            self.status_label.setText(f"🔍搜索到{len(self.filtered_projects)}个匹配项目")
        else:
            self.status_label.setText(f"📊共找到{len(self.projects_data)}个项目")
    
    def update_table(self):
        """更新项目列表表格"""
        self.projects_table.setRowCount(len(self.filtered_projects))
        for row, project in enumerate(self.filtered_projects):
            # 项目名称 - 去掉加粗，统一字体大小
            name_item = QTableWidgetItem(project['name'])
            name_item.setFont(QFont("楷体", 12))  # 去掉Bold
            name_item.setTextAlignment(Qt.AlignCenter)
            self.projects_table.setItem(row, 0, name_item)
            # 保存时间 - 统一字体大小
            time_item = QTableWidgetItem(project['save_time'])
            time_item.setFont(QFont("Times New Roman", 12))  # 统一为12号字体
            time_item.setTextAlignment(Qt.AlignCenter)
            self.projects_table.setItem(row, 1, time_item)
            # 文件名 - 统一字体大小
            filename_item = QTableWidgetItem(project['filename'])
            filename_item.setFont(QFont("Times New Roman", 12))  # 统一为12号字体
            filename_item.setTextAlignment(Qt.AlignCenter)
            self.projects_table.setItem(row, 2, filename_item)
        # 调整行高，让表格更美观
        for row in range(len(self.filtered_projects)):
            self.projects_table.setRowHeight(row, 40)  # 增加行高
    
    def init_empty_score_table(self):
        """初始化空的评分表格"""
        self.score_table.setRowCount(4)
        # 设置行标题和默认数据
        structures = ["桥面系", "上部结构", "下部结构", "全桥总评分"]
        for row, structure in enumerate(structures):
            # 结构部位
            structure_item = QTableWidgetItem(structure)
            structure_item.setFont(QFont("楷体", 12, QFont.Bold))
            structure_item.setTextAlignment(Qt.AlignCenter)
            if structure == "全桥总评分":
                structure_item.setData(Qt.BackgroundRole, QColor("#e8f5e8"))
            self.score_table.setItem(row, 0, structure_item)
            # 技术状况评分
            score_item = QTableWidgetItem("--")
            score_item.setFont(QFont("Times New Roman", 12))
            score_item.setTextAlignment(Qt.AlignCenter)
            if structure == "全桥总评分":
                score_item.setData(Qt.BackgroundRole, QColor("#e8f5e8"))
            self.score_table.setItem(row, 1, score_item)
            # 技术状况等级
            grade_item = QTableWidgetItem("--")
            grade_item.setFont(QFont("Times New Roman", 12))
            grade_item.setTextAlignment(Qt.AlignCenter)
            if structure == "全桥总评分":
                grade_item.setData(Qt.BackgroundRole, QColor("#e8f5e8"))
            self.score_table.setItem(row, 2, grade_item)
        # 设置行高
        for row in range(4):
            self.score_table.setRowHeight(row, 35)

    def update_score_preview(self, project_data):
        """更新评分预览面板"""
        try:
            # 获取评分数据
            evaluation_results = project_data.get('evaluation_results', {})
            bridge_structure_scores = evaluation_results.get('bridge_structure_scores', {})
            bridge_structure_grades = evaluation_results.get('bridge_structure_grades', {})  # 添加这行
            total_bridge_score = evaluation_results.get('total_bridge_score', 0.00)
            total_bridge_grade = evaluation_results.get('total_bridge_grade', '未评定')
            # 更新结构系统评分
            structures = ["桥面系", "上部结构", "下部结构"]
            for row, structure in enumerate(structures):
                if structure in bridge_structure_scores:
                    score = bridge_structure_scores[structure]
                    grade = bridge_structure_grades.get(structure, "--")  # 直接从保存的数据获取评级
                    # 更新评分
                    score_item = QTableWidgetItem(f"{score:.2f}")
                    score_item.setFont(QFont("Times New Roman", 12, QFont.Bold))
                    score_item.setTextAlignment(Qt.AlignCenter)
                    score_item.setData(Qt.BackgroundRole, self.get_score_color(score))
                    self.score_table.setItem(row, 1, score_item)
                    # 更新等级
                    grade_item = QTableWidgetItem(grade)
                    grade_item.setFont(QFont("Times New Roman", 12, QFont.Bold))
                    grade_item.setTextAlignment(Qt.AlignCenter)
                    grade_item.setData(Qt.BackgroundRole, self.get_grade_color(grade))
                    self.score_table.setItem(row, 2, grade_item)
                else:
                    # 无数据
                    score_item = QTableWidgetItem("--")
                    score_item.setTextAlignment(Qt.AlignCenter)
                    self.score_table.setItem(row, 1, score_item)
                    grade_item = QTableWidgetItem("--")
                    grade_item.setTextAlignment(Qt.AlignCenter)
                    self.score_table.setItem(row, 2, grade_item)
            # 更新全桥总评分
            if total_bridge_score > 0:
                total_score_item = QTableWidgetItem(f"{total_bridge_score:.2f}")
                total_score_item.setFont(QFont("Times New Roman", 12, QFont.Bold))
                total_score_item.setTextAlignment(Qt.AlignCenter)
                total_score_item.setData(Qt.BackgroundRole, self.get_score_color(total_bridge_score))
                self.score_table.setItem(3, 1, total_score_item)
                
                total_grade_item = QTableWidgetItem(total_bridge_grade)
                total_grade_item.setFont(QFont("Times New Roman", 12, QFont.Bold))
                total_grade_item.setTextAlignment(Qt.AlignCenter)
                total_grade_item.setData(Qt.BackgroundRole, self.get_grade_color(total_bridge_grade))
                self.score_table.setItem(3, 2, total_grade_item)
            else:
                # 无数据
                total_score_item = QTableWidgetItem("--")
                total_score_item.setTextAlignment(Qt.AlignCenter)
                total_score_item.setData(Qt.BackgroundRole, QColor("#e8f5e8"))
                self.score_table.setItem(3, 1, total_score_item)
                
                total_grade_item = QTableWidgetItem("--")
                total_grade_item.setTextAlignment(Qt.AlignCenter)
                total_grade_item.setData(Qt.BackgroundRole, QColor("#e8f5e8"))
                self.score_table.setItem(3, 2, total_grade_item)
        except Exception as e:
            QMessageBox.warning(self, "警告", f"更新评分预览时出错:{str(e)}")
            self.init_empty_score_table()

    def get_score_color(self, score):
        """根据评分获取背景颜色"""
        if 95 <= score <= 100:
            return QColor("#d4edda")  # 绿色 - 1类
        elif 80 <= score < 95:
            return QColor("#d1ecf1")  # 蓝色 - 2类
        elif 60 <= score < 80:
            return QColor("#fff3cd")  # 黄色 - 3类
        elif 40 <= score < 60:
            return QColor("#f8d7da")  # 红色 - 4类
        elif 0 <= score < 40:
            return QColor("#f5c6cb")  # 深红色 - 5类
        else:
            return QColor("#f8f9fa")  # 默认灰色

    def get_grade_color(self, grade):
        """根据等级获取背景颜色"""
        color_map = {
            "1类": QColor("#d4edda"),  # 绿色
            "2类": QColor("#d1ecf1"),  # 蓝色
            "3类": QColor("#fff3cd"),  # 黄色
            "4类": QColor("#f8d7da"),  # 红色
            "5类": QColor("#f5c6cb"),  # 深红色
            "未评定": QColor("#f8f9fa"),  # 灰色
            "未知": QColor("#f8f9fa")   # 灰色
        }
        return color_map.get(grade, QColor("#f8f9fa"))

    def on_selection_changed(self):
        """表格选择变化事件"""
        selected_rows = self.projects_table.selectionModel().selectedRows()
        has_selection = len(selected_rows) > 0
        self.load_button.setEnabled(has_selection)
        self.delete_button.setEnabled(has_selection)
        # **新增：更新评分预览**
        if has_selection:
            project = self.get_selected_project()
            if project:
                self.update_score_preview(project['data'])
        else:
            self.init_empty_score_table()
    
    def get_selected_project(self):
        """获取选中的项目数据"""
        selected_rows = self.projects_table.selectionModel().selectedRows()
        if not selected_rows:
            return None
        row = selected_rows[0].row()
        if 0 <= row < len(self.filtered_projects):
            return self.filtered_projects[row]
        return None
    def load_selected_project(self):
        """加载选中的项目"""
        project = self.get_selected_project()
        if not project:
            QMessageBox.warning(self, "警告", "请先选择一个项目")
            return
        try:
            # 直接加载项目数据到主界面
            success, load_info = self.load_project_to_main_window(project['data'])
            if success:
                # **新增：显示加载成功提示框**
                project_name = project['data'].get('project_info', {}).get('name', '未知项目')
                success_message = f"加载完成！\n" + \
                                f"桥梁名：{project_name}\n" + \
                                f"成功加载{load_info['total_components']}个部件，{load_info['total_damage_records']}条病害记录\n" + \
                                f"数据自动重新分配：\n" + \
                                f"下部结构{load_info['structure_details']['下部结构']}\n" + \
                                f"上部结构{load_info['structure_details']['上部结构']}\n" + \
                                f"桥面系{load_info['structure_details']['桥面系']}\n" + \
                                f"注意事项（{load_info['notes_count']}项）：\n" + \
                                f"【桥面系】病害数据分布：{load_info['bridge_deck_details']}\n" + \
                                f"【上部结构】病害数据分布：{load_info['superstructure_details']}\n" + \
                                f"【下部结构】病害数据分布：{load_info['substructure_details']}"
                QMessageBox.information(self, "加载完成", success_message)
                # **新增：设置当前项目名称**
                self.parent_system.current_project_name = project_name
                # 关闭项目管理窗口
                self.close()
                # 自动开始矩阵运算，传递auto_confirm=True参数
                if hasattr(self.parent_system, 'start_matrix_calculation'):
                    self.parent_system.start_matrix_calculation(auto_confirm=True)
        except Exception as e:
            QMessageBox.critical(self, "加载失败", f"加载项目时发生错误：\n{str(e)}")
    
    def load_project_to_main_window(self, project_data):
        """将项目数据加载到主窗口"""
        try:
            if not self.parent_system:
                return False, {}
            # 首先加载桥梁基础数据（从配置文件）
            self.parent_system.load_bridge_data()
            # 恢复部件病害数据
            component_damage_data = project_data.get('component_damage_data', {})
            self.parent_system.component_damage_data = component_damage_data
            # 恢复部件构件数量
            component_quantities = project_data.get('component_quantities', {})
            self.parent_system.component_quantities = component_quantities
            # 恢复原始权重
            original_weights = project_data.get('original_weights', {})
            self.parent_system.original_weights = original_weights
            # 恢复部件存在性状态
            component_existence = project_data.get('component_existence', {})
            # **新增：统计加载信息**
            total_components = sum(len(comps) for comps in component_existence.values())
            total_damage_records = sum(len(damages) for damages in component_damage_data.values())
            structure_details = {}
            bridge_deck_count = 0
            superstructure_count = 0
            substructure_count = 0
            
            for structure_name in ['桥面系', '上部结构', '下部结构']:
                count = len(component_existence.get(structure_name, {}))
                structure_details[structure_name] = f"{count}个部件"
                
                if structure_name == '桥面系':
                    bridge_deck_count = sum(len(damages) for comp_name, damages in component_damage_data.items() 
                                        if comp_name.startswith('桥面系-'))
                elif structure_name == '上部结构':
                    superstructure_count = sum(len(damages) for comp_name, damages in component_damage_data.items() 
                                            if comp_name.startswith('上部结构-'))
                elif structure_name == '下部结构':
                    substructure_count = sum(len(damages) for comp_name, damages in component_damage_data.items() 
                                        if comp_name.startswith('下部结构-'))
            
            # 更新UI界面的复选框状态和数量输入框
            for structure_name, components in component_existence.items():
                if structure_name in self.parent_system.structure_tables:
                    table = self.parent_system.structure_tables[structure_name]
                    for row in range(table.rowCount()):
                        # 获取部件名称
                        checkbox_widget = table.cellWidget(row, 0)
                        if checkbox_widget:
                            checkbox = None
                            name_label = None
                            for child in checkbox_widget.findChildren(QCheckBox):
                                checkbox = child
                                break
                            for child in checkbox_widget.findChildren(QLabel):
                                name_label = child
                                break
                            if checkbox and name_label:
                                component_name = name_label.text()
                                # 恢复复选框状态
                                if component_name in components:
                                    checkbox.setChecked(components[component_name])
                                # **修复构件数量键名格式**：使用横线而不是下划线
                                quantity_key = f"{structure_name}-{component_name}"
                                if quantity_key in component_quantities:
                                    quantity_widget = table.cellWidget(row, 2)
                                    if quantity_widget and hasattr(quantity_widget, 'setText'):
                                        quantity_widget.setText(str(component_quantities[quantity_key]))
            
            # **关键修复**：初始化current_weights
            # 为每个结构系统调用update_structure_weights来初始化current_weights
            for structure_name in ['桥面系', '上部结构', '下部结构']:
                if structure_name in self.parent_system.structure_tables:
                    self.parent_system.update_structure_weights(structure_name)
            
            # **新增：准备返回的加载信息**
            load_info = {
                'total_components': total_components,
                'total_damage_records': total_damage_records,
                'structure_details': structure_details,
                'notes_count': 3,
                'bridge_deck_details': f"{bridge_deck_count}条",
                'superstructure_details': f"{superstructure_count}条",
                'substructure_details': f"{substructure_count}条"
            }
            
            print(f"项目数据加载成功：")
            print(f"-部件病害数据：{len(component_damage_data)} 个部件")
            print(f"-部件构件数量：{len(component_quantities)} 个部件")
            print(f"-部件存在性状态：{sum(len(comps) for comps in component_existence.values())} 个部件")
            return True, load_info
        except Exception as e:
            QMessageBox.critical(self, "加载失败", f"加载项目数据到主窗口时发生错误：\n{str(e)}")
            return False, load_info
    
    def delete_selected_project(self):
        """删除选中的项目"""
        project = self.get_selected_project()
        if not project:
            QMessageBox.warning(self, "警告", "请先选择一个项目")
            return
        # 确认删除
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除项目'{project['name']}'吗？\n"
            f"文件：{project['filename']}\n"
            f"保存时间：{project['save_time']}\n",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                os.remove(project['filepath'])
                QMessageBox.information(self, "删除成功", f"项目'{project['name']}'已删除")
                self.load_projects()  # 重新加载项目列表
            except Exception as e:
                QMessageBox.critical(self, "删除失败", f"删除项目时发生错误：\n{str(e)}")
class CollapsibleGroupBox(QWidget):
    """可折叠的组件框"""
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.toggle_button = QPushButton(title)
        self.toggle_button.setCheckable(True)
        self.toggle_button.setChecked(False)
        self.toggle_button.setStyleSheet("""
            QPushButton {
                text-align: left;
                border: 2px solid #3498db;
                padding: 12px 16px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ecf0f1, stop:1 #bdc3c7);
                font-weight: bold;
                font-size: 14px;
                color: #2c3e50;
                border-radius: 8px;
                margin: 4px 0px;
                min-height: 20px;
            }
            QPushButton:checked {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: 2px solid #2980b9;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5dade2, stop:1 #3498db);
                color: white;
                border: 2px solid #2980b9;
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2980b9, stop:1 #1f618d);
            }
        """)
        self.content_area = QWidget()
        self.content_area.setVisible(False)
        self.content_area.setStyleSheet("""
            QWidget {
                border: 1px solid #bdc3c7;
                border-top: none;
                border-radius: 0px 0px 8px 8px;
                background-color: #f8f9fa;
                margin: 0px;
                padding: 8px;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 4)
        layout.setSpacing(0)
        layout.addWidget(self.toggle_button)
        layout.addWidget(self.content_area)
        self.toggle_button.clicked.connect(self.toggle_content)
        # 设置初始文本
        self.update_button_text()
    def toggle_content(self):
        self.content_area.setVisible(self.toggle_button.isChecked())
        self.update_button_text()
    def update_button_text(self):
        """更新按钮文本显示展开/折叠状态"""
        current_text = self.toggle_button.text()
        clean_text = current_text.replace("▶ ", "").replace("▼ ", "")
        if self.toggle_button.isChecked():
            new_text = f"▼ {clean_text}"
        else:
            new_text = f"▶ {clean_text}"
        self.toggle_button.setText(new_text)
    def set_content_layout(self, layout):
        self.content_area.setLayout(layout)
class CalculationProcessWindow(QDialog):
    """矩阵运算过程展示窗口"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_system = parent
        self.calculation_data = getattr(parent, 'calculation_process_data', {})
        self.init_ui()
    def init_ui(self):
        self.setWindowTitle("桥梁技术状况矩阵运算过程详览")
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        # 设置窗口大小和居中显示
        window_width = 800
        window_height = 800
        # 获取屏幕尺寸并计算居中位置
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        # 设置窗口图标
        try:
            self.setWindowIcon(QIcon("Basic_data/Highway_bridges.ico"))
        except:
            pass
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                font-family: "楷体", "Times New Roman", serif;
            }
            QTabWidget::pane {
                border: 1px solid #bdc3c7;
                background-color: white;
                border-radius: 8px;
            }
            QTabBar::tab {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ecf0f1, stop:1 #bdc3c7);
                border: 1px solid #95a5a6;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-family: "楷体", "Times New Roman", serif;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                font-weight: bold;
            }
            QTextEdit {
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: white;
                font-family: "楷体", "Times New Roman", serif;
                font-size: 14px;
                line-height: 1.4;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
                font-family: "楷体", "Times New Roman", serif;
                min-width: 100px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5dade2, stop:1 #3498db);
            }
            QLabel {
                font-family: "楷体", "Times New Roman", serif;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        # 标题
        title_label = QLabel("🧮桥梁技术状况矩阵运算过程详览")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 18, QFont.Bold))
        title_label.setStyleSheet("""
            color: #2c3e50;
            padding: 15px;
            background-color: rgba(52, 152, 219, 0.1);
            border-radius: 8px;
            margin-bottom: 10px;
            font-family: "楷体", "Times New Roman", serif;
            font-size: 18px;
        """)
        layout.addWidget(title_label)
        # 创建选项卡控件
        self.tab_widget = QTabWidget()
        # 创建各个选项卡
        self.create_overview_tab()
        self.create_structure_tabs()
        self.create_final_summary_tab()
        layout.addWidget(self.tab_widget)
        # 只保留关闭按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        layout.addLayout(button_layout)
    def create_overview_tab(self):
        """创建总览选项卡"""
        overview_widget = QWidget()
        layout = QVBoxLayout(overview_widget)
        # 基本信息区域
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        basic_info = self.calculation_data.get('basic_info', {})
        overview_content = f"""
🔍运算基本信息
{'=' * 60}
矩阵维度: {basic_info.get('matrix_dimensions', '未知')}
运算开始时间: {basic_info.get('start_time', '未知')}
运算结束时间: {basic_info.get('end_time', '未知')}
处理的结构部位数量: {len(self.calculation_data.get('structure_systems', {}))}
处理的部件总数: {basic_info.get('total_components', 0)}
成功计算的部件数: {basic_info.get('calculated_components', 0)}
跳过的部件数(不存在的部件数量): {basic_info.get('skipped_components', 0)}
📋矩阵运算流程详解
{'=' * 60}
1.🔢初始化矩阵A (全1矩阵)
   • 作用:为所有构件设置初始满分状态
   • 维度:根据构件数量和病害类型确定
   • 初始值:所有元素均为1
2.📝创建矩阵B(病害数据填充矩阵)
   • 作用:根据实际病害数据填充评定标度
   • 数据来源:Excel病害数据表
   • 填充规则:按病害等级填入对应的评定标度值
3.⚖️创建矩阵C(扣分值矩阵)
   • 作用:将评定标度转换为具体扣分值
   • 计算方法:基于病害等级和严重程度
   • 扣分原则:参考公路桥梁技术状况评定标准
4.📈创建矩阵D(排序扣分值矩阵)
   • 作用:按扣分值大小排序，确保严重病害优先计算
   • 排序原则:从大到小排列
   • 目的:保证计算的合理性和准确性
5.🧮创建矩阵E(计算矩阵)
   • 作用:对多个病害进行加权平均计算
   • 处理方式:单个病害直接计算，多个病害加权处理
   • 算法:考虑病害数量和严重程度的综合影响
6.✅创建矩阵F(最终得分矩阵)
   • 作用:计算每个构件的最终技术状况得分
   • 计算公式:100-总扣分值
   • 结果范围:0-100分
7.📊计算部件技术状况评分
   • 作用:综合所有构件得分，计算部件整体评分
   • 权重分配:基于部件存在性重新分配权重
8.🏗️结构部位评分
   • 作用:将各部件评分汇总为结构部位评分
   • 权重分配:基于部件重要性和规范要求
   • 计算方法:加权平均
9.🌉全桥技术状况评分
   • 作用:综合各结构部位评分得出全桥评分
   • 最终结果:反映桥梁整体技术状况
   • 评级标准:按照规范要求划分技术状况等级
💡计算原理说明
{'=' * 60}
• 🎯评分原则:基于《公路桥梁技术状况评定标准》(JTG/T H21-2011)
• 📐矩阵算法:采用矩阵运算确保计算的系统性和准确性
• ⚡病害排序:严重病害优先计算，确保评分的合理性
• 🔄权重分配:根据部件存在性动态调整权重
• 📊评分范围:0-100分制，对应5个技术状况等级
• 🛡️容错机制:对缺失数据和异常情况进行合理处理
• 📈统计分析:提供详细的计算过程和统计信息
        """
        info_text.setPlainText(overview_content)
        layout.addWidget(info_text)
        self.tab_widget.addTab(overview_widget, "📊运算总览")
    def create_structure_tabs(self):
        """创建各结构系统的详细过程选项卡"""
        structure_systems = self.calculation_data.get('structure_systems', {})
        component_details = self.calculation_data.get('component_details', {})
        for structure_name, structure_data in structure_systems.items():
            structure_widget = QWidget()
            layout = QVBoxLayout(structure_widget)
            # 创建滚动区域
            scroll_area = QScrollArea()
            scroll_content = QWidget()
            scroll_layout = QVBoxLayout(scroll_content)
            # 结构部位信息
            info_text = QTextEdit()
            info_text.setReadOnly(True)
            info_text.setMaximumHeight(200)  # 增加高度以显示更多信息
            # 分析部件状态，获取具体的部件名称列表
            component_keys = structure_data.get('components', [])
            calculated_components = []
            skipped_components = []
            all_components = []  # 记录所有部件
            for component_key in component_keys:
                component_data = component_details.get(component_key, {})
                component_name = component_data.get('component_name', component_key)
                all_components.append(component_name)
                # 修改这里：不跳过没有数据的部件，而是将其归类为跳过
                if not component_data:
                    skipped_components.append(component_name)
                    continue
                calculation_status = component_data.get('calculation_status', '')
                h_value = component_data.get('h_value')
                # 更严格的判断逻辑
                if (calculation_status == '计算成功' or 
                    (h_value is not None and isinstance(h_value, (int, float)))):
                    calculated_components.append(component_name)
                else:
                    skipped_components.append(component_name)
            # 直接使用原始统计数据，不再重新计算
            original_total = structure_data.get('total_components', 0)
            original_calculated = structure_data.get('calculated_components', 0)
            original_skipped = structure_data.get('skipped_components', 0)
            # 构建详细的结构部位信息
            structure_info = f"""🏗️{structure_name}运算过程
{'=' * 60}
部件总数: {original_total}
成功计算: {original_calculated}个
跳过部件数量: {original_skipped}个
📋成功计算的部件:\n"""
            if calculated_components:
                for i, component_name in enumerate(calculated_components, 1):
                    structure_info += f"{i}.{component_name}\n"
            else:
                structure_info += "无\n"
            info_text.setPlainText(structure_info)
            scroll_layout.addWidget(info_text)
            # 分类显示部件：有病害的和无病害的
            component_keys = structure_data.get('components', [])
            # 分离有病害和无病害的部件
            damaged_components = []
            undamaged_components = []
            for component_key in component_keys:
                component_data = component_details.get(component_key, {})
                if not component_data:
                    continue
                if component_data.get('has_damage_data', False):
                    damaged_components.append((component_key, component_data))
                else:
                    undamaged_components.append((component_key, component_data))
            # 创建有病害部件的折叠组
            if damaged_components:
                damaged_group = CollapsibleGroupBox(f"▶ 有病害部件 ({len(damaged_components)}个)")
                damaged_layout = QVBoxLayout()
                for component_key, component_data in damaged_components:
                    component_text = QTextEdit()
                    component_text.setReadOnly(True)
                    component_text.setMaximumHeight(400)  # 限制高度
                    process_content = self.build_component_process_content(component_key, component_data)
                    component_text.setPlainText(process_content)
                    damaged_layout.addWidget(component_text)
                damaged_group.set_content_layout(damaged_layout)
                scroll_layout.addWidget(damaged_group)
            # 创建无病害部件的折叠组
            if undamaged_components:
                undamaged_group = CollapsibleGroupBox(f"▶ 无病害部件 ({len(undamaged_components)}个)")
                undamaged_layout = QVBoxLayout()
                for component_key, component_data in undamaged_components:
                    component_text = QTextEdit()
                    component_text.setReadOnly(True)
                    component_text.setMaximumHeight(200)  # 无病害部件显示更简洁
                    process_content = self.build_component_process_content(component_key, component_data)
                    component_text.setPlainText(process_content)
                    undamaged_layout.addWidget(component_text)
                undamaged_group.set_content_layout(undamaged_layout)
                scroll_layout.addWidget(undamaged_group)
            scroll_area.setWidget(scroll_content)
            scroll_area.setWidgetResizable(True)
            layout.addWidget(scroll_area)
            self.tab_widget.addTab(structure_widget, f"🏗️{structure_name}")
    def build_component_process_content(self, component_key, component_data):
        """构建部件详细运算过程内容"""
        component_name = component_data.get('component_name', component_key)
        content = f"""
🔧部件:{component_name}
{'=' * 60}
构件数量:{component_data.get('component_quantity', 'N/A')}
计算状态:{component_data.get('calculation_status', 'N/A')}
        """
        # 检查是否有病害数据
        if component_data.get('has_damage_data', False):
            content += "\n📊矩阵运算详细过程:\n"
            # 矩阵运算过程
            matrix_operations = component_data.get('matrix_operations', {})
            # 矩阵B信息 - 重点显示填充过程
            if 'matrix_B' in matrix_operations:
                matrix_b = matrix_operations['matrix_B']
                content += f"创建矩阵B-{component_name}，维度:{matrix_b.get('dimensions', 'N/A')}\n"
                content += f"开始根据部件{component_name}的病害数据填充矩阵B...\n"
                # 显示填充详情
                fill_details = matrix_b.get('fill_details', [])
                if fill_details:
                    for detail in fill_details:
                        content += f"{detail}\n"
                filled_cells = matrix_b.get('filled_cells', 0)
                processed_components = matrix_b.get('processed_components', 0)
                content += f"部件{component_name}共填充了{filled_cells}个单元格，处理了{processed_components}个构件\n\n"
            # 矩阵C信息 - 显示创建过程和填充详情
            if 'matrix_C' in matrix_operations:
                matrix_c = matrix_operations['matrix_C']
                content += f"创建矩阵C-{component_name}-（扣分值矩阵），维度: {matrix_c.get('dimensions', 'N/A')}\n"
                content += f"基于矩阵B，开始根据部件 {component_name} 的标度值填充扣分值...\n"
                # 显示填充详情
                fill_details = matrix_c.get('fill_details', [])
                if fill_details:
                    for detail in fill_details:
                        content += f"{detail}\n"
                processed_cells = matrix_c.get('processed_cells', 0)
                content += f"部件{component_name}共处理了{processed_cells}个单元格\n\n"
            # 矩阵D信息 - 显示排序过程
            if 'matrix_D' in matrix_operations:
                matrix_d = matrix_operations['matrix_D']
                content += f"创建了矩阵D（排序扣分值矩阵），维度: {matrix_d.get('dimensions', 'N/A')}\n"
                content += f"基于矩阵C，开始对每行扣分值进行从大到小排序...\n"
                # 显示排序详情
                sort_details = matrix_d.get('sort_details', [])
                if sort_details:
                    for detail in sort_details:
                        content += f"{detail}\n"
                sorted_rows = matrix_d.get('sorted_rows', 0)
                total_non_zero_values = matrix_d.get('total_non_zero_values', 0)
                content += f"排序统计信息：\n"
                content += f"-共处理了{sorted_rows}行数据\n"
                content += f"-总计非零扣分值:{total_non_zero_values}个\n"
                content += f"-每行扣分值已按从大到小顺序排列\n\n"
            # 矩阵E信息 - 显示计算过程
            if 'matrix_E' in matrix_operations:
                matrix_e = matrix_operations['matrix_E']
                content += f"创建了矩阵E（计算矩阵），维度: {matrix_e.get('dimensions', 'N/A')}\n"
                content += f"基于矩阵D，开始通过公式计算矩阵E...\n"
                # 显示计算详情
                calculation_details = matrix_e.get('calculation_details', [])
                if calculation_details:
                    for detail in calculation_details:
                        content += f"{detail}\n"
                # 显示格式化过程
                content += f"\n开始对矩阵E进行数值格式化处理...\n"
                format_details = matrix_e.get('format_details', [])
                if format_details:
                    for detail in format_details:
                        content += f"{detail}\n"
                # 显示统计信息
                processed_rows = matrix_e.get('processed_rows', 0)
                single_value_rows = matrix_e.get('single_value_rows', 0)
                multi_value_rows = matrix_e.get('multi_value_rows', 0)
                formatted_count = matrix_e.get('formatted_count', 0)
                content += f"\n计算统计信息：\n"
                content += f"-共处理了{processed_rows}行数据\n"
                content += f"-单个非零值行:{single_value_rows}行\n"
                content += f"-多个非零值行:{multi_value_rows}行\n"
                content += f"-格式化处理了{formatted_count}个非零值\n"
                content += f"-矩阵E已通过公式计算并格式化完成（所有数值保留两位小数）\n\n"
            # 矩阵F信息 - 显示计算过程
            if 'matrix_F' in matrix_operations:
                matrix_f = matrix_operations['matrix_F']
                content += f"创建了矩阵F（构件实际得分矩阵），维度: {matrix_f.get('dimensions', 'N/A')}\n"
                content += f"基于矩阵E，开始计算各构件的实际得分...\n"
                # 显示计算详情
                calculation_details = matrix_f.get('calculation_details', [])
                if calculation_details:
                    for detail in calculation_details:
                        content += f"{detail}\n"
                # 显示统计信息
                processed_components = matrix_f.get('processed_components', 0)
                defective_components = matrix_f.get('defective_components', 0)
                min_score = matrix_f.get('min_score', 100)
                content += f"\n矩阵F统计信息：\n"
                content += f"-处理构件数量:{processed_components}个\n"
                content += f"-有病害构件数量:{defective_components}个\n"
                content += f"-最低得分:{min_score:.2f}\n\n"
            # 添加h值计算过程详细信息
            h_value_calculation = component_data.get('h_value_calculation', {})
            if h_value_calculation:
                content += f"🧮开始计算部件技术状况评分值...\n"
                # 构件总数量和t值信息
                component_quantity = h_value_calculation.get('component_quantity', 0)
                t_value = h_value_calculation.get('t_value', 0)
                t_value_details = h_value_calculation.get('t_value_details', '')
                content += f"构件总数量n2={component_quantity}\n"
                content += f"{t_value_details}\n"
                content += f"构件总数{component_quantity}->t值={t_value:.2f}\n"
                # 有缺陷构件数量
                defective_count = h_value_calculation.get('defective_count', 0)
                content += f"有缺陷的构件数量n1={defective_count}\n"
                
                # 最低得分和平均得分
                min_score = h_value_calculation.get('min_score', 100)
                avg_score = h_value_calculation.get('avg_score', 100)
                content += f"最低得分fmin={min_score:.2f}\n"
                content += f"平均得分average_f={avg_score:.2f}\n"
                
                # 计算方法和过程
                calculation_method = h_value_calculation.get('calculation_method', '')
                calculation_process = h_value_calculation.get('calculation_process', '')
                content += f"{calculation_method}\n"
                content += f"{calculation_process}\n"
                # 最终结果
                h_value = component_data.get('h_value', 0)
                content += f"部件{component_name}计算完成，评定分数={h_value:.2f}\n\n"
        else:
            content += "\n📝无病害数据，直接设置h值 = 100.00\n"
            content += "说明:该部件所有构件均无病害记录，按满分计算\n"
        # 最终状态
        h_value = component_data.get('h_value', 0)
        content += f"✅最终状态:\n"
        content += f"部件评分:{h_value}\n"
        # 添加评级说明
        if isinstance(h_value, (int, float)):
            if 95 <= h_value <= 100:
                grade_desc = "1类"
            elif 80 <= h_value < 95:
                grade_desc = "2类"
            elif 60 <= h_value < 80:
                grade_desc = "3类"
            elif 40 <= h_value < 60:
                grade_desc = "4类"
            elif 0 <= h_value < 40:
                grade_desc = "5类"
            content += f"技术状况:{grade_desc}\n"
        content += f"计算状态:{component_data.get('calculation_status', 'N/A')}\n"
        if component_data.get('error_message'):
            content += f"错误信息:{component_data.get('error_message', 'N/A')}\n"
        content += f"\n{'=' * 60}\n"
        return content
    def create_final_summary_tab(self):
        """创建最终汇总选项卡"""
        summary_widget = QWidget()
        layout = QVBoxLayout(summary_widget)
        summary_text = QTextEdit()
        summary_text.setReadOnly(True)
        # 获取计算汇总数据
        calculation_summary = self.calculation_data.get('calculation_summary', {})
        structure_systems_summary = calculation_summary.get('structure_systems_summary', {})
        # 构建结构系统评分结果
        structure_scores_text = ""
        if hasattr(self.parent_system, 'bridge_structure_scores'):
            bridge_scores = self.parent_system.bridge_structure_scores
            for structure_name, score in bridge_scores.items():
                # 评级判定
                if 95 <= score <= 100:
                    grade = "1类"
                elif 80 <= score < 95:
                    grade = "2类"
                elif 60 <= score < 80:
                    grade = "3类"
                elif 40 <= score < 60:
                    grade = "4类"
                elif 0 <= score < 40:
                    grade = "5类"
                structure_scores_text += f"{structure_name}: {score:.2f}分-技术状况等级:{grade}\n"
        else:
            structure_scores_text = "暂无结构系统评分数据"
        # 构建全桥评分结果
        bridge_score_text = ""
        if hasattr(self.parent_system, 'total_bridge_score'):
            total_score = self.parent_system.total_bridge_score
            bridge_score_text = f"全桥技术状况评分: {total_score:.2f}分\n"
            # 评级判定
            if 95 <= total_score <= 100:
                grade = "1类"
            elif 80 <= total_score < 95:
                grade = "2类"
            elif 60 <= total_score < 80:
                grade = "3类"
            elif 40 <= total_score < 60:
                grade = "4类"
            elif 0 <= total_score < 40:
                grade = "5类"
            bridge_score_text += f"技术状况等级:{grade}\n"
        else:
            bridge_score_text = "暂无全桥评分数据"
        # 构建结构系统统计
        systems_summary_text = ""
        for structure_name, summary in structure_systems_summary.items():
            systems_summary_text += f"{structure_name}:\n"
            systems_summary_text += f"  总部件数: {summary.get('total', 0)}\n"
            systems_summary_text += f"  成功计算: {summary.get('calculated', 0)}\n"
            systems_summary_text += f"  跳过部件: {summary.get('skipped', 0)}\n"
        summary_content = f"""
🎯最终评分汇总
{'=' * 60}
📊各结构部位评分结果
{'=' * 60}
{structure_scores_text}
🏆全桥技术状况评分
{'=' * 60}
{bridge_score_text}
📈计算统计信息
{'=' * 60}  
总部件数: {calculation_summary.get('total_components', 0)}
成功计算: {calculation_summary.get('calculated_components', 0)}
跳过部件: {calculation_summary.get('skipped_components', 0)}

📋各结构部位统计
{'=' * 60}
{systems_summary_text}
📋评分标准说明
{'=' * 60}
• 95-100分: 1类 (全新状态，功能完好)
• 80-95分:  2类 (有轻微缺损，对桥梁使用功能无影响)
• 60-80分:  3类 (有中度缺损，尚能维持正常使用功能)
• 40-60分:  4类 (主要构件有较大缺损，严重影响桥梁使用功能；或影响承载能力，不能保证正常使用)
• 0-40分:   5类 (主要构件存在严重缺损，不能正常使用，危机桥梁安全，桥梁处于危险状态)

📋5类桥梁技术状况单项控制指标
{'=' * 60}
1.上部结构有落梁；或有梁、板断裂现象
2.梁式桥上部承重构件控制截面出现全截面开裂；或组合结构上部承重构件结合面开裂贯通，造成截面组合作用严重降低
3.梁式桥上部承重构件有严重的异常位移，存在失稳现象
4.结构出现明显的永久变形，变形大于规范值
5.关键部位混凝土出现压碎或杆件失稳倾向；或桥面板出现严重塌陷
6.拱式桥拱脚严重错台、位移，造成拱顶挠度大于限值；或拱圈严重变形
7.圬工拱桥拱圈大范围砌体断裂，脱落现象严重
8.腹拱、侧墙、立墙或立柱产生破坏造成桥面板严重塌落
9.系杆或吊杆出现严重锈蚀或断裂现象
10.悬索桥主缆或多根吊索出现严重锈蚀、断丝
11.斜拉桥拉索钢丝出现严重锈蚀、断丝，主梁出现严重变形
12.扩大基础冲刷深度大于设计值，冲空面积达20%以上
13.桥墩(桥台或基础)不稳定,出现严重滑动、下沉、位移、倾斜等现象
14.悬索桥、斜拉桥索塔基础出现严重沉降或位移；或悬索桥锚碇有水平位移或沉降
        """
        summary_text.setPlainText(summary_content)
        layout.addWidget(summary_text)
        self.tab_widget.addTab(summary_widget, "🎯最终汇总")
class MatrixDimensionDialog(QDialog):
    """矩阵维度设置对话框"""
    def __init__(self, parent=None, auto_confirm=False):
        super().__init__(parent)
        self.matrix_rows = 5000
        self.matrix_cols = 20
        self.auto_confirm = auto_confirm  # 新增自动确认标志
        self.init_ui()
        # 如果设置了自动确认，则在初始化完成后自动点击确定
        if self.auto_confirm:
            # 使用QTimer延迟执行，确保UI完全初始化
            from PySide6.QtCore import QTimer
            QTimer.singleShot(100, self.auto_accept)
    def auto_accept(self):
        """自动点击确定按钮"""
        # 使用默认值直接接受
        self.matrix_rows = 5000
        self.matrix_cols = 20
        self.accept()
    def init_ui(self):
        self.setWindowTitle("矩阵维度设置")
        # 设置窗口大小
        window_width = 450
        window_height = 320
        # 获取屏幕尺寸并计算居中位置
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        # 设置窗口图标
        try:
            self.setWindowIcon(QIcon("Basic_data/Highway_bridges.ico"))
        except:
            pass
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                border-radius: 10px;
            }
            QLabel {
                color: #2c3e50;
                font-size: 14px;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
            }
            QLineEdit {
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                background-color: white;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
                font-size: 14px;
                padding: 8px 12px;
                selection-background-color: #3498db;
            }
            QLineEdit:focus {
                border: 2px solid #3498db;
                background-color: #f8f9fa;
            }
            QLineEdit:hover {
                border: 2px solid #85c1e9;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-size: 14px;
                font-weight: bold;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
                min-width: 80px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5dade2, stop:1 #3498db);
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2980b9, stop:1 #21618c);
                transform: translateY(1px);
            }
            QPushButton#cancelButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #95a5a6, stop:1 #7f8c8d);
            }
            QPushButton#cancelButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #bdc3c7, stop:1 #95a5a6);
            }
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 25, 30, 25)
        # 标题区域
        title_container = QWidget()
        title_layout = QVBoxLayout(title_container)
        title_layout.setSpacing(8)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("🔢设置矩阵A的维度")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 18, QFont.Bold))
        title_label.setStyleSheet("""
            font-size: 18px; 
            font-weight: bold; 
            color: #2c3e50; 
            margin: 5px;
            padding: 10px;
            background-color: rgba(52, 152, 219, 0.1);
            border-radius: 8px;
        """)
        title_layout.addWidget(title_label)
        subtitle_label = QLabel("请根据实际需要设置矩阵的行数和列数")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("color: #7f8c8d; font-size: 12px; margin-top: 5px;")
        title_layout.addWidget(subtitle_label)
        layout.addWidget(title_container)
        # 维度设置区域
        form_container = QWidget()
        form_container.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                border: 1px solid #e0e0e0;
            }
        """)
        form_layout = QFormLayout(form_container)
        form_layout.setSpacing(15)
        form_layout.setContentsMargins(20, 20, 20, 20)
        # 行数设置
        rows_label = QLabel("矩阵行数:")
        rows_label.setFont(QFont("楷体", 12, QFont.Bold))
        self.rows_input = QLineEdit()
        self.rows_input.setText("5000")
        self.rows_input.setPlaceholderText("请输入行数")
        form_layout.addRow(rows_label, self.rows_input)
        # 列数设置
        cols_label = QLabel("矩阵列数:")
        cols_label.setFont(QFont("楷体", 12, QFont.Bold))
        self.cols_input = QLineEdit()
        self.cols_input.setText("20")
        self.cols_input.setPlaceholderText("请输入列数")
        form_layout.addRow(cols_label, self.cols_input)
        layout.addWidget(form_container)
        # 详细说明区域
        info_container = QWidget()
        info_container.setStyleSheet("""
            QWidget {
                background-color: rgba(241, 196, 15, 0.1);
                border-radius: 8px;
                border-left: 4px solid #f1c40f;
            }
        """)
        info_layout = QVBoxLayout(info_container)
        info_layout.setContentsMargins(15, 12, 15, 12)
        info_title = QLabel("💡参数说明:")
        info_title.setFont(QFont("楷体", 12, QFont.Bold))
        info_title.setStyleSheet("color: #f39c12; margin-bottom: 5px;")
        info_layout.addWidget(info_title)
        info_text = QLabel("""
• 矩阵行数: 表示某部件中构件的最大数量
• 矩阵列数: 表示某构件病害类型的最大数量
""")
        info_text.setWordWrap(True)
        info_text.setStyleSheet("""
            color: #8b4513; 
            font-size: 11px; 
            line-height: 1.4;
            padding: 5px;
        """)
        info_layout.addWidget(info_text)
        layout.addWidget(info_container)
        # 按钮区域
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setSpacing(15)
        button_layout.setContentsMargins(0, 10, 0, 0)
        button_layout.addStretch()
        ok_button = QPushButton("✓确定")
        ok_button.clicked.connect(self.validate_and_accept)
        button_layout.addWidget(ok_button)
        cancel_button = QPushButton("✗取消")
        cancel_button.setObjectName("cancelButton")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        layout.addWidget(button_container)
    def validate_and_accept(self):
        """验证输入并接受对话框"""
        try:
            # 验证行数
            rows_text = self.rows_input.text().strip()
            if not rows_text:
                QMessageBox.warning(self, "输入错误", "请输入矩阵行数！")
                self.rows_input.setFocus()
                return
            rows = int(rows_text)
            if rows < 1 or rows > 10000:
                QMessageBox.warning(self, "输入错误",
                                    "矩阵行数必须在1到10000之间！\n\n"
                                    "建议值：\n"
                                    "•小型桥梁: 500-1000\n"
                                    "•中型桥梁: 1000-2000\n"
                                    "•大型桥梁: 2000-5000")
                self.rows_input.selectAll()
                self.rows_input.setFocus()
                return
            # 验证列数
            cols_text = self.cols_input.text().strip()
            if not cols_text:
                QMessageBox.warning(self, "输入错误", "请输入矩阵列数！")
                self.cols_input.setFocus()
                return
            cols = int(cols_text)
            if cols < 1 or cols > 100:
                QMessageBox.warning(self, "输入错误",
                                    "矩阵列数必须在1到100之间！\n\n"
                                    "建议值：\n"
                                    "•标准病害类型: 12列\n"
                                    "•扩展病害类型: 15-20列\n"
                                    "•自定义分析: 根据实际需要")
                self.cols_input.selectAll()
                self.cols_input.setFocus()
                return
            # 内存使用警告
            memory_mb = (rows * cols * 4) / (1024 * 1024)
            if memory_mb > 100:  # 超过100MB时警告
                reply = QMessageBox.question(self, "内存使用提醒",
                                             f"当前设置的矩阵维度为 {rows}×{cols}\n"
                                             f"预计内存占用: {memory_mb:.1f} MB\n\n"
                                             f"较大的矩阵可能会影响计算速度。\n"
                                             f"是否继续使用此设置？",
                                             QMessageBox.Yes | QMessageBox.No,
                                             QMessageBox.No)
                if reply == QMessageBox.No:
                    return
            self.matrix_rows = rows
            self.matrix_cols = cols
            self.accept()
        except ValueError:
            QMessageBox.warning(self, "输入错误",
                                "请输入有效的数字！\n\n"
                                "行数和列数必须是正整数。")
    def get_dimensions(self):
        """获取设置的矩阵维度"""
        return self.matrix_rows, self.matrix_cols
class DamageInputDialog(QDialog):
    """病害录入对话框"""
    def __init__(self, component_name, structure_name=None, parent=None):
        super().__init__(parent)
        self.component_name = component_name
        self.structure_name = structure_name  # 新增结构系统名称
        self.damage_data = []
        self.parent_system = parent  # 保存父窗口引用以访问桥梁数据
        self.init_ui()
    def init_ui(self):
        # 根据结构系统名称调整窗口标题
        if self.structure_name:
            window_title = f"{self.structure_name}-{self.component_name}病害录入"
        else:
            window_title = f"{self.component_name}病害录入"
        self.setWindowTitle(window_title)
        # 设置窗口大小
        window_width = 900
        window_height = 450
        # 获取屏幕尺寸并计算居中位置
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        layout = QVBoxLayout(self)
        # 标题
        if self.structure_name:
            title_text = f"{self.structure_name}-{self.component_name}病害信息录入"
        else:
            title_text = f"{self.component_name}病害信息录入"
        title_label = QLabel(title_text)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 16, QFont.Bold))
        layout.addWidget(title_label)
        # 病害表格
        self.damage_table = QTableWidget()
        self.damage_table.setColumnCount(7)
        self.damage_table.setHorizontalHeaderLabels(["序号", "构件编号", "构件名称", "病害类型", "评定标度", "病害描述", "扣分值"])
        # 设置列宽
        self.damage_table.setColumnWidth(0, 50)  # 序号
        self.damage_table.setColumnWidth(1, 80)  # 构件编号
        self.damage_table.setColumnWidth(2, 80)  # 构件名称
        self.damage_table.setColumnWidth(3, 150)  # 病害类型
        self.damage_table.setColumnWidth(4, 80)  # 评定标度
        self.damage_table.setColumnWidth(5, 300)  # 病害描述
        self.damage_table.horizontalHeader().setStretchLastSection(True)
        # 设置表格行高以改善显示效果
        self.damage_table.verticalHeader().setDefaultSectionSize(35)  # 设置默认行高为35像素
        self.damage_table.setAlternatingRowColors(True)  # 设置交替行颜色
        # ===== 新增：启用自动换行功能 =====
        self.damage_table.setWordWrap(True)  # 启用自动换行
        self.damage_table.resizeRowsToContents()  # 自动调整行高以适应内容
        # ===== 自动换行设置结束 =====
        # 在init_ui方法中的表格样式部分
        self.damage_table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f0f8ff;  /* 浅蓝色 */
                background-color: white;              /* 白色 */
                gridline-color: #d0d0d0;             /* 网格线颜色 */
                selection-background-color: #4a90e2;  /* 选中行背景色 */
            }
            QTableWidget::item {
                padding: 5px;
                border: none;
                word-wrap: break-word;  /* 支持自动换行 */
                text-align: center;     /* 默认居中对齐 */
            }
            QTableWidget::item:selected {
                background-color: #4a90e2;
                color: white;
            }
            QHeaderView::section {
                background-color: #5a7a9a;
                color: white;
                padding: 8px;
                border: 1px solid #4a6a8a;
                font-weight: bold;
                font-family: "楷体";
                font-size: 11px;  /* 调整表头字体大小 */
            }
        """)
        # 设置垂直表头居中
        vertical_header = self.damage_table.verticalHeader()
        vertical_header.setDefaultAlignment(Qt.AlignCenter)
        # 设置垂直表头字体
        vertical_header.setFont(QFont("Times New Roman", 11))
        # 确保垂直表头可见
        vertical_header.setVisible(True)
        layout.addWidget(self.damage_table)
        # 按钮区域
        button_layout = QHBoxLayout()
        add_button = QPushButton("添加病害")
        add_button.clicked.connect(self.add_damage_row)
        button_layout.addWidget(add_button)
        delete_button = QPushButton("删除选中")
        delete_button.clicked.connect(self.delete_selected_row)
        button_layout.addWidget(delete_button)
        button_layout.addStretch()
        ok_button = QPushButton("确定")
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    def get_disease_types(self):
        """获取当前部件的病害类型列表"""
        if not self.parent_system or not hasattr(self.parent_system, 'bridge_data'):
            QMessageBox.warning(self, "系统错误", "无法访问桥梁数据，请检查系统初始化是否正常！")
            return []
        component_data = {}
        # 如果指定了结构系统名称，从对应的结构系统中获取数据
        if self.structure_name and self.structure_name in self.parent_system.bridge_data:
            structure_data = self.parent_system.bridge_data[self.structure_name]
            component_data = structure_data.get(self.component_name, {})
            if not component_data:
                QMessageBox.warning(self, "配置错误",
                                    f"在结构系统'{self.structure_name}'中未找到部件'{self.component_name}'的配置信息！")
                return []
        else:
            # 兼容旧版本：在所有结构系统中查找部件
            for system_name, system_data in self.parent_system.bridge_data.items():
                if self.component_name in system_data:
                    component_data = system_data[self.component_name]
                    break
            if not component_data:
                QMessageBox.warning(self, "配置错误",
                                    f"在所有结构系统中都未找到部件'{self.component_name}'的配置信息！")
                return []
        diseases = component_data.get('diseases', [])
        if not diseases:
            QMessageBox.warning(self, "配置错误",
                                f"部件'{self.component_name}'没有配置病害类型信息！")
        return [disease['name'] for disease in diseases]
    def get_disease_max_scale(self, disease_name):
        """获取指定病害类型的最大标度值"""
        if not self.parent_system or not hasattr(self.parent_system, 'bridge_data'):
            QMessageBox.warning(self, "系统错误", "无法访问桥梁数据，请检查系统初始化是否正常！")
            return None
        # 如果指定了结构系统名称，从对应的结构系统中获取数据
        if self.structure_name and self.structure_name in self.parent_system.bridge_data:
            structure_data = self.parent_system.bridge_data[self.structure_name]
            component_data = structure_data.get(self.component_name, {})
            if not component_data:
                QMessageBox.warning(self, "配置错误",
                                    f"在结构系统'{self.structure_name}'中未找到部件'{self.component_name}'的配置信息！")
                return None
        else:
            # 兼容旧版本：在所有结构系统中查找部件
            component_data = {}
            for system_name, system_data in self.parent_system.bridge_data.items():
                if self.component_name in system_data:
                    component_data = system_data[self.component_name]
                    break
            if not component_data:
                QMessageBox.warning(self, "配置错误",
                                    f"在所有结构系统中都未找到部件'{self.component_name}'的配置信息！")
                return None
        diseases = component_data.get('diseases', [])
        if not diseases:
            QMessageBox.warning(self, "配置错误",
                                f"部件'{self.component_name}'没有配置病害类型信息！")
            return None
        # 查找指定的病害类型
        for disease in diseases:
            if disease['name'] == disease_name:
                max_scale = disease.get('max_scale')
                if max_scale is None:
                    QMessageBox.warning(self, "配置错误",
                                        f"病害类型'{disease_name}'缺少最大标度值配置！")
                    return None
                return max_scale
        # 如果找不到对应病害类型，给出详细警告信息
        structure_info = f"结构系统'{self.structure_name}'-" if self.structure_name else ""
        available_diseases = [disease['name'] for disease in diseases]
        QMessageBox.warning(self, "配置文件警告",
                            f"在{structure_info}部件'{self.component_name}'中未找到病害类型'{disease_name}'的配置信息！\n\n"
                            f"可用的病害类型有：{', '.join(available_diseases)}\n\n"
                            f"请检查配置文件 'Basic_data/Steel-concrete_composite_arch_bridge.json' 中是否包含该病害类型。")
        return None
    def calculate_deduction_score(self, scale, max_scale):
        """根据评定标度和最大标度计算扣分值"""
        try:
            scale_int = int(scale)
            # 如果max_scale为None，说明病害类型无效
            if max_scale is None:
                QMessageBox.warning(self, "警告", "无效的病害类型，请重新选择！")
                return 0
            max_scale_int = int(max_scale)
            # 验证标度范围
            if scale_int < 1 or scale_int > max_scale_int:
                QMessageBox.warning(self, "警告", f"评定标度必须在1到{max_scale_int}之间，请重新输入！")
                return 0
            # 根据最大标度确定扣分规则
            if max_scale_int == 3:
                # 3类最高等级：1类-0分，2类-20分，3类-35分
                score_map = {1: 0, 2: 20, 3: 35}
            elif max_scale_int == 4:
                # 4类最高等级：1类-0分，2类-25分，3类-40分，4类-50分
                score_map = {1: 0, 2: 25, 3: 40, 4: 50}
            elif max_scale_int == 5:
                # 5类最高等级：1类-0分，2类-35分，3类-45分，4类-60分，5类-100分
                score_map = {1: 0, 2: 35, 3: 45, 4: 60, 5: 100}
            else:
                # 如果最大标度不是3、4、5，给出警告
                QMessageBox.warning(self, "警告", f"不支持的最大标度值：{max_scale_int}，请检查配置文件！")
                return 0
            return score_map.get(scale_int, 0)
        except (ValueError, TypeError):
            QMessageBox.warning(self, "警告", "评定标度必须是有效的数字！")
            return 0
    def on_scale_changed(self, row):
        """当评定标度改变时自动计算扣分值"""
        scale_item = self.damage_table.item(row, 4)
        score_item = self.damage_table.item(row, 6)
        disease_combo = self.damage_table.cellWidget(row, 3)
        if scale_item and score_item and disease_combo:
            scale_text = scale_item.text().strip()
            disease_name = disease_combo.currentText()
            if scale_text and disease_name:
                max_scale = self.get_disease_max_scale(disease_name)
                deduction_score = self.calculate_deduction_score(scale_text, max_scale)
                score_item.setText(str(deduction_score))
    def on_disease_changed(self, row):
        """当病害类型改变时重新计算扣分值"""
        self.on_scale_changed(row)
    def add_damage_row(self):
        """添加病害行"""
        row = self.damage_table.rowCount()
        self.damage_table.insertRow(row)
        
        # 序号 - Times New Roman字体
        seq_item = QTableWidgetItem(str(row + 1))
        seq_item.setTextAlignment(Qt.AlignCenter)
        seq_item.setFlags(seq_item.flags() & ~Qt.ItemIsEditable)
        seq_item.setFont(QFont("Times New Roman", 10))
        self.damage_table.setItem(row, 0, seq_item)
        
        # 构件编号 - Times New Roman字体
        code_item = QTableWidgetItem("")
        code_item.setTextAlignment(Qt.AlignCenter)
        code_item.setFont(QFont("Times New Roman", 10))
        self.damage_table.setItem(row, 1, code_item)
        
        # 构件名称 - 楷体
        name_item = QTableWidgetItem("")
        name_item.setTextAlignment(Qt.AlignCenter)
        name_item.setFont(QFont("楷体", 10))
        self.damage_table.setItem(row, 2, name_item)
        
        # 病害类型 - 楷体
        disease_combo = QComboBox()
        disease_types = self.get_disease_types()
        disease_combo.addItems(disease_types)
        disease_combo.currentTextChanged.connect(lambda: self.on_disease_changed(row))
        disease_combo.setStyleSheet("""
            QComboBox {
                padding: 4px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                font-family: "楷体";
                font-size: 12px;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                width: 12px;
                height: 12px;
            }
        """)
        self.damage_table.setCellWidget(row, 3, disease_combo)
        
        # 评定标度 - Times New Roman字体
        scale_item = QTableWidgetItem("")
        scale_item.setTextAlignment(Qt.AlignCenter)
        scale_item.setFont(QFont("Times New Roman", 10))
        self.damage_table.setItem(row, 4, scale_item)
        
        # 病害描述 - 混合字体，居中显示，支持自动换行
        description_item = QTableWidgetItem("")
        description_item.setTextAlignment(Qt.AlignCenter)
        # 设置混合字体（优先楷体，备选Times New Roman）
        description_item.setFont(QFont("楷体, Times New Roman", 10))
        self.damage_table.setItem(row, 5, description_item)
        
        # 扣分值 - Times New Roman字体
        score_item = QTableWidgetItem("")
        score_item.setTextAlignment(Qt.AlignCenter)
        score_item.setFlags(score_item.flags() & ~Qt.ItemIsEditable)
        score_item.setBackground(Qt.lightGray)
        score_item.setFont(QFont("Times New Roman", 10))
        self.damage_table.setItem(row, 6, score_item)
        
        # 连接评定标度的变化事件
        self.damage_table.itemChanged.connect(self.on_item_changed)
    def on_item_changed(self, item):
        """当表格项改变时的处理"""
        if item.column() == 4:  # 评定标度列
            row = item.row()
            self.on_scale_changed(row)
    def delete_selected_row(self):
        """删除选中的行"""
        current_row = self.damage_table.currentRow()
        if current_row >= 0:
            self.damage_table.removeRow(current_row)
            # 重新编号
            for row in range(self.damage_table.rowCount()):
                seq_item = self.damage_table.item(row, 0)
                if seq_item:
                    seq_item.setText(str(row + 1))
    def get_damage_data(self):
        """获取病害数据"""
        damage_list = []
        incomplete_rows = []
        for row in range(self.damage_table.rowCount()):
            damage_info = {}
            # 获取各字段数据
            seq_item = self.damage_table.item(row, 0)
            if seq_item:
                damage_info["序号"] = seq_item.text()

            code_item = self.damage_table.item(row, 1)
            if code_item:
                damage_info["构件编号"] = code_item.text()

            name_item = self.damage_table.item(row, 2)
            if name_item:
                damage_info["构件名称"] = name_item.text()

            disease_combo = self.damage_table.cellWidget(row, 3)
            if disease_combo and isinstance(disease_combo, QComboBox):
                damage_info["病害类型"] = disease_combo.currentText()

            scale_item = self.damage_table.item(row, 4)
            if scale_item:
                damage_info["评定标度"] = scale_item.text()

            description_item = self.damage_table.item(row, 5)
            # 病害描述 - 新增
            if description_item:
                damage_info["病害描述"] = description_item.text()
                
            score_item = self.damage_table.item(row, 6)    
            if score_item:
                damage_info["扣分值"] = score_item.text()
            # 检查数据完整性
            if any(value.strip() for value in damage_info.values() if value):
                # 检查关键字段是否完整
                if (not damage_info.get("病害类型", "").strip() or
                        not damage_info.get("评定标度", "").strip()):
                    incomplete_rows.append(row + 1)
                else:
                    damage_list.append(damage_info)
        # 如果有不完整的行，提示用户
        if incomplete_rows:
            QMessageBox.warning(self, "数据不完整",
                                f"第 {', '.join(map(str, incomplete_rows))} 行的病害类型或评定标度未填写完整！")
        return damage_list
    def load_damage_data(self, damage_data):
        """加载已有的病害数据到对话框"""
        try:
            if not damage_data:
                return
            if not isinstance(damage_data, (list, tuple)):
                QMessageBox.warning(self, "数据格式错误", "病害数据格式不正确，应为列表格式！")
                return
            
            loaded_count = 0
            error_count = 0
            for i, damage_info in enumerate(damage_data):
                try:
                    if not isinstance(damage_info, dict):
                        error_count += 1
                        continue
                    self.add_damage_row()
                    row = self.damage_table.rowCount() - 1
                    
                    # 设置序号（自动生成）- Times New Roman字体
                    seq_item = self.damage_table.item(row, 0)
                    if seq_item and "序号" in damage_info:
                        try:
                            seq_item.setText(str(damage_info["序号"]))
                            seq_item.setFont(QFont("Times New Roman", 10))
                        except Exception:
                            seq_item.setText(str(row + 1))
                    
                    # 设置构件编号 - Times New Roman字体
                    if "构件编号" in damage_info:
                        try:
                            code_item = QTableWidgetItem(str(damage_info["构件编号"]))
                            code_item.setTextAlignment(Qt.AlignCenter)
                            code_item.setFont(QFont("Times New Roman", 10))
                            self.damage_table.setItem(row, 1, code_item)
                        except Exception as e:
                            print(f"设置构件编号失败 (行{i + 1}): {e}")
                    
                    # 设置构件名称 - 楷体
                    if "构件名称" in damage_info:
                        try:
                            name_item = QTableWidgetItem(str(damage_info["构件名称"]))
                            name_item.setTextAlignment(Qt.AlignCenter)
                            name_item.setFont(QFont("楷体", 10))
                            self.damage_table.setItem(row, 2, name_item)
                        except Exception as e:
                            print(f"设置构件名称失败 (行{i + 1}): {e}")
                    
                    # 设置病害类型（下拉框）- 楷体已在样式中设置
                    disease_combo = self.damage_table.cellWidget(row, 3)
                    if disease_combo and isinstance(disease_combo, QComboBox) and "病害类型" in damage_info:
                        try:
                            disease_text = str(damage_info["病害类型"])
                            index = disease_combo.findText(disease_text)
                            if index >= 0:
                                disease_combo.setCurrentIndex(index)
                            else:
                                print(f"未找到病害类型 '{disease_text}' (行{i + 1})")
                        except Exception as e:
                            print(f"设置病害类型失败 (行{i + 1}): {e}")
                    
                    # 设置评定标度 - Times New Roman字体
                    if "评定标度" in damage_info:
                        try:
                            scale_item = QTableWidgetItem(str(damage_info["评定标度"]))
                            scale_item.setTextAlignment(Qt.AlignCenter)
                            scale_item.setFont(QFont("Times New Roman", 10))
                            self.damage_table.setItem(row, 4, scale_item)
                        except Exception as e:
                            print(f"设置评定标度失败 (行{i + 1}): {e}")
                    
                    # 设置病害描述 - 混合字体，居中显示
                    if "病害描述" in damage_info:
                        try:
                            description_item = QTableWidgetItem(str(damage_info["病害描述"]))
                            description_item.setTextAlignment(Qt.AlignCenter)
                            # 设置混合字体
                            description_item.setFont(QFont("楷体, Times New Roman", 10))
                            self.damage_table.setItem(row, 5, description_item)
                        except Exception as e:
                            print(f"设置病害描述失败 (行{i + 1}): {e}")
                    
                    # 设置扣分值 - Times New Roman字体
                    if "扣分值" in damage_info:
                        try:
                            score_item = QTableWidgetItem(str(damage_info["扣分值"]))
                            score_item.setTextAlignment(Qt.AlignCenter)
                            score_item.setFont(QFont("Times New Roman", 10))
                            self.damage_table.setItem(row, 6, score_item)
                        except Exception as e:
                            print(f"设置扣分值失败 (行{i + 1}): {e}")
                    
                    loaded_count += 1
                except Exception as e:
                    error_count += 1
                    print(f"加载第{i + 1}行数据时发生错误: {e}")
                    # 如果添加行失败，尝试删除可能已添加的空行
                    try:
                        if self.damage_table.rowCount() > 0:
                            self.damage_table.removeRow(self.damage_table.rowCount() - 1)
                    except Exception:
                        pass
            
            # 显示加载结果
            if error_count > 0:
                QMessageBox.information(self, "数据加载完成",
                                        f"成功加载 {loaded_count} 条记录\n失败 {error_count} 条记录\n详细错误信息请查看控制台输出")
            elif loaded_count > 0:
                print(f"成功加载 {loaded_count} 条病害数据")
        except Exception as e:
            QMessageBox.critical(self, "加载失败", f"加载病害数据时发生严重错误：\n{str(e)}")
            print(f"load_damage_data 发生异常: {e}")
class BridgeEvaluationSystemBase(QMainWindow):
    """桥梁评定系统基类，包含通用方法"""
    def __init__(self):
        super().__init__()
        self.bridge_data = {}
        self.original_weights = {}  # 存储原始权重
        self.current_weights = {}  # 存储当前重新分
        self.component_damage_data = {}  # 存储各部件的病害数据
        self.structure_tables = {}  # 存储各结构系统的表格
        self.component_quantities = {}  # 存储各部件的构件总数量，键格式："结构系统-部件名称"
        self.current_project_name = ""  # **新增：当前项目名称**
        self.load_bridge_data()
        self.init_ui()
    def load_bridge_data(self):
        """从Basic_data目录加载桥梁数据"""
        try:
            config_path = os.path.join('Basic_data', 'Steel-concrete_composite_arch_bridge.json')
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 加载所有结构系统数据
                self.bridge_data = {
                    '桥面系': data.get('桥面系', {}),
                    '上部结构': data.get('上部结构', {}),
                    '下部结构': data.get('下部结构', {})
                }
                # 保存原始权重
                for system_name, system_data in self.bridge_data.items():
                    self.original_weights[system_name] = {}
                    for component_name, component_data in system_data.items():
                        if isinstance(component_data, dict):
                            self.original_weights[system_name][component_name] = component_data.get('weight', 0)
        except FileNotFoundError:
            QMessageBox.warning(self, "警告", "配置文件未找到")
            self.bridge_data = {'桥面系': {}, '上部结构': {}, '下部结构': {}}
        except json.JSONDecodeError:
            QMessageBox.warning(self, "警告", "配置文件格式错误")
            self.bridge_data = {'桥面系': {}, '上部结构': {}, '下部结构': {}}
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("公路桥梁技术状况评定程序")
        # 设置窗口大小和位置
        window_width = 550
        window_height = 835
        # 获取屏幕尺寸并计算居中位置
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        # 统一使用setGeometry设置位置和大小
        self.setGeometry(x, y, window_width, window_height)
        # 设置窗口图标
        try:
            self.setWindowIcon(QIcon("Basic_data/Highway_bridges.ico"))
        except:
            pass
        # 设置样式
        self.set_application_style()
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(5, 5, 5, 5)
        # 添加标题
        title_label = QLabel("钢-混组合拱桥技术状况评定")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 20, QFont.Bold))
        title_label.setStyleSheet(
            "color: #2c3e50; margin-bottom: 10px; font-size: 20px; font-weight: bold; font-family: '楷体';")
        main_layout.addWidget(title_label)
        # ==================== 数据操作区容器 ====================
        data_ops_container = QWidget()
        data_ops_container.setStyleSheet("""
            QWidget {
                background-color: rgba(39, 174, 96, 0.05);
                border: 1px solid rgba(39, 174, 96, 0.3);
                border-radius: 8px;
                margin: 2px;
            }
        """)
        data_ops_layout = QHBoxLayout(data_ops_container)
        data_ops_layout.setContentsMargins(15, 8, 15, 8)
        data_ops_layout.setSpacing(20)
        
        # 数据操作区标题
        data_ops_title = QLabel("📊病害录入")
        data_ops_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        data_ops_title.setFont(QFont("楷体", 12, QFont.Bold))
        data_ops_title.setStyleSheet("""
            QLabel {
                color: #27ae60;
                font-family: "楷体";
                font-weight: bold;
                min-width: 90px;
                max-width: 90px;
                background-color: transparent;
                border: none;
            }
        """)
        data_ops_layout.addWidget(data_ops_title)
        
        # 按钮容器（居中）
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(15)
        
        # 左侧弹性空间
        button_layout.addStretch()
        
        # 导入病害数据按钮
        import_button = QPushButton("📊导入病害数据")
        import_button.setFont(QFont("楷体", 11, QFont.Bold))
        import_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #229954;
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
        """)
        import_button.clicked.connect(self.import_all_bridge_data)
        button_layout.addWidget(import_button)
        
        # 导出模板按钮
        export_button = QPushButton("📋导出模板")
        export_button.setFont(QFont("楷体", 11, QFont.Bold))
        export_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2980b9;
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
        """)
        export_button.clicked.connect(self.export_all_bridge_template)
        button_layout.addWidget(export_button)
        
        # 右侧弹性空间
        button_layout.addStretch()
        
        data_ops_layout.addWidget(button_container)
        main_layout.addWidget(data_ops_container)
        
        # ==================== 计算分析区容器 ====================
        calc_container = QWidget()
        calc_container.setStyleSheet("""
            QWidget {
                background-color: rgba(231, 76, 60, 0.05);
                border: 1px solid rgba(231, 76, 60, 0.3);
                border-radius: 8px;
                margin: 2px;
            }
        """)
        calc_layout = QHBoxLayout(calc_container)
        calc_layout.setContentsMargins(15, 8, 15, 8)
        calc_layout.setSpacing(20)
        
        # 计算分析区标题
        calc_title = QLabel("🧮计算分析")
        calc_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        calc_title.setFont(QFont("楷体", 12, QFont.Bold))
        calc_title.setStyleSheet("""
            QLabel {
                color: #e74c3c;
                font-family: "楷体";
                font-weight: bold;
                min-width: 90px;
                max-width: 90px;
                background-color: transparent;
                border: none;
            }
        """)
        calc_layout.addWidget(calc_title)
        
        # 按钮容器（居中）
        calc_button_container = QWidget()
        calc_button_layout = QHBoxLayout(calc_button_container)
        calc_button_layout.setContentsMargins(0, 0, 0, 0)
        calc_button_layout.setSpacing(15)
        
        # 左侧弹性空间
        calc_button_layout.addStretch()
        
        # 计算按钮
        calculate_button = QPushButton("🧮矩阵运算")
        calculate_button.setFont(QFont("楷体", 11, QFont.Bold))
        calculate_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #ff6f61, stop:1 #ff8e53);
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #ff9a74, stop:1 #ffaa66);
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #e65b3d, stop:1 #f47b3a);
                transform: translateY(1px);
            }
        """)
        calculate_button.clicked.connect(self.start_matrix_calculation)
        calc_button_layout.addWidget(calculate_button)
        # 导出计算过程按钮
        export_process_button = QPushButton("🔢运算过程")
        export_process_button.setFont(QFont("楷体", 11, QFont.Bold))
        export_process_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #3498db, stop:1 #9b59b6);
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #2980b9, stop:1 #8e44ad);
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #2471a3, stop:1 #7d3c98);
            }
        """)
        export_process_button.clicked.connect(self.export_calculation_process)
        calc_button_layout.addWidget(export_process_button)
        
        # 右侧弹性空间
        calc_button_layout.addStretch()
        
        calc_layout.addWidget(calc_button_container)
        main_layout.addWidget(calc_container)
        
        # ==================== 数据管理区容器 ====================
        data_mgmt_container = QWidget()
        data_mgmt_container.setStyleSheet("""
            QWidget {
                background-color: rgba(243, 156, 18, 0.05);
                border: 1px solid rgba(243, 156, 18, 0.3);
                border-radius: 8px;
                margin: 2px;
            }
        """)
        data_mgmt_layout = QHBoxLayout(data_mgmt_container)
        data_mgmt_layout.setContentsMargins(15, 8, 15, 8)
        data_mgmt_layout.setSpacing(20)
        
        # 数据管理区标题
        data_mgmt_title = QLabel("💾数据管理")
        data_mgmt_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        data_mgmt_title.setFont(QFont("楷体", 12, QFont.Bold))
        data_mgmt_title.setStyleSheet("""
            QLabel {
                color: #f39c12;
                font-family: "楷体";
                font-weight: bold;
                min-width: 90px;
                max-width: 90px;
                background-color: transparent;
                border: none;
            }
        """)
        data_mgmt_layout.addWidget(data_mgmt_title)
        
        # 按钮容器（居中）
        mgmt_button_container = QWidget()
        mgmt_button_layout = QHBoxLayout(mgmt_button_container)
        mgmt_button_layout.setContentsMargins(0, 0, 0, 0)
        mgmt_button_layout.setSpacing(15)
        
        # 左侧弹性空间
        mgmt_button_layout.addStretch()
        
        # 保存数据按钮
        save_data_button = QPushButton("💾保存数据")
        save_data_button.setFont(QFont("楷体", 11, QFont.Bold))
        save_data_button.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #e67e22;
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background-color: #d35400;
            }
        """)
        save_data_button.clicked.connect(self.save_project_data)
        mgmt_button_layout.addWidget(save_data_button)
        
        # 历史数据按钮
        history_data_button = QPushButton("📂历史数据")
        history_data_button.setFont(QFont("楷体", 11, QFont.Bold))
        history_data_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #0099cc, stop:1 #33cccc);
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体", "Times New Roman", serif;
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #33cccc, stop:1 #66ccff);
                transform: translateY(-1px);
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #009999, stop:1 #006666);
                transform: translateY(1px);
            }
        """)
        history_data_button.setToolTip("查看和加载历史保存的项目数据")
        history_data_button.clicked.connect(self.open_project_manager)  # 连接点击事件
        mgmt_button_layout.addWidget(history_data_button)
        
        # 右侧弹性空间
        mgmt_button_layout.addStretch()
        
        data_mgmt_layout.addWidget(mgmt_button_container)
        main_layout.addWidget(data_mgmt_container)
        
        # ==================== 报告管理区容器 ====================
        report_container = QWidget()
        report_container.setStyleSheet("""
            QWidget {
                background-color: rgba(155, 89, 182, 0.05);
                border: 1px solid rgba(155, 89, 182, 0.3);
                border-radius: 8px;
                margin: 2px;
            }
        """)
        report_layout = QHBoxLayout(report_container)
        report_layout.setContentsMargins(15, 8, 15, 8)
        report_layout.setSpacing(20)
        
        # 报告管理区标题
        report_title = QLabel("📋病害分析")
        report_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        report_title.setFont(QFont("楷体", 12, QFont.Bold))
        report_title.setStyleSheet("""
            QLabel {
                color: #9b59b6;
                font-family: "楷体";
                font-weight: bold;
                min-width: 90px;
                max-width: 90px;
                background-color: transparent;
                border: none;
            }
        """)
        report_layout.addWidget(report_title)
        # 按钮容器（居中）
        report_button_container = QWidget()
        report_button_layout = QHBoxLayout(report_button_container)
        report_button_layout.setContentsMargins(0, 0, 0, 0)
        report_button_layout.setSpacing(15)
        # 左侧弹性空间
        report_button_layout.addStretch()
        # 展示病害热力图按钮
        heatmap_button = QPushButton("📈病害分布")
        heatmap_button.setFont(QFont("楷体", 11, QFont.Bold))
        heatmap_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #8e44ad, stop:1 #3498db);
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #9b59b6, stop:1 #5dade2);
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #7d3c98, stop:1 #2980b9);
            }
        """)
        heatmap_button.clicked.connect(self.show_damage_heatmap)
        report_button_layout.addWidget(heatmap_button)
        # 导出报告按钮
        export_report_button = QPushButton("🧾导出报告")
        export_report_button.setFont(QFont("楷体", 11, QFont.Bold))
        export_report_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #ff4444, stop:1 #4444ff);
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #cc3333, stop:1 #3333cc);
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #992222, stop:1 #222299);
            }
        """)
        export_report_button.clicked.connect(self.export_inspection_report)
        report_button_layout.addWidget(export_report_button)
        
        # 右侧弹性空间
        report_button_layout.addStretch()
        
        report_layout.addWidget(report_button_container)
        main_layout.addWidget(report_container)

        # 添加分隔线
        separator = QLabel()
        separator.setFixedHeight(1)
        separator.setStyleSheet("background-color: #bdc3c7; margin: 10px 20px;")
        main_layout.addWidget(separator)
        
        # 创建选项卡控件
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # 创建各个结构系统的选项卡
        self.create_structure_tabs()
    #============数据管理模块=============

    def open_project_manager(self):
        """打开项目管理窗口"""
        try:
            project_manager = ProjectManagerWindow(self)
            project_manager.exec()
        except Exception as e:
            QMessageBox.critical(
                self,
                "错误",
                f"打开项目管理器时发生错误：\n{str(e)}"
            )
            print(f"打开项目管理器时发生错误：{e}")

    def save_project_data(self):
        """保存项目数据（仅保存矩阵运算必需的核心数据）"""
        try:
            # **修改：使用当前项目名称作为默认值**
            default_name = self.current_project_name if hasattr(self, 'current_project_name') and self.current_project_name else ""
            # 获取项目名称
            project_name, ok = QInputDialog.getText(
                self,
                "保存项目",
                "请输入项目名称：",
                text=default_name  # **使用当前项目名称作为默认值**
            )
            if not ok or not project_name.strip():
                return
            project_name = project_name.strip()
            # **新增：更新当前项目名称**
            self.current_project_name = project_name
            # 检查项目名称是否合法
            invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
            if any(char in project_name for char in invalid_chars):
                QMessageBox.warning(
                    self, 
                    "项目名称无效", 
                    f"项目名称不能包含以下字符：{' '.join(invalid_chars)}"
                )
                return
            # 创建保存目录
            save_dir = "Saved_Projects/Steel-concrete_composite_arch_bridge"
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            # 检查是否存在同名项目
            existing_files = []
            if os.path.exists(save_dir):
                for filename in os.listdir(save_dir):
                    if filename.startswith(f"{project_name}_") and filename.endswith('.json'):
                        existing_files.append(filename)
            # 如果存在同名项目，询问用户操作
            if existing_files:
                # 显示现有版本信息
                version_info = "\n".join([f"• {f}" for f in sorted(existing_files)])
                # 创建自定义消息框，提供三个选项
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("项目名称已存在")
                msg_box.setText(f"发现同名项目'{project_name}'的以下版本：\n\n{version_info}\n\n您希望如何处理？")
                # 添加三个按钮
                new_version_btn = msg_box.addButton("另存为新版本", QMessageBox.YesRole)
                overwrite_btn = msg_box.addButton("覆盖最新版本", QMessageBox.AcceptRole)
                rename_btn = msg_box.addButton("重新输入名称", QMessageBox.RejectRole)
                msg_box.setDefaultButton(new_version_btn)
                msg_box.exec()
                clicked_button = msg_box.clickedButton()
                if clicked_button == rename_btn:
                    # 用户选择重新输入，递归调用
                    return self.save_project_data()
                elif clicked_button == overwrite_btn:
                    # **修复：覆盖最新版本时，先删除旧文件，然后用新时间戳创建文件**
                    if existing_files:
                        # 找到最新的文件（按时间戳排序）
                        latest_file = sorted(existing_files)[-1]
                        old_filepath = os.path.join(save_dir, latest_file)
                        # 删除旧文件
                        try:
                            os.remove(old_filepath)
                        except OSError as e:
                            print(f"删除旧文件时发生错误：{e}")
                        # 生成新的文件名（使用当前时间戳）
                        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"{project_name}_{timestamp}.json"
                        filepath = os.path.join(save_dir, filename)
                        overwrite_mode = True
                    else:
                        # 如果没有现有文件，按新版本处理
                        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"{project_name}_{timestamp}.json"
                        filepath = os.path.join(save_dir, filename)
                        overwrite_mode = False
                else:
                    # 默认为另存为新版本
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{project_name}_{timestamp}.json"
                    filepath = os.path.join(save_dir, filename)
                    overwrite_mode = False
            else:
                # 没有同名项目，直接保存
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"{project_name}_{timestamp}.json"
                filepath = os.path.join(save_dir, filename)
                overwrite_mode = False
            
            # 收集部件存在性状态（复选框勾选状态）
            component_existence = {}
            for structure_name, table in self.structure_tables.items():
                component_existence[structure_name] = {}
                for row in range(table.rowCount()):
                    checkbox_widget = table.cellWidget(row, 0)
                    if checkbox_widget:
                        # 找到复选框
                        checkbox = None
                        name_label = None
                        for child in checkbox_widget.findChildren(QCheckBox):
                            checkbox = child
                            break
                        for child in checkbox_widget.findChildren(QLabel):
                            name_label = child
                            break
                        if checkbox and name_label:
                            component_name = name_label.text()
                            component_existence[structure_name][component_name] = checkbox.isChecked()
            
            # **新增：收集部件评分结果（h值）**
            component_scores = {}
            component_calculation_details = {}  # 新增：存储详细计算信息
            if hasattr(self, 'component_calculation_results') and self.component_calculation_results:
                for component_key, result in self.component_calculation_results.items():
                    # component_key格式为"结构系统-部件名称"
                    if '-' in component_key:
                        structure_name, component_name = component_key.split('-', 1)
                        if structure_name not in component_scores:
                            component_scores[structure_name] = {}
                            component_calculation_details[structure_name] = {}  # 新增
                        # 存储h值（评分值）
                        component_scores[structure_name][component_name] = result.get('h_value', 0.0)
                        # 新增：存储详细计算信息
                        if hasattr(self, 'calculation_process_data') and self.calculation_process_data:
                            component_detail = self.calculation_process_data.get('component_details', {}).get(component_key, {})
                            h_value_calc = component_detail.get('h_value_calculation', {})
                            component_calculation_details[structure_name][component_name] = {
                                'avg_score': h_value_calc.get('avg_score', 100),
                                't_value': h_value_calc.get('t_value', None),
                                't_value_details': h_value_calc.get('t_value_details', ''),
                                'min_score': h_value_calc.get('min_score', 100),
                                'calculation_method': h_value_calc.get('calculation_method', ''),
                                'calculation_process': h_value_calc.get('calculation_process', '')
                            }
            
            # **新增：收集重新分配后的权重**
            adjusted_weights = {}
            if hasattr(self, 'current_weights') and self.current_weights:
                adjusted_weights = self.current_weights.copy()
            
            # **修改：先初始化evaluation_results，然后根据条件更新**
            evaluation_results = {
                "bridge_structure_scores": {},  # 各结构系统评分
                "bridge_structure_grades": {},  # 各结构系统评级（新增）
                "total_bridge_score": 0.00,  # 全桥总评分
                "total_bridge_grade": '未评定'  # 全桥技术状况等级
            }
            
            # 检查是否有有效的评分结果
            if (hasattr(self, 'bridge_structure_scores') and 
                self.bridge_structure_scores and 
                len(self.bridge_structure_scores) > 0):
                evaluation_results["bridge_structure_scores"] = self.bridge_structure_scores

            # **新增：自动计算各结构部位的评级**
            if (hasattr(self, 'bridge_structure_grades') and 
                self.bridge_structure_grades and 
                len(self.bridge_structure_grades) > 0):
                evaluation_results["bridge_structure_grades"] = self.bridge_structure_grades
                
            if (hasattr(self, 'total_bridge_score') and 
                self.total_bridge_score is not None and 
                self.total_bridge_score > 0):
                evaluation_results["total_bridge_score"] = self.total_bridge_score
                
            if (hasattr(self, 'total_bridge_grade') and 
                self.total_bridge_grade and 
                self.total_bridge_grade != "未评定"):
                evaluation_results["total_bridge_grade"] = self.total_bridge_grade
            
            # 收集需要保存的数据（仅保存矩阵运算必需的核心数据）
            project_data = {
                "project_info": {
                    "name": project_name,
                    "save_time": datetime.datetime.now().isoformat()
                },
                # 矩阵运算的核心输入数据（移除了bridge_data，因为可从配置文件加载）
                "component_damage_data": self.component_damage_data,  # 部件病害数据
                "component_quantities": self.component_quantities,  # 部件构件数量
                "component_existence": component_existence,  # 部件存在性状态
                "original_weights": self.original_weights,  # 原始权重配置
                "component_scores": component_scores,  # **新增：部件评分结果（h值）**
                "component_calculation_details": component_calculation_details,  # 新增部件计算详细过程
                "adjusted_weights": adjusted_weights,  # **新增：重新分配后的权重**
                "evaluation_results": evaluation_results  # 使用改进后的评分结果
            }
            # 处理无穷大值的函数
            def convert_infinity(obj):
                if isinstance(obj, dict):
                    return {k: convert_infinity(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [convert_infinity(item) for item in obj]
                elif isinstance(obj, float) and np.isinf(obj):
                    return "Infinity"
                else:
                    return obj
            # 转换project_data中的无穷大值
            project_data = convert_infinity(project_data)
            # 保存到JSON文件
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
            # **修改：改进成功消息，包含评分结果信息**
            evaluation_info = ""
            if evaluation_results["total_bridge_score"] > 0:
                evaluation_info = f"\n•评分结果（全桥评分：{evaluation_results['total_bridge_score']:.2f}分，等级：{evaluation_results['total_bridge_grade']}）"
            
            
            if overwrite_mode:
                success_message = f"项目数据已成功覆盖保存！\n\n" + \
                                f"项目名称：{project_name}\n" + \
                                f"操作类型：覆盖最新版本（已更新文件名）\n" + \
                                f"保存位置：{filepath}\n\n" + \
                                f"保存内容（矩阵运算核心数据）：\n" + \
                                f"•部件病害数据\n" + \
                                f"•部件构件数量\n" + \
                                f"•部件存在性状态\n" + \
                                f"•权重配置信息{evaluation_info}\n\n" + \
                                f"注：桥梁基础配置将从配置文件自动加载\n" + \
                                f"加载项目数据后可直接进行矩阵运算！"
            else:
                version_count = len(existing_files) + 1
                success_message = f"项目数据已成功保存！\n\n" + \
                                f"项目名称：{project_name}\n" + \
                                f"版本信息：第{version_count}个版本\n" + \
                                f"保存位置：{filepath}\n\n" + \
                                f"保存内容（矩阵运算核心数据）：\n" + \
                                f"•部件病害数据\n" + \
                                f"•部件构件数量\n" + \
                                f"•部件存在性状态\n" + \
                                f"•权重配置信息{evaluation_info}\n\n" + \
                                f"注：桥梁基础配置将从配置文件自动加载\n" + \
                                f"加载项目数据后可直接进行矩阵运算！"
            QMessageBox.information(
                self,
                "保存成功",
                success_message
            )
            print(f"项目数据已保存到：{filepath}")
        except Exception as e:
            QMessageBox.critical(
                self,
                "保存失败",
                f"保存项目数据时发生错误：\n{str(e)}"
            )
            print(f"保存项目数据时发生错误：{e}")


    def set_application_style(self):
        """设置应用程序样式"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QTabWidget::pane {
                border: none;  /* 移除边框 */
                background-color: white;
                padding: 0px;
            }
            QTabWidget::tab-bar {
                alignment: center;
            }
            QTabBar::tab {
                background-color: #bdc3c7;
                color: #2c3e50;
                padding: 8px 12px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-family: "楷体";
                font-size: 14px;
                font-weight: bold;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background-color: #3498db;
                color: white;
            }
            QTabBar::tab:hover {
                background-color: #85c1e9;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #34495e;  /* 减小边框宽度 */
                border-radius: 4px;
                gridline-color: #7f8c8d;
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
                show-decoration-selected: 1;
            }
            QTableWidget::item {
                padding: 4px;
                border-right: 1px solid #7f8c8d;
                border-bottom: 1px solid #7f8c8d;
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
                text-align: center;
            }
            QTableWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QHeaderView::section {
                background-color: #34495e;
                color: white;
                padding: 8px;
                border-right: 1px solid #2c3e50;
                border-bottom: 2px solid #2c3e50;
                font-weight: bold;
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
                text-align: center;
            }
            QHeaderView::section:last {
                border-right: none;
            }
            QLineEdit {
                padding: 4px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
                text-align: center;
            }
            QLineEdit:disabled {
                background-color: #ecf0f1;
                color: #7f8c8d;
                border-color: #d5dbdb;
            }
            QCheckBox {
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
            }
            QLabel {
                color: #2c3e50;
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
                text-align: center;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-family: "楷体";
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
    def create_structure_tabs(self):
        """创建结构系统选项卡"""
        structure_names = ['桥面系', '上部结构', '下部结构']
        for structure_name in structure_names:
            # 创建选项卡页面
            tab_widget = self.create_structure_widget(structure_name)
            self.tab_widget.addTab(tab_widget, structure_name)
        # 添加全桥技术状况选项卡
        bridge_status_widget = self.create_bridge_status_widget()
        self.tab_widget.addTab(bridge_status_widget, "全桥技术状况")
    def create_bridge_status_widget(self):
        """创建全桥技术状况选项卡页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        # 添加标题
        title_label = QLabel("各部件技术状况评定")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 16, QFont.Bold))
        title_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                background-color: rgba(52, 152, 219, 0.1);
                border-radius: 8px;
                padding: 10px;
                margin: 5px;
            }
        """)
        layout.addWidget(title_label)
        # 创建详细结果表格
        self.bridge_status_table = self.create_bridge_status_table()
        layout.addWidget(self.bridge_status_table)
        # 添加汇总结果表格
        summary_label = QLabel("全桥技术状况评定")
        summary_label.setAlignment(Qt.AlignCenter)
        summary_label.setFont(QFont("楷体", 14, QFont.Bold))
        summary_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                background-color: rgba(46, 204, 113, 0.1);
                border-radius: 6px;
                padding: 8px;
                margin: 5px 0;
            }
        """)
        layout.addWidget(summary_label)
        # 创建汇总表格
        self.bridge_summary_table = self.create_bridge_summary_table()
        layout.addWidget(self.bridge_summary_table)
        # 添加说明文字
        info_label = QLabel("注：此表格显示各部件的技术状况评分结果，数据来源于矩阵运算。")
        info_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 12px;
                padding: 5px;
                background-color: rgba(241, 196, 15, 0.1);
                border-radius: 5px;
            }
        """)
        layout.addWidget(info_label)
        return widget
    def create_bridge_status_table(self):
        """创建各部件技术状况评分结果表格"""
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["结构部位", "部件名称", "构件总数量", "评分"])
        # 设置列宽
        table.setColumnWidth(0, 120)  # 结构部位
        table.setColumnWidth(1, 150)  # 部件名称
        table.setColumnWidth(2, 100)  # 构件总数量
        table.setColumnWidth(3, 100)  # 评分
        # 设置表格样式
        table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f0f8ff;
                background-color: white;
                gridline-color: #d0d0d0;
                selection-background-color: #4a90e2;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 8px;
                border: none;
            }
            QTableWidget::item:selected {
                background-color: #4a90e2;
                color: white;
            }
            QHeaderView::section {
                background-color: #5a7a9a;
                color: white;
                padding: 10px;
                border: 1px solid #4a6a8a;
                font-weight: bold;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
            }
        """)
        # 设置表格属性
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(True)
        return table
    def update_bridge_status_table(self):
        """更新各部件技术状况评分结果表格"""
        if not hasattr(self, 'bridge_status_table'):
            return
        # 检查是否有计算结果
        if not hasattr(self, 'component_calculation_results') or not self.component_calculation_results:
            self.bridge_status_table.setRowCount(0)
            return
        # 清空表格
        self.bridge_status_table.setRowCount(0)
        # 按结构系统分组整理数据
        structure_data = {
            '桥面系': [],
            '上部结构': [],
            '下部结构': []
        }
        # 遍历计算结果，按结构系统分组
        for component_key, result in self.component_calculation_results.items():
            # 解析组件键格式："结构系统-部件名称"
            if '-' in component_key:
                structure_name, component_name = component_key.split('-', 1)
                if structure_name in structure_data:
                    structure_data[structure_name].append({
                        'component_name': component_name,
                        'component_count': result.get('component_quantity', 0),
                        'h_value': result.get('h_value', 0.00)
                    })
        # 添加数据到表格
        current_row = 0
        for structure_name in ['桥面系', '上部结构', '下部结构']:
            components = structure_data[structure_name]
            if not components:
                continue
            # 为每个结构系统的第一行添加合并的结构系统名称
            start_row = current_row
            for i, component_data in enumerate(components):
                self.bridge_status_table.insertRow(current_row)
                # 第一行显示结构系统名称，其他行留空
                if i == 0:
                    structure_item = QTableWidgetItem(structure_name)
                    structure_item.setTextAlignment(Qt.AlignCenter)
                    structure_item.setFont(QFont("楷体", 12, QFont.Bold))
                    self.bridge_status_table.setItem(current_row, 0, structure_item)
                else:
                    self.bridge_status_table.setItem(current_row, 0, QTableWidgetItem(""))
                # 部件名称
                component_item = QTableWidgetItem(component_data['component_name'])
                component_item.setTextAlignment(Qt.AlignCenter)
                component_item.setFont(QFont("楷体", 12))
                self.bridge_status_table.setItem(current_row, 1, component_item)
                # 构件数量
                count_item = QTableWidgetItem(str(component_data['component_count']))
                count_item.setTextAlignment(Qt.AlignCenter)
                count_item.setFont(QFont("Times New Roman", 12))  # 添加这行
                self.bridge_status_table.setItem(current_row, 2, count_item)
                # 评分（h值）
                score_item = QTableWidgetItem(f"{component_data['h_value']:.2f}")
                score_item.setTextAlignment(Qt.AlignCenter)
                score_item.setFont(QFont("Times New Roman", 12))  # 添加这行
                self.bridge_status_table.setItem(current_row, 3, score_item)
                current_row += 1
            # 合并结构系统列的单元格
            if len(components) > 1:
                self.bridge_status_table.setSpan(start_row, 0, len(components), 1)
    def create_bridge_summary_table(self):
        """创建全桥技术状况汇总表格"""
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["结构部位", "技术状况评分", "技术状况等级"])
        # 设置列宽
        table.setColumnWidth(0, 150)  # 结构部位
        table.setColumnWidth(1, 150)  # 技术状况评分
        table.setColumnWidth(2, 150)  # 技术状况等级
        # 设置表格样式
        table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #e8f5e8;
                background-color: white;
                gridline-color: #d0d0d0;
                selection-background-color: #2ecc71;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
                font-size: 12px;
                border: 2px solid #27ae60;
            }
            QTableWidget::item {
                padding: 10px;
                border: none;
            }
            QTableWidget::item:selected {
                background-color: #2ecc71;
                color: white;
            }
            QHeaderView::section {
                background-color: #27ae60;
                color: white;
                padding: 12px;
                border: 1px solid #229954;
                font-weight: bold;
                font-family: "楷体", "Microsoft YaHei", "Times New Roman";
            }
        """)
        # 设置表格属性
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(True)
        table.setMaximumHeight(200)  # 限制表格高度
        return table
    def update_bridge_summary_table(self):
        """更新全桥技术状况汇总表格"""
        if not hasattr(self, 'bridge_summary_table'):
            return
        # 清空表格
        self.bridge_summary_table.setRowCount(0)
        # 检查是否有计算结果
        if not hasattr(self, 'bridge_structure_scores') or not self.bridge_structure_scores:
            return
        # 技术状况等级评估函数
        def evaluate_technical_condition(score):
            if 95 <= score <= 100:
                return "1类"
            elif 80 <= score < 95:
                return "2类"
            elif 60 <= score < 80:
                return "3类"
            elif 40 <= score < 60:
                return "4类"
            elif 0 <= score < 40:
                return "5类"
            else:
                raise ValueError(f"无效的评分值:{score}，评分应在0到100之间")
        # 添加各结构系统数据
        structure_order = ['桥面系', '上部结构', '下部结构']
        for i, structure_name in enumerate(structure_order):
            if structure_name in self.bridge_structure_scores:
                self.bridge_summary_table.insertRow(i)
                # 结构系统名称
                structure_item = QTableWidgetItem(structure_name)
                structure_item.setTextAlignment(Qt.AlignCenter)
                structure_item.setFont(QFont("楷体", 12, QFont.Bold))
                self.bridge_summary_table.setItem(i, 0, structure_item)
                # 技术状况评分
                score = self.bridge_structure_scores[structure_name]
                score_item = QTableWidgetItem(f"{score:.2f}")
                score_item.setTextAlignment(Qt.AlignCenter)
                score_item.setFont(QFont("Times New Roman", 12))
                self.bridge_summary_table.setItem(i, 1, score_item)
                # 技术状况等级
                grade = evaluate_technical_condition(score)
                grade_item = QTableWidgetItem(grade)
                grade_item.setTextAlignment(Qt.AlignCenter)
                grade_item.setFont(QFont("Times New Roman", 12))
                self.bridge_summary_table.setItem(i, 2, grade_item)
        # 添加全桥总评分行
        if hasattr(self, 'total_bridge_score') and hasattr(self, 'total_bridge_grade'):
            total_row = len(structure_order)
            self.bridge_summary_table.insertRow(total_row)
            # 全桥总评分
            total_structure_item = QTableWidgetItem("全桥总评分")
            total_structure_item.setTextAlignment(Qt.AlignCenter)
            total_structure_item.setFont(QFont("楷体", 12, QFont.Bold))
            total_structure_item.setBackground(Qt.lightGray)
            self.bridge_summary_table.setItem(total_row, 0, total_structure_item)
            total_score_item = QTableWidgetItem(f"{self.total_bridge_score:.2f}")
            total_score_item.setTextAlignment(Qt.AlignCenter)
            total_score_item.setFont(QFont("Times New Roman", 12))
            total_score_item.setBackground(Qt.lightGray)
            self.bridge_summary_table.setItem(total_row, 1, total_score_item)
            total_grade_item = QTableWidgetItem(self.total_bridge_grade)
            total_grade_item.setTextAlignment(Qt.AlignCenter)
            total_grade_item.setFont(QFont("Times New Roman", 12))
            total_grade_item.setBackground(Qt.lightGray)
            self.bridge_summary_table.setItem(total_row, 2, total_grade_item)
    def create_structure_widget(self, structure_name):
        """创建结构系统选项卡页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 10, 10, 10)
        # 创建表格
        table = self.create_structure_table(structure_name)
        # 设置表格尺寸策略为可扩展
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.structure_tables[structure_name] = table
        # 直接添加表格到布局，让它填充整个区域
        layout.addWidget(table)
        return widget
    def create_structure_table(self, structure_name):
        """为指定结构系统创建表格"""
        # 创建表格
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["部件", "权重", "构件总数量", "病害录入", "备注"])
        # 添加统一的表格样式
        table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f0f8ff;  /* 浅蓝色 */
                background-color: white;              /* 白色 */
                gridline-color: #d0d0d0;             /* 网格线颜色 */
                selection-background-color: #4a90e2;  /* 选中行背景色 */
                font-family: "楷体", "Times New Roman";
                font-size: 12px;
                border: 1px solid #bdc3c7;          /* 表格外边框 */
            }
            QTableWidget::item {
                padding: 5px;
                border: 1px solid #e0e0e0;          /* 单元格边框 */
            }
            QTableWidget::item:selected {
                background-color: #4a90e2;
                color: white;
            }
            QHeaderView::section {
                background-color: #5a7a9a;
                color: white;
                padding: 8px;
                border: 1px solid #4a6a8a;
                font-weight: bold;
                font-family: "楷体", "Times New Roman";
            }
        """)
        # 设置表格属性 - 使用自适应列宽
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # 部件列自适应
        header.setSectionResizeMode(1, QHeaderView.Fixed)  # 权重列固定
        header.setSectionResizeMode(2, QHeaderView.Fixed)  # 构件总数量列适应内容
        header.setSectionResizeMode(3, QHeaderView.Fixed)  # 病害录入列固定
        header.setSectionResizeMode(4, QHeaderView.Stretch)  # 备注列自适应
        # 设置固定列的宽度
        table.setColumnWidth(0, 120)  # 部件列
        table.setColumnWidth(1, 60)  # 权重列
        table.setColumnWidth(2, 80)  # 构件总数量列
        table.setColumnWidth(3, 100)  # 病害录入列
        # 设置表格尺寸策略
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # 设置行高
        table.verticalHeader().setDefaultSectionSize(40)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        # 设置表头字体
        header_font = QFont("楷体", 12, QFont.Bold)
        table.horizontalHeader().setFont(header_font)
        # 初始化表格内容
        self.populate_structure_table(table, structure_name)
        return table
    def populate_structure_table(self, table, structure_name):
        """填充指定结构系统的表格数据"""
        structure_data = self.bridge_data.get(structure_name, {})
        if not structure_data:
            # 如果没有数据，显示提示信息
            table.setRowCount(1)
            info_item = QTableWidgetItem(f"暂无{structure_name}数据")
            info_item.setTextAlignment(Qt.AlignCenter)
            info_item.setFlags(info_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(0, 0, info_item)
            table.setSpan(0, 0, 1, 4)  # 合并单元格
            return
        # 设置行数
        table.setRowCount(len(structure_data))
        # 填充表格数据
        row = 0
        for component_name, component_data in structure_data.items():
            if isinstance(component_data, dict):
                weight = component_data.get('weight', 0)
                # 创建包含复选框和部件名称的容器
                checkbox_widget = QWidget()
                checkbox_layout = QHBoxLayout(checkbox_widget)
                checkbox_layout.setContentsMargins(5, 0, 5, 0)
                checkbox_layout.setSpacing(5)
                # 复选框
                checkbox = QCheckBox()
                checkbox.setChecked(True)  # 默认勾选
                checkbox.stateChanged.connect(lambda state, sn=structure_name: self.update_structure_weights(sn))
                checkbox_layout.addWidget(checkbox)
                # 部件名称标签
                name_label = QLabel(component_name)
                name_label.setFont(QFont("楷体", 12, QFont.Bold))
                name_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                checkbox_layout.addWidget(name_label)
                checkbox_layout.addStretch()
                table.setCellWidget(row, 0, checkbox_widget)
                # 权重 - 保留两位小数
                weight_item = QTableWidgetItem(f"{weight:.2f}")
                weight_item.setFlags(weight_item.flags() & ~Qt.ItemIsEditable)  # 设置为只读
                weight_item.setTextAlignment(Qt.AlignCenter)
                weight_item.setFont(QFont("Times New Roman", 12))  # 数字使用新罗马
                table.setItem(row, 1, weight_item)
                # 构件总数量输入框
                quantity_input = QLineEdit()
                quantity_input.setAlignment(Qt.AlignCenter)
                quantity_input.setFont(QFont("Times New Roman", 14))
                quantity_input.setPlaceholderText("")  # 无占位符
                quantity_input.setStyleSheet("font-family: 'Times New Roman'; font-size: 14px;")
                table.setCellWidget(row, 2, quantity_input)
                # 病害录入按钮
                damage_button = QPushButton("病害录入")
                damage_button.clicked.connect(
                    lambda checked, name=component_name, sn=structure_name: self.open_structure_damage_input(name, sn))
                table.setCellWidget(row, 3, damage_button)
                # 备注输入框
                remark_input = QLineEdit()
                remark_input.setAlignment(Qt.AlignCenter)
                remark_input.setFont(QFont("Times New Roman", 14))
                remark_input.setPlaceholderText("")
                remark_input.setStyleSheet("font-family: 'Times New Roman'; font-size: 14px;")
                table.setCellWidget(row, 4, remark_input)
                row += 1
    def update_structure_weights(self, structure_name):
        """更新指定结构系统的权重分配"""
        table = self.structure_tables.get(structure_name)
        if not table:
            return
        structure_data = self.bridge_data.get(structure_name, {})
        component_names = list(structure_data.keys())
        # 初始化当前结构系统的权重存储
        if structure_name not in self.current_weights:
            self.current_weights[structure_name] = {}
        # 获取选中的部件
        selected_components = []
        for row in range(table.rowCount()):
            if row < len(component_names):
                checkbox_widget = table.cellWidget(row, 0)
                if checkbox_widget:
                    checkbox = checkbox_widget.findChild(QCheckBox)
                    if checkbox and checkbox.isChecked():
                        selected_components.append(component_names[row])
        if not selected_components:
            # 如果没有选中任何部件，所有权重都设为0
            for row in range(table.rowCount()):
                if row < len(component_names):
                    component_name = component_names[row]
                    weight_item = table.item(row, 1)
                    if weight_item:
                        weight_item.setText("0.00")
                    # 保存权重为0（使用round固定精度）
                    self.current_weights[structure_name][component_name] = round(0.00, 2)
            # 输出权重分配情况
            print(f"\n=== {structure_name} 权重分配情况 ===")
            total_weight = 0
            for component_name, weight in self.current_weights[structure_name].items():
                print(f"{structure_name}-{component_name}: {weight:.2f}")
                total_weight += weight
            print(f"权重总和: {total_weight:.2f}")
            return
        # 计算选中部件的原始权重总和
        total_original_weight = sum(
            self.original_weights.get(structure_name, {}).get(comp, 0)
            for comp in selected_components
        )
        if total_original_weight == 0:
            return
        # 计算新权重（保持精度）
        new_weights = []
        for comp in selected_components:
            original_weight = self.original_weights.get(structure_name, {}).get(comp, 0)
            new_weight = original_weight / total_original_weight
            new_weights.append(new_weight)
        # 确保权重总和精确等于1（处理浮点数精度问题）
        weight_sum = sum(new_weights)
        if weight_sum > 0:
            # 按比例调整，确保总和为1
            new_weights = [w / weight_sum for w in new_weights]
            # 直接使用round固定到小数点后两位
            new_weights = [round(w, 2) for w in new_weights]
            # 处理舍入后的总和偏差
            current_sum = sum(new_weights)
            if current_sum != 1.0:
                # 找到最大权重的索引来调整偏差
                max_index = new_weights.index(max(new_weights))
                adjustment = round(1.0 - current_sum, 2)
                new_weights[max_index] = round(new_weights[max_index] + adjustment, 2)
        # 更新表格中的权重和控件状态
        selected_index = 0
        for row in range(table.rowCount()):
            if row < len(component_names):
                component_name = component_names[row]
                checkbox_widget = table.cellWidget(row, 0)
                if checkbox_widget:
                    checkbox = checkbox_widget.findChild(QCheckBox)
                    weight_item = table.item(row, 1)
                    quantity_input = table.cellWidget(row, 2)  # 构件总数量输入框
                    damage_button = table.cellWidget(row, 3)  # 病害录入按钮
                    if checkbox and weight_item:
                        if checkbox.isChecked():
                            # 选中的部件：设置新权重，启用控件
                            if selected_index < len(new_weights):
                                new_weight = new_weights[selected_index]
                                weight_item.setText(f"{new_weight:.2f}")
                                # 保存重新分配后的权重（已经是round后的值）
                                self.current_weights[structure_name][component_name] = new_weight
                                selected_index += 1
                            # 启用构件总数量输入框
                            if quantity_input and isinstance(quantity_input, QLineEdit):
                                quantity_input.setEnabled(True)
                                # 添加这行：如果当前值是"0"，则清空
                                if quantity_input.text() == "0":
                                    quantity_input.clear()
                                quantity_input.setStyleSheet("""
                                    QLineEdit {
                                        padding: 4px;
                                        border: 1px solid #bdc3c7;
                                        border-radius: 4px;
                                        background-color: white;
                                        font-family: "Times New Roman";
                                        font-size: 14px;
                                        text-align: center;
                                    }
                                """)
                            # 启用病害录入按钮
                            if damage_button and isinstance(damage_button, QPushButton):
                                damage_button.setEnabled(True)
                                damage_button.setStyleSheet("""
                                    QPushButton {
                                        background-color: #3498db;
                                        color: white;
                                        border: none;
                                        padding: 6px 12px;
                                        border-radius: 4px;
                                        font-family: "楷体";
                                        font-size: 11px;
                                        font-weight: bold;
                                    }
                                    QPushButton:hover {
                                        background-color: #2980b9;
                                    }
                                    QPushButton:pressed {
                                        background-color: #21618c;
                                    }
                                """)
                        else:
                            # 未选中的部件：权重为0，禁用控件
                            weight_item.setText("0.00")
                            # 保存权重为0（使用round固定精度）
                            self.current_weights[structure_name][component_name] = round(0.00, 2)
                            # 禁用并清空构件总数量输入框
                            if quantity_input and isinstance(quantity_input, QLineEdit):
                                quantity_input.setText("0")  # 设置默认值为0
                                quantity_input.setEnabled(False)
                                quantity_input.setStyleSheet("""
                                    QLineEdit {
                                        padding: 4px;
                                        border: 1px solid #d5dbdb;
                                        border-radius: 4px;
                                        background-color: #ecf0f1;
                                        color: #7f8c8d;
                                        font-family: "Times New Roman";
                                        font-size: 14px;
                                        text-align: center;
                                    }
                                """)
                            # 禁用病害录入按钮
                            if damage_button and isinstance(damage_button, QPushButton):
                                damage_button.setEnabled(False)
                                damage_button.setStyleSheet("""
                                    QPushButton {
                                        background-color: #bdc3c7;
                                        color: #7f8c8d;
                                        border: none;
                                        padding: 6px 12px;
                                        border-radius: 4px;
                                        font-family: "楷体";
                                        font-size: 11px;
                                        font-weight: bold;
                                    }
                                """)
        # 输出当前权重分配情况
        print(f"\n=== {structure_name} 权重分配情况 ===")
        total_weight = 0
        for component_name, weight in self.current_weights[structure_name].items():
            print(f"{structure_name}-{component_name}: {weight:.2f}")
            total_weight += weight
        print(f"权重总和: {total_weight:.2f}")
    def open_structure_damage_input(self, component_name, structure_name):
        """打开指定结构系统部件的病害录入对话框"""
        dialog = DamageInputDialog(component_name, structure_name, self)
        # 加载已有病害数据
        damage_key = f"{structure_name}-{component_name}"
        if damage_key in self.component_damage_data:
            damage_data = self.component_damage_data[damage_key]
            dialog.load_damage_data(damage_data)
        if dialog.exec() == QDialog.Accepted:
            # 保存病害数据
            damage_data = dialog.get_damage_data()
            self.component_damage_data[damage_key] = damage_data
            print(f"{structure_name}-{component_name} 病害数据已保存：", damage_data)
    def import_all_bridge_data(self):
        """导入全桥病害数据（支持三个结构系统）"""
        try:
            import pandas as pd
            from PySide6.QtWidgets import QFileDialog
            # 检查Import_data文件夹是否存在
            import_folder = "Import_data/Steel-concrete_composite_arch_bridge"
            if not os.path.exists(import_folder):
                QMessageBox.information(self, "提示",
                                        f"未找到'{import_folder}'文件夹，请先创建该文件夹并放入Excel文件。")
                return
            # 检查文件夹是否为空
            excel_files = [f for f in os.listdir(import_folder) if f.endswith(('.xlsx', '.xls'))]
            if not excel_files:
                QMessageBox.information(self, "提示",
                                        f"'{import_folder}'文件夹中没有找到Excel文件，请先放入病害数据文件。")
                return
            # 打开文件选择对话框
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "选择全桥病害数据Excel文件",
                import_folder,
                "Excel文件 (*.xlsx *.xls)"
            )
            if not file_path:
                return
            # 验证文件是否存在且可读
            if not os.path.exists(file_path):
                QMessageBox.critical(self, "文件错误", "选择的文件不存在！")
                return
            if not os.access(file_path, os.R_OK):
                QMessageBox.critical(self, "文件错误", "无法读取选择的文件，请检查文件权限！")
                return
            # 读取Excel文件的所有工作表
            try:
                excel_file = pd.ExcelFile(file_path)
            except Exception as e:
                QMessageBox.critical(self, "文件格式错误",
                                     f"无法读取Excel文件，请确保文件格式正确！\n错误详情：{str(e)}")
                return
            # 定义结构系统名称映射（Excel工作表名 -> 系统名称）
            structure_mapping = {
                '桥面系病害数据': '桥面系',
                '上部结构病害数据': '上部结构',
                '下部结构病害数据': '下部结构'
            }
            # 检查必要的工作表是否存在
            missing_sheets = []
            for sheet_name in structure_mapping.keys():
                if sheet_name not in excel_file.sheet_names:
                    missing_sheets.append(sheet_name)
            if missing_sheets:
                missing_info = "\n".join([f"- {sheet}" for sheet in missing_sheets])
                available_info = "\n".join([f"- {sheet}" for sheet in excel_file.sheet_names])
                QMessageBox.warning(self, "工作表缺失",
                                    f"Excel文件中缺少以下必要的工作表：\n{missing_info}\n\n"
                                    f"当前文件包含的工作表：\n{available_info}\n\n"
                                    f"程序将只处理存在的工作表。")
            # 存储所有结构系统的导入结果
            import_results = {}
            total_imported_components = 0
            total_imported_damages = 0
            all_warnings = []
            all_errors = []
            # 逐个处理每个结构系统
            for sheet_name, structure_name in structure_mapping.items():
                if sheet_name not in excel_file.sheet_names:
                    import_results[structure_name] = {
                        'status': 'missing',
                        'component_count': 0,
                        'damage_count': 0,
                        'message': f'未找到工作表"{sheet_name}"'
                    }
                    continue
                try:
                    # 读取指定工作表的所有数据
                    df_all = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    # 检查工作表是否为空
                    if df_all.empty:
                        import_results[structure_name] = {
                            'status': 'error',
                            'component_count': 0,
                            'damage_count': 0,
                            'message': f'工作表"{sheet_name}"为空'
                        }
                        continue
                    # 处理单个结构系统的数据
                    result = self.process_structure_data(df_all, structure_name)
                    import_results[structure_name] = result
                    if result['status'] == 'success':
                        total_imported_components += result['component_count']
                        total_imported_damages += result['damage_count']
                    # 收集警告和错误信息
                    if 'warnings' in result and result['warnings']:
                        for warning in result['warnings']:
                            all_warnings.append(f"[{structure_name}] {warning}")
                    if 'errors' in result and result['errors']:
                        for error in result['errors']:
                            all_errors.append(f"[{structure_name}] {error}")
                except Exception as e:
                    import traceback
                    error_detail = f"处理工作表时发生错误：{str(e)}\n{traceback.format_exc()}"
                    import_results[structure_name] = {
                        'status': 'error',
                        'component_count': 0,
                        'damage_count': 0,
                        'message': error_detail
                    }
                    all_errors.append(f"[{structure_name}] {error_detail}")
            # 生成导入结果消息
            result_messages = []
            success_count = 0
            for structure_name, result in import_results.items():
                if result['status'] == 'success':
                    success_count += 1
                    result_messages.append(
                        f"{structure_name}：成功导入{result['component_count']}个部件，{result['damage_count']}条病害记录"
                    )
                elif result['status'] == 'missing':
                    result_messages.append(f"{structure_name}：{result['message']}")
                else:
                    result_messages.append(f"{structure_name}：导入失败 - {result['message']}")
            # 显示导入结果
            if success_count > 0:
                result_msg = f"导入完成！\n" + "\n".join(result_messages)
                result_msg += f"\n总计：{total_imported_components}个部件，{total_imported_damages}条病害记录"
                result_msg += "\n权重已自动重新分配。"
                # 如果有警告信息，添加到结果中
                if all_warnings:
                    warning_count = len(all_warnings)
                    if warning_count <= 10:  # 显示所有警告
                        result_msg += f"\n注意事项（{warning_count}项）：\n" + "\n".join(all_warnings)
                    else:  # 只显示前10个警告
                        result_msg += f"\n注意事项（显示前10项，共{warning_count}项）：\n" + "\n".join(all_warnings[:10])
                        result_msg += f"\n...还有{warning_count-10}项警告未显示"
                QMessageBox.information(self, "导入完成", result_msg)
            else:
                result_msg = "导入失败！\n" + "\n".join(result_messages)
                # 如果有错误信息，添加到结果中
                if all_errors:
                    error_count = len(all_errors)
                    if error_count <= 5:  # 显示所有错误
                        result_msg += f"\n错误详情（{error_count}项）：\n" + "\n".join(all_errors)
                    else:  # 只显示前5个错误
                        result_msg += f"\n错误详情（显示前5项，共{error_count}项）：\n" + "\n".join(all_errors[:5])
                        result_msg += f"\n...还有{error_count-5}项错误未显示"
                QMessageBox.warning(self, "导入失败", result_msg)
        except ImportError:
            QMessageBox.warning(self, "缺少依赖",
                                "请先安装pandas库：pip install pandas openpyxl")
        except Exception as e:
            import traceback
            error_detail = f"导入过程中发生未预期的错误：{str(e)}\n{traceback.format_exc()}"
            QMessageBox.critical(self, "系统错误", error_detail)
    def get_disease_order_for_import(self, structure_name, component_name, disease_type, warnings_list=None):
        """获取病害类型在配置中的排序号（用于导入时）"""
        try:
            if structure_name not in self.bridge_data:
                error_msg = f"结构系统'{structure_name}'在配置文件中不存在，无法获取病害排序号"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            structure_data = self.bridge_data[structure_name]
            if component_name not in structure_data:
                error_msg = f"部件'{component_name}'在结构系统'{structure_name}'中不存在，无法获取病害排序号"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            diseases = structure_data[component_name].get('diseases', [])
            if not diseases:
                error_msg = f"部件'{component_name}'没有配置病害类型列表"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            for index, disease in enumerate(diseases):
                if disease.get('name') == disease_type:
                    return index + 1  # 排序号从1开始
            # 如果没有找到对应的病害类型
            available_diseases = [d.get('name', '未知') for d in diseases]
            error_msg = f"病害类型'{disease_type}'在部件'{component_name}'中不存在，可用类型：{', '.join(available_diseases)}"
            if warnings_list is not None:
                warnings_list.append(error_msg)
            raise ValueError(error_msg)
        except ValueError:
            # 重新抛出已知的业务逻辑错误
            raise
        except Exception as e:
            error_msg = f"获取病害排序号时发生未知错误：{str(e)}"
            if warnings_list is not None:
                warnings_list.append(error_msg)
            raise RuntimeError(error_msg) from e
    def get_disease_max_scale_for_import(self, structure_name, component_name, disease_type, warnings_list=None):
        """获取病害类型的最大标度（用于导入时）"""
        try:
            if structure_name not in self.bridge_data:
                error_msg = f"结构系统'{structure_name}'在配置文件中不存在，无法获取最大标度"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            structure_data = self.bridge_data[structure_name]
            if component_name not in structure_data:
                error_msg = f"部件'{component_name}'在结构系统'{structure_name}'中不存在，无法获取最大标度"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            diseases = structure_data[component_name].get('diseases', [])
            if not diseases:
                error_msg = f"部件'{component_name}'没有配置病害类型列表"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            for disease in diseases:
                if disease.get('name') == disease_type:
                    max_scale = disease.get('max_scale')
                    if max_scale is None:
                        error_msg = f"病害类型'{disease_type}'缺少最大标度配置"
                        if warnings_list is not None:
                            warnings_list.append(error_msg)
                        raise ValueError(error_msg)
                    if not isinstance(max_scale, int) or max_scale < 1:
                        error_msg = f"病害类型'{disease_type}'的最大标度配置无效：{max_scale}"
                        if warnings_list is not None:
                            warnings_list.append(error_msg)
                        raise ValueError(error_msg)
                    return max_scale
            # 如果没有找到对应的病害类型
            available_diseases = [d.get('name', '未知') for d in diseases]
            error_msg = f"病害类型'{disease_type}'在部件'{component_name}'中不存在，可用类型：{', '.join(available_diseases)}"
            if warnings_list is not None:
                warnings_list.append(error_msg)
            raise ValueError(error_msg)
        except ValueError:
            # 重新抛出已知的业务逻辑错误
            raise
        except Exception as e:
            error_msg = f"获取最大标度时发生未知错误：{str(e)}"
            if warnings_list is not None:
                warnings_list.append(error_msg)
            raise RuntimeError(error_msg) from e
    def calculate_deduction_score_for_import(self, scale, max_scale, warnings_list=None):
        """根据评定标度和最大标度计算扣分值（用于导入时）"""
        try:
            # 参数验证
            if not isinstance(scale, int) or scale < 1:
                error_msg = f"评定标度无效：{scale}（必须是大于0的整数）"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            if not isinstance(max_scale, int) or max_scale < 1:
                error_msg = f"最大标度无效：{max_scale}（必须是大于0的整数）"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            if scale > max_scale:
                error_msg = f"评定标度{scale}超出最大标度{max_scale}"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            
            # 根据最大标度选择对应的扣分映射表（修正后的映射表）
            if max_scale == 3:
                # 3类最高等级：1类-0分，2类-20分，3类-35分
                score_map = {1: 0, 2: 20, 3: 35}
            elif max_scale == 4:
                # 4类最高等级：1类-0分，2类-25分，3类-40分，4类-50分
                score_map = {1: 0, 2: 25, 3: 40, 4: 50}
            elif max_scale == 5:
                # 5类最高等级：1类-0分，2类-35分，3类-45分，4类-60分，5类-100分
                score_map = {1: 0, 2: 35, 3: 45, 4: 60, 5: 100}
            else:
                # **修改：对于其他最大标度，直接抛出错误而不是继续处理**
                error_msg = f"不支持的最大标度：{max_scale}。当前仅支持3、4、5级标度系统"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            
            deduction_score = score_map.get(scale)
            if deduction_score is None:
                error_msg = f"标度{scale}在最大标度{max_scale}的映射表中不存在"
                if warnings_list is not None:
                    warnings_list.append(error_msg)
                raise ValueError(error_msg)
            
            return deduction_score
            
        except ValueError:
            # 重新抛出已知的业务逻辑错误
            raise
        except Exception as e:
            error_msg = f"计算扣分值时发生未知错误：{str(e)}"
            if warnings_list is not None:
                warnings_list.append(error_msg)
            raise RuntimeError(error_msg) from e
    def process_structure_data(self, df_all, structure_name):
        """处理单个结构系统的数据"""
        try:
            # 初始化结果收集器
            warnings = []
            errors = []
            # === 第一步：验证结构系统配置 ===
            if structure_name not in self.bridge_data:
                errors.append(f"结构系统'{structure_name}'在配置文件中不存在")
                return {
                    'status': 'error',
                    'component_count': 0,
                    'damage_count': 0,
                    'errors': errors,
                    'warnings': warnings,
                    'message': f'结构系统验证失败：{"; ".join(errors)}'
                }
            structure_data = self.bridge_data[structure_name]
            if not structure_data:
                errors.append(f"结构系统'{structure_name}'的配置数据为空")
                return {
                    'status': 'error',
                    'component_count': 0,
                    'damage_count': 0,
                    'errors': errors,
                    'warnings': warnings,
                    'message': f'结构系统配置错误：{"; ".join(errors)}'
                }
            # === 第二步：读取部件存在性确认部分 ===
            component_existence = {}
            # 查找"一、XX部件存在性确认"标题行
            existence_start_row = None
            for i, row in df_all.iterrows():
                if pd.notna(row.iloc[0]) and "部件存在性确认" in str(row.iloc[0]):
                    existence_start_row = i
                    break
            if existence_start_row is None:
                warnings.append(f"未找到'{structure_name}'的部件存在性确认部分，将使用默认设置")
            else:
                # 查找表头行（部件名称、是否存在、构件总数量、备注）
                header_row = None
                for i in range(existence_start_row + 1, min(existence_start_row + 10, len(df_all))):
                    row = df_all.iloc[i]
                    if (pd.notna(row.iloc[0]) and "部件名称" in str(row.iloc[0]) and
                            pd.notna(row.iloc[1]) and "是否存在" in str(row.iloc[1])):
                        header_row = i
                        break
                if header_row is None:
                    warnings.append(f"'{structure_name}'部件存在性确认部分缺少正确的表头，将使用默认设置")
                else:
                    # 读取部件存在性数据
                    processed_components = set()
                    for i in range(header_row + 1, len(df_all)):
                        row = df_all.iloc[i]
                        # 如果遇到空行或下一部分标题，停止读取
                        if (pd.isna(row.iloc[0]) or
                                (pd.notna(row.iloc[0]) and "病害数据录入" in str(row.iloc[0]))):
                            break
                        component_name = str(row.iloc[0]).strip()
                        existence_status = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "是"
                        quantity_str = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "0"
                        if not component_name:
                            continue
                        # 检查部件是否在配置中存在
                        if component_name not in structure_data:
                            warnings.append(f"部件'{component_name}'在配置文件中不存在，已跳过")
                            continue
                        # 检查重复定义
                        if component_name in processed_components:
                            warnings.append(f"部件'{component_name}'重复定义，使用最后一次的值")
                        processed_components.add(component_name)
                        # 判断是否存在（更宽松的判断条件）
                        is_existing = existence_status.lower() in ["是", "存在", "true", "1", "yes", "y"]
                        component_existence[component_name] = is_existing
                        # 解析构件总数量
                        try:
                            if "请填写" in quantity_str or quantity_str == "" or quantity_str.lower() in ["nan",
                                                                                                          "null"]:
                                quantity = 0
                            else:
                                quantity = int(float(quantity_str))
                                if quantity < 0:
                                    warnings.append(f"部件'{component_name}'的构件数量为负数({quantity})，已设为0")
                                    quantity = 0
                        except (ValueError, TypeError):
                            warnings.append(f"部件'{component_name}'的构件数量'{quantity_str}'格式错误，已设为0")
                            quantity = 0
                        # 使用"结构系统-部件名称"作为键存储构件数量
                        quantity_key = f"{structure_name}-{component_name}"
                        self.component_quantities[quantity_key] = quantity
                        # 添加print语句输出构件数量
                        print(f"存储构件数量 - {quantity_key}: {quantity}")
                    # 检查配置文件中的部件是否都有对应的存在性确认
                    missing_components = set(structure_data.keys()) - processed_components
                    if missing_components:
                        warnings.append(
                            f"以下部件在配置中存在但未在Excel中找到：{', '.join(sorted(missing_components))}")
            # === 第三步：更新UI界面 ===
            table = self.structure_tables.get(structure_name)
            if table:
                component_names = list(structure_data.keys())
                for row in range(table.rowCount()):
                    if row < len(component_names):
                        component_name = component_names[row]
                        # 更新复选框状态
                        checkbox_widget = table.cellWidget(row, 0)
                        if checkbox_widget:
                            checkbox = checkbox_widget.findChild(QCheckBox)
                            if checkbox:
                                if component_name in component_existence:
                                    checkbox.setChecked(component_existence[component_name])
                                else:
                                    checkbox.setChecked(False)
                        # 更新构件总数量
                        quantity_widget = table.cellWidget(row, 2)
                        if quantity_widget and isinstance(quantity_widget, QLineEdit):
                            quantity_key = f"{structure_name}-{component_name}"
                            if quantity_key in self.component_quantities:
                                quantity_widget.setText(str(self.component_quantities[quantity_key]))
                            else:
                                quantity_widget.setText("0")
                # 触发权重更新
                try:
                    self.update_structure_weights(structure_name)
                except Exception as e:
                    warnings.append(f"更新权重时发生错误：{str(e)}")
            # === 第四步：读取病害数据录入部分 ===
            damage_start_row = None
            # 查找"二、病害数据录入"标题行
            for i, row in df_all.iterrows():
                if pd.notna(row.iloc[0]) and "病害数据录入" in str(row.iloc[0]):
                    damage_start_row = i
                    break
            valid_damage_count = 0
            if damage_start_row is None:
                warnings.append(f"未找到'{structure_name}'的病害数据录入部分")
            else:
                # 查找病害数据表头行
                damage_header_row = None
                for i in range(damage_start_row + 1, min(damage_start_row + 10, len(df_all))):
                    row = df_all.iloc[i]
                    if (pd.notna(row.iloc[0]) and "部件名称" in str(row.iloc[0]) and
                            pd.notna(row.iloc[3]) and "病害类型" in str(row.iloc[3])):
                        damage_header_row = i
                        break
                if damage_header_row is None:
                    warnings.append(f"'{structure_name}'病害数据录入部分缺少正确的表头")
                else:
                    # 创建病害数据列表
                    damage_data_rows = []
                    for i in range(damage_header_row + 1, len(df_all)):
                        row = df_all.iloc[i]
                        # 如果遇到空行，跳过
                        if pd.isna(row.iloc[0]) and pd.isna(row.iloc[3]):
                            continue
                        # 如果遇到其他部分标题，停止读取
                        if (pd.notna(row.iloc[0]) and
                                ("各部件病害类型说明" in str(row.iloc[0]) or "三、" in str(row.iloc[0]))):
                            break
                        # 提取病害数据
                        component_name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        component_code = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                        component_detail_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                        disease_type = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
                        scale_str = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
                        damage_description = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) and len(row) > 5 else ""  # 新增病害描述读取
                        # 基本数据完整性检查
                        if not component_name:
                            continue
                        if not disease_type:
                            warnings.append(f"数据行{i + 1}：部件'{component_name}'缺少病害类型")
                            continue
                        if not scale_str:
                            warnings.append(f"数据行{i + 1}：部件'{component_name}'的病害'{disease_type}'缺少评定标度")
                            continue
                        # 验证评定标度格式
                        try:
                            scale = int(float(scale_str))
                            if scale < 1:
                                warnings.append(f"数据行{i + 1}：评定标度{scale}无效（必须≥1），已跳过")
                                continue
                        except (ValueError, TypeError):
                            warnings.append(f"数据行{i + 1}：评定标度'{scale_str}'格式错误，已跳过")
                            continue
                        damage_data_rows.append({
                            "行号": i + 1,
                            "部件名称": component_name,
                            "构件编号": component_code,
                            "构件名称": component_detail_name,
                            "病害类型": disease_type,
                            "评定标度": scale,
                            "病害描述": damage_description  # 新增病害描述字段
                        })
                    if damage_data_rows:
                        # 验证病害数据
                        valid_data = []
                        for row_data in damage_data_rows:
                            row_num = row_data["行号"]
                            component_name = row_data["部件名称"]
                            disease_type = row_data["病害类型"]
                            scale = row_data["评定标度"]
                            # 检查部件是否存在
                            if component_name not in structure_data:
                                warnings.append(f"数据行{row_num}：部件'{component_name}'在配置中不存在")
                                continue
                            # 检查部件是否被标记为存在
                            if component_name not in component_existence:
                                warnings.append(f"数据行{row_num}：部件'{component_name}'未在存在性确认中定义")
                                continue
                            if not component_existence[component_name]:
                                warnings.append(f"数据行{row_num}：部件'{component_name}'在存在性确认中标记为不存在")
                                continue
                            # 检查病害类型是否存在
                            component_diseases = [d['name'] for d in structure_data[component_name].get('diseases', [])]
                            if disease_type not in component_diseases:
                                available_diseases = ', '.join(component_diseases) if component_diseases else '无'
                                warnings.append(
                                    f"数据行{row_num}：部件'{component_name}'中不存在病害类型'{disease_type}'，可用类型：{available_diseases}")
                                continue
                            # 检查评定标度范围
                            max_scale = None
                            for disease in structure_data[component_name].get('diseases', []):
                                if disease['name'] == disease_type:
                                    max_scale = disease.get('max_scale')
                                    break
                            if max_scale is None:
                                warnings.append(f"数据行{row_num}：病害类型'{disease_type}'缺少最大标度配置")
                                continue
                            if scale < 1 or scale > max_scale:
                                warnings.append(
                                    f"数据行{row_num}：病害类型'{disease_type}'的评定标度{scale}超出有效范围(1-{max_scale})")
                                continue
                            valid_data.append(row_data)
                        # 处理有效的病害数据
                        if valid_data:
                            # 清空该结构系统的现有病害数据
                            structure_components = list(structure_data.keys())
                            cleared_count = 0
                            for comp_name in structure_components:
                                damage_key = f"{structure_name}-{comp_name}"
                                if damage_key in self.component_damage_data:
                                    del self.component_damage_data[damage_key]
                                    cleared_count += 1
                            if cleared_count > 0:
                                warnings.append(f"已清空{cleared_count}个部件的原有病害数据")
                            # 按部件分组存储数据（包含排序号和扣分值）
                            component_damage_counts = {}
                            for row_data in valid_data:
                                component_name = row_data["部件名称"]
                                disease_type = row_data["病害类型"]
                                scale = row_data["评定标度"]
                                damage_key = f"{structure_name}-{component_name}"
                                if damage_key not in self.component_damage_data:
                                    self.component_damage_data[damage_key] = []
                                try:
                                    # 获取病害类型排序号
                                    disease_order = self.get_disease_order_for_import(
                                        structure_name, component_name, disease_type, warnings
                                    )
                                    # 获取最大标度并计算扣分值
                                    max_scale = self.get_disease_max_scale_for_import(
                                        structure_name, component_name, disease_type, warnings
                                    )
                                    deduction_score = self.calculate_deduction_score_for_import(
                                        scale, max_scale, warnings
                                    )
                                    # 存储完整的病害信息
                                    damage_info = {
                                        "序号": len(self.component_damage_data[damage_key]) + 1,
                                        "构件编号": row_data["构件编号"],
                                        "构件名称": row_data["构件名称"],
                                        "病害类型": disease_type,
                                        "病害排序号": disease_order,
                                        "评定标度": str(scale),
                                        "病害描述": row_data.get("病害描述", ""),  # 更安全的写法
                                        "扣分值": deduction_score
                                    }
                                    self.component_damage_data[damage_key].append(damage_info)
                                    # 打印存储的病害数据
                                    print(f"存储病害数据 - {structure_name}-{component_name}: {damage_info}")
                                    # 统计每个部件的病害数量
                                    component_damage_counts[component_name] = component_damage_counts.get(
                                        component_name, 0) + 1
                                except (ValueError, RuntimeError) as e:
                                    # 记录具体的错误信息，但继续处理其他数据
                                    row_num = row_data.get("行号", "未知")
                                    warnings.append(f"数据行{row_num}处理失败：{str(e)}")
                                    continue
                            valid_damage_count = len(valid_data)
                            # 生成详细的导入统计信息
                            if component_damage_counts:
                                damage_summary = []
                                for comp_name, count in sorted(component_damage_counts.items()):
                                    damage_summary.append(f"{comp_name}({count}条)")
                                warnings.append(f"病害数据分布：{', '.join(damage_summary)}")
            # === 第五步：生成最终结果 ===
            existing_count = sum(1 for exists in component_existence.values() if exists)
            total_components = len(component_existence) if component_existence else 0
            # 计算该结构系统的总构件数量
            structure_total_quantity = 0
            for component_name in component_existence.keys():
                quantity_key = f"{structure_name}-{component_name}"
                if quantity_key in self.component_quantities:
                    structure_total_quantity += self.component_quantities[quantity_key]
            # 构建结果消息
            result_parts = []
            if existing_count > 0:
                result_parts.append(f"成功处理{existing_count}个部件")
            if valid_damage_count > 0:
                result_parts.append(f"{valid_damage_count}条病害记录")
            if structure_total_quantity > 0:
                result_parts.append(f"总构件数量{structure_total_quantity}")
            message = ', '.join(result_parts) if result_parts else '无有效数据'
            return {
                'status': 'success' if not errors else 'error',
                'component_count': existing_count,
                'damage_count': valid_damage_count,
                'total_components': total_components,
                'total_quantity': structure_total_quantity,  # 使用structure_total_quantity
                'errors': errors,
                'warnings': warnings,
                'message': message
            }
        except Exception as e:
            import traceback
            error_detail = f"处理{structure_name}数据时发生异常：{str(e)}\n{traceback.format_exc()}"
            return {
                'status': 'error',
                'component_count': 0,
                'damage_count': 0,
                'total_components': 0,
                'total_quantity': 0,  # 保持一致
                'errors': [error_detail],
                'warnings': [],
                'message': f'系统异常：{str(e)}'
            }
    def _create_worksheet_for_structure(self, wb, structure_name, structure_data):
        """为指定结构系统创建工作表的通用函数"""
        from openpyxl.styles import Font, Alignment, Border, Side
        # 创建新的工作表
        ws = wb.create_sheet(title=f'{structure_name}病害数据')
        # 设置字体（仿宋）
        fangsong_font = Font(name='仿宋', size=12)
        header_font = Font(name='仿宋', size=12, bold=True)
        title_font = Font(name='仿宋', size=14, bold=True)
        # 设置对齐方式（水平垂直居中）
        center_alignment = Alignment(horizontal='center', vertical='center')
        # 设置边框样式
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        current_row = 1
        # === 第一部分：部件存在性确认表 ===
        ws.cell(row=current_row, column=1, value=f"一、{structure_name}部件存在性确认")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        ws.cell(row=current_row, column=1).font = title_font
        ws.cell(row=current_row, column=1).alignment = center_alignment
        current_row += 1
        # 部件存在性确认表头
        existence_headers = ["部件名称", "是否存在", "构件总数量"]
        for i, header in enumerate(existence_headers):
            ws.cell(row=current_row, column=i + 1, value=header)
            ws.cell(row=current_row, column=i + 1).font = header_font
            ws.cell(row=current_row, column=i + 1).alignment = center_alignment
            ws.cell(row=current_row, column=i + 1).border = thin_border
        current_row += 1
        # 添加部件存在性确认数据
        for component_name in structure_data.keys():
            if isinstance(structure_data[component_name], dict):
                # 部件名称
                ws.cell(row=current_row, column=1, value=component_name)
                ws.cell(row=current_row, column=1).font = fangsong_font
                ws.cell(row=current_row, column=1).alignment = center_alignment
                ws.cell(row=current_row, column=1).border = thin_border
                # 是否存在（下拉选择：是/否）
                ws.cell(row=current_row, column=2, value="是")
                ws.cell(row=current_row, column=2).font = fangsong_font
                ws.cell(row=current_row, column=2).alignment = center_alignment
                ws.cell(row=current_row, column=2).border = thin_border
                # 构件总数量（待填写）
                ws.cell(row=current_row, column=3, value="请填写数量")
                ws.cell(row=current_row, column=3).font = fangsong_font
                ws.cell(row=current_row, column=3).alignment = center_alignment
                ws.cell(row=current_row, column=3).border = thin_border
                current_row += 1
        # 设置部件确认表的列宽
        ws.column_dimensions['A'].width = 15  # 部件名称
        ws.column_dimensions['B'].width = 12  # 是否存在
        ws.column_dimensions['C'].width = 15  # 构件总数量
        current_row += 1
        # === 第二部分：病害数据录入表 ===
        ws.cell(row=current_row, column=1, value="二、病害数据录入")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)  # 从5改为6
        ws.cell(row=current_row, column=1).font = title_font
        ws.cell(row=current_row, column=1).alignment = center_alignment
        current_row += 1
        # 病害数据录入表头
        damage_headers = ["部件名称", "构件编号", "构件名称", "病害类型", "评定标度", "病害描述"]  # 添加病害描述列
        for i, header in enumerate(damage_headers):
            ws.cell(row=current_row, column=i + 1, value=header)
            ws.cell(row=current_row, column=i + 1).font = header_font
            ws.cell(row=current_row, column=i + 1).alignment = center_alignment
            ws.cell(row=current_row, column=i + 1).border = thin_border
        current_row += 1
        # 创建病害录入示例数据
        for component_name, component_data in structure_data.items():
            if isinstance(component_data, dict):
                diseases = component_data.get('diseases', [])
                # 为每个病害类型创建一行示例数据
                for i, disease in enumerate(diseases[:2]):  # 每个部件只显示2个病害示例
                    # 部件名称
                    ws.cell(row=current_row, column=1, value=component_name)
                    ws.cell(row=current_row, column=1).font = fangsong_font
                    ws.cell(row=current_row, column=1).alignment = center_alignment
                    ws.cell(row=current_row, column=1).border = thin_border
                    # 构件编号（保持空白）
                    ws.cell(row=current_row, column=2, value="")
                    ws.cell(row=current_row, column=2).font = fangsong_font
                    ws.cell(row=current_row, column=2).alignment = center_alignment
                    ws.cell(row=current_row, column=2).border = thin_border
                    # 构件名称（保持空白）
                    ws.cell(row=current_row, column=3, value="")
                    ws.cell(row=current_row, column=3).font = fangsong_font
                    ws.cell(row=current_row, column=3).alignment = center_alignment
                    ws.cell(row=current_row, column=3).border = thin_border
                    # 病害类型
                    ws.cell(row=current_row, column=4, value=disease['name'])
                    ws.cell(row=current_row, column=4).font = fangsong_font
                    ws.cell(row=current_row, column=4).alignment = center_alignment
                    ws.cell(row=current_row, column=4).border = thin_border
                    # 评定标度
                    ws.cell(row=current_row, column=5, value="1")
                    ws.cell(row=current_row, column=5).font = fangsong_font
                    ws.cell(row=current_row, column=5).alignment = center_alignment
                    ws.cell(row=current_row, column=5).border = thin_border
                    # 病害描述（新增）
                    ws.cell(row=current_row, column=6, value="")
                    ws.cell(row=current_row, column=6).font = fangsong_font
                    ws.cell(row=current_row, column=6).alignment = center_alignment
                    ws.cell(row=current_row, column=6).border = thin_border
                    current_row += 1
        # 设置病害录入表的列宽
        ws.column_dimensions['A'].width = 15  # 部件名称
        ws.column_dimensions['B'].width = 12  # 构件编号
        ws.column_dimensions['C'].width = 15  # 构件名称
        ws.column_dimensions['D'].width = 25  # 病害类型
        ws.column_dimensions['E'].width = 12  # 评定标度
        ws.column_dimensions['F'].width = 30  # 病害描述（新增）
        # === 第三部分：在右侧添加病害类型说明 ===
        start_col = 8  # H列
        explanation_row = 1
        # 添加说明标题（跨四列）
        ws.cell(row=explanation_row, column=start_col, value="三、各部件病害类型说明")
        ws.merge_cells(start_row=explanation_row, start_column=start_col,
                       end_row=explanation_row, end_column=start_col + 3)
        ws.cell(row=explanation_row, column=start_col).font = title_font
        ws.cell(row=explanation_row, column=start_col).alignment = center_alignment
        explanation_row += 1
        # 为每个部件添加病害类型说明
        for component_name, component_data in structure_data.items():
            if isinstance(component_data, dict):
                diseases = component_data.get('diseases', [])
                # 部件名称标题（跨四列）
                ws.cell(row=explanation_row, column=start_col, value=f"【{component_name}】")
                ws.merge_cells(start_row=explanation_row, start_column=start_col,
                               end_row=explanation_row, end_column=start_col + 3)
                ws.cell(row=explanation_row, column=start_col).font = Font(name='仿宋', size=12, bold=True)
                ws.cell(row=explanation_row, column=start_col).alignment = center_alignment
                explanation_row += 1
                # 添加表头
                headers = ["序号", "病害类型名称", "最大标度", "备注"]
                for i, header in enumerate(headers):
                    ws.cell(row=explanation_row, column=start_col + i, value=header)
                    ws.cell(row=explanation_row, column=start_col + i).font = Font(name='仿宋', size=11, bold=True)
                    ws.cell(row=explanation_row, column=start_col + i).alignment = center_alignment
                    ws.cell(row=explanation_row, column=start_col + i).border = thin_border
                explanation_row += 1
                # 添加病害类型数据
                for i, disease in enumerate(diseases):
                    # 序号
                    ws.cell(row=explanation_row, column=start_col, value=i + 1)
                    ws.cell(row=explanation_row, column=start_col).font = fangsong_font
                    ws.cell(row=explanation_row, column=start_col).alignment = center_alignment
                    ws.cell(row=explanation_row, column=start_col).border = thin_border
                    # 病害类型名称
                    ws.cell(row=explanation_row, column=start_col + 1, value=disease['name'])
                    ws.cell(row=explanation_row, column=start_col + 1).font = fangsong_font
                    ws.cell(row=explanation_row, column=start_col + 1).alignment = center_alignment
                    ws.cell(row=explanation_row, column=start_col + 1).border = thin_border
                    # 最大标度
                    ws.cell(row=explanation_row, column=start_col + 2, value=disease.get('max_scale', 'N/A'))
                    ws.cell(row=explanation_row, column=start_col + 2).font = fangsong_font
                    ws.cell(row=explanation_row, column=start_col + 2).alignment = center_alignment
                    ws.cell(row=explanation_row, column=start_col + 2).border = thin_border
                    # 备注说明
                    remark = ""
                    # 为桥面铺装添加特殊备注
                    if "桥面铺装" in component_name:
                        if i + 1 <= 4:  # 序号1-4
                            remark = "适用于沥青混凝土桥面铺装"
                        elif i + 1 >= 5 and i + 1 <= 11:  # 序号5-11
                            remark = "适用于混凝土桥面铺装"
                    if "支座" in component_name:
                        if i + 1 <= 3:  # 序号1-3
                            remark = "适用于板式支座"
                        elif i + 1 == 4 or i + 1 == 6:  # 序号4或6
                            remark = "适用于盆式支座"
                        elif i + 1 >= 7 and i + 1 <= 9:  # 序号7-9
                            remark = "适用于钢支座"
                        elif i + 1 == 10:  # 序号10
                            remark = "适用于摆式支座"
                        elif i + 1 == 11:  # 序号11
                            remark = "适用于活动支座"
                    ws.cell(row=explanation_row, column=start_col + 3, value=remark)
                    ws.cell(row=explanation_row, column=start_col + 3).font = fangsong_font
                    ws.cell(row=explanation_row, column=start_col + 3).alignment = center_alignment
                    ws.cell(row=explanation_row, column=start_col + 3).border = thin_border
                    explanation_row += 1
        # 设置说明区域的列宽
        ws.column_dimensions['H'].width = 8  # 序号列
        ws.column_dimensions['I'].width = 30  # 病害类型名称列
        ws.column_dimensions['J'].width = 12  # 最大标度列
        ws.column_dimensions['K'].width = 30  # 备注说明列
        return ws
    def export_all_bridge_template(self):
        """导出全桥病害数据Excel模板（包含三个工作簿）"""
        try:
            import pandas as pd
            from PySide6.QtWidgets import QFileDialog
            import subprocess
            import platform
            # 检查Import_data文件夹是否存在，不存在则创建
            import_folder = "Import_data/Steel-concrete_composite_arch_bridge"
            if not os.path.exists(import_folder):
                os.makedirs(import_folder)
            # 选择保存位置
            default_filename = os.path.join(import_folder, "全桥病害数据模板.xlsx")
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存Excel模板文件",
                default_filename,
                "Excel文件 (*.xlsx)"
            )
            if file_path:
                # 保存Excel文件并设置格式
                from openpyxl import Workbook
                # 创建工作簿
                wb = Workbook()
                # 删除默认工作表
                wb.remove(wb.active)
                # 创建三个工作表
                structures = [
                    ("桥面系", self.bridge_data.get('桥面系', {})),
                    ("上部结构", self.bridge_data.get('上部结构', {})),
                    ("下部结构", self.bridge_data.get('下部结构', {}))
                ]
                for structure_name, structure_data in structures:
                    if structure_data:  # 只有当数据存在时才创建工作表
                        self._create_worksheet_for_structure(wb, structure_name, structure_data)
                # 保存文件
                wb.save(file_path)
                # 询问是否立即打开文件
                reply = QMessageBox.question(
                    self,
                    "导出成功",
                    f"全桥病害数据模板已成功导出到：\n{file_path}\n\n模板包含：\n1. 桥面系病害数据工作表\n2. 上部结构病害数据工作表\n3. 下部结构病害数据工作表\n\n每个工作表都包含：\n• 部件存在性确认表\n• 病害数据录入表\n• 病害类型说明表\n\n是否立即打开文件？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes:
                    try:
                        # 根据操作系统选择打开方式
                        if platform.system() == 'Windows':
                            os.startfile(file_path)
                        elif platform.system() == 'Darwin':  # macOS
                            subprocess.run(['open', file_path])
                        else:  # Linux
                            subprocess.run(['xdg-open', file_path])
                    except Exception as e:
                        QMessageBox.warning(self, "打开失败", f"无法打开文件：{str(e)}")
                else:
                    QMessageBox.information(self, "提示", "模板已保存，您可以稍后手动打开进行编辑。")
        except ImportError:
            QMessageBox.warning(self, "缺少依赖",
                                "请先安装pandas和openpyxl库：\npip install pandas openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中发生错误：{str(e)}")
    def print_matrix_preview(self, matrix, matrix_name, rows, cols, description=""):
        """打印矩阵预览，显示前20行和12列
        Args:
            matrix (np.ndarray): 要打印的矩阵
            matrix_name (str): 矩阵名称（如"A"、"B"、"C"等）
            rows (int): 矩阵总行数
            cols (int): 矩阵总列数
            description (str, optional): 矩阵的额外描述信息
        """
        # 确定描述文本
        if description:
            print(f"\n矩阵{matrix_name}的内容（{description}）：")
        else:
            print(f"\n矩阵{matrix_name}的内容：")
        # 设置显示的最大行列数
        max_display_rows = 20
        max_display_cols = 12
        display_rows = min(max_display_rows, rows)
        display_cols = min(max_display_cols, cols)
        # 打印前20行和12列
        for i in range(display_rows):
            # 获取当前行的前12列
            row_preview = matrix[i][:display_cols]
            # 格式化打印
            if i == 0:
                print(f"[[{' '.join([f'{x}' for x in row_preview])}]]")
            elif i == display_rows - 1 and rows > display_rows:
                print(f"[{' '.join([f'{x}' for x in row_preview])}]]")
            else:
                print(f" [{' '.join([f'{x}' for x in row_preview])}]")
        # 如果行数超过显示行数，添加省略号
        if rows > display_rows:
            print("...")
    # ----------矩阵运行逻辑----------
    def start_matrix_calculation(self, auto_confirm=False):
        """开始矩阵计算"""
        # 先显示矩阵维度设置对话框
        dimension_dialog = MatrixDimensionDialog(self, auto_confirm=auto_confirm)
        if dimension_dialog.exec() == QDialog.Accepted:
            # 获取用户设置的矩阵维度
            matrix_rows, matrix_cols = dimension_dialog.get_dimensions()
            # 开始矩阵运算逻辑
            self.perform_matrix_calculation(matrix_rows, matrix_cols)
    def perform_matrix_calculation(self, matrix_rows, matrix_cols):
        """执行矩阵运算
        对所有存在的部件分别执行矩阵计算，从矩阵A到矩阵F，最终得到各部件的技术状况评分值
        Args:
            matrix_rows (int): 矩阵行数
            matrix_cols (int): 矩阵列数
        """
        try:
            print(f"\n=== 开始矩阵运算 ===")
            print(f"矩阵维度: {matrix_rows}x{matrix_cols}")
            # 初始化计算过程数据收集器
            self.calculation_process_data = {
                'basic_info': {
                    'matrix_dimensions': f"{matrix_rows}x{matrix_cols}",
                    'start_time': None,
                    'end_time': None,
                    'total_components': 0,
                    'calculated_components': 0,
                    'skipped_components': 0
                },
                'component_details': {},
                'structure_systems': {},
                'calculation_summary': {}
            }
            import datetime
            self.calculation_process_data['basic_info']['start_time'] = datetime.datetime.now().strftime(
                '%Y-%m-%d %H:%M:%S')
            # 存储所有部件的计算结果
            component_results = {}
            # 遍历所有结构系统和部件
            for structure_name in self.structure_tables.keys():
                print(f"\n--- 处理结构系统: {structure_name} ---")
                # 初始化结构系统数据
                self.calculation_process_data['structure_systems'][structure_name] = {
                    'components': [],
                    'total_components': 0,
                    'calculated_components': 0,
                    'skipped_components': 0
                }
                # 获取该结构系统的表格
                table = self.structure_tables[structure_name]
                # 遍历表格中的所有行，获取部件信息
                for row in range(table.rowCount()):
                    # 获取第0列的复合控件（包含复选框和部件名称）
                    checkbox_widget = table.cellWidget(row, 0)
                    if not checkbox_widget:
                        continue
                    # 从复合控件中找到复选框和标签
                    checkbox = None
                    name_label = None
                    # 遍历复合控件的子控件
                    for child in checkbox_widget.findChildren(QCheckBox):
                        checkbox = child
                        break
                    for child in checkbox_widget.findChildren(QLabel):
                        name_label = child
                        break
                    if not checkbox or not name_label:
                        continue
                    component_name = name_label.text()
                    if not component_name:
                        continue
                    # 统计总部件数
                    self.calculation_process_data['basic_info']['total_components'] += 1
                    self.calculation_process_data['structure_systems'][structure_name]['total_components'] += 1
                    # 检查部件是否存在（复选框是否勾选）
                    if not checkbox.isChecked():
                        print(f"  跳过部件 {component_name}：未勾选")
                        self.calculation_process_data['basic_info']['skipped_components'] += 1
                        self.calculation_process_data['structure_systems'][structure_name]['skipped_components'] += 1
                        continue
                    # 构造唯一键
                    component_key = f"{structure_name}-{component_name}"
                    # 检查是否有构件数量数据
                    if component_key not in self.component_quantities:
                        print(f"  跳过部件 {component_name}：无构件数量数据")
                        self.calculation_process_data['basic_info']['skipped_components'] += 1
                        self.calculation_process_data['structure_systems'][structure_name]['skipped_components'] += 1
                        continue
                    component_quantity = self.component_quantities[component_key]
                    if component_quantity <= 0:
                        print(f"  跳过部件 {component_name}：构件数量为0")
                        self.calculation_process_data['basic_info']['skipped_components'] += 1
                        self.calculation_process_data['structure_systems'][structure_name]['skipped_components'] += 1
                        continue
                    print(f"\n  开始计算部件: {component_name} (构件数量: {component_quantity})")
                    # 初始化部件详细数据
                    component_detail = {
                        'structure_name': structure_name,
                        'component_name': component_name,
                        'component_quantity': component_quantity,
                        'calculation_type': '',
                        'has_damage_data': False,
                        'damage_data': [],
                        'matrix_operations': {},
                        'h_value': None,
                        'calculation_status': 'processing',
                        'error_message': None
                    }
                    # 检查是否有病害数据
                    has_damage_data = component_key in self.component_damage_data and len(
                        self.component_damage_data[component_key]) > 0
                    component_detail['has_damage_data'] = has_damage_data
                    if has_damage_data:
                        component_detail['damage_data'] = self.component_damage_data[component_key].copy()
                    if not has_damage_data:
                        # 没有病害数据，直接设置h值为100.00
                        h_value = 100.00
                        print(f"  部件 {component_name} 无病害数据，直接设置h值 = {h_value:.2f}")
                        component_detail['calculation_type'] = '无病害数据'
                        component_detail['h_value'] = h_value
                        component_detail['calculation_status'] = 'completed'
                        # 存储计算结果
                        component_results[component_key] = {
                            'structure_name': structure_name,
                            'component_name': component_name,
                            'component_quantity': component_quantity,
                            'h_value': h_value,
                            'matrix_F': None  # 无需矩阵计算
                        }
                        self.calculation_process_data['basic_info']['calculated_components'] += 1
                        self.calculation_process_data['structure_systems'][structure_name]['calculated_components'] += 1
                    else:
                        # 有病害数据，执行完整的矩阵计算流程
                        print(f"  部件 {component_name} 存在病害数据，开始矩阵计算...")
                        component_detail['calculation_type'] = '完整矩阵计算'
                        try:
                            # 创建初始矩阵A
                            print(f"    创建矩阵A...")
                            matrix_A = self.create_matrix_A(matrix_rows, matrix_cols)
                            component_detail['matrix_operations']['matrix_A'] = {
                                'description': '初始矩阵A（全部填充100）',
                                'dimensions': f"{matrix_rows}x{matrix_cols}",
                                'data': matrix_A.tolist() if hasattr(matrix_A, 'tolist') else str(matrix_A)
                            }
                            # 创建矩阵B（根据当前部件的病害数据填充）
                            print(f"    创建矩阵B（病害数据填充）...")
                            matrix_B, matrix_B_stats = self.create_matrix_B_for_component_with_stats(matrix_A, matrix_rows, matrix_cols, component_key)
                            component_detail['matrix_operations']['matrix_B'] = {
                                'description': '矩阵B（根据病害数据填充）',
                                'dimensions': f"{matrix_rows}x{matrix_cols}",
                                'filled_cells': matrix_B_stats.get('filled_cells', 0),
                                'processed_components': matrix_B_stats.get('processed_components', 0),
                                'fill_details': matrix_B_stats.get('fill_details', []),
                                'data': matrix_B.tolist() if hasattr(matrix_B, 'tolist') else str(matrix_B)
                            }
                            # 创建矩阵C（扣分值矩阵）
                            print(f"    创建矩阵C（扣分值矩阵）...")
                            matrix_C, matrix_C_stats = self.create_matrix_C_for_component_with_stats(matrix_B, matrix_rows, matrix_cols, component_key)
                            component_detail['matrix_operations']['matrix_C'] = {
                                'description': '矩阵C（扣分值矩阵）',
                                'dimensions': f"{matrix_rows}x{matrix_cols}",
                                'processed_cells': matrix_C_stats.get('processed_cells', 0),
                                'filled_cells': matrix_C_stats.get('filled_cells', 0),
                                'zero_cells': matrix_C_stats.get('zero_cells', 0),
                                'fill_details': matrix_C_stats.get('fill_details', []),
                                'data': matrix_C.tolist() if hasattr(matrix_C, 'tolist') else str(matrix_C)
                            }
                            # 创建矩阵D（排序扣分值矩阵）
                            print(f"    创建矩阵D（排序扣分值矩阵）...")
                            matrix_D, matrix_D_stats = self.create_matrix_D_with_stats(matrix_C, matrix_rows, matrix_cols)
                            component_detail['matrix_operations']['matrix_D'] = {
                                'description': '矩阵D（排序扣分值矩阵）',
                                'dimensions': f"{matrix_rows}x{matrix_cols}",
                                'sorted_rows': matrix_D_stats.get('sorted_rows', 0),
                                'total_non_zero_values': matrix_D_stats.get('total_non_zero_values', 0),
                                'sort_details': matrix_D_stats.get('sort_details', []),
                                'data': matrix_D.tolist() if hasattr(matrix_D, 'tolist') else str(matrix_D)
                            }
                            # 创建矩阵E（计算矩阵）
                            print(f"    创建矩阵E（计算矩阵）...")
                            matrix_E, matrix_E_stats = self.create_matrix_E_with_stats(matrix_D, matrix_rows, matrix_cols)
                            component_detail['matrix_operations']['matrix_E'] = {
                                'description': '矩阵E（计算矩阵）',
                                'dimensions': f"{matrix_rows}x{matrix_cols}",
                                'processed_rows': matrix_E_stats.get('processed_rows', 0),
                                'single_value_rows': matrix_E_stats.get('single_value_rows', 0),
                                'multi_value_rows': matrix_E_stats.get('multi_value_rows', 0),
                                'formatted_count': matrix_E_stats.get('formatted_count', 0),
                                'calculation_details': matrix_E_stats.get('calculation_details', []),
                                'format_details': matrix_E_stats.get('format_details', []),
                                'data': matrix_E.tolist() if hasattr(matrix_E, 'tolist') else str(matrix_E)
                            }
                            # 创建矩阵F（构件实际得分矩阵）
                            print(f"    创建矩阵F（构件实际得分矩阵）...")
                            matrix_F, matrix_F_stats = self.create_matrix_F_with_stats(matrix_E, matrix_rows)
                            component_detail['matrix_operations']['matrix_F'] = {
                                'description': '矩阵F（构件实际得分矩阵）',
                                'dimensions': f"{matrix_rows}x1",
                                'processed_components': matrix_F_stats.get('processed_components', 0),
                                'defective_components': matrix_F_stats.get('defective_components', 0),
                                'total_deduction': matrix_F_stats.get('total_deduction', 0),
                                'min_score': matrix_F_stats.get('min_score', 100),
                                'max_score': matrix_F_stats.get('max_score', 100),
                                'avg_score': matrix_F_stats.get('avg_score', 100),
                                'calculation_details': matrix_F_stats.get('calculation_details', []),
                                'data': matrix_F.tolist() if hasattr(matrix_F, 'tolist') else str(matrix_F)
                            }
                            # 计算h值（部件综合评定分数）
                            print(f"计算h值（部件综合评定分数）...")
                            h_value, h_value_stats = self.calculate_h_value_with_stats(matrix_F, structure_name, component_name)
                            if h_value is not None:
                                # 确保h值保留两位小数
                                h_value = round(h_value, 2)
                                component_detail['h_value'] = h_value
                                component_detail['calculation_status'] = 'completed'
                                # 存储h值计算的详细统计信息
                                component_detail['h_value_calculation'] = h_value_stats
                                # 存储计算结果
                                component_results[component_key] = {
                                    'structure_name': structure_name,
                                    'component_name': component_name,
                                    'component_quantity': component_quantity,
                                    'h_value': h_value,
                                    'matrix_F': matrix_F
                                }
                                print(f"  部件 {component_name} 计算完成，h值 = {h_value:.2f}")
                                self.calculation_process_data['basic_info']['calculated_components'] += 1
                                self.calculation_process_data['structure_systems'][structure_name][
                                    'calculated_components'] += 1
                            else:
                                print(f"部件{component_name}计算h值失败")
                                component_detail['calculation_status'] = 'failed'
                                component_detail['error_message'] = '计算h值失败'
                        except Exception as component_error:
                            error_msg = str(component_error)
                            print(f"  部件 {component_name} 计算过程中发生错误: {error_msg}")
                            component_detail['calculation_status'] = 'error'
                            component_detail['error_message'] = error_msg
                            continue
                    # 保存部件详细数据
                    self.calculation_process_data['component_details'][component_key] = component_detail
                    self.calculation_process_data['structure_systems'][structure_name]['components'].append(
                        component_key)
            # 存储所有计算结果
            self.component_calculation_results = component_results
            # 记录结束时间
            self.calculation_process_data['basic_info']['end_time'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            # 生成计算汇总
            self.calculation_process_data['calculation_summary'] = {
                'total_components': self.calculation_process_data['basic_info']['total_components'],
                'calculated_components': self.calculation_process_data['basic_info']['calculated_components'],
                'skipped_components': self.calculation_process_data['basic_info']['skipped_components'],
                'success_rate': round(self.calculation_process_data['basic_info']['calculated_components'] / max(1,self.calculation_process_data['basic_info']['total_components']) * 100,2),'structure_systems_summary': {}}
            for structure_name, structure_data in self.calculation_process_data['structure_systems'].items():
                self.calculation_process_data['calculation_summary']['structure_systems_summary'][structure_name] = {
                    'total': structure_data['total_components'],
                    'calculated': structure_data['calculated_components'],
                    'skipped': structure_data['skipped_components'],
                    'success_rate': round(
                        structure_data['calculated_components'] / max(1, structure_data['total_components']) * 100, 2)
                }
            # 打印最终结果汇总
            print(f"\n=== 矩阵运算完成 ===")
            print(f"成功计算的部件数量: {len(component_results)}")
            if component_results:
                print("\n各部件技术状况评分结果:")
                print("-" * 60)
                print(f"{'结构系统':<15} {'部件名称':<15} {'构件数量':<8} {'h值':<8}")
                print("-" * 60)
                for component_key, result in component_results.items():
                    structure_name = result['structure_name']
                    component_name = result['component_name']
                    quantity = result['component_quantity']
                    h_value = result['h_value']
                    print(f"{structure_name:<15} {component_name:<15} {quantity:<8} {h_value:<8.2f}")
                # 计算桥梁整体技术状况评分
                print("\n=== 开始计算桥梁整体技术状况评分 ===")
                self.calculate_bridge_technical_condition_score()
            else:
                print("未找到可计算的部件数据")
        except Exception as e:
            print(f"\n矩阵运算过程中发生错误: {str(e)}")
            # 记录错误信息到计算过程数据中
            if hasattr(self, 'calculation_process_data'):
                self.calculation_process_data['basic_info']['error'] = str(e)
                self.calculation_process_data['basic_info']['end_time'] = datetime.datetime.now().strftime(
                    '%Y-%m-%d %H:%M:%S')
    def create_matrix_A(self, rows, cols):
        """创建初始矩阵A，所有值为1
        Args:
            rows (int): 矩阵行数
            cols (int): 矩阵列数
        Returns:
            np.ndarray: 创建的矩阵A
        """
        # 使用NumPy创建整数矩阵，而非浮点数
        matrix_A = np.ones((rows, cols), dtype=np.int32)
        # 打印矩阵A的基本信息
        print(f"\n创建了矩阵A，维度: {rows}x{cols}")
        # 使用通用方法打印矩阵预览
        self.print_matrix_preview(matrix_A, "A", rows, cols)
        return matrix_A
    def create_matrix_B_for_component_with_stats(self, matrix_A, rows, cols, component_key):
        """为特定部件创建矩阵B，并返回详细统计信息"""
        # 复制矩阵A作为矩阵B的基础
        matrix_B = np.copy(matrix_A)
        print(f"\n    创建矩阵B for {component_key}，维度: {rows}x{cols}")
        print(f"    开始根据部件 {component_key} 的病害数据填充矩阵B...")
        # 记录已处理的构件和填充的单元格数量
        component_to_row = {}  # 构件ID到行索引的映射
        filled_cells = 0
        fill_details = []
        # 只处理当前部件的病害数据
        if component_key in self.component_damage_data:
            damage_list = self.component_damage_data[component_key]
            for damage_info in damage_list:
                # 获取构件编号和构件名称作为唯一识别码
                component_code = damage_info.get("构件编号", "")
                component_name_from_info = damage_info.get("构件名称", "")
                # 构造唯一识别码
                component_id = f"{component_code}-{component_name_from_info}"
                # 获取病害排序号作为列索引
                disease_order = damage_info.get("病害排序号", 0)
                if not disease_order:
                    continue
                # 获取评定标度值
                scale_value = damage_info.get("评定标度", "")
                if not scale_value:
                    continue
                try:
                    # 将评定标度转换为整数
                    scale_int = int(scale_value)
                    # 确定行索引 - 为每个唯一构件分配一个行索引
                    if component_id not in component_to_row:
                        row_index = len(component_to_row)
                        component_to_row[component_id] = row_index
                    else:
                        row_index = component_to_row[component_id]
                    # 列索引就是病害排序号-1（因为索引从0开始）
                    col_index = int(disease_order) - 1
                    # 确保索引在有效范围内
                    if 0 <= row_index < rows and 0 <= col_index < cols:
                        # 填充矩阵B的对应单元格
                        matrix_B[row_index, col_index] = scale_int
                        filled_cells += 1
                        if filled_cells <= 5:  # 只记录前5个，避免输出过多
                            fill_details.append(f"填充矩阵B[{row_index}, {col_index}] = {scale_int} (构件: {component_id})")
                            print(f"填充矩阵B[{row_index}, {col_index}] = {scale_int} (构件: {component_id})")
                except (ValueError, IndexError) as e:
                    print(f"处理病害数据时出错: {e} (构件: {component_id})")
        print(f"部件 {component_key} 共填充了 {filled_cells} 个单元格，处理了 {len(component_to_row)} 个构件")
        # 返回矩阵和统计信息
        stats = {
            'filled_cells': filled_cells,
            'processed_components': len(component_to_row),
            'fill_details': fill_details
        }
        return matrix_B, stats
    def create_matrix_C_for_component_with_stats(self, matrix_B, rows, cols, component_key):
        """为特定部件创建矩阵C，并返回详细统计信息"""
        # 复制矩阵B作为矩阵C的基础
        matrix_C = matrix_B.copy()
        print(f"创建矩阵C for {component_key}（扣分值矩阵），维度: {rows}x{cols}")
        print(f"基于矩阵B，开始根据部件 {component_key} 的标度值填充扣分值...")
        # 记录统计信息
        filled_cells = 0
        zero_cells = 0
        fill_details = []  # 记录填充详情
        # 创建病害数据的快速查找字典
        damage_lookup = {}
        component_to_row = {}  # 构件ID到行索引的映射
        row_to_component = {}  # 行索引到构件ID的反向映射
        # 只处理当前部件的病害数据
        if component_key in self.component_damage_data:
            damage_list = self.component_damage_data[component_key]
            for damage_info in damage_list:
                component_code = damage_info.get("构件编号", "")
                component_name_from_info = damage_info.get("构件名称", "")
                component_id = f"{component_code}-{component_name_from_info}"
                disease_order = damage_info.get("病害排序号", 0)
                deduction_score = damage_info.get("扣分值", "")
                if disease_order and deduction_score:
                    try:
                        # 确定行索引
                        if component_id not in component_to_row:
                            row_index = len(component_to_row)
                            component_to_row[component_id] = row_index
                            row_to_component[row_index] = component_id
                        else:
                            row_index = component_to_row[component_id]
                        col_index = int(disease_order) - 1
                        score_int = int(deduction_score)
                        if 0 <= row_index < rows and 0 <= col_index < cols:
                            damage_lookup[(row_index, col_index)] = score_int
                    except (ValueError, IndexError):
                        continue
        # 遍历矩阵B的所有位置，根据标度值填充矩阵C
        for i in range(rows):
            for j in range(cols):
                scale_value = matrix_B[i, j]
                if scale_value == 1:
                    # 标度为1，扣分为0，但不记录到fill_details中
                    matrix_C[i, j] = 0
                    zero_cells += 1
                elif scale_value > 1:
                    # 标度大于1，查找对应的扣分值
                    if (i, j) in damage_lookup:
                        deduction_score = damage_lookup[(i, j)]
                        matrix_C[i, j] = deduction_score
                        filled_cells += 1
                        # 只记录有病害的单元格填充详情
                        component_id = row_to_component.get(i, f"构件-{i}")
                        fill_details.append(f"填充矩阵C[{i}, {j}] = {deduction_score} (构件: {component_id}, 标度值{scale_value}对应扣分{deduction_score})")
                        print(f"      填充矩阵C[{i}, {j}] = {deduction_score} (构件: {component_id}, 标度值{scale_value}对应扣分{deduction_score})")
                    else:
                        # 如果找不到对应的扣分值，设为0
                        matrix_C[i, j] = 0
        processed_cells = filled_cells + zero_cells
        print(f"    部件 {component_key} 共处理了 {processed_cells} 个单元格")
        # 返回矩阵和统计信息
        stats = {
            'processed_cells': processed_cells,
            'filled_cells': filled_cells,
            'zero_cells': zero_cells,
            'fill_details': fill_details
        }
        return matrix_C, stats
    def create_matrix_D_with_stats(self, matrix_C, rows, cols):
        """创建矩阵D，基于矩阵C，对每行扣分值进行从大到小排序，并返回详细统计信息
        Args:
            matrix_C (np.ndarray): 矩阵C（扣分值矩阵）
            rows (int): 矩阵行数
            cols (int): 矩阵列数
        Returns:
            tuple: (matrix_D, stats) 矩阵D和统计信息
        """
        # 复制矩阵C作为矩阵D的基础
        matrix_D = matrix_C.copy()
        print(f"\n创建了矩阵D（排序扣分值矩阵），维度: {rows}x{cols}")
        print("基于矩阵C，开始对每行扣分值进行从大到小排序...")
        
        # 记录统计信息和详细过程
        sorted_rows = 0
        total_non_zero_values = 0
        sort_details = []  # 记录排序详情
        
        # 遍历每一行进行排序
        for i in range(rows):
            # 获取当前行的数据
            row_data = matrix_D[i, :].copy()
            # 统计非零值数量
            non_zero_count = np.count_nonzero(row_data)
            
            if non_zero_count > 0:
                total_non_zero_values += non_zero_count
                # 记录原始非零值
                original_non_zero = row_data[row_data > 0]
                
                # 对当前行进行从大到小排序
                sorted_row = np.sort(row_data)[::-1]  # 降序排序
                # 将排序后的数据写回矩阵D
                matrix_D[i, :] = sorted_row
                sorted_rows += 1
                
                # 记录排序后的非零值
                sorted_non_zero = sorted_row[sorted_row > 0]
                
                # 记录排序详情
                if sorted_rows <= 10:  # 只记录前10行详情，避免数据过多
                    sort_detail = f"第{i + 1}行排序: 原始非零值{original_non_zero.tolist()} -> 排序后{sorted_non_zero.tolist()}"
                    sort_details.append(sort_detail)
                    print(f"  {sort_detail}")
                elif sorted_rows == 11:
                    print("  ... (更多行的排序信息已省略)")
                    sort_details.append("... (更多行的排序信息已省略)")
        
        # 使用通用方法打印矩阵预览
        self.print_matrix_preview(matrix_D, "D", rows, cols, "排序扣分值矩阵")
        
        print(f"\n排序统计信息：")
        print(f"  - 共处理了 {sorted_rows} 行数据")
        print(f"  - 总计非零扣分值: {total_non_zero_values} 个")
        print(f"  - 每行扣分值已按从大到小顺序排列")
        
        # 返回矩阵和统计信息
        stats = {
            'sorted_rows': sorted_rows,
            'total_non_zero_values': total_non_zero_values,
            'sort_details': sort_details
        }
        
        return matrix_D, stats
    def create_matrix_E_with_stats(self, matrix_D, rows, cols):
        """创建矩阵E，基于矩阵D，通过公式计算得到E矩阵，并返回详细统计信息
        Args:
            matrix_D (np.ndarray): 矩阵D（排序后的扣分值矩阵）
            rows (int): 矩阵行数
            cols (int): 矩阵列数
        Returns:
            tuple: (matrix_E, stats) 矩阵E和统计信息
        """
        # 创建与矩阵D相同维度的零矩阵作为矩阵E
        matrix_E = np.zeros_like(matrix_D, dtype=np.float64)
        print(f"\n创建了矩阵E（计算矩阵），维度: {rows}x{cols}")
        print("基于矩阵D，开始通过公式计算矩阵E...")
        
        # 记录统计信息和详细过程
        processed_rows = 0
        single_value_rows = 0
        multi_value_rows = 0
        calculation_details = []  # 记录计算详情
        format_details = []  # 记录格式化详情
        
        # 遍历每一行进行计算
        for i in range(rows):
            # 获取当前行的数据
            values = matrix_D[i]
            # 统计非零值数量
            non_zero_count = np.count_nonzero(values)
            
            if non_zero_count == 0:
                # 如果该行全为0，跳过
                continue
            elif non_zero_count == 1:
                # 如果只有一个非零值，直接赋值给E[i][0]
                matrix_E[i][0] = values[0]
                single_value_rows += 1
                processed_rows += 1
                
                if processed_rows <= 10:  # 只记录前10行，避免输出过多
                    calc_detail = f"第{i + 1}行: 单个非零值 {values[0]} -> E[{i}][0] = {matrix_E[i][0]}"
                    calculation_details.append(calc_detail)
                    print(f"  {calc_detail}")
            elif non_zero_count > 1:
                # 如果有多个非零值，使用公式计算
                matrix_E[i][0] = values[0]  # 第一个值直接赋值
                sum_e = matrix_E[i][0]  # 累计和
                
                # 计算后续值
                for j in range(1, non_zero_count):
                    factor = 100 - sum_e
                    denominator = 100 * (j + 1) ** (1 / 2)
                    matrix_E[i][j] = (values[j] * factor) / denominator
                    sum_e += matrix_E[i][j]
                
                multi_value_rows += 1
                processed_rows += 1
                
                if processed_rows <= 10:  # 只记录前10行，避免输出过多
                    non_zero_values = values[values > 0]
                    calculated_values = matrix_E[i][matrix_E[i] > 0]
                    calc_detail = f"第{i + 1}行: 多个非零值 {non_zero_values.tolist()} -> 计算结果 {calculated_values.tolist()}"
                    calculation_details.append(calc_detail)
                    print(f"  {calc_detail}")
            
            if processed_rows == 11:
                print("  ... (更多行的计算信息已省略)")
                calculation_details.append("... (更多行的计算信息已省略)")
        
        # 对矩阵E进行数值格式化处理：所有数值都保留两位小数
        print("\n开始对矩阵E进行数值格式化处理...")
        formatted_count = 0
        
        for i in range(rows):
            for j in range(cols):
                if matrix_E[i][j] != 0:  # 只处理非零值
                    original_value = matrix_E[i][j]
                    # 所有非零值都保留两位小数
                    matrix_E[i][j] = round(original_value, 2)
                    formatted_count += 1
                    
                    # 记录格式化详情
                    if formatted_count <= 5:
                        format_detail = f"格式化 E[{i}][{j}]: {original_value} -> {matrix_E[i][j]}"
                        format_details.append(format_detail)
                        print(f"  {format_detail}")
                    elif formatted_count == 6:
                        print("  ... (更多格式化信息已省略)")
                        format_details.append("... (更多格式化信息已省略)")
        
        # 使用通用方法打印矩阵预览
        self.print_matrix_preview(matrix_E, "E", rows, cols, "计算矩阵")
        
        print(f"\n计算统计信息：")
        print(f"  - 共处理了 {processed_rows} 行数据")
        print(f"  - 单个非零值行: {single_value_rows} 行")
        print(f"  - 多个非零值行: {multi_value_rows} 行")
        print(f"  - 格式化处理了 {formatted_count} 个非零值")
        print(f"  - 矩阵E已通过公式计算并格式化完成（所有数值保留两位小数）")
        
        # 返回矩阵和统计信息
        stats = {
            'processed_rows': processed_rows,
            'single_value_rows': single_value_rows,
            'multi_value_rows': multi_value_rows,
            'formatted_count': formatted_count,
            'calculation_details': calculation_details,
            'format_details': format_details
        }
        
        return matrix_E, stats
    def create_matrix_F_with_stats(self, matrix_E, rows):
        """创建矩阵F，基于矩阵E，计算每个构件的实际得分值（带详细统计信息）
        Args:
            matrix_E (np.ndarray): 矩阵E（计算后的扣分值矩阵）
            rows (int): 矩阵行数
            cols (int): 矩阵列数
        Returns:
            tuple: (matrix_F, stats) - 矩阵F和统计信息
        """
        # 创建列向量矩阵F，维度为 rows x 1
        matrix_F = np.zeros((rows, 1), dtype=np.float64)
        print(f"\n创建了矩阵F（构件实际得分矩阵），维度: {rows}x1")
        print("基于矩阵E，开始计算每个构件的实际得分...")
        # 初始化统计信息
        stats = {
            'processed_components': 0,
            'total_deduction': 0,
            'calculation_details': [],
            'min_score': 100.00,
            'max_score': 100.00,
            'defective_components': 0
        }
        
        # 记录计算详情
        calculation_details = []
        processed_components = 0
        total_deduction = 0
        
        # 遍历每一行（每个构件）
        for i in range(rows):
            # 计算当前行的扣分值总和
            row_sum = np.sum(matrix_E[i])
            # 计算实际得分：100 - 扣分值总和
            actual_score = 100 - row_sum
            # 保留两位小数
            matrix_F[i][0] = round(actual_score, 2)
            
            # 统计信息
            if row_sum > 0:  # 只统计有扣分的构件
                processed_components += 1
                total_deduction += row_sum
                # 记录前10个构件的计算过程
                if processed_components <= 10:
                    detail = f"构件{i + 1}: 扣分总和 {row_sum:.2f} -> 实际得分 {matrix_F[i][0]:.2f}"
                    calculation_details.append(detail)
                    print(f"  {detail}")
                elif processed_components == 11:
                    print("  ... (更多构件的计算信息已省略)")
        
        # 打印矩阵F的预览（列向量格式）
        print(f"\n矩阵F的内容（构件实际得分）:")
        print("构件编号 | 实际得分")
        print("-" * 20)
        
        # 显示前10行和后5行
        display_rows = min(10, rows)
        for i in range(display_rows):
            component_id = i + 1
            score = matrix_F[i][0]
            print(f"构件{component_id:3d}   | {score:6.2f}")
        
        if rows > 15:
            print("   ...   |   ...")
            for i in range(max(display_rows, rows - 5), rows):
                component_id = i + 1
                score = matrix_F[i][0]
                print(f"构件{component_id:3d}   | {score:6.2f}")
        elif rows > display_rows:
            for i in range(display_rows, rows):
                component_id = i + 1
                score = matrix_F[i][0]
                print(f"构件{component_id:3d}   | {score:6.2f}")
        # 计算统计信息
        min_score = np.min(matrix_F)
        max_score = np.max(matrix_F)
        print(f"\n计算统计信息：")
        print(f"  - 共计算了 {rows} 个构件的实际得分")
        print(f"  - 有扣分的构件数量: {processed_components} 个")
        print(f"  - 总扣分值: {total_deduction:.2f}")
        print(f"  - 最低得分: {min_score:.2f}")
        print(f"  - 最高得分: {max_score:.2f}")
        print(f"  - 矩阵F已完成计算（构件实际得分列向量）")
        # 更新统计信息
        stats.update({
            'processed_components': processed_components,
            'total_deduction': round(total_deduction, 2),
            'calculation_details': calculation_details,
            'min_score': round(min_score, 2),
            'max_score': round(max_score, 2),
            'defective_components': processed_components
        })
        return matrix_F, stats
    def get_t_value(self, structure_name, component_name):
        """根据结构系统和部件名称获取构件总数量，并计算对应的t值
        Args:
            structure_name (str): 结构系统名称
            component_name (str): 部件名称
        Returns:
            float: 对应的t值，如果无法获取构件数量则返回None
        """
        # t值表格
        t_values_table = {
            1: np.inf, 2: 10, 3: 9.7, 4: 9.5, 5: 9.2, 6: 8.9, 7: 8.7, 8: 8.5, 9: 8.3, 10: 8.1,
            11: 7.9, 12: 7.7, 13: 7.5, 14: 7.3, 15: 7.2, 16: 7.08, 17: 6.96, 18: 6.84, 19: 6.72,
            20: 6.6, 21: 6.48, 22: 6.36, 23: 6.24, 24: 6.12, 25: 6, 26: 5.88, 27: 5.76, 28: 5.64,
            29: 5.52, 30: 5.4, 40: 4.9, 50: 4.4, 60: 4, 70: 3.6, 80: 3.2, 90: 2.8, 100: 2.5, 200: 2.3
        }
        # 构建唯一键
        quantity_key = f"{structure_name}-{component_name}"
        # 检查是否存在构件总数量
        if quantity_key not in self.component_quantities:
            print(f"  警告: 无法找到 '{quantity_key}' 的构件总数量，无法计算t值")
            return None
        n2 = self.component_quantities[quantity_key]
        # 检查构件数量是否有效
        if n2 <= 0:
            print(f"  警告: '{quantity_key}' 的构件总数量为 {n2}，无法计算t值")
            return None
        print(f"  获取构件数量: {quantity_key} -> {n2}")
        # 特殊情况处理
        # 当构件数量等于1时，返回无穷大
        if n2 == 1:
            t_value = np.inf
            print(f"  构件总数 {n2} -> t值 = {t_value}（构件数量为1）")
            return t_value
        # 当构件数量大于等于200时，返回2.3
        if n2 >= 200:
            t_value = 2.3
            print(f"  构件总数 {n2} >= 200 -> t值 = {t_value:.2f}")
            return round(t_value, 2)
        # 获取表格中的所有键值，并排序（排除1，因为已经单独处理）
        keys = sorted([k for k in t_values_table.keys() if k > 1])
        # 如果n2小于等于最小值（2），返回最小值对应的t值
        if n2 <= keys[0]:
            t_value = t_values_table[keys[0]]
            print(f"  构件总数 {n2} <= {keys[0]} -> t值 = {t_value:.2f}")
            return round(t_value, 2)
        # 线性插值（在1-200范围内）
        for i in range(len(keys) - 1):
            if keys[i] < n2 < keys[i + 1]:
                x1, x2 = keys[i], keys[i + 1]
                y1, y2 = t_values_table[x1], t_values_table[x2]
                # 线性插值公式: y = y1 + (y2 - y1) * (x - x1) / (x2 - x1)
                t_value = y1 + (y2 - y1) * (n2 - x1) / (x2 - x1)
                print(f"  构件总数 {n2} 在 [{x1}, {x2}] 之间 -> 线性插值: t值 = {t_value:.2f}")
                return round(t_value, 2)
            elif n2 == keys[i]:
                # 如果正好等于某个键值
                t_value = t_values_table[keys[i]]
                print(f"  构件总数 {n2} -> t值 = {t_value:.2f}（精确匹配）")
                return round(t_value, 2)
        # 理论上不应该到达这里
        print(f"  警告: 无法为构件总数 {n2} 计算t值")
        return None
    def calculate_h_value_with_stats(self, matrix_F, structure_name, component_name):
        """计算部件的技术状况评分h值并收集详细统计信息
        Args:
            matrix_F (np.ndarray): F矩阵（构件得分列向量）
            structure_name (str): 结构系统名称
            component_name (str): 部件名称
        Returns:
            tuple: (h_value, stats_dict) h值和详细统计信息
        """
        print(f"\n开始计算部件技术状况评分h值...")
        
        # 初始化统计信息字典
        stats = {
            'component_quantity': 0,
            't_value': None,
            't_value_details': '',
            'defective_count': 0,
            'min_score': 100,
            'avg_score': 100,
            'calculation_method': '',
            'calculation_process': '',
            'final_h_value': None
        }
        
        # 从component_quantities获取构件总数量
        quantity_key = f"{structure_name}-{component_name}"
        if quantity_key not in self.component_quantities:
            print(f"警告: 无法找到 '{quantity_key}' 的构件总数量")
            return None, stats
        
        n2 = self.component_quantities[quantity_key]
        stats['component_quantity'] = n2
        print(f"构件总数量 n2 = {n2}")
        
        # 获取t值和详细信息
        t_value = self.get_t_value(structure_name, component_name)
        if t_value is None:
            print(f"无法获取t值，无法计算h值")
            return None, stats
        
        stats['t_value'] = t_value
        
        # 生成t值详细说明
        if n2 == 1:
            stats['t_value_details'] = f"构件数量为1，t值 = ∞"
        elif n2 >= 200:
            stats['t_value_details'] = f"构件数量 ≥ 200，t值 = 2.3"
        else:
            stats['t_value_details'] = f"根据构件数量{n2}查表或插值得到t值 = {t_value}"
        
        # 统计得分小于100的构件数量（有缺陷的构件）
        n1 = np.count_nonzero(matrix_F < 100)
        stats['defective_count'] = n1
        print(f"有缺陷的构件数量 n1 = {n1}")
        
        # 获取所有有缺陷构件的得分
        F_defects = matrix_F[matrix_F < 100]
        
        # 计算最低得分
        if n1 > 0:
            fmin = np.min(F_defects)
            stats['min_score'] = round(fmin, 2)
            print(f"最低得分 fmin = {fmin:.2f}")
        else:
            fmin = 100
            stats['min_score'] = 100
            print(f"无缺陷构件，最低得分 fmin = {fmin}")
        
        # 计算平均得分
        if n1 > 0:
            defects_sum = np.sum(F_defects)
            perfect_count = n2 - n1
            average_f = (defects_sum + perfect_count * 100) / n2
        else:
            average_f = 100
        
        stats['avg_score'] = round(average_f, 2)
        print(f"平均得分 average_f = {average_f:.2f}")
        
        # 判断是否为主要部件
        is_major_component = self.is_major_component(structure_name, component_name)
        
        # 计算h值
        if fmin < 40 and is_major_component:
            h = fmin
            stats['calculation_method'] = '主要部件特殊规则'
            stats['calculation_process'] = f"由于是主要部件且最低得分{fmin:.2f}<40，直接使用最低得分作为h值\nh={fmin:.2f}"
            print(f"主要部件最低得分{fmin:.2f}<40，直接使用最低得分作为h值")
        else:
            h = average_f - (100 - fmin) / t_value
            stats['calculation_method'] = '标准计算公式'
            if t_value == np.inf:
                stats['calculation_process'] = f"\nh = 平均得分 - (100 - 最低得分) / t值\nh = {average_f:.2f} - (100 - {fmin:.2f}) / ∞\nh = {average_f:.2f} - 0 = {h:.2f}"
            else:
                stats['calculation_process'] = f"\nh = 平均得分 - (100 - 最低得分) / t值\nh = {average_f:.2f} - (100 - {fmin:.2f}) / {t_value:.2f}\nh = {average_f:.2f} - {(100 - fmin) / t_value:.2f} = {h:.2f}"
            
            if fmin < 40 and not is_major_component:
                print(f"非主要部件最低得分 {fmin:.2f} < 40，但仍使用修正公式计算h值")
            else:
                print(f"最低得分 {fmin:.2f} >= 40，使用修正公式计算h值")
            print(f"h = {average_f:.2f} - (100 - {fmin:.2f}) / {t_value:.2f} = {h:.2f}")
        
        h = round(h, 2)
        stats['final_h_value'] = h
        print(f"\n部件综合评定分数 h = {h:.2f}")
        
        return h, stats
    def is_major_component(self, structure_name, component_name):
        """判断是否为主要部件
        Args:
            structure_name (str): 结构系统名称
            component_name (str): 部件名称
        Returns:
            bool: 是否为主要部件
        """
        # 定义主要部件
        major_components = {
            '上部结构': ['拱肋、系杆'],
            '下部结构': ['桥墩', '桥台', '墩台基础']
        }
        # 桥面系的所有部件都不是主要部件
        if structure_name == '桥面系':
            return False
        # 检查是否在主要部件列表中
        return component_name in major_components.get(structure_name, [])
    def calculate_bridge_technical_condition_score(self):
        """计算全桥技术状况评分
        根据各结构系统的权重分配和部件h值，分别计算桥面系、上部结构、下部结构的技术状况评分
        """
        try:
            print(f"\n=== 开始计算全桥技术状况评分 ===")
            # 检查是否有权重数据
            if not hasattr(self, 'current_weights') or not self.current_weights:
                QMessageBox.warning(self, "计算错误", "请先导入病害数据以获取权重分配信息！")
                return
            # 检查是否有h值计算结果
            if not hasattr(self, 'component_calculation_results') or not self.component_calculation_results:
                QMessageBox.warning(self, "计算错误", "请先执行矩阵运算以获取各部件的h值！")
                return
            # 存储各结构系统的评分结果
            structure_scores = {}
            # 遍历三个结构系统
            for structure_name in ['桥面系', '上部结构', '下部结构']:
                print(f"\n--- 计算结构系统: {structure_name} ---")
                # 检查该结构系统是否有权重数据
                if structure_name not in self.current_weights:
                    print(f"  警告: 结构系统 '{structure_name}' 没有权重数据，跳过计算")
                    structure_scores[structure_name] = 0.00
                    continue
                # 获取该结构系统的权重分配
                structure_weights = self.current_weights[structure_name]
                # 计算该结构系统的加权得分
                weighted_score_sum = 0.00
                total_weight = 0.00
                calculated_components = 0
                print(f"  {structure_name} 部件权重和h值计算:")
                print(f"  {'部件名称':<15} {'权重':<8} {'h值':<8} {'加权得分':<10}")
                print(f"  {'-' * 45}")
                # 遍历该结构系统下的所有部件
                for component_name, weight in structure_weights.items():
                    # 构造唯一键
                    component_key = f"{structure_name}-{component_name}"
                    # 检查权重是否为0（部件不存在）
                    if weight == 0.00:
                        print(f"  {component_name:<15} {weight:<8.2f} {'N/A':<8} {'0.00':<10} (部件不存在)")
                        continue
                    # 检查是否有h值计算结果
                    if component_key not in self.component_calculation_results:
                        print(f"  {component_name:<15} {weight:<8.2f} {'N/A':<8} {'0.00':<10} (无h值数据)")
                        continue
                    # 获取h值
                    h_value = self.component_calculation_results[component_key]['h_value']
                    # 计算加权得分
                    weighted_score = h_value * weight
                    weighted_score_sum += weighted_score
                    total_weight += weight
                    calculated_components += 1
                    print(f"  {component_name:<15} {weight:<8.2f} {h_value:<8.2f} {weighted_score:<10.2f}")
                # 计算该结构系统的最终评分
                if total_weight > 0:
                    # 这里使用加权和作为结构系统评分
                    structure_score = weighted_score_sum
                    structure_scores[structure_name] = round(structure_score, 2)
                    print(f"  {'-' * 45}")
                    print(f"  {structure_name} 加权得分总和: {weighted_score_sum:.2f}")
                    print(f"  {structure_name} 权重总和: {total_weight:.2f}")
                    print(f"  {structure_name} 技术状况评分: {structure_score:.2f}")
                    print(f"  计算的部件数量: {calculated_components}")
                else:
                    structure_scores[structure_name] = 0.00
                    print(f"  {structure_name} 无有效部件，评分为: 0.00")
            # 技术状况等级评估函数
            def evaluate_technical_condition(score):
                """根据得分评估技术状况等级"""
                if 95 <= score <= 100:
                    return "1类"
                elif 80 <= score < 95:
                    return "2类"
                elif 60 <= score < 80:
                    return "3类"
                elif 40 <= score < 60:
                    return "4类"
                elif 0 <= score < 40:
                    return "5类"
                else:
                    raise ValueError(f"无效的评分值:{score}，评分应在0到100之间")
            # 输出最终结果
            print(f"\n=== 全桥技术状况评分结果 ===")
            print(f"{'结构系统':<15} {'技术状况评分':<15} {'技术状况等级':<10}")
            print(f"{'-' * 45}")
            for structure_name, score in structure_scores.items():
                grade = evaluate_technical_condition(score)
                print(f"{structure_name:<15} {score:<15.2f} {grade:<10}")
            # 存储结果供后续使用
            self.bridge_structure_scores = structure_scores
            self.bridge_structure_grades = {}
            for structure_name, score in structure_scores.items():
                grade = evaluate_technical_condition(score)
                self.bridge_structure_grades[structure_name] = grade
            # 计算全桥技术状况评分
            print(f"\n=== 计算全桥技术状况评分 ===")
            bridge_weights = {
                '桥面系': 0.20,
                '上部结构': 0.40,
                '下部结构': 0.40
            }
            # 计算全桥技术状况评分 Dc = BDCI × WD + SPCI × WSP + SBCI × WSB
            total_bridge_score = 0.00
            print(f"全桥技术状况评分计算过程:")
            print(f"{'结构系统':<15} {'评分值':<10} {'权重':<8} {'加权得分':<10}")
            print(f"{'-' * 45}")
            for structure_name, weight in bridge_weights.items():
                if structure_name in structure_scores:
                    structure_score = structure_scores[structure_name]
                    weighted_score = structure_score * weight
                    total_bridge_score += weighted_score
                    print(f"{structure_name:<15} {structure_score:<10.2f} {weight:<8.2f} {weighted_score:<10.2f}")
                else:
                    print(f"{structure_name:<15} {'0.00':<10} {weight:<8.2f} {'0.00':<10}")
            print(f"{'-' * 45}")
            print(f"全桥技术状况评分 (Dr): {total_bridge_score:.2f}")
            # 评估全桥技术状况等级
            bridge_grade = evaluate_technical_condition(total_bridge_score)
            print(f"全桥技术状况等级: {bridge_grade}")
            # 存储全桥评分结果
            self.total_bridge_score = round(total_bridge_score, 2)
            self.total_bridge_grade = bridge_grade
            print(f"\n=== 全桥技术状况评分计算完成 ===")
            print(f"最终结果已保存至 self.total_bridge_score: {self.total_bridge_score}")
            print(f"全桥技术状况等级已保存至 self.total_bridge_grade: {self.total_bridge_grade}")
            # 更新全桥技术状况表格显示
            self.update_bridge_status_table()
            # 更新汇总表格
            self.update_bridge_summary_table()
            return structure_scores
        except Exception as e:
            error_msg = f"计算全桥技术状况评分时发生错误: {str(e)}"
            print(f"\n{error_msg}")
            return None
    def export_calculation_process(self):
        """导出并展示完整的矩阵运算过程"""
        # 检查是否已执行矩阵运算 - 修改检查条件
        if not hasattr(self, 'component_calculation_results') or not self.component_calculation_results:
            QMessageBox.warning(self, "提示",
                                "请先执行矩阵运算后再查看运算过程！\n\n"
                                "操作步骤：\n"
                                "1.录入各部件的病害数据\n"
                                "2.点击'🧮矩阵运算'按钮\n"
                                "3.运算完成后即可查看详细过程")
            return
        # 检查是否有运算过程数据
        if not hasattr(self, 'calculation_process_data'):
            # 如果没有详细的过程数据，但有运算结果，显示基本信息
            self.calculation_process_data = {
                'basic_info': {
                    'matrix_dimensions': f"{getattr(self, 'matrix_rows', 1000)}x{getattr(self, 'matrix_cols', 12)}",
                    'start_time': '运算已完成',
                    'total_structures': len([s for s in ['桥面系', '上部结构', '下部结构'] if
                                             hasattr(self, 'component_counts') and s in self.component_counts]),
                    'total_components': sum(
                        len(components) for components in getattr(self, 'component_counts', {}).values())
                },
                'structures': {},
                'weight_allocation': {
                    'allocation_process': '权重已按构件数量动态分配',
                    'normalization_process': '权重已归一化处理'
                },
                'final_results': {
                    'structure_scores': self._format_structure_scores(),
                    'bridge_score': self._format_bridge_score(),
                    'grade_determination': '评级已根据评分标准确定'
                }
            }
        # 创建并显示运算过程窗口
        try:
            process_window = CalculationProcessWindow(self)
            process_window.exec()
        except Exception as e:
            QMessageBox.critical(self, "显示错误", f"显示运算过程时发生错误:\n{str(e)}")
    def _format_structure_scores(self):
        """格式化结构系统评分信息"""
        if not hasattr(self, 'bridge_structure_scores'):
            return "暂无结构系统评分数据"
        result = []
        for structure, data in self.bridge_structure_scores.items():
            score = data.get('score', 0)
            grade = data.get('grade', '未知')
            result.append(f"{structure}: {score:.2f}分 ({grade})")
        return "\n".join(result)
    def _format_bridge_score(self):
        """格式化全桥评分信息"""
        if not hasattr(self, 'total_bridge_score'):
            return "暂无全桥评分数据"
        score = getattr(self, 'total_bridge_score', 0)
        grade = getattr(self, 'total_bridge_grade', '未知')
        return f"全桥总评分: {score:.2f}分\n技术状况等级: {grade}"
    
    def show_damage_heatmap(self):
        """显示病害分布热力图"""
        try:
            # 检查是否有病害数据
            if not hasattr(self, 'component_damage_data') or not self.component_damage_data:
                QMessageBox.information(self, "提示", "当前项目暂无病害数据，无法生成热力图。\n请先录入病害信息。")
                return
            # 创建并显示热力图窗口
            heatmap_window = DamageHeatmapWindow(self.component_damage_data, self)
            heatmap_window.exec()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"显示病害热力图时发生错误：\n{str(e)}")
    def export_inspection_report(self):
        """导出检查报告表"""
        try:
            # 打开项目选择窗口
            export_window = ExportReportWindow(self)
            export_window.exec()
        except Exception as e:
            QMessageBox.critical(
                self,
                "导出失败",
                f"导出检查报告时发生错误：\n{str(e)}"
            )
            print(f"导出检查报告时发生错误：{e}")


class DamageHeatmapWindow(QDialog):
    """病害分布热力图窗口"""
    def __init__(self, damage_data, parent=None):
        super().__init__(parent)
        self.damage_data = damage_data
        self.parent_system = parent
        self.init_ui()
        self.generate_heatmap()
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("病害分布热力图分析")
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        # 设置窗口大小和居中显示
        window_width = 1400
        window_height = 900
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        # 设置窗口图标
        try:
            self.setWindowIcon(QIcon("Basic_data/Highway_bridges.ico"))
        except:
            pass
        
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                font-family: "楷体", "Times New Roman", serif;
            }
            QLabel {
                font-family: "楷体", "Times New Roman", serif;
                color: #2c3e50;
                font-size: 13px;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
                font-family: "楷体", "Times New Roman", serif;
                min-width: 100px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5dade2, stop:1 #3498db);
            }
            QTabWidget::pane {
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background: white;
            }
            QTabBar::tab {
                background: #ecf0f1;
                color: #2c3e50;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-family: "楷体";
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background: #3498db;
                color: white;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel("病害分布热力图分析")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("楷体", 18, QFont.Bold))
        title_label.setStyleSheet("""
            color: #e67e22;
            padding: 10px;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(230, 126, 34, 0.1), 
                stop:0.5 rgba(230, 126, 34, 0.05),
                stop:1 rgba(230, 126, 34, 0.1));
            border-radius: 10px;
            margin-bottom: 12px;
            font-family: "楷体", "Times New Roman", serif;
            border: 2px solid rgba(230, 126, 34, 0.3);
        """)
        layout.addWidget(title_label)
        # 创建标签页
        self.tab_widget = QTabWidget()
        # 病害严重程度热力图标签页
        self.severity_tab = QWidget()
        self.create_severity_heatmap_tab()
        self.tab_widget.addTab(self.severity_tab, "🌡️病害严重程度分布")
        # 病害类型分布标签页
        self.type_tab = QWidget()
        self.create_type_distribution_tab()
        self.tab_widget.addTab(self.type_tab, "📊病害类型统计")
        # 构件病害密度标签页
        self.density_tab = QWidget()
        self.create_density_heatmap_tab()
        self.tab_widget.addTab(self.density_tab, "🎯构件病害密度")
        # 扣分影响分析标签页
        self.impact_tab = QWidget()
        self.create_impact_analysis_tab()
        self.tab_widget.addTab(self.impact_tab, "⚡扣分影响分析")
        layout.addWidget(self.tab_widget)
        
    def create_severity_heatmap_tab(self):
        """创建病害严重程度热力图标签页"""
        layout = QVBoxLayout(self.severity_tab)
        # 创建matplotlib图形
        self.severity_figure = Figure(figsize=(12, 8))
        self.severity_canvas = FigureCanvas(self.severity_figure)
        layout.addWidget(self.severity_canvas)
        # 添加说明文本
        info_label = QLabel("📋说明：颜色越深表示病害越严重")
        info_label.setStyleSheet("color: #7f8c8d; font-style: italic; padding: 10px;")
        layout.addWidget(info_label)
    
    def create_type_distribution_tab(self):
        """创建病害类型分布标签页"""
        layout = QVBoxLayout(self.type_tab)
        # 创建matplotlib图形
        self.type_figure = Figure(figsize=(12, 8))
        self.type_canvas = FigureCanvas(self.type_figure)
        layout.addWidget(self.type_canvas)
        # 添加说明文本
        info_label = QLabel("📋说明：展示各种病害类型的出现频次和分布情况")
        info_label.setStyleSheet("color: #7f8c8d; font-style: italic; padding: 10px;")
        layout.addWidget(info_label)
    
    def create_density_heatmap_tab(self):
        """创建构件病害密度标签页"""
        layout = QVBoxLayout(self.density_tab)
        # 创建matplotlib图形
        self.density_figure = Figure(figsize=(12, 8))
        self.density_canvas = FigureCanvas(self.density_figure)
        layout.addWidget(self.density_canvas)
        # 添加说明文本
        info_label = QLabel("📋说明：显示各构件的病害数量密度，帮助识别问题集中区域")
        info_label.setStyleSheet("color: #7f8c8d; font-style: italic; padding: 10px;")
        layout.addWidget(info_label)
    
    def create_impact_analysis_tab(self):
        """创建扣分影响分析标签页"""
        layout = QVBoxLayout(self.impact_tab)
        # 创建matplotlib图形
        self.impact_figure = Figure(figsize=(12, 8))
        self.impact_canvas = FigureCanvas(self.impact_figure)
        layout.addWidget(self.impact_canvas)
        # 添加说明文本
        info_label = QLabel("📋说明：分析各病害对总体评分的影响程度，指导维护优先级")
        info_label.setStyleSheet("color: #7f8c8d; font-style: italic; padding: 10px;")
        layout.addWidget(info_label)
    
    def generate_heatmap(self):
        """生成所有热力图"""
        try:
            self.generate_severity_heatmap()
            self.generate_type_distribution()
            self.generate_density_heatmap()
            self.generate_impact_analysis()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成热力图时发生错误：\n{str(e)}")
    
    def generate_severity_heatmap(self):
        """生成病害严重程度热力图"""
        self.severity_figure.clear()
        # 统一设置中英文字体
        plt.rcParams['font.sans-serif'] = ['KaiTi', 'Times New Roman']  # 楷体和新罗马
        plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
        # 准备数据
        severity_data = defaultdict(lambda: defaultdict(list))
        for component_key, damages in self.damage_data.items():
            structure_system = component_key.split('-')[0]  # 桥面系、上部结构、下部结构
            component_name = component_key.split('-')[1] if '-' in component_key else component_key
            for damage in damages:
                severity = int(damage.get('评定标度', '1'))
                severity_data[structure_system][component_name].append(severity)
        # 计算平均严重程度
        avg_severity = {}
        for structure in severity_data:
            avg_severity[structure] = {}
            for component in severity_data[structure]:
                severities = severity_data[structure][component]
                avg_severity[structure][component] = sum(severities) / len(severities) if severities else 0
        if not avg_severity:
            ax = self.severity_figure.add_subplot(111)
            ax.text(0.5, 0.5, '暂无病害数据', ha='center', va='center', fontsize=14, 
                fontfamily='KaiTi', transform=ax.transAxes)
            self.severity_canvas.draw()
            return
        # 创建热力图数据矩阵
        structures = list(avg_severity.keys())
        all_components = set()
        for structure in structures:
            all_components.update(avg_severity[structure].keys())
        components = sorted(list(all_components))
        # 构建数据矩阵
        matrix = []
        for structure in structures:
            row = []
            for component in components:
                value = avg_severity[structure].get(component, 0)
                row.append(value)
            matrix.append(row)
        # 绘制热力图
        ax = self.severity_figure.add_subplot(111)
        im = ax.imshow(matrix, cmap='Reds', aspect='auto', vmin=0, vmax=4)
        # 设置坐标轴 - 统一字体大小
        ax.set_xticks(range(len(components)))
        ax.set_yticks(range(len(structures)))
        ax.set_xticklabels(components, rotation=45, ha='right', fontsize=11, fontfamily='KaiTi')
        ax.set_yticklabels(structures, fontsize=12, fontfamily='KaiTi')
        # 添加数值标注 - 统一字体
        for i in range(len(structures)):
            for j in range(len(components)):
                if matrix[i][j] > 0:
                    text = ax.text(j, i, f'{matrix[i][j]:.1f}',
                                ha="center", va="center", 
                                color="black" if matrix[i][j] < 2 else "white",
                                fontweight='bold', fontsize=10, fontfamily='Times New Roman')
        # 添加颜色条 - 标签使用楷体，刻度使用 Times New Roman
        cbar = self.severity_figure.colorbar(im, ax=ax, shrink=0.8)
        cbar.set_label('平均评定标度', rotation=270, labelpad=20, fontsize=12, fontfamily='KaiTi')  # 标签字体为楷体

        # 设置刻度标签的字体为 Times New Roman
        cbar.ax.tick_params(labelsize=10)
        for l in cbar.ax.get_yticklabels():
            l.set_fontname('Times New Roman')
        # 设置标题和标签 - 统一字体
        ax.set_title('病害严重程度分布热力图', fontsize=16, fontweight='bold', pad=20, fontfamily='KaiTi')
        ax.set_xlabel('部件类型', fontsize=12, fontfamily='KaiTi')
        ax.set_ylabel('结构部位', fontsize=12, fontfamily='KaiTi')
        
        # 调整布局避免重叠
        self.severity_figure.tight_layout(pad=2.0)
        self.severity_canvas.draw()

    def generate_type_distribution(self):
        """生成病害类型分布图"""
        self.type_figure.clear()
        
        # 统计病害类型
        damage_types = Counter()
        damage_by_structure = defaultdict(Counter)
        
        for component_key, damages in self.damage_data.items():
            structure_system = component_key.split('-')[0]
            
            for damage in damages:
                damage_type = damage.get('病害类型', '未知')
                damage_types[damage_type] += 1
                damage_by_structure[structure_system][damage_type] += 1
        
        if not damage_types:
            ax = self.type_figure.add_subplot(111)
            ax.text(0.5, 0.5, '暂无病害数据', ha='center', va='center', fontsize=14, 
                fontfamily='KaiTi', transform=ax.transAxes)
            self.type_canvas.draw()
            return
        
        # 统一设置中英文字体
        plt.rcParams['font.sans-serif'] = ['KaiTi', 'Times New Roman']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 创建子图 - 调整间距
        ax1 = self.type_figure.add_subplot(221)  # 总体病害类型饼图
        ax2 = self.type_figure.add_subplot(222)  # 病害类型柱状图
        ax3 = self.type_figure.add_subplot(212)  # 各结构系统病害类型堆叠柱状图
        
        # 1. 总体病害类型饼图
        top_damages = damage_types.most_common(8)  # 取前8种病害
        labels = [item[0] for item in top_damages]
        sizes = [item[1] for item in top_damages]
        colors = plt.cm.Set3(range(len(labels)))
        
        wedges, texts, autotexts = ax1.pie(sizes, labels=labels, autopct='%1.1f%%', 
                                        colors=colors, startangle=90)
        # 统一饼图字体
        for text in texts:
            text.set_fontsize(10)
            text.set_fontfamily('KaiTi')
        for autotext in autotexts:
            autotext.set_fontsize(9)
            autotext.set_fontfamily('Times New Roman')
            autotext.set_color('black')
            autotext.set_weight('bold')
        
        ax1.set_title('病害类型分布', fontsize=13, fontweight='bold', fontfamily='KaiTi')
        
        # 2. 病害类型柱状图
        damage_names = [item[0] for item in damage_types.most_common(10)]
        damage_counts = [item[1] for item in damage_types.most_common(10)]
        
        bars = ax2.bar(range(len(damage_names)), damage_counts, color='skyblue', alpha=0.7)
        ax2.set_xlabel('病害类型', fontsize=11, fontfamily='KaiTi')
        ax2.set_ylabel('出现次数', fontsize=11, fontfamily='KaiTi')
        ax2.set_title('病害类型频次统计', fontsize=13, fontweight='bold', fontfamily='KaiTi')
        ax2.set_xticks(range(len(damage_names)))
        ax2.set_xticklabels(damage_names, rotation=45, ha='right', fontsize=9, fontfamily='KaiTi')
        ax2.tick_params(axis='y', labelsize=10, labelfontfamily='Times New Roman')
        
        # 添加数值标注
        for bar in bars:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.2,
                    f'{int(height)}', ha='center', va='bottom', 
                    fontsize=9, fontfamily='Times New Roman', fontweight='bold')
        
        # 3. 各结构系统病害类型堆叠柱状图
        structures = list(damage_by_structure.keys())
        all_damage_types = set()
        for structure in structures:
            all_damage_types.update(damage_by_structure[structure].keys())
        
        damage_type_list = sorted(list(all_damage_types))
        
        # 准备堆叠数据
        bottom = [0] * len(structures)
        colors = plt.cm.tab20(range(len(damage_type_list)))
        
        for i, damage_type in enumerate(damage_type_list):
            values = [damage_by_structure[structure][damage_type] for structure in structures]
            ax3.bar(structures, values, bottom=bottom, label=damage_type, 
                color=colors[i], alpha=0.8)
            bottom = [b + v for b, v in zip(bottom, values)]
        
        ax3.set_xlabel('结构部位', fontsize=12, fontfamily='KaiTi')
        ax3.set_ylabel('病害数量', fontsize=12, fontfamily='KaiTi')
        ax3.set_title('各结构部位病害类型分布', fontsize=14, fontweight='bold', fontfamily='KaiTi')
        ax3.tick_params(axis='x', labelsize=11, labelfontfamily='KaiTi')
        ax3.tick_params(axis='y', labelsize=10, labelfontfamily='Times New Roman')
        
        # 调整图例位置和字体
        legend = ax3.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=9)
        for text in legend.get_texts():
            text.set_fontfamily('KaiTi')
        
        # 调整布局避免重叠
        self.type_figure.tight_layout(pad=3.0)
        self.type_canvas.draw()

    def generate_density_heatmap(self):
        """生成构件病害密度热力图"""
        self.density_figure.clear()
        
        # 统计各构件的病害数量
        component_damage_count = defaultdict(int)
        component_by_structure = defaultdict(list)
        
        for component_key, damages in self.damage_data.items():
            structure_system = component_key.split('-')[0]
            component_name = component_key.split('-')[1] if '-' in component_key else component_key
            
            damage_count = len(damages)
            component_damage_count[component_key] = damage_count
            component_by_structure[structure_system].append((component_name, damage_count))
        
        if not component_damage_count:
            ax = self.density_figure.add_subplot(111)
            ax.text(0.5, 0.5, '暂无病害数据', ha='center', va='center', fontsize=14, 
                fontfamily='KaiTi', transform=ax.transAxes)
            self.density_canvas.draw()
            return
        
        # 统一设置中英文字体
        plt.rcParams['font.sans-serif'] = ['KaiTi', 'Times New Roman']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 创建树状图风格的可视化
        ax = self.density_figure.add_subplot(111)
        
        y_pos = 0
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
        structure_colors = {}
        
        for i, (structure, components) in enumerate(component_by_structure.items()):
            structure_colors[structure] = colors[i % len(colors)]
            
            # 绘制结构系统标题 - 统一字体
            ax.text(-0.8, y_pos, structure, fontsize=13, fontweight='bold', 
                ha='right', va='center', color=structure_colors[structure], fontfamily='KaiTi')
            
            # 绘制构件病害密度
            for component_name, damage_count in sorted(components, key=lambda x: x[1], reverse=True):
                # 绘制矩形，宽度表示病害数量
                rect = patches.Rectangle((0, y_pos-0.3), damage_count, 0.6, 
                                    facecolor=structure_colors[structure], 
                                    alpha=0.7, edgecolor='black', linewidth=1)
                ax.add_patch(rect)
                
                # 添加构件名称和数量标注 - 统一字体
                ax.text(0.2, y_pos, f'{component_name}', 
                    fontsize=11, va='center', ha='left', fontfamily='KaiTi')
                ax.text(damage_count + 0.2, y_pos, f'({damage_count})', 
                    fontsize=10, va='center', ha='left', fontfamily='Times New Roman', fontweight='bold')
                
                y_pos -= 1
            
            y_pos -= 0.8  # 增加结构系统间距
        
        # 设置坐标轴 - 统一字体
        max_damage = max(component_damage_count.values()) if component_damage_count else 1
        ax.set_xlim(-2, max_damage + 2)
        ax.set_ylim(y_pos, 1)
        ax.set_xlabel('病害数量', fontsize=12, fontfamily='KaiTi')
        ax.set_title('结构部位病害密度分布', fontsize=16, fontweight='bold', pad=20, fontfamily='KaiTi')
        
        # 设置x轴刻度字体
        ax.tick_params(axis='x', labelsize=10, labelfontfamily='Times New Roman')
        
        # 隐藏y轴刻度
        ax.set_yticks([])
        
        # 添加网格
        ax.grid(True, axis='x', alpha=0.3)
        
        # 调整布局
        self.density_figure.tight_layout(pad=2.0)
        self.density_canvas.draw()

    def generate_impact_analysis(self):
        """生成扣分影响分析图"""
        self.impact_figure.clear()
        
        # 统计扣分数据
        component_scores = defaultdict(list)
        damage_type_scores = defaultdict(list)
        severity_scores = defaultdict(list)
        
        for component_key, damages in self.damage_data.items():
            structure_system = component_key.split('-')[0]
            
            for damage in damages:
                score = damage.get('扣分值', 0)
                damage_type = damage.get('病害类型', '未知')
                severity = damage.get('评定标度', '1')
                
                component_scores[structure_system].append(score)
                damage_type_scores[damage_type].append(score)
                severity_scores[severity].append(score)
        
        if not component_scores:
            ax = self.impact_figure.add_subplot(111)
            ax.text(0.5, 0.5, '暂无病害数据', ha='center', va='center', fontsize=14, 
                fontfamily='KaiTi', transform=ax.transAxes)
            self.impact_canvas.draw()
            return
        
        # 统一设置中英文字体
        plt.rcParams['font.sans-serif'] = ['KaiTi', 'Times New Roman']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 创建子图 - 调整间距
        ax1 = self.impact_figure.add_subplot(221)  # 结构系统扣分分布
        ax2 = self.impact_figure.add_subplot(222)  # 病害类型扣分分布
        ax3 = self.impact_figure.add_subplot(223)  # 严重程度扣分分布
        ax4 = self.impact_figure.add_subplot(224)  # 扣分累积影响
        
        # 1. 结构系统扣分分布箱线图
        structure_data = []
        structure_labels = []
        for structure, scores in component_scores.items():
            structure_data.append(scores)
            structure_labels.append(structure)
        
        if structure_data:
            bp1 = ax1.boxplot(structure_data, labels=structure_labels, patch_artist=True)
            colors = ['lightblue', 'lightgreen', 'lightcoral']
            for patch, color in zip(bp1['boxes'], colors):
                patch.set_facecolor(color)
            ax1.set_title('结构部位扣分分布', fontsize=12, fontweight='bold', fontfamily='KaiTi')
            ax1.set_ylabel('扣分值', fontsize=11, fontfamily='KaiTi')
            ax1.tick_params(axis='x', labelsize=10, labelfontfamily='KaiTi')
            ax1.tick_params(axis='y', labelsize=10, labelfontfamily='Times New Roman')
        
        # 2. 病害类型扣分影响
        damage_avg_scores = {dt: sum(scores)/len(scores) for dt, scores in damage_type_scores.items()}
        top_damages = sorted(damage_avg_scores.items(), key=lambda x: x[1], reverse=True)[:8]
        
        damage_names = [item[0] for item in top_damages]
        damage_scores = [item[1] for item in top_damages]
        
        bars = ax2.barh(range(len(damage_names)), damage_scores, color='orange', alpha=0.7)
        ax2.set_yticks(range(len(damage_names)))
        ax2.set_yticklabels(damage_names, fontsize=10, fontfamily='KaiTi')
        ax2.set_xlabel('平均扣分值', fontsize=11, fontfamily='KaiTi')
        ax2.set_title('病害类型平均扣分', fontsize=12, fontweight='bold', fontfamily='KaiTi')
        ax2.tick_params(axis='x', labelsize=10, labelfontfamily='Times New Roman')
        
        # 添加数值标注
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax2.text(width, bar.get_y() + bar.get_height()/2,
                    f'{width:.1f}', ha='left', va='center', 
                    fontsize=9, fontfamily='Times New Roman', fontweight='bold')
        
        # 3. 严重程度扣分分布
        severity_avg = {}
        for severity, scores in severity_scores.items():
            severity_avg[severity] = sum(scores) / len(scores) if scores else 0
        
        severities = sorted(severity_avg.keys())
        avg_scores = [severity_avg[s] for s in severities]
        
        bars3 = ax3.bar(severities, avg_scores, color='red', alpha=0.6)
        ax3.set_xlabel('评定标度', fontsize=11, fontfamily='KaiTi')
        ax3.set_ylabel('平均扣分值', fontsize=11, fontfamily='KaiTi')
        ax3.set_title('严重程度与扣分关系', fontsize=12, fontweight='bold', fontfamily='KaiTi')
        ax3.tick_params(axis='both', labelsize=10, labelfontfamily='Times New Roman')
        
        # 添加数值标注
        for bar in bars3:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:.1f}', ha='center', va='bottom', 
                    fontsize=9, fontfamily='Times New Roman', fontweight='bold')
        
        # 4. 扣分累积影响分析
        all_scores = []
        for damages in self.damage_data.values():
            for damage in damages:
                all_scores.append(damage.get('扣分值', 0))
        
        if all_scores:
            all_scores.sort(reverse=True)
            cumulative_scores = []
            cumulative_sum = 0
            for score in all_scores:
                cumulative_sum += score
                cumulative_scores.append(cumulative_sum)
            
            ax4.plot(range(1, len(cumulative_scores)+1), cumulative_scores, 
                    marker='o', markersize=3, linewidth=2, color='purple')
            ax4.set_xlabel('病害数量（按扣分排序）', fontsize=11, fontfamily='KaiTi')
            ax4.set_ylabel('累积扣分值', fontsize=11, fontfamily='KaiTi')
            ax4.set_title('扣分累积影响曲线', fontsize=12, fontweight='bold', fontfamily='KaiTi')
            ax4.tick_params(axis='both', labelsize=10, labelfontfamily='Times New Roman')
            ax4.grid(True, alpha=0.3)
        # 调整布局避免重叠
        self.impact_figure.tight_layout(pad=3.0)
        self.impact_canvas.draw()

class ExportReportWindow(QDialog):
    """导出报告窗口"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_system = parent
        self.projects_data = []
        self.filtered_projects = []
        self.init_ui()
        self.load_projects()
    
    def init_ui(self):
        self.setWindowTitle("导出桥梁技术状况评定报告")
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        # 设置窗口大小和居中显示
        window_width = 950
        window_height = 700
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        # 设置窗口图标
        try:
            self.setWindowIcon(QIcon("Basic_data/Highway_bridges.ico"))
        except:
            pass
        # 优雅简洁的样式设计
        self.setStyleSheet("""
            QDialog {
                background-color: #f8f9fa;
                font-family: "KaiTi", "Times New Roman", serif;
                border: 1px solid #dee2e6;
            }
            QLabel {
                font-family: "KaiTi", "Times New Roman", serif;
                color: #495057;
                font-size: 14px;
                background: transparent;
            }
            .title-label {
                font-family: "KaiTi", "Times New Roman", serif;
                color: #2c3e50;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                background-color: #ffffff;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin: 10px;
            }
            .info-label {
                font-family: "KaiTi", "Times New Roman", serif;
                color: #6c757d;
                font-size: 14px;
                padding: 12px;
                background-color: #e9ecef;
                border-radius: 6px;
                border-left: 4px solid #17a2b8;
            }
            .search-label {
                font-family: "KaiTi", "Times New Roman", serif;
                color: #2c3e50;
                font-size: 14px;
                font-weight: bold;
            }
            QLineEdit {
                border: 2px solid #ced4da;
                border-radius: 6px;
                padding: 10px 15px;
                font-size: 14px;
                font-family: "Times New Roman", "KaiTi", serif;
                background-color: #ffffff;
                color: #495057;
                min-height: 20px;
                selection-background-color: #007bff;
            }
            QLineEdit:focus {
                border-color: #007bff;
                outline: none;
                box-shadow: 0 0 0 3px rgba(0, 123, 255, 0.25);
            }
            QLineEdit::placeholder {
                color: #6c757d;
                font-style: italic;
            }
            QTableWidget {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: #ffffff;
                gridline-color: #e9ecef;
                font-family: "Times New Roman", "KaiTi", serif;
                font-size: 14px;
                selection-background-color: #e3f2fd;
                color: #495057;
                alternate-background-color: #f8f9fa;
            }
            QTableWidget::item {
                padding: 12px 8px;
                border-bottom: 1px solid #e9ecef;
                text-align: center;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: #1976d2;
                border: 1px solid #2196f3;
            }
            QTableWidget::item:hover {
                background-color: #f5f5f5;
            }
            QHeaderView::section {
                background-color: #3498db;
                color: #ffffff;
                padding: 12px 8px;
                border: none;
                font-weight: bold;
                font-family: "KaiTi", "Times New Roman", serif;
                font-size: 14px;
                text-align: center;
                min-height: 35px;
                border-right: 1px solid rgba(255, 255, 255, 0.3);
            }
            QHeaderView::section:last {
                border-right: none;
            }
            QPushButton {
                background-color: #007bff;
                color: #ffffff;
                border: 2px solid #007bff;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
                font-family: "KaiTi", "Times New Roman", serif;
                min-width: 100px;
                min-height: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
                border-color: #0056b3;
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background-color: #004085;
                border-color: #004085;
                transform: translateY(0px);
            }
            QPushButton:disabled {
                background-color: #6c757d;
                color: #ffffff;
                border-color: #6c757d;
            }
            .export-button {
                background-color: #28a745;
                border-color: #28a745;
            }
            .export-button:hover {
                background-color: #218838;
                border-color: #218838;
            }
            .export-button:pressed {
                background-color: #1e7e34;
                border-color: #1e7e34;
            }
            .refresh-button {
                background-color: #fd7e14;
                border-color: #fd7e14;
            }
            .refresh-button:hover {
                background-color: #e8650e;
                border-color: #e8650e;
            }
            .refresh-button:pressed {
                background-color: #d15502;
                border-color: #d15502;
            }
            .status-label {
                font-family: "KaiTi", "Times New Roman", serif;
                color: #6c757d;
                font-size: 14px;
                font-style: italic;
                padding: 10px 15px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                border-left: 4px solid #6c757d;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(25, 25, 25, 25)
        # 标题
        title_label = QLabel("导出桥梁技术状况评定报告")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("KaiTi", 14, QFont.Bold))
        title_label.setProperty("class", "title-label")
        layout.addWidget(title_label)
        # 说明文字
        info_label = QLabel("请选择需要导出报告的项目，程序将自动生成详细报告")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setFont(QFont("KaiTi", 12))
        info_label.setProperty("class", "info-label")
        layout.addWidget(info_label)
        # 搜索区域
        search_layout = QHBoxLayout()
        search_layout.setSpacing(5)  # 设置间距为5px
        search_label = QLabel("项目检索：")
        search_label.setFont(QFont("KaiTi", 12, QFont.Bold))
        search_label.setProperty("class", "search-label")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入项目名称进行搜索...")
        self.search_input.textChanged.connect(self.filter_projects)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)
        # 项目列表表格
        self.projects_table = QTableWidget()
        self.projects_table.setColumnCount(3)
        self.projects_table.setHorizontalHeaderLabels(["项目名称", "保存时间", "文件名"])
        # ===== 添加以下配置 =====
        vertical_header = self.projects_table.verticalHeader()
        vertical_header.setFont(QFont("Times New Roman", 10))       # 设置字体
        vertical_header.setDefaultAlignment(Qt.AlignCenter)        # 设置居中对齐
        vertical_header.setDefaultSectionSize(30)
        # 设置列宽
        header = self.projects_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # 项目名称 - 固定宽度
        self.projects_table.setColumnWidth(0, 300)        
        header.setSectionResizeMode(1, QHeaderView.Fixed)  # 保存时间 - 固定宽度
        self.projects_table.setColumnWidth(1, 180)        
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # 文件名 - 自适应剩余空间
        # 设置表格属性
        self.projects_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.projects_table.setAlternatingRowColors(True)
        self.projects_table.setSortingEnabled(True)
        # 表格选择事件
        self.projects_table.itemSelectionChanged.connect(self.on_selection_changed)
        layout.addWidget(self.projects_table)
        # 添加双击事件
        self.projects_table.itemDoubleClicked.connect(self.export_selected_project)
        layout.addWidget(self.projects_table)
        # 状态标签
        self.status_label = QLabel("正在初始化系统...")
        self.status_label.setProperty("class", "status-label")
        layout.addWidget(self.status_label)
        # 按钮区域
        button_layout = QHBoxLayout()
        self.export_button = QPushButton("导出选中项目")
        self.export_button.setProperty("class", "export-button")
        self.export_button.clicked.connect(self.export_selected_project)
        self.export_button.setEnabled(False)
        self.refresh_button = QPushButton("刷新列表")
        self.refresh_button.setProperty("class", "refresh-button")
        self.refresh_button.clicked.connect(self.load_projects)
        button_layout.addWidget(self.export_button)
        button_layout.addStretch()
        button_layout.addWidget(self.refresh_button)
        layout.addLayout(button_layout)
    
    def load_projects(self):
        """加载所有保存的项目"""
        try:
            self.projects_data = []
            save_dir = "Saved_Projects/Steel-concrete_composite_arch_bridge"
            if not os.path.exists(save_dir):
                self.status_label.setText("保存目录不存在，暂无项目数据")
                self.update_table()
                return
            # 扫描保存目录中的所有JSON文件
            json_files = [f for f in os.listdir(save_dir) if f.endswith('.json')]
            if not json_files:
                self.status_label.setText("暂无保存的项目")
                self.update_table()
                return

            # 读取每个项目文件的信息
            for filename in json_files:
                filepath = os.path.join(save_dir, filename)
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)
                    
                    project_info = project_data.get('project_info', {})
                    project_name = project_info.get('name', '未知项目')
                    save_time = project_info.get('save_time', '未知时间')
                    
                    # 格式化保存时间
                    try:
                        if save_time != '未知时间':
                            dt = datetime.datetime.fromisoformat(save_time)
                            formatted_time = dt.strftime('%Y-%m-%d %H:%M:%S')
                        else:
                            formatted_time = save_time
                    except:
                        formatted_time = save_time
                    
                    self.projects_data.append({
                        'name': project_name,
                        'save_time': formatted_time,
                        'filename': filename,
                        'filepath': filepath,
                        'data': project_data
                    })
                    
                except Exception as e:
                    QMessageBox.warning(self, "读取文件错误", f"读取项目文件{filename}时发生错误：\n{str(e)}")
                    continue
            
            # 按保存时间排序（最新的在前）
            self.projects_data.sort(key=lambda x: x['save_time'], reverse=True)
            self.filtered_projects = self.projects_data.copy()
            self.update_table()
            self.status_label.setText(f"系统检测到{len(self.projects_data)}个项目")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载项目列表时发生错误：\n{str(e)}")
            self.status_label.setText("系统加载失败")
    
    def filter_projects(self):
        """根据搜索关键词过滤项目"""
        search_text = self.search_input.text().strip().lower()
        if not search_text:
            self.filtered_projects = self.projects_data.copy()
        else:
            self.filtered_projects = [
                project for project in self.projects_data
                if search_text in project['name'].lower()
            ]
        self.update_table()
        
        if search_text:
            self.status_label.setText(f"检索到{len(self.filtered_projects)}个匹配项目")
        else:
            self.status_label.setText(f"系统检测到{len(self.projects_data)}个项目")
    
    def update_table(self):
        """更新项目列表表格"""
        self.projects_table.setRowCount(len(self.filtered_projects))
        for row, project in enumerate(self.filtered_projects):
            # 项目名称
            name_item = QTableWidgetItem(project['name'])
            name_item.setFont(QFont("KaiTi", 12, QFont.Bold))
            name_item.setTextAlignment(Qt.AlignCenter)
            self.projects_table.setItem(row, 0, name_item)
            # 保存时间
            time_item = QTableWidgetItem(project['save_time'])
            time_item.setFont(QFont("Times New Roman", 12))
            time_item.setTextAlignment(Qt.AlignCenter)
            self.projects_table.setItem(row, 1, time_item)
            # 文件名
            filename_item = QTableWidgetItem(project['filename'])
            filename_item.setFont(QFont("Times New Roman", 12))
            filename_item.setTextAlignment(Qt.AlignCenter)
            self.projects_table.setItem(row, 2, filename_item)
        # 调整行高
        for row in range(len(self.filtered_projects)):
            self.projects_table.setRowHeight(row, 35)
    
    def on_selection_changed(self):
        """表格选择变化事件"""
        selected_rows = self.projects_table.selectionModel().selectedRows()
        has_selection = len(selected_rows) > 0
        self.export_button.setEnabled(has_selection)
    
    def get_selected_project(self):
        """获取选中的项目数据"""
        selected_rows = self.projects_table.selectionModel().selectedRows()
        if not selected_rows:
            return None
        row = selected_rows[0].row()
        if 0 <= row < len(self.filtered_projects):
            return self.filtered_projects[row]
        return None
    
    def export_selected_project(self):
        """导出选中项目的检查报告"""
        project = self.get_selected_project()
        if not project:
            QMessageBox.warning(self, "警告", "请先选择一个项目")
            return
        try:
            # 确保Export_data目录存在
            export_dir = "Export_data/Steel-concrete_composite_arch_bridge"
            if not os.path.exists(export_dir):
                os.makedirs(export_dir)
            # 生成文件名
            project_name = project['data'].get('project_info', {}).get('name', '未知项目')
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{project_name}_评分报告_{timestamp}.xlsx"
            filepath = os.path.join(export_dir, filename)
            # 生成报告
            generator = BridgeReportGenerator(project['data'])
            generator.generate_report(filepath)
            # 导出成功提示
            reply = QMessageBox.question(
                self,
                "导出成功",
                f"报告已成功导出！\n"
                f"项目名称：{project_name}\n"
                f"文件路径：{filepath}\n"
                f"是否立即打开报告文件？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                # 打开文件
                import subprocess
                try:
                    os.startfile(filepath)  # Windows
                except AttributeError:
                    try:
                        subprocess.call(['open', filepath])  # macOS
                    except:
                        subprocess.call(['xdg-open', filepath])  # Linux
            # 关闭窗口
            self.close()
        except Exception as e:
            QMessageBox.critical(
                self,
                "导出失败",
                f"导出检查报告时发生错误：\n{str(e)}"
            )

class BridgeReportGenerator:
    """桥梁检查报告生成器"""
    def __init__(self, project_data):
        self.project_data = project_data
        self.project_info = project_data.get('project_info', {})
        self.component_damage_data = project_data.get('component_damage_data', {})
        self.evaluation_results = project_data.get('evaluation_results', {})
    def generate_report(self, filepath):
        """生成完整的检查报告"""
        # 创建工作簿
        wb = openpyxl.Workbook()
        # 删除默认工作表
        wb.remove(wb.active)
        # 依次创建各个结构系统的工作表
        self.create_deck_system_sheet(wb.create_sheet(title="桥面系评分表"))  # 桥面系
        self.create_superstructure_sheet(wb.create_sheet(title="上部结构评分表"))  # 上部结构
        self.create_substructure_sheet(wb.create_sheet(title="下部结构评分表"))  # 下部结构
        self.create_total_bridge_sheet(wb.create_sheet(title="全桥评分表"))  # 全桥技术状况评分表
        self.create_summary_sheet(wb.create_sheet(title="技术状况汇总表"))  # 技术状况总表
        # 保存文件
        wb.save(filepath)
        print(f"检查报告已生成：{filepath}")
    def create_deck_system_sheet(self, ws):
        """创建桥面系技术状况BDCI评分表"""
        structure_name = '桥面系'
        # 设置标题
        title = f"表7.1 {structure_name}技术状况BDCI评分表"
        ws.merge_cells('A1:F1')
        ws['A1'] = title
        ws['A1'].font = Font(name='Times New Roman', size=12, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        # 设置表头 - 使用LaTeX格式
        headers = ['部件名称', '构件数量', '损坏类型', '标度', '病害描述', '扣分值$DP_{ij}$']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        current_row = 3
        # 从JSON文件中获取桥面系存在的构件
        deck_existence = self.project_data.get('component_existence', {}).get('桥面系', {})
        existing_components = []  # 记录存在的构件顺序
        # 遍历桥面系中存在的构件
        for component_name, exists in deck_existence.items():
            if not exists:  # 跳过不存在的构件
                continue
            existing_components.append(component_name)  # 记录存在的构件
            component_key = f"{structure_name}-{component_name}"
            damages = self.component_damage_data.get(component_key, [])
            total_quantity = self.project_data.get('component_quantities', {}).get(component_key, 0)
            if total_quantity == 0:
                continue  # 跳过数量为0的部件
            # 按构件编号分组病害数据
            damage_by_component = {}
            for damage in damages:
                component_id = damage.get('构件编号', '')
                if component_id not in damage_by_component:
                    damage_by_component[component_id] = []
                damage_by_component[component_id].append(damage)
            damaged_components_count = len(damage_by_component)  # 有病害的构件数量
            undamaged_components_count = total_quantity - damaged_components_count  # 无病害的构件数量
            start_row = current_row
            # 处理有病害的构件
            for component_id, component_damages in damage_by_component.items():
                component_start_row = current_row
                # 为该构件的每个病害创建一行
                for i, damage in enumerate(component_damages):
                    # 部件名称（只在该构件的第一行显示）
                    cell = ws.cell(row=current_row, column=1)
                    if current_row == start_row:  # 整个部件的第一行
                        cell.value = component_name
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                    # 构件数量（只在该构件的第一行显示1）
                    cell = ws.cell(row=current_row, column=2)
                    if i == 0:  # 该构件的第一行
                        cell.value = 1
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 损坏类型
                    cell = ws.cell(row=current_row, column=3)
                    cell.value = damage.get('病害类型', '')
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 标度
                    cell = ws.cell(row=current_row, column=4)
                    cell.value = f"{damage.get('评定标度', '1')}类"
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 病害描述
                    cell = ws.cell(row=current_row, column=5)
                    cell.value = damage.get('病害描述', '')
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 扣分值
                    cell = ws.cell(row=current_row, column=6)
                    cell.value = damage.get('扣分值', 0)
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    current_row += 1
                # 合并该构件的构件数量列（如果有多个病害）
                if len(component_damages) > 1:
                    ws.merge_cells(f'B{component_start_row}:B{current_row-1}')
            # 处理无病害的构件（如果有）
            if undamaged_components_count > 0:
                # 部件名称（如果没有病害构件，这是第一行）
                cell = ws.cell(row=current_row, column=1)
                if damaged_components_count == 0:
                    cell.value = component_name
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 构件数量（显示无病害构件的数量）
                cell = ws.cell(row=current_row, column=2)
                cell.value = undamaged_components_count
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 损坏类型
                cell = ws.cell(row=current_row, column=3)
                cell.value = "无"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 标度
                cell = ws.cell(row=current_row, column=4)
                cell.value = "1类"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 病害描述
                cell = ws.cell(row=current_row, column=5)
                cell.value = "——"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 扣分值
                cell = ws.cell(row=current_row, column=6)
                cell.value = 0
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                current_row += 1
            # 合并部件名称单元格（如果该部件有多行）
            total_rows = sum(len(damages) for damages in damage_by_component.values()) + (1 if undamaged_components_count > 0 else 0)
            if total_rows > 1:
                ws.merge_cells(f'A{start_row}:A{current_row-1}')
        # 添加部件评分表格
        component_scores = self.project_data.get('component_scores', {}).get('桥面系', {})
        adjusted_weights = self.project_data.get('adjusted_weights', {}).get('桥面系', {})
        bridge_score = self.project_data.get('evaluation_results', {}).get('bridge_structure_scores', {}).get('桥面系', 0)
        total_components = len(existing_components)  # 记录桥面系中存在的部件总数
        # 为每个存在的部件添加评分行
        dcci_index = 1
        for component_name in existing_components:
            if component_name in component_scores:
                # 部件名称评分行
                cell = ws.cell(row=current_row, column=1)
                cell.value = f"{component_name}评分"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 合并B到F列显示DCCI公式
                ws.merge_cells(f'B{current_row}:F{current_row}')
                cell = ws.cell(row=current_row, column=2)
                # 获取详细计算信息
                calculation_details = self.project_data.get('component_calculation_details', {}).get('桥面系', {}).get(component_name, {})
                calculation_method = calculation_details.get('calculation_method', '')
                t_value = calculation_details.get('t_value', None)
                final_score = component_scores[component_name]
                # 根据t_value和calculation_method决定显示方式
                if t_value is None:
                    # 无病害部件：t_value为null，calculation_method为空，直接显示最终分数
                    cell.value = f"$DCCI_{{{dcci_index}}}=\\overline{{DMCI}}-\\frac{{100-DMCI_{{min}}}}{{t}}={final_score:.2f}$"
                elif str(t_value) == "Infinity":
                    # 构件数量为1的特殊情况：t_value为无穷大
                    if calculation_method == '标准计算公式':
                        avg_score = calculation_details.get('avg_score', 100)
                        min_score = calculation_details.get('min_score', 100)
                        cell.value = f"$DCCI_{{{dcci_index}}}=\\overline{{DMCI}}-\\frac{{100-DMCI_{{min}}}}{{t}}={avg_score:.2f}-\\frac{{100-{min_score:.2f}}}{{\\infty}}={final_score:.2f}$"
                    else:
                        cell.value = f"$DCCI_{{{dcci_index}}}=DMCI_{{min}}={final_score:.2f}$"
                elif calculation_method == '标准计算公式' and t_value is not None:
                    # 标准计算公式：显示详细计算过程
                    avg_score = calculation_details.get('avg_score', 100)
                    min_score = calculation_details.get('min_score', 100)
                    cell.value = f"$DCCI_{{{dcci_index}}}=\\overline{{DMCI}}-\\frac{{100-DMCI_{{min}}}}{{t}}={avg_score:.2f}-\\frac{{100-{min_score:.2f}}}{{{float(t_value):.2f}}}={final_score:.2f}$"
                elif calculation_method == '主要部件特殊规则':
                    # 主要部件特殊规则：直接显示最终分数
                    cell.value = f"$DCCI_{{{dcci_index}}}=DMCI_{{min}}={final_score:.2f}$"
                else:
                    # 其他情况：显示公式前缀和最终分数
                    cell.value = f"$DCCI_{{{dcci_index}}}=\\overline{{DMCI}}-\\frac{{100-DMCI_{{min}}}}{{t}}={final_score:.2f}$"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                # 为合并的单元格添加边框
                for col in range(2, 7):
                    border_cell = ws.cell(row=current_row, column=col)
                    border_cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                current_row += 1
                dcci_index += 1
        # 添加桥面系总评分行
        cell = ws.cell(row=current_row, column=1)
        cell.value = f"{structure_name}评分"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 合并B到F列显示BDCI公式 - 使用LaTeX格式
        ws.merge_cells(f'B{current_row}:F{current_row}')
        cell = ws.cell(row=current_row, column=2)
        # 构建BDCI公式字符串
        formula_parts = []
        for i, component_name in enumerate(existing_components):
            if component_name in component_scores and component_name in adjusted_weights:
                score = component_scores[component_name]
                weight = adjusted_weights[component_name]
                formula_parts.append(f"{score:.2f}\\times{weight:.2f}")
        # 弹性换行：根据部件数量决定换行策略
        if len(formula_parts) <= 2:
            # 部件数量少，不换行
            formula_str = "+".join(formula_parts)
        elif len(formula_parts) <= 4:
            # 部件数量中等，每2项换行
            formula_str = "+".join(formula_parts[:2])+"\\\\+"+"+".join(formula_parts[2:])
        else:
            # 部件数量多，每3项换行
            formula_lines = []
            for i in range(0, len(formula_parts), 3):
                line_parts = formula_parts[i:i+3]
                if i == 0:
                    # 第一行不需要前置+号
                    formula_lines.append("+".join(line_parts))
                else:
                    # 后续行需要前置+号
                    formula_lines.append("+"+"+".join(line_parts))
            formula_str = "\\\\".join(formula_lines)
        # 使用LaTeX格式的求和公式，包含完整的上下标
        cell.value = f"$BDCI=\\sum_{{i=1}}^{{{total_components}}}DCCI_i\\timesW_i={formula_str}={bridge_score:.2f}$"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        # 为合并的单元格添加边框
        for col in range(2, 7):
            border_cell = ws.cell(row=current_row, column=col)
            border_cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        # 记录表格的最后一行
        last_row = current_row
        # 修正双线外边框位置 - 从表头行（第2行）开始而不是标题行
        # 顶部边框（从第2行开始）
        for col in range(1, 7):
            cell = ws.cell(row=2, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=Side(style='double'),
                bottom=current_border.bottom
            )
        # 底部边框
        for col in range(1, 7):
            cell = ws.cell(row=last_row, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=Side(style='double')
            )
        # 左侧边框（从第2行开始）
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=1)
            current_border = cell.border
            cell.border = Border(
                left=Side(style='double'),
                right=current_border.right,
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 右侧边框（从第2行开始）
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=6)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=Side(style='double'),
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 设置列宽 - 增加扣分值列的宽度
        column_widths = [18, 10, 15, 8, 60, 18]  
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        # 设置行高
        for row in range(1, current_row + 1):
            ws.row_dimensions[row].height = 25
    
    def create_superstructure_sheet(self, ws):
        """创建上部结构技术状况SPCI评分表"""
        structure_name = '上部结构'
        # 设置标题
        title = f"表7.2 {structure_name}技术状况SPCI评分表"
        ws.merge_cells('A1:F1')
        ws['A1'] = title
        ws['A1'].font = Font(name='Times New Roman', size=12, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        # 设置表头 - 使用LaTeX格式
        headers = ['部件名称', '构件数量', '损坏类型', '标度', '病害描述', '扣分值$DP_{ij}$']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        current_row = 3
        # 从JSON文件中获取上部结构存在的构件
        structure_existence = self.project_data.get('component_existence', {}).get('上部结构', {})
        existing_components = []  # 记录存在的构件顺序
        # 遍历上部结构中存在的构件
        for component_name, exists in structure_existence.items():
            if not exists:  # 跳过不存在的构件
                continue
            existing_components.append(component_name)  # 记录存在的构件
            component_key = f"{structure_name}-{component_name}"
            damages = self.component_damage_data.get(component_key, [])
            total_quantity = self.project_data.get('component_quantities', {}).get(component_key, 0)
            if total_quantity == 0:
                continue  # 跳过数量为0的部件
            # 按构件编号分组病害数据
            damage_by_component = {}
            for damage in damages:
                component_id = damage.get('构件编号', '')
                if component_id not in damage_by_component:
                    damage_by_component[component_id] = []
                damage_by_component[component_id].append(damage)
            damaged_components_count = len(damage_by_component)  # 有病害的构件数量
            undamaged_components_count = total_quantity - damaged_components_count  # 无病害的构件数量
            start_row = current_row
            # 处理有病害的构件
            for component_id, component_damages in damage_by_component.items():
                component_start_row = current_row
                # 为该构件的每个病害创建一行
                for i, damage in enumerate(component_damages):
                    # 部件名称（只在该构件的第一行显示）
                    cell = ws.cell(row=current_row, column=1)
                    if current_row == start_row:  # 整个部件的第一行
                        cell.value = component_name
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                    # 构件数量（只在该构件的第一行显示1）
                    cell = ws.cell(row=current_row, column=2)
                    if i == 0:  # 该构件的第一行
                        cell.value = 1
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 损坏类型
                    cell = ws.cell(row=current_row, column=3)
                    cell.value = damage.get('病害类型', '')
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 标度
                    cell = ws.cell(row=current_row, column=4)
                    cell.value = f"{damage.get('评定标度', '1')}类"
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 病害描述
                    cell = ws.cell(row=current_row, column=5)
                    cell.value = damage.get('病害描述', '')
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 扣分值
                    cell = ws.cell(row=current_row, column=6)
                    cell.value = damage.get('扣分值', 0)
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    current_row += 1
                # 合并该构件的构件数量列（如果有多个病害）
                if len(component_damages) > 1:
                    ws.merge_cells(f'B{component_start_row}:B{current_row-1}')
            # 处理无病害的构件（如果有）
            if undamaged_components_count > 0:
                # 部件名称（如果没有病害构件，这是第一行）
                cell = ws.cell(row=current_row, column=1)
                if damaged_components_count == 0:
                    cell.value = component_name
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 构件数量（显示无病害构件的数量）
                cell = ws.cell(row=current_row, column=2)
                cell.value = undamaged_components_count
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 损坏类型
                cell = ws.cell(row=current_row, column=3)
                cell.value = "无"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 标度
                cell = ws.cell(row=current_row, column=4)
                cell.value = "1类"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 病害描述
                cell = ws.cell(row=current_row, column=5)
                cell.value = "——"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 扣分值
                cell = ws.cell(row=current_row, column=6)
                cell.value = 0
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                current_row += 1
            # 合并部件名称单元格（如果该部件有多行）
            total_rows = sum(len(damages) for damages in damage_by_component.values()) + (1 if undamaged_components_count > 0 else 0)
            if total_rows > 1:
                ws.merge_cells(f'A{start_row}:A{current_row-1}')
        # 添加部件评分表格
        component_scores = self.project_data.get('component_scores', {}).get('上部结构', {})
        adjusted_weights = self.project_data.get('adjusted_weights', {}).get('上部结构', {})
        bridge_score = self.project_data.get('evaluation_results', {}).get('bridge_structure_scores', {}).get('上部结构', 0)
        total_components = len(existing_components)  # 记录上部结构中存在的部件总数
        # 为每个存在的部件添加评分行
        scci_index = 1
        for component_name in existing_components:
            if component_name in component_scores:
                # 部件名称评分行
                cell = ws.cell(row=current_row, column=1)
                cell.value = f"{component_name}评分"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 合并B到F列显示SCCI公式
                ws.merge_cells(f'B{current_row}:F{current_row}')
                cell = ws.cell(row=current_row, column=2)
                
                # 获取详细计算信息
                calculation_details = self.project_data.get('component_calculation_details', {}).get('上部结构', {}).get(component_name, {})
                calculation_method = calculation_details.get('calculation_method', '')
                t_value = calculation_details.get('t_value', None)
                final_score = component_scores[component_name]
                
                # 根据t_value和calculation_method决定显示方式
                if t_value is None:
                    # 无病害部件：t_value为null，calculation_method为空，直接显示最终分数
                    cell.value = f"$PCCI_{{{scci_index}}}=\\overline{{PMCI}}-\\frac{{100-PMCI_{{min}}}}{{t}}={final_score:.2f}$"
                elif str(t_value) == "Infinity":
                    # 构件数量为1的特殊情况：t_value为无穷大
                    if calculation_method == '标准计算公式':
                        avg_score = calculation_details.get('avg_score', 100)
                        min_score = calculation_details.get('min_score', 100)
                        cell.value = f"$PCCI_{{{scci_index}}}=\\overline{{PMCI}}-\\frac{{100-PMCI_{{min}}}}{{t}}={avg_score:.2f}-\\frac{{100-{min_score:.2f}}}{{\\infty}}={final_score:.2f}$"
                    else:
                        cell.value = f"$PCCI_{{{scci_index}}}=PMCI_{{min}}={final_score:.2f}$"
                elif calculation_method == '标准计算公式' and t_value is not None:
                    # 标准计算公式：显示详细计算过程
                    avg_score = calculation_details.get('avg_score', 100)
                    min_score = calculation_details.get('min_score', 100)
                    cell.value = f"$PCCI_{{{scci_index}}}=\\overline{{PMCI}}-\\frac{{100-PMCI_{{min}}}}{{t}}={avg_score:.2f}-\\frac{{100-{min_score:.2f}}}{{{float(t_value):.2f}}}={final_score:.2f}$"
                elif calculation_method == '主要部件特殊规则':
                    # 主要部件特殊规则：直接显示最终分数
                    cell.value = f"$PCCI_{{{scci_index}}}=PMCI_{{min}}={final_score:.2f}$"
                else:
                    # 其他情况：显示公式前缀和最终分数
                    cell.value = f"$PCCI_{{{scci_index}}}=\\overline{{PMCI}}-\\frac{{100-PMCI_{{min}}}}{{t}}={final_score:.2f}$"
                    
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                # 为合并的单元格添加边框
                for col in range(2, 7):
                    border_cell = ws.cell(row=current_row, column=col)
                    border_cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                current_row += 1
                scci_index += 1
        # 添加上部结构总评分行
        cell = ws.cell(row=current_row, column=1)
        cell.value = f"{structure_name}评分"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 合并B到F列显示SSCI公式 - 使用LaTeX格式
        ws.merge_cells(f'B{current_row}:F{current_row}')
        cell = ws.cell(row=current_row, column=2)
        # 构建SPCI公式字符串
        formula_parts = []
        for i, component_name in enumerate(existing_components):
            if component_name in component_scores and component_name in adjusted_weights:
                score = component_scores[component_name]
                weight = adjusted_weights[component_name]
                formula_parts.append(f"{score:.2f}\\times{weight:.2f}")

        # 弹性换行：根据部件数量决定换行策略
        if len(formula_parts) <= 2:
            # 部件数量少，不换行
            formula_str = "+".join(formula_parts)
        elif len(formula_parts) <= 4:
            # 部件数量中等，每2项换行
            formula_str = "+".join(formula_parts[:2])+"\\\\+"+"+".join(formula_parts[2:])
        else:
            # 部件数量多，每3项换行
            formula_lines = []
            for i in range(0, len(formula_parts), 3):
                line_parts = formula_parts[i:i+3]
                if i == 0:
                    # 第一行不需要前置+号
                    formula_lines.append("+".join(line_parts))
                else:
                    # 后续行需要前置+号
                    formula_lines.append("+"+"+".join(line_parts))
            formula_str = "\\\\".join(formula_lines)

        # 使用LaTeX格式的求和公式，包含完整的上下标
        cell.value = f"$SPCI=\\sum_{{i=1}}^{{{total_components}}}PCCI_i\\timesW_i={formula_str}={bridge_score:.2f}$"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        # 为合并的单元格添加边框
        for col in range(2, 7):
            border_cell = ws.cell(row=current_row, column=col)
            border_cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        # 记录表格的最后一行
        last_row = current_row
        # 修正双线外边框位置 - 从表头行（第2行）开始而不是标题行
        # 顶部边框（从第2行开始）
        for col in range(1, 7):
            cell = ws.cell(row=2, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=Side(style='double'),
                bottom=current_border.bottom
            )
        # 底部边框
        for col in range(1, 7):
            cell = ws.cell(row=last_row, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=Side(style='double')
            )
        # 左侧边框（从第2行开始）
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=1)
            current_border = cell.border
            cell.border = Border(
                left=Side(style='double'),
                right=current_border.right,
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 右侧边框（从第2行开始）
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=6)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=Side(style='double'),
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 设置列宽 - 增加扣分值列的宽度
        column_widths = [18, 10, 15, 8, 60, 18]  
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        # 设置行高
        for row in range(1, current_row + 1):
            ws.row_dimensions[row].height = 25

    def create_substructure_sheet(self, ws):
        """创建下部结构技术状况SBCI评分表"""
        structure_name = '下部结构'
        # 设置标题
        title = f"表7.3 {structure_name}技术状况SBCI评分表"
        ws.merge_cells('A1:F1')
        ws['A1'] = title
        ws['A1'].font = Font(name='Times New Roman', size=12, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        # 设置表头 - 使用LaTeX格式
        headers = ['部件名称', '构件数量', '损坏类型', '标度', '病害描述', '扣分值$DP_{ij}$']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        current_row = 3
        # 从JSON文件中获取下部结构存在的构件
        structure_existence = self.project_data.get('component_existence', {}).get('下部结构', {})
        existing_components = []  # 记录存在的构件顺序
        # 遍历下部结构中存在的构件
        for component_name, exists in structure_existence.items():
            if not exists:  # 跳过不存在的构件
                continue
            existing_components.append(component_name)  # 记录存在的构件
            component_key = f"{structure_name}-{component_name}"
            damages = self.component_damage_data.get(component_key, [])
            total_quantity = self.project_data.get('component_quantities', {}).get(component_key, 0)
            if total_quantity == 0:
                continue  # 跳过数量为0的部件
            # 按构件编号分组病害数据
            damage_by_component = {}
            for damage in damages:
                component_id = damage.get('构件编号', '')
                if component_id not in damage_by_component:
                    damage_by_component[component_id] = []
                damage_by_component[component_id].append(damage)
            damaged_components_count = len(damage_by_component)  # 有病害的构件数量
            undamaged_components_count = total_quantity - damaged_components_count  # 无病害的构件数量
            start_row = current_row
            # 处理有病害的构件
            for component_id, component_damages in damage_by_component.items():
                component_start_row = current_row
                # 为该构件的每个病害创建一行
                for i, damage in enumerate(component_damages):
                    # 部件名称（只在该构件的第一行显示）
                    cell = ws.cell(row=current_row, column=1)
                    if current_row == start_row:  # 整个部件的第一行
                        cell.value = component_name
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                    # 构件数量（只在该构件的第一行显示1）
                    cell = ws.cell(row=current_row, column=2)
                    if i == 0:  # 该构件的第一行
                        cell.value = 1
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 损坏类型
                    cell = ws.cell(row=current_row, column=3)
                    cell.value = damage.get('病害类型', '')
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 标度
                    cell = ws.cell(row=current_row, column=4)
                    cell.value = f"{damage.get('评定标度', '1')}类"
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 病害描述
                    cell = ws.cell(row=current_row, column=5)
                    cell.value = damage.get('病害描述', '')
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    # 扣分值
                    cell = ws.cell(row=current_row, column=6)
                    cell.value = damage.get('扣分值', 0)
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    current_row += 1
                # 合并该构件的构件数量列（如果有多个病害）
                if len(component_damages) > 1:
                    ws.merge_cells(f'B{component_start_row}:B{current_row-1}')
            # 处理无病害的构件（如果有）
            if undamaged_components_count > 0:
                # 部件名称（如果没有病害构件，这是第一行）
                cell = ws.cell(row=current_row, column=1)
                if damaged_components_count == 0:
                    cell.value = component_name
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 构件数量（显示无病害构件的数量）
                cell = ws.cell(row=current_row, column=2)
                cell.value = undamaged_components_count
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 损坏类型
                cell = ws.cell(row=current_row, column=3)
                cell.value = "无"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 标度
                cell = ws.cell(row=current_row, column=4)
                cell.value = "1类"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 病害描述
                cell = ws.cell(row=current_row, column=5)
                cell.value = "——"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 扣分值
                cell = ws.cell(row=current_row, column=6)
                cell.value = 0
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                current_row += 1
            # 合并部件名称单元格（如果该部件有多行）
            total_rows = sum(len(damages) for damages in damage_by_component.values()) + (1 if undamaged_components_count > 0 else 0)
            if total_rows > 1:
                ws.merge_cells(f'A{start_row}:A{current_row-1}')
        # 添加部件评分表格
        component_scores = self.project_data.get('component_scores', {}).get('下部结构', {})
        adjusted_weights = self.project_data.get('adjusted_weights', {}).get('下部结构', {})
        bridge_score = self.project_data.get('evaluation_results', {}).get('bridge_structure_scores', {}).get('下部结构', 0)
        total_components = len(existing_components)  # 记录下部结构中存在的部件总数
        # 为每个存在的部件添加评分行
        pcci_index = 1
        for component_name in existing_components:
            if component_name in component_scores:
                # 部件名称评分行
                cell = ws.cell(row=current_row, column=1)
                cell.value = f"{component_name}评分"
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                # 合并B到F列显示PCCI公式
                ws.merge_cells(f'B{current_row}:F{current_row}')
                cell = ws.cell(row=current_row, column=2)
                
                # 获取详细计算信息
                calculation_details = self.project_data.get('component_calculation_details', {}).get('下部结构', {}).get(component_name, {})
                calculation_method = calculation_details.get('calculation_method', '')
                t_value = calculation_details.get('t_value', None)
                final_score = component_scores[component_name]
                
                # 根据t_value和calculation_method决定显示方式
                if t_value is None:
                    # 无病害部件：t_value为null，calculation_method为空，直接显示最终分数
                    cell.value = f"$BCCI_{{{pcci_index}}}=\\overline{{BMCI}}-\\frac{{100-BMCI_{{min}}}}{{t}}={final_score:.2f}$"
                elif str(t_value) == "Infinity":
                    # 构件数量为1的特殊情况：t_value为无穷大
                    if calculation_method == '标准计算公式':
                        avg_score = calculation_details.get('avg_score', 100)
                        min_score = calculation_details.get('min_score', 100)
                        cell.value = f"$BCCI_{{{pcci_index}}}=\\overline{{BMCI}}-\\frac{{100-BMCI_{{min}}}}{{t}}={avg_score:.2f}-\\frac{{100-{min_score:.2f}}}{{\\infty}}={final_score:.2f}$"
                    else:
                        cell.value = f"$BCCI_{{{pcci_index}}}=BMCI_{{min}}={final_score:.2f}$"
                elif calculation_method == '标准计算公式' and t_value is not None:
                    # 标准计算公式：显示详细计算过程
                    avg_score = calculation_details.get('avg_score', 100)
                    min_score = calculation_details.get('min_score', 100)
                    cell.value = f"$BCCI_{{{pcci_index}}}=\\overline{{BMCI}}-\\frac{{100-BMCI_{{min}}}}{{t}}={avg_score:.2f}-\\frac{{100-{min_score:.2f}}}{{{float(t_value):.2f}}}={final_score:.2f}$"
                elif calculation_method == '主要部件特殊规则':
                    # 主要部件特殊规则：直接显示最终分数
                    cell.value = f"$BCCI_{{{pcci_index}}}=BMCI_{{min}}={final_score:.2f}$"
                else:
                    # 其他情况：显示公式前缀和最终分数
                    cell.value = f"$BCCI_{{{pcci_index}}}=\\overline{{BMCI}}-\\frac{{100-BMCI_{{min}}}}{{t}}={final_score:.2f}$"
                    
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                # 为合并的单元格添加边框
                for col in range(2, 7):
                    border_cell = ws.cell(row=current_row, column=col)
                    border_cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                current_row += 1
                pcci_index += 1
        # 添加下部结构总评分行
        cell = ws.cell(row=current_row, column=1)
        cell.value = f"{structure_name}评分"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 合并B到F列显示PSCI公式 - 使用LaTeX格式
        ws.merge_cells(f'B{current_row}:F{current_row}')
        cell = ws.cell(row=current_row, column=2)
        # 构建SBCI公式字符串
        formula_parts = []
        for i, component_name in enumerate(existing_components):
            if component_name in component_scores and component_name in adjusted_weights:
                score = component_scores[component_name]
                weight = adjusted_weights[component_name]
                formula_parts.append(f"{score:.2f}\\times{weight:.2f}")

        # 弹性换行：根据部件数量决定换行策略
        if len(formula_parts) <= 2:
            # 部件数量少，不换行
            formula_str = "+".join(formula_parts)
        elif len(formula_parts) <= 4:
            # 部件数量中等，每2项换行
            formula_str = "+".join(formula_parts[:2])+"\\\\+"+"+".join(formula_parts[2:])
        else:
            # 部件数量多，每3项换行
            formula_lines = []
            for i in range(0, len(formula_parts), 3):
                line_parts = formula_parts[i:i+3]
                if i == 0:
                    # 第一行不需要前置+号
                    formula_lines.append("+".join(line_parts))
                else:
                    # 后续行需要前置+号
                    formula_lines.append("+"+"+".join(line_parts))
            formula_str = "\\\\".join(formula_lines)
        # 使用LaTeX格式的求和公式，包含完整的上下标
        cell.value = f"$SBCI=\\sum_{{i=1}}^{{{total_components}}}SCCI_i\\timesW_i={formula_str}={bridge_score:.2f}$"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        # 为合并的单元格添加边框
        for col in range(2, 7):
            border_cell = ws.cell(row=current_row, column=col)
            border_cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        # 记录表格的最后一行
        last_row = current_row
        # 修正双线外边框位置 - 从表头行（第2行）开始而不是标题行
        # 顶部边框（从第2行开始）
        for col in range(1, 7):
            cell = ws.cell(row=2, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=Side(style='double'),
                bottom=current_border.bottom
            )
        # 底部边框
        for col in range(1, 7):
            cell = ws.cell(row=last_row, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=Side(style='double')
            )
        # 左侧边框（从第2行开始）
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=1)
            current_border = cell.border
            cell.border = Border(
                left=Side(style='double'),
                right=current_border.right,
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 右侧边框（从第2行开始）
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=6)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=Side(style='double'),
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 设置列宽 - 增加扣分值列的宽度
        column_widths = [18, 10, 15, 8, 60, 18]  
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        # 设置行高
        for row in range(1, current_row + 1):
            ws.row_dimensions[row].height = 25

    def create_total_bridge_sheet(self, ws):
        """创建全桥技术状况Dr评分表"""
        # 设置标题
        title = "表7.4 桥梁技术状况Dr评分计算表"
        ws.merge_cells('A1:C1')
        ws['A1'] = title
        ws['A1'].font = Font(name='Times New Roman', size=12, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        # 设置表头
        headers = ['Dr计算值', '技术状况等级', '技术状况描述']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        # 从JSON数据中获取全桥评分信息
        evaluation_results = self.project_data.get('evaluation_results', {})
        total_score = evaluation_results.get('total_bridge_score', 0)
        total_grade = evaluation_results.get('total_bridge_grade', '1类')
        # 获取各结构部位的评分
        bridge_scores = evaluation_results.get('bridge_structure_scores', {})
        bdci = bridge_scores.get('桥面系', 0)
        spci = bridge_scores.get('上部结构', 0) 
        sbci = bridge_scores.get('下部结构', 0)
        # 根据评级确定技术状况描述
        grade_descriptions = {
            '1类': '全新状态，功能完好。',
            '2类': '桥梁有轻微缺损，对使用功能无影响。',
            '3类': '有中度缺损，尚能维持正常使用功能。',
            '4类': '主要构件有较大缺损，严重影响桥梁使用功能；或影响承载能力，不能保证正常使用。',
            '5类': '主要构件存在严重缺损，不能正常使用，危机桥梁安全，桥梁处于危险状态。'
        }
        description = grade_descriptions.get(total_grade, '未知状态')
        # 填充数据行
        current_row = 3
        # Dr计算值列 - 显示完整的计算公式
        cell = ws.cell(row=current_row, column=1)
        cell.value = f"$Dr={bdci:.2f}\\times0.2+{spci:.2f}\\times0.4+{sbci:.2f}\\times0.4={total_score:.2f}$"
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 技术状况等级列
        cell = ws.cell(row=current_row, column=2)
        cell.value = total_grade
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 技术状况描述列
        cell = ws.cell(row=current_row, column=3)
        cell.value = description
        cell.font = Font(name='Times New Roman', size=10.5)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 设置双线外边框
        # 顶部边框（从第2行开始）
        for col in range(1, 4):
            cell = ws.cell(row=2, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=Side(style='double'),
                bottom=current_border.bottom
            )
        # 底部边框
        for col in range(1, 4):
            cell = ws.cell(row=current_row, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=Side(style='double')
            )
        # 左侧边框（从第2行开始）
        for row in range(2, current_row + 1):
            cell = ws.cell(row=row, column=1)
            current_border = cell.border
            cell.border = Border(
                left=Side(style='double'),
                right=current_border.right,
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 右侧边框（从第2行开始）
        for row in range(2, current_row + 1):
            cell = ws.cell(row=row, column=3)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=Side(style='double'),
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 设置列宽
        column_widths = [60, 15, 50]  # Dr计算值列较宽，描述列最宽
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        # 设置行高
        for row in range(1, current_row + 1):
            ws.row_dimensions[row].height = 25

    def create_summary_sheet(self, ws):
        """创建技术状况总表"""
        # 设置标题
        title = "表8.1 桥梁技术状况评定表"
        ws.merge_cells('A1:E1')
        ws['A1'] = title
        ws['A1'].font = Font(name='Times New Roman', size=12, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        # 设置表头
        headers = ['序号', '结构部位', '技术状况评分', '技术状况等级', '主要病害']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        # 获取评分和等级数据
        evaluation_results = self.project_data.get('evaluation_results', {})
        bridge_scores = evaluation_results.get('bridge_structure_scores', {})
        bridge_grades = evaluation_results.get('bridge_structure_grades', {})
        # 获取病害数据
        component_damage_data = self.project_data.get('component_damage_data', {})
        # 结构部位列表
        structures = ['桥面系', '上部结构', '下部结构']
        current_row = 3
        # 填充各结构部位数据
        for idx, structure in enumerate(structures, 1):
            # 序号
            ws.cell(row=current_row, column=1).value = idx
            ws.cell(row=current_row, column=1).font = Font(name='Times New Roman', size=10.5)
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
            # 位置
            ws.cell(row=current_row, column=2).value = structure
            ws.cell(row=current_row, column=2).font = Font(name='Times New Roman', size=10.5)
            ws.cell(row=current_row, column=2).alignment = Alignment(horizontal='center', vertical='center')
            # 技术状况评分
            score = bridge_scores.get(structure, '--')
            ws.cell(row=current_row, column=3).value = score
            ws.cell(row=current_row, column=3).font = Font(name='Times New Roman', size=10.5)
            ws.cell(row=current_row, column=3).alignment = Alignment(horizontal='center', vertical='center')
            # 技术状况等级
            grade = bridge_grades.get(structure, '--')
            ws.cell(row=current_row, column=4).value = grade
            ws.cell(row=current_row, column=4).font = Font(name='Times New Roman', size=10.5)
            ws.cell(row=current_row, column=4).alignment = Alignment(horizontal='center', vertical='center')
            # 主要病害 - 使用富文本格式
            damage_description = self._get_structure_damage_simple_format(structure, component_damage_data)
            damage_cell = ws.cell(row=current_row, column=5)
            damage_cell.value = damage_description
            damage_cell.font = Font(name='Times New Roman', size=10.5)
            damage_cell.alignment = Alignment(horizontal='justify', vertical='top', wrap_text=True)  # 两端对齐
            # 设置边框
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            current_row += 1
        # 最后一行：合并所有列显示全桥评分信息
        ws.merge_cells(f'A{current_row}:E{current_row}')
        # 获取全桥评分数据
        total_score = evaluation_results.get('total_bridge_score', '--')
        total_grade = evaluation_results.get('total_bridge_grade', '--')
        # 技术状况描述
        status_description = self._get_bridge_status_description(total_grade)
        # 组合全桥评分信息
        bridge_summary = f"桥梁技术状况评分Dr={total_score}，属于{total_grade}桥，即{status_description}"
        cell = ws.cell(row=current_row, column=1)
        cell.value = bridge_summary
        cell.font = Font(name='Times New Roman', size=10.5, bold=True)  # 加粗字体
        cell.alignment = Alignment(horizontal='left', vertical='center')  # 左对齐
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        # 设置双线外边框
        # 顶部边框（从第2行开始）
        for col in range(1, 6):
            cell = ws.cell(row=2, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=Side(style='double'),
                bottom=current_border.bottom
            )
        # 底部边框
        for col in range(1, 6):
            cell = ws.cell(row=current_row, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=Side(style='double')
            )
        # 左侧边框（从第2行开始）
        for row in range(2, current_row + 1):
            cell = ws.cell(row=row, column=1)
            current_border = cell.border
            cell.border = Border(
                left=Side(style='double'),
                right=current_border.right,
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 右侧边框（从第2行开始）
        for row in range(2, current_row + 1):
            cell = ws.cell(row=row, column=5)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=Side(style='double'),
                top=current_border.top,
                bottom=current_border.bottom
            )
        # 设置列宽
        ws.column_dimensions['A'].width = 8   # 序号
        ws.column_dimensions['B'].width = 15  # 位置
        ws.column_dimensions['C'].width = 18  # 技术状况评分
        ws.column_dimensions['D'].width = 15  # 技术状况等级
        ws.column_dimensions['E'].width = 75  # 主要病害
        # 设置行高 - 让行高自动调整以适应内容
        ws.row_dimensions[1].height = 25  # 标题行
        ws.row_dimensions[2].height = 25  # 表头行
        # 数据行不设置固定高度，让其自动调整
        ws.row_dimensions[current_row].height = 25  # 最后一行

    def _get_structure_damage_simple_format(self, structure_name, component_damage_data):
        """获取指定结构部位的病害汇总描述（简单格式，用特殊符号区分部件名称）"""
        # 根据结构名称筛选相关的部件
        structure_components = {}
        for key, damages in component_damage_data.items():
            if key.startswith(structure_name + '-'):
                component_name = key.split('-', 1)[1]  # 提取部件名称
                if damages:  # 只处理有病害的部件
                    structure_components[component_name] = damages
        if not structure_components:
            return '无明显病害'
        damage_summary = []
        # 按部件组织病害描述
        for component_name, damages in structure_components.items():
            if damages:
                component_damages = []
                # 按构件编号分组病害
                component_groups = {}
                for damage in damages:
                    component_id = damage.get('构件编号', '')
                    if component_id not in component_groups:
                        component_groups[component_id] = []
                    component_groups[component_id].append(damage.get('病害描述', ''))
                # 组织每个构件的病害描述
                for component_id, descriptions in component_groups.items():
                    if descriptions:
                        # 将同一构件的多个病害用逗号连接
                        component_damage_text = '、'.join(descriptions)
                        component_damages.append(component_damage_text)
                # 将同一部件下不同构件的病害用分号分隔，最后加句号
                if component_damages:
                    # 使用★符号来突出显示部件名称
                    component_text = f"➤{component_name}：{'; '.join(component_damages)}。"
                    damage_summary.append(component_text)
        return '\n'.join(damage_summary)

    def _get_bridge_status_description(self, grade):
        """根据技术状况等级获取描述"""
        descriptions = {
            '1类': '全新状态，功能完好。',
            '2类': '桥梁有轻微缺损，对使用功能无影响。',
            '3类': '有中度缺损，尚能维持正常使用功能。',
            '4类': '主要构件有较大缺损，严重影响桥梁使用功能；或影响承载能力，不能保证正常使用。',
            '5类': '主要构件存在严重缺损，不能正常使用，危机桥梁安全，桥梁处于危险状态。'
        }
        return descriptions.get(grade, '技术状况等级未知')


def main():
    app = QApplication(sys.argv)
    # 设置应用程序字体
    font = QFont("楷体", 12)
    app.setFont(font)
    window = BridgeEvaluationSystemBase()  
    window.show()
    sys.exit(app.exec())
if __name__ == "__main__":
    main()