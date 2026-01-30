"""
智能抽签系统 - 图片配色版
功能：从 Excel 中按省区随机抽取人员，生成标记结果的新 Excel
"""

import sys
import pandas as pd
from datetime import datetime
import random
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QListWidget,
    QTextEdit, QMessageBox, QFileDialog, QFrame,
    QScrollArea, QGridLayout, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor


# 根据图片提取的配色方案（清新浅色风格）
COLORS = {
    # 主色调 - 浅蓝紫色系
    'primary': '#76C5FF',           # 浅蓝色（主按钮）
    'primary_dark': '#5BA8E8',      # 深蓝色
    'primary_light': '#A6DCFF',     # 浅蓝色（悬停）
    'secondary': '#7465EB',         # 紫蓝色（辅助）
    'secondary_dark': '#5D4FD1',    # 深紫色
    'secondary_light': '#9A8FF3',   # 浅紫色

    # 功能色
    'success': '#D4EDDA',           # 浅绿色（成功背景）
    'success_text': '#155724',      # 深绿色（成功文字）
    'success_dark': '#C3E6CB',      # 深绿色背景

    'warning': '#FFF3CD',           # 浅黄色（警告背景）
    'warning_text': '#856404',      # 深黄色（警告文字）

    'danger': '#FE767F',            # 浅红色（危险按钮）
    'danger_dark': '#F45560',       # 深红色
    'danger_text': '#721C24',       # 深红色文字

    # 背景色
    'bg_main': '#FAFAFA',           # 主背景（浅灰）
    'bg_card': '#FFFFFF',           # 卡片背景（白色）
    'bg_input': '#F8F9FA',          # 输入框背景
    'bg_hover': '#E6F0F7',          # 悬停背景（浅蓝灰）
    'bg_selected': '#E8F2F9',       # 选中背景

    # 边框色
    'border_light': '#E8F2F9',      # 浅边框
    'border': '#D8EBF3',            # 边框色
    'border_dark': '#C4D7E3',       # 深边框

    # 文字色
    'text_primary': '#2C3E50',      # 主文本（深灰蓝）
    'text_secondary': '#6C757D',    # 次要文本（灰）
    'text_light': '#ADB5BD',        # 浅色文本
    'text_white': '#FFFFFF',        # 白色文字
}


class CleanButton(QPushButton):
    """清新按钮"""
    def __init__(self, text, color_type='primary', parent=None):
        super().__init__(text, parent)
        self.color_type = color_type
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(32)
        self._apply_style()

    def _apply_style(self):
        styles = {
            'primary': {
                'bg': COLORS['primary'],
                'bg_hover': COLORS['primary_light'],
                'text': '#FFFFFF',
                'shadow': '#5BA8E8',
            },
            'secondary': {
                'bg': COLORS['secondary'],
                'bg_hover': COLORS['secondary_light'],
                'text': '#FFFFFF',
                'shadow': COLORS['secondary_dark'],
            },
            'success': {
                'bg': COLORS['success'],
                'bg_hover': COLORS['success_dark'],
                'text': COLORS['success_text'],
                'shadow': '#B0D9B6',
            },
            'warning': {
                'bg': COLORS['warning'],
                'bg_hover': '#FFE69C',
                'text': COLORS['warning_text'],
                'shadow': '#F0E5A8',
            },
            'danger': {
                'bg': COLORS['danger'],
                'bg_hover': '#FF8A92',
                'text': '#FFFFFF',
                'shadow': COLORS['danger_dark'],
            },
            'outline': {
                'bg': '#FFFFFF',
                'bg_hover': COLORS['bg_hover'],
                'text': COLORS['primary'],
                'border': COLORS['border'],
            },
        }

        s = styles.get(self.color_type, styles['primary'])

        if self.color_type == 'outline':
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {s['bg']};
                    color: {s['text']};
                    border: 2px solid {s['border']};
                    border-radius: 6px;
                    padding: 6px 16px;
                    font-size: 12px;
                    font-weight: 600;
                }}
                QPushButton:hover {{
                    background-color: {s['bg_hover']};
                    border-color: {COLORS['primary']};
                }}
                QPushButton:pressed {{
                    background-color: {COLORS['bg_selected']};
                }}
                QPushButton:disabled {{
                    background-color: #F8F9FA;
                    color: {COLORS['text_light']};
                    border-color: {COLORS['border_light']};
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {s['bg']};
                    color: {s['text']};
                    border: none;
                    border-radius: 6px;
                    padding: 6px 16px;
                    font-size: 12px;
                    font-weight: 600;
                }}
                QPushButton:hover {{
                    background-color: {s['bg_hover']};
                }}
                QPushButton:pressed {{
                    background-color: {s['bg']};
                }}
                QPushButton:disabled {{
                    background-color: #E9ECEF;
                    color: {COLORS['text_light']};
                }}
            """)


class CleanCard(QFrame):
    """清新卡片"""
    def __init__(self, title, icon='', parent=None):
        super().__init__(parent)
        self.title = title
        self.icon = icon
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['bg_card']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                padding: 0px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 标题栏
        title_widget = QWidget()
        title_widget.setStyleSheet(f"""
            QWidget {{
                background-color: {COLORS['bg_hover']};
                border-top-left-radius: 7px;
                border-top-right-radius: 7px;
            }}
        """)

        title_layout = QHBoxLayout(title_widget)
        title_layout.setContentsMargins(12, 10, 12, 10)

        icon_label = QLabel(self.icon)
        icon_label.setStyleSheet("font-size: 14px;")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_label.setFixedSize(20, 20)

        title_label = QLabel(self.title)
        title_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 13px;
                font-weight: 700;
            }}
        """)

        title_layout.addWidget(icon_label)
        title_layout.addWidget(title_label)
        title_layout.addStretch()

        layout.addWidget(title_widget)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        separator.setStyleSheet(f"QFrame {{ background-color: {COLORS['border']}; max-height: 1px; }}")
        layout.addWidget(separator)

        # 内容区域
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(12, 12, 12, 12)
        self.content_layout.setSpacing(10)

        layout.addWidget(self.content_widget)

    def add_widget(self, widget):
        self.content_layout.addWidget(widget)

    def add_layout(self, layout):
        self.content_layout.addLayout(layout)


class CleanLineEdit(QLineEdit):
    """清新输入框"""
    def __init__(self, placeholder='', parent=None):
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setMinimumHeight(30)
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet(f"""
            QLineEdit {{
                background-color: {COLORS['bg_input']};
                border: 2px solid {COLORS['border']};
                border-radius: 6px;
                padding: 6px 10px;
                font-size: 12px;
                color: {COLORS['text_primary']};
            }}
            QLineEdit:focus {{
                border-color: {COLORS['primary']};
                background-color: #FFFFFF;
            }}
            QLineEdit::placeholder {{
                color: {COLORS['text_light']};
            }}
        """)


class CleanListWidget(QListWidget):
    """清新列表"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet(f"""
            QListWidget {{
                background-color: {COLORS['bg_input']};
                border: 2px solid {COLORS['border']};
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }}
            QListWidget::item {{
                padding: 6px 10px;
                border-radius: 6px;
                margin: 1px;
                background-color: transparent;
            }}
            QListWidget::item:hover {{
                background-color: {COLORS['bg_hover']};
            }}
            QListWidget::item:selected {{
                background-color: {COLORS['primary']};
                color: white;
            }}
        """)


class CleanTableWidget(QTableWidget):
    """清新表格"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_table()

    def _setup_table(self):
        self.setColumnCount(5)
        self.setHorizontalHeaderLabels(['序号', 'Excel行号', 'ID', '姓名', '省区'])

        # 设置表格样式
        self.setStyleSheet(f"""
            QTableWidget {{
                background-color: {COLORS['bg_input']};
                border: 2px solid {COLORS['border']};
                border-radius: 6px;
                gridline-color: {COLORS['border_light']};
            }}
            QTableWidget::item {{
                padding: 3px;
                border-bottom: 1px solid {COLORS['border_light']};
            }}
            QTableWidget::item:selected {{
                background-color: {COLORS['bg_selected']};
                color: {COLORS['text_primary']};
            }}
            QHeaderView::section {{
                background-color: {COLORS['bg_hover']};
                color: {COLORS['text_primary']};
                padding: 5px;
                border: none;
                border-bottom: 2px solid {COLORS['border']};
                font-size: 12px;
                font-weight: 700;
            }}
            QTableCornerButton::section {{
                background-color: {COLORS['bg_hover']};
                border: none;
            }}
        """)

        # 设置行高
        vertical_header = self.verticalHeader()
        vertical_header.setVisible(False)
        vertical_header.setDefaultSectionSize(24)

        # 设置列宽
        horizontal_header = self.horizontalHeader()
        horizontal_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        horizontal_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        horizontal_header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        horizontal_header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        horizontal_header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        self.setColumnWidth(0, 45)
        self.setColumnWidth(1, 80)

        # 设置选择行为
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        # 设置编辑模式
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        # 设置交替行颜色
        self.setAlternatingRowColors(True)
        self.setStyleSheet(self.styleSheet() + f"""
            QTableWidget {{
                alternate-background-color: {COLORS['bg_card']};
            }}
        """)

    def add_result_row(self, index, row_num, id_num, name, province):
        """添加结果行"""
        row_position = self.rowCount()
        self.insertRow(row_position)

        # 序号
        item_num = QTableWidgetItem(str(index))
        item_num.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        item_num.setForeground(QColor(COLORS['primary']))
        font = item_num.font()
        font.setBold(True)
        item_num.setFont(font)
        self.setItem(row_position, 0, item_num)

        # Excel行号
        item_row = QTableWidgetItem(str(row_num))
        item_row.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setItem(row_position, 1, item_row)

        # ID
        item_id = QTableWidgetItem(str(id_num))
        item_id.setForeground(QColor(COLORS['text_primary']))
        self.setItem(row_position, 2, item_id)

        # 姓名
        item_name = QTableWidgetItem(str(name))
        item_name.setForeground(QColor(COLORS['text_primary']))
        font_name = item_name.font()
        font_name.setBold(True)
        item_name.setFont(font_name)
        self.setItem(row_position, 3, item_name)

        # 省区
        item_prov = QTableWidgetItem(str(province))
        item_prov.setForeground(QColor(COLORS['text_secondary']))
        self.setItem(row_position, 4, item_prov)


class RandomDrawApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.provinces = []
        self.drawn_result = None
        self.original_df = None
        self.all_drawn_people = None  # 所有已抽中的人员
        self.draw_count = 0  # 抽签次数
        self.export_file_path = None  # 导出文件路径
        self.is_ended = False  # 是否已结束抽签

        self._setup_window()
        self._setup_ui()

        # 自动加载默认文件
        default_file = "工作簿1.xlsx"
        if os.path.exists(default_file):
            self.load_excel(default_file)

    def _setup_window(self):
        self.setWindowTitle('🎲 抽签')
        self.setGeometry(100, 100, 760, 700)
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {COLORS['bg_main']};
            }}
        """)

    def _setup_ui(self):
        # 主容器
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(12)

        # 标题区域
        title_container = QWidget()
        title_container.setStyleSheet(f"""
            QWidget {{
                background-color: {COLORS['primary']};
                border-radius: 6px;
                padding: 6px 12px;
            }}
        """)

        title_layout = QVBoxLayout(title_container)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(0)

        title_label = QLabel('🎲 抽签')
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet(f"""
            QLabel {{
                color: white;
                font-size: 16px;
                font-weight: 700;
                letter-spacing: 1px;
            }}
        """)

        title_layout.addWidget(title_label)

        main_layout.addWidget(title_container)

        # 网格布局
        grid_layout = QGridLayout()
        grid_layout.setSpacing(12)

        # 文件上传卡片
        file_card = CleanCard('📁 数据源', '')
        file_input_layout = QHBoxLayout()
        self.file_path_edit = CleanLineEdit('点击浏览选择 Excel 文件...')
        self.file_path_edit.setReadOnly(True)

        browse_btn = CleanButton('浏览', 'outline')
        browse_btn.setMinimumWidth(70)
        browse_btn.clicked.connect(self.browse_file)

        load_btn = CleanButton('加载', 'primary')
        load_btn.setMinimumWidth(70)
        load_btn.clicked.connect(self.load_selected_file)

        file_input_layout.addWidget(self.file_path_edit, 1)
        file_input_layout.addWidget(browse_btn)
        file_input_layout.addWidget(load_btn)
        file_card.add_layout(file_input_layout)

        # 状态标签
        self.status_label = QLabel('⏳ 等待加载文件...')
        self.status_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_secondary']};
                font-size: 11px;
                padding: 6px 10px;
                background-color: {COLORS['bg_input']};
                border-radius: 6px;
                border: 1px solid {COLORS['border']};
            }}
        """)
        file_card.add_widget(self.status_label)

        grid_layout.addWidget(file_card, 0, 0, 1, 2)

        # 省区选择卡片
        province_card = CleanCard('🏢 选择省区', '✓')

        # 按钮行
        btn_row_widget = QWidget()
        btn_layout = QHBoxLayout(btn_row_widget)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        select_all_btn = CleanButton('全选', 'outline')
        select_all_btn.setMinimumWidth(55)
        select_all_btn.clicked.connect(self.select_all)

        clear_btn = CleanButton('清空', 'warning')
        clear_btn.setMinimumWidth(55)
        clear_btn.clicked.connect(self.clear_selection)

        self.selected_count_label = QLabel('已选: 0 个省区')
        self.selected_count_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_white']};
                font-size: 11px;
                font-weight: 600;
                padding: 4px 10px;
                background-color: {COLORS['primary']};
                border-radius: 12px;
            }}
        """)

        btn_layout.addWidget(select_all_btn)
        btn_layout.addWidget(clear_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.selected_count_label)

        province_card.add_widget(btn_row_widget)

        # 省区列表
        self.province_list = CleanListWidget()
        self.province_list.setMaximumHeight(120)
        self.province_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.province_list.itemSelectionChanged.connect(self.on_selection_changed)
        province_card.add_widget(self.province_list)

        # 添加弹性空间，使内容向上对齐
        province_card.content_layout.addStretch()

        grid_layout.addWidget(province_card, 1, 0, 1, 1)

        # 抽取设置卡片
        count_card = CleanCard('🎯 抽取设置', '⚙️')

        count_row = QWidget()
        count_layout = QHBoxLayout(count_row)
        count_layout.setContentsMargins(0, 0, 0, 0)
        count_layout.setSpacing(8)

        count_label = QLabel('📊 抽取人数：')
        count_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 13px;
                font-weight: 700;
                padding: 4px 0px;
            }}
        """)

        self.count_input = CleanLineEdit('5')
        self.count_input.setFixedWidth(80)

        count_layout.addWidget(count_label)
        count_layout.addWidget(self.count_input)
        count_layout.addStretch()

        count_card.add_widget(count_row)

        # 操作按钮
        action_row = QWidget()
        action_layout = QVBoxLayout(action_row)
        action_layout.setContentsMargins(0, 12, 0, 0)
        action_layout.setSpacing(10)
        action_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 第一行：开始抽签和导出结果
        first_row_widget = QWidget()
        first_row_layout = QHBoxLayout(first_row_widget)
        first_row_layout.setContentsMargins(0, 0, 0, 0)
        first_row_layout.setSpacing(10)

        self.draw_btn = CleanButton('🎲 开始抽签', 'primary')
        self.draw_btn.setMinimumWidth(165)
        self.draw_btn.setMinimumHeight(40)
        self.draw_btn.clicked.connect(self.start_draw)
        self.draw_btn.setEnabled(False)

        self.export_btn = CleanButton('📥 导出结果', 'success')
        self.export_btn.setMinimumWidth(165)
        self.export_btn.setMinimumHeight(40)
        self.export_btn.clicked.connect(self.export_result)
        self.export_btn.setEnabled(False)

        first_row_layout.addWidget(self.draw_btn)
        first_row_layout.addWidget(self.export_btn)

        # 第二行：结束抽签（居中）
        self.end_btn = CleanButton('⏹ 结束抽签', 'danger')
        self.end_btn.setMinimumWidth(165)
        self.end_btn.setMinimumHeight(40)
        self.end_btn.clicked.connect(self.end_draw)
        self.end_btn.setEnabled(False)

        action_layout.addWidget(first_row_widget)
        action_layout.addWidget(self.end_btn)
        action_layout.setAlignment(self.end_btn, Qt.AlignmentFlag.AlignCenter)

        count_card.add_widget(action_row)

        # 添加弹性空间，使内容向上对齐
        count_card.content_layout.addStretch()

        grid_layout.addWidget(count_card, 1, 1, 1, 1)

        main_layout.addLayout(grid_layout)

        # 结果展示卡片
        result_card = CleanCard('📊 抽签结果', '🏆')

        # 结果统计
        self.result_stats_label = QLabel('💡 提示：请先选择省区并开始抽签')
        self.result_stats_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_secondary']};
                font-size: 11px;
                padding: 6px 10px;
                background-color: {COLORS['bg_input']};
                border-radius: 6px;
                border: 1px solid {COLORS['border']};
            }}
        """)
        result_card.add_widget(self.result_stats_label)

        # 结果表格
        self.result_table = CleanTableWidget()
        result_card.add_widget(self.result_table)

        main_layout.addWidget(result_card, 8)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            '选择 Excel 文件',
            '',
            'Excel 文件 (*.xlsx *.xls);;所有文件 (*)'
        )
        if file_path:
            self.file_path_edit.setText(file_path)

    def load_selected_file(self):
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, '提示', '请先选择文件')
            return

        self.load_excel(file_path)

    def load_excel(self, file_path):
        try:
            # 读取 Excel 文件
            self.df = pd.read_excel(file_path)
            self.original_df = self.df.copy()

            # 获取省区列表
            fourth_level_provinces = [
                dept for dept in self.df['四级部门'].dropna().unique().tolist()
                if '省区' in dept
            ]
            third_level_provinces = [
                dept for dept in self.df['三级部门'].dropna().unique().tolist()
                if '独立省区' in dept
            ]

            # 合并并去重
            all_provinces = fourth_level_provinces + third_level_provinces
            self.provinces = sorted(list(set(all_provinces)))

            # 更新省区列表
            self.province_list.clear()
            for province in self.provinces:
                if province in fourth_level_provinces:
                    count = len(self.df[self.df['四级部门'] == province])
                else:
                    count = len(self.df[self.df['三级部门'] == province])

                item_text = f"  {province}  ({count} 人)"
                self.province_list.addItem(item_text)

            # 更新状态
            total_count = len(self.df)
            self.status_label.setText(f'✅ 已加载：{total_count} 人，{len(self.provinces)} 个省区')
            self.status_label.setStyleSheet(f"""
                QLabel {{
                    color: {COLORS['success_text']};
                    font-size: 13px;
                    padding: 10px 14px;
                    background-color: {COLORS['success']};
                    border-radius: 8px;
                    border: 1px solid {COLORS['success_dark']};
                    font-weight: 600;
                }}
            """)

            # 清空结果
            self.result_table.setRowCount(0)
            self.result_stats_label.setText(f'📊 数据已加载，共 {total_count} 人，{len(self.provinces)} 个省区')
            self.result_stats_label.setStyleSheet(f"""
                QLabel {{
                    color: {COLORS['primary_dark']};
                    font-size: 13px;
                    padding: 10px 14px;
                    background-color: {COLORS['bg_selected']};
                    border-radius: 8px;
                    border: 1px solid {COLORS['primary']};
                    font-weight: 600;
                }}
            """)

            QMessageBox.information(
                self,
                '✅ 加载成功',
                f'成功加载 Excel 文件！\n\n📊 总人数：{total_count}\n🏢 省区数：{len(self.provinces)}'
            )

        except Exception as e:
            QMessageBox.critical(self, '❌ 加载失败', f'加载 Excel 文件失败：\n{str(e)}')

    def on_selection_changed(self):
        """处理选择变化"""
        selected_items = self.province_list.selectedItems()
        count = len(selected_items)
        self.selected_count_label.setText(f'已选: {count} 个省区')

        # 更新按钮状态
        self.draw_btn.setEnabled(count > 0 and self.df is not None)

    def select_all(self):
        self.province_list.selectAll()

    def clear_selection(self):
        self.province_list.clearSelection()

    def start_draw(self):
        if self.df is None:
            QMessageBox.warning(self, '⚠️ 提示', '请先加载 Excel 文件')
            return

        if self.is_ended:
            QMessageBox.warning(self, '⚠️ 提示', '抽签已结束，如需重新开始请重新加载文件')
            return

        try:
            draw_count = int(self.count_input.text())
            if draw_count < 1:
                QMessageBox.warning(self, '⚠️ 提示', '抽取人数必须大于 0')
                return
        except ValueError:
            QMessageBox.warning(self, '⚠️ 提示', '请输入有效的抽取人数')
            return

        # 获取选中的省区
        selected_items = self.province_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, '⚠️ 提示', '请至少选择一个省区')
            return

        selected_provinces = []
        for item in selected_items:
            # 解析省区名称
            text = item.text()
            province = text.split('(')[0].strip()
            selected_provinces.append(province)

        # 筛选数据
        filtered_df = None
        fourth_level_provinces = [
            dept for dept in self.df['四级部门'].dropna().unique().tolist()
            if '省区' in dept
        ]

        for province in selected_provinces:
            if province in fourth_level_provinces:
                temp_df = self.df[self.df['四级部门'] == province]
            else:
                temp_df = self.df[self.df['三级部门'] == province]

            if filtered_df is None:
                filtered_df = temp_df
            else:
                filtered_df = pd.concat([filtered_df, temp_df], ignore_index=True)

        if filtered_df is None or len(filtered_df) == 0:
            QMessageBox.warning(self, '⚠️ 提示', '选中的省区中没有数据')
            return

        # 排除已抽中的人员
        if self.all_drawn_people is not None and len(self.all_drawn_people) > 0:
            drawn_ids = self.all_drawn_people['员工 ID'].tolist()
            filtered_df = filtered_df[~filtered_df['员工 ID'].isin(drawn_ids)]

        if len(filtered_df) == 0:
            QMessageBox.warning(
                self,
                '⚠️ 提示',
                '选中的省区中已无未抽中的人员'
            )
            return

        if len(filtered_df) < draw_count:
            QMessageBox.warning(
                self,
                '⚠️ 提示',
                f'选中省区中只有 {len(filtered_df)} 人未抽中，无法抽取 {draw_count} 人'
            )
            return

        # 随机抽取
        self.drawn_result = filtered_df.sample(n=draw_count)

        # 累加到已抽中人员列表
        self.draw_count += 1
        if self.all_drawn_people is None:
            self.all_drawn_people = self.drawn_result.copy()
        else:
            self.all_drawn_people = pd.concat([self.all_drawn_people, self.drawn_result], ignore_index=True)

        # 显示结果
        self._show_result(selected_provinces, draw_count)

        # 启用导出和结束按钮
        self.export_btn.setEnabled(True)
        self.end_btn.setEnabled(True)

        # 自动更新导出文件
        self._auto_update_export()

        QMessageBox.information(
            self,
            '🎉 抽签成功',
            f'✅ 抽签完成！\n\n🎯 本次抽取：{draw_count} 人\n📊 累计抽取：{len(self.all_drawn_people)} 人\n📋 结果已显示在下方'
        )

    def _show_result(self, selected_provinces, draw_count):
        """显示抽签结果"""
        # 清空表格
        self.result_table.setRowCount(0)

        # 显示所有累计抽取的结果（倒序显示，最新的在前面）
        if self.all_drawn_people is not None and len(self.all_drawn_people) > 0:
            for i, (idx, row) in enumerate(self.all_drawn_people.iloc[::-1].iterrows(), 1):
                # 判断省区级别
                if pd.notna(row.get('四级部门')) and '省区' in row['四级部门']:
                    province = row['四级部门']
                elif pd.notna(row.get('三级部门')) and '独立省区' in row['三级部门']:
                    province = row['三级部门']
                else:
                    province = '未知'

                self.result_table.add_result_row(
                    index=i,
                    row_num=idx + 2,
                    id_num=row['员工 ID'],
                    name=row['姓名'],
                    province=province
                )

        # 更新统计
        provinces_str = ', '.join(selected_provinces[:2])
        if len(selected_provinces) > 2:
            provinces_str += f' 等 {len(selected_provinces)} 个省区'

        self.result_stats_label.setText(
            f'🎉 第{self.draw_count}次抽签完成！从 {provinces_str} 中抽取了 {draw_count} 人\n📊 累计抽取：{len(self.all_drawn_people)} 人'
        )
        self.result_stats_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['success_text']};
                font-size: 13px;
                padding: 10px 14px;
                background-color: {COLORS['success']};
                border-radius: 8px;
                border: 1px solid {COLORS['success_dark']};
                font-weight: 600;
            }}
        """)

    def _auto_update_export(self):
        """自动更新导出文件"""
        try:
            # 如果还没有导出文件路径，创建一个
            if self.export_file_path is None:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                self.export_file_path = f'抽签结果_自动更新_{timestamp}.xlsx'

            # 导出原文件，并在"是否被抽中"列标记
            export_df = self.original_df.copy()

            # 确保有"是否被抽中"列
            if '是否被抽中' not in export_df.columns:
                export_df['是否被抽中'] = ''

            # 将所有行的"是否被抽中"设置为空
            export_df['是否被抽中'] = ''

            # 获取所有已抽中人员的ID列表
            drawn_ids = self.all_drawn_people['员工 ID'].tolist()

            # 标记所有抽中的人员
            export_df.loc[export_df['员工 ID'].isin(drawn_ids), '是否被抽中'] = '是'

            # 保存到 Excel
            export_df.to_excel(self.export_file_path, index=False, engine='openpyxl')

        except Exception as e:
            print(f"自动更新导出文件失败：{str(e)}")

    def end_draw(self):
        """结束抽签"""
        if self.all_drawn_people is None or len(self.all_drawn_people) == 0:
            QMessageBox.warning(self, '⚠️ 提示', '还没有进行抽签')
            return

        self.is_ended = True

        # 禁用开始抽签按钮
        self.draw_btn.setEnabled(False)

        # 选择最终导出路径
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f'抽签结果最终_{timestamp}.xlsx'

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            '保存最终结果',
            default_filename,
            'Excel 文件 (*.xlsx);;所有文件 (*)'
        )

        if file_path:
            self.export_file_path = file_path

        try:
            # 导出原文件，并在"是否被抽中"列标记
            export_df = self.original_df.copy()

            # 确保有"是否被抽中"列
            if '是否被抽中' not in export_df.columns:
                export_df['是否被抽中'] = ''

            # 将所有行的"是否被抽中"设置为空
            export_df['是否被抽中'] = ''

            # 获取所有已抽中人员的ID列表
            drawn_ids = self.all_drawn_people['员工 ID'].tolist()

            # 标记所有抽中的人员
            export_df.loc[export_df['员工 ID'].isin(drawn_ids), '是否被抽中'] = '是'

            # 保存到 Excel
            export_df.to_excel(self.export_file_path, index=False, engine='openpyxl')

            QMessageBox.information(
                self,
                '🎊 抽签结束',
                f'✅ 抽签已结束！\n\n📊 总共抽签次数：{self.draw_count} 次\n🎯 累计抽取人数：{len(self.all_drawn_people)} 人\n\n📁 结果已保存到：\n{self.export_file_path}'
            )

        except Exception as e:
            QMessageBox.critical(self, '❌ 导出失败', f'导出失败：\n{str(e)}')

    def export_result(self):
        """导出结果"""
        if self.all_drawn_people is None or len(self.all_drawn_people) == 0:
            QMessageBox.warning(self, '⚠️ 提示', '请先进行抽签')
            return

        # 选择保存路径
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f'抽签结果_{timestamp}.xlsx'

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            '保存结果',
            default_filename,
            'Excel 文件 (*.xlsx);;所有文件 (*)'
        )

        if not file_path:
            return

        try:
            # 导出原文件，并在"是否被抽中"列标记
            export_df = self.original_df.copy()

            # 确保有"是否被抽中"列
            if '是否被抽中' not in export_df.columns:
                export_df['是否被抽中'] = ''

            # 将所有行的"是否被抽中"设置为空
            export_df['是否被抽中'] = ''

            # 获取所有已抽中人员的ID列表
            drawn_ids = self.all_drawn_people['员工 ID'].tolist()

            # 标记所有抽中的人员
            export_df.loc[export_df['员工 ID'].isin(drawn_ids), '是否被抽中'] = '是'

            # 保存到 Excel
            export_df.to_excel(file_path, index=False, engine='openpyxl')

            QMessageBox.information(
                self,
                '✅ 导出成功',
                f'结果已成功导出到：\n{file_path}\n\n📊 共导出 {len(export_df)} 条记录\n✅ 抽中 {len(drawn_ids)} 人'
            )

        except Exception as e:
            QMessageBox.critical(self, '❌ 导出失败', f'导出失败：\n{str(e)}')


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # 设置全局字体
    font = QFont('Microsoft YaHei', 10)
    app.setFont(font)

    window = RandomDrawApp()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
