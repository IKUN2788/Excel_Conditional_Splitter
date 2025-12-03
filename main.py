import sys
import os
import pandas as pd
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox, 
                             QTableWidget, QTableWidgetItem, QMessageBox, QGroupBox, 
                             QRadioButton, QButtonGroup, QStackedWidget, QFormLayout,
                             QHeaderView, QAbstractItemView, QFrame, QCheckBox)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette

# Modern Dark/Light Theme Stylesheet
STYLESHEET = """
QMainWindow {
    background-color: #f5f7fa;
}

QGroupBox {
    font-weight: bold;
    border: 1px solid #c0c4cc;
    border-radius: 6px;
    margin-top: 12px;
    background-color: #ffffff;
    padding-top: 15px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 5px;
    color: #2b85e4;
}

QLabel {
    color: #000000;
    font-size: 14px;
    font-weight: 500;
}

QLineEdit {
    border: 1px solid #c0c4cc;
    border-radius: 4px;
    padding: 5px;
    color: #000000;
    background-color: #ffffff;
    font-weight: 500;
}

QLineEdit:focus {
    border: 1px solid #409eff;
}

QLineEdit:read-only {
    background-color: #f5f7fa;
    color: #333333;
}

QComboBox {
    border: 1px solid #c0c4cc;
    border-radius: 4px;
    padding: 5px;
    min-width: 6em;
    color: #000000;
    font-weight: 500;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left-width: 0px;
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
}

QPushButton {
    background-color: #409eff;
    color: white;
    border: none;
    border-radius: 4px;
    padding: 8px 15px;
    font-weight: bold;
    font-size: 13px;
}

QPushButton:hover {
    background-color: #66b1ff;
}

QPushButton:pressed {
    background-color: #3a8ee6;
}

QPushButton#delete_btn {
    background-color: #f56c6c;
}

QPushButton#delete_btn:hover {
    background-color: #f78989;
}

QPushButton#clear_btn {
    background-color: #909399;
}

QPushButton#clear_btn:hover {
    background-color: #a6a9ad;
}

QTableWidget {
    border: 1px solid #c0c4cc;
    gridline-color: #ebeef5;
    background-color: #ffffff;
    selection-background-color: #ecf5ff;
    selection-color: #000000;
    color: #000000;
}

QHeaderView::section {
    background-color: #f5f7fa;
    padding: 4px;
    border: none;
    border-bottom: 1px solid #c0c4cc;
    color: #000000;
    font-weight: bold;
}

QRadioButton {
    spacing: 5px;
    color: #000000;
    font-weight: 500;
}

QRadioButton::indicator {
    width: 15px;
    height: 15px;
}
"""

class DropArea(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.main_window = parent
        
        layout = QVBoxLayout()
        self.label = QLabel("拖拽Excel文件到此处 或 点击浏览")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color: #333333; font-size: 15px; font-weight: bold; border: 2px dashed #909399; border-radius: 6px; padding: 20px;")
        layout.addWidget(self.label)
        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
            self.label.setStyleSheet("color: #2b85e4; font-size: 15px; font-weight: bold; border: 2px dashed #2b85e4; border-radius: 6px; padding: 20px; background-color: #ecf5ff;")
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.label.setStyleSheet("color: #333333; font-size: 15px; font-weight: bold; border: 2px dashed #909399; border-radius: 6px; padding: 20px;")

    def dropEvent(self, event):
        self.label.setStyleSheet("color: #333333; font-size: 15px; font-weight: bold; border: 2px dashed #909399; border-radius: 6px; padding: 20px;")
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            # Take the first file
            fpath = files[0]
            if fpath.endswith(('.xlsx', '.xls')):
                self.main_window.process_file(fpath)
            else:
                QMessageBox.warning(self, "格式错误", "请拖入Excel文件 (.xlsx, .xls)")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.main_window.open_file_dialog()

class ExcelSplitterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel条件拆分器")
        self.resize(900, 1050)
        
        # Data storage
        self.df = None
        self.file_path = ""
        self.conditions = [] 

        # Apply Styles
        self.setStyleSheet(STYLESHEET)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 1. File Selection Area (Drag & Drop)
        file_group = QGroupBox("1. 选择Excel文件")
        file_layout = QVBoxLayout()
        
        self.drop_area = DropArea(self)
        file_layout.addWidget(self.drop_area)
        
        self.path_display = QLabel("未选择文件")
        self.path_display.setStyleSheet("color: #000000; padding-left: 5px; font-weight: 500;")
        file_layout.addWidget(self.path_display)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # 2. Column Selection
        col_group = QGroupBox("2. 选择拆分依据列")
        col_layout = QHBoxLayout()
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        self.sheet_combo.setPlaceholderText("选择工作表")
        
        self.col_combo = QComboBox()
        self.col_combo.setPlaceholderText("选择目标列")
        
        col_layout.addWidget(QLabel("工作表:"))
        col_layout.addWidget(self.sheet_combo, 1)
        col_layout.addWidget(QLabel("目标列:"))
        col_layout.addWidget(self.col_combo, 1)
        
        col_group.setLayout(col_layout)
        main_layout.addWidget(col_group)

        # 3. Condition Configuration
        cond_group = QGroupBox("3. 添加拆分条件")
        cond_layout = QVBoxLayout()
        cond_layout.setSpacing(10)
        
        # Condition Type Selection
        type_layout = QHBoxLayout()
        self.type_group = QButtonGroup(self)
        self.rb_numeric = QRadioButton("数值范围")
        self.rb_text = QRadioButton("文本包含")
        self.rb_regex = QRadioButton("正则表达式")
        self.rb_numeric.setChecked(True)
        self.type_group.addButton(self.rb_numeric, 0)
        self.type_group.addButton(self.rb_text, 1)
        self.type_group.addButton(self.rb_regex, 2)
        self.type_group.buttonToggled.connect(self.on_type_changed)
        
        type_layout.addWidget(QLabel("条件类型:"))
        type_layout.addWidget(self.rb_numeric)
        type_layout.addWidget(self.rb_text)
        type_layout.addWidget(self.rb_regex)
        
        # Negate Checkbox
        self.chk_negate = QCheckBox("取反 (Not)")
        self.chk_negate.setStyleSheet("QCheckBox { color: #f56c6c; font-weight: bold; font-size: 14px; } QCheckBox::indicator { width: 18px; height: 18px; }")
        type_layout.addStretch()
        type_layout.addWidget(self.chk_negate)
        
        cond_layout.addLayout(type_layout)

        # Stacked Widget for different inputs
        self.stack = QStackedWidget()
        self.stack.setFixedHeight(40) # Fixed height for neatness
        
        # Page 0: Numeric
        page_numeric = QWidget()
        num_layout = QHBoxLayout()
        num_layout.setContentsMargins(0,0,0,0)
        self.num_op = QComboBox()
        self.num_op.addItems([">=", ">", "<=", "<", "==", "介于(Range)"])
        self.num_val1 = QLineEdit()
        self.num_val1.setPlaceholderText("值1 (如: 90)")
        self.num_val2 = QLineEdit()
        self.num_val2.setPlaceholderText("值2 (仅'介于'时使用)")
        self.num_val2.setEnabled(False)
        self.num_op.currentTextChanged.connect(lambda t: self.num_val2.setEnabled(t == "介于(Range)"))
        
        num_layout.addWidget(QLabel("关系:"))
        num_layout.addWidget(self.num_op)
        num_layout.addWidget(self.num_val1)
        num_layout.addWidget(QLabel("-"))
        num_layout.addWidget(self.num_val2)
        page_numeric.setLayout(num_layout)
        self.stack.addWidget(page_numeric)

        # Page 1: Text
        page_text = QWidget()
        text_layout = QHBoxLayout()
        text_layout.setContentsMargins(0,0,0,0)
        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("包含的文本")
        text_layout.addWidget(QLabel("文本内容:"))
        text_layout.addWidget(self.text_input)
        page_text.setLayout(text_layout)
        self.stack.addWidget(page_text)

        # Page 2: Regex
        page_regex = QWidget()
        regex_layout = QHBoxLayout()
        regex_layout.setContentsMargins(0,0,0,0)
        self.regex_input = QLineEdit()
        self.regex_input.setPlaceholderText(r"正则表达式 (如: ^\d{3}$)")
        regex_layout.addWidget(QLabel("表达式:"))
        regex_layout.addWidget(self.regex_input)
        page_regex.setLayout(regex_layout)
        self.stack.addWidget(page_regex)

        cond_layout.addWidget(self.stack)

        # Output Name and Add Button
        add_layout = QHBoxLayout()
        self.output_name = QLineEdit()
        self.output_name.setPlaceholderText("输出表名称 (如: 优秀名单)")
        self.add_btn = QPushButton("添加条件")
        self.add_btn.setIcon(QIcon.fromTheme("list-add")) # Try system icon
        self.add_btn.clicked.connect(self.add_condition)
        
        add_layout.addWidget(QLabel("输出名称:"))
        add_layout.addWidget(self.output_name)
        add_layout.addWidget(self.add_btn)
        cond_layout.addLayout(add_layout)

        cond_group.setLayout(cond_layout)
        main_layout.addWidget(cond_group)

        # 4. Condition List
        list_group = QGroupBox("4. 已添加条件列表")
        list_layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["列名", "类型", "详细条件", "输出名称"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        
        btn_layout = QHBoxLayout()
        self.del_btn = QPushButton("删除选中条件")
        self.del_btn.setObjectName("delete_btn")
        self.del_btn.clicked.connect(self.delete_condition)
        self.clear_btn = QPushButton("清空所有条件")
        self.clear_btn.setObjectName("clear_btn")
        self.clear_btn.clicked.connect(self.clear_conditions)
        btn_layout.addWidget(self.del_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()

        list_layout.addWidget(self.table)
        list_layout.addLayout(btn_layout)
        list_group.setLayout(list_layout)
        main_layout.addWidget(list_group)

        # 5. Output Settings & Action
        action_group = QGroupBox("5. 输出设置与执行")
        action_wrapper_layout = QVBoxLayout()

        # Output Mode
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("保存方式:"))
        self.rb_single_file = QRadioButton("合并为一个Excel文件 (不同Sheet)")
        self.rb_multi_files = QRadioButton("拆分为多个独立Excel文件")
        self.rb_single_file.setChecked(True)
        
        self.out_mode_group = QButtonGroup()
        self.out_mode_group.addButton(self.rb_single_file, 0)
        self.out_mode_group.addButton(self.rb_multi_files, 1)
        
        mode_layout.addWidget(self.rb_single_file)
        mode_layout.addWidget(self.rb_multi_files)
        mode_layout.addStretch()
        action_wrapper_layout.addLayout(mode_layout)

        # Action Button
        self.split_btn = QPushButton("🚀 开始拆分")
        self.split_btn.setFixedHeight(50)
        self.split_btn.setStyleSheet("""
            QPushButton {
                font-size: 18px; 
                font-weight: bold; 
                background-color: #67c23a; 
                border-radius: 6px;
                color: white;
            }
            QPushButton:hover {
                background-color: #85ce61;
            }
            QPushButton:pressed {
                background-color: #5daf34;
            }
        """)
        self.split_btn.clicked.connect(self.start_split)
        action_wrapper_layout.addWidget(self.split_btn)
        
        action_group.setLayout(action_wrapper_layout)
        main_layout.addWidget(action_group)

    def open_file_dialog(self):
        fname, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if fname:
            self.process_file(fname)

    def process_file(self, fname):
        try:
            self.path_display.setText(fname)
            self.file_path = fname
            # Load Excel to get sheet names
            xl = pd.ExcelFile(fname,engine='calamine')
            self.sheet_combo.clear()
            self.sheet_combo.addItems(xl.sheet_names)
            self.drop_area.label.setText(f"已选择: {os.path.basename(fname)}\n(点击或拖拽更换)")
            self.drop_area.label.setStyleSheet("color: #67c23a; font-size: 15px; font-weight: bold; border: 2px solid #67c23a; border-radius: 6px; padding: 20px; background-color: #f0f9eb;")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法读取文件: {str(e)}")
            self.path_display.setText("读取失败")

    def on_sheet_changed(self, text):
        if not self.file_path or not text:
            return
        try:
            # Read only headers
            df = pd.read_excel(self.file_path, sheet_name=text, nrows=0,engine='calamine')
            self.col_combo.clear()
            self.col_combo.addItems(df.columns.astype(str))
        except Exception as e:
            print(f"Error loading columns: {e}")

    def on_type_changed(self, btn):
        id = self.type_group.id(btn)
        self.stack.setCurrentIndex(id)

    def add_condition(self):
        col = self.col_combo.currentText()
        if not col:
            QMessageBox.warning(self, "警告", "请先选择一列")
            return
        
        out_name = self.output_name.text().strip()
        if not out_name:
            QMessageBox.warning(self, "警告", "请输入输出表名称")
            return

        cond_type_id = self.type_group.checkedId()
        cond_desc = ""
        params = {}

        if cond_type_id == 0: # Numeric
            op = self.num_op.currentText()
            v1 = self.num_val1.text().strip()
            v2 = self.num_val2.text().strip()
            
            if not v1:
                QMessageBox.warning(self, "警告", "请输入数值")
                return
            
            try:
                float(v1)
                if op == "介于(Range)":
                    if not v2:
                        QMessageBox.warning(self, "警告", "请输入第二个数值")
                        return
                    float(v2)
                    cond_desc = f"{v1} <= x <= {v2}"
                    params = {'op': 'range', 'v1': float(v1), 'v2': float(v2)}
                else:
                    cond_desc = f"x {op} {v1}"
                    params = {'op': op, 'v1': float(v1)}
            except ValueError:
                QMessageBox.warning(self, "警告", "输入必须是数值")
                return
            
            c_type = "数值范围"

        elif cond_type_id == 1: # Text
            txt = self.text_input.text()
            if not txt:
                QMessageBox.warning(self, "警告", "请输入文本")
                return
            cond_desc = f"包含 '{txt}'"
            params = {'text': txt}
            c_type = "文本包含"

        elif cond_type_id == 2: # Regex
            pat = self.regex_input.text()
            if not pat:
                QMessageBox.warning(self, "警告", "请输入正则表达式")
                return
            try:
                re.compile(pat)
            except re.error:
                QMessageBox.warning(self, "警告", "无效的正则表达式")
                return
            cond_desc = f"Regex: {pat}"
            params = {'pattern': pat}
            c_type = "正则表达式"

        # Handle Negate
        is_negate = self.chk_negate.isChecked()
        if is_negate:
            cond_desc = f"[NOT] {cond_desc}"
            c_type = f"{c_type} (取反)"

        # Add to list
        self.conditions.append({
            'col': col,
            'type': c_type,
            'desc': cond_desc,
            'params': params,
            'output_name': out_name,
            'type_id': cond_type_id,
            'is_negate': is_negate
        })
        
        self.refresh_table()
        self.output_name.clear()
        self.chk_negate.setChecked(False)

    def refresh_table(self):
        self.table.setRowCount(0)
        for c in self.conditions:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(c['col']))
            self.table.setItem(row, 1, QTableWidgetItem(c['type']))
            self.table.setItem(row, 2, QTableWidgetItem(c['desc']))
            self.table.setItem(row, 3, QTableWidgetItem(c['output_name']))

    def delete_condition(self):
        rows = sorted(set(index.row() for index in self.table.selectedIndexes()), reverse=True)
        for row in rows:
            del self.conditions[row]
        self.refresh_table()

    def clear_conditions(self):
        self.conditions = []
        self.refresh_table()

    def start_split(self):
        if not self.file_path:
            QMessageBox.warning(self, "警告", "请先选择文件")
            return
        if not self.conditions:
            QMessageBox.warning(self, "警告", "请至少添加一个条件")
            return

        try:
            sheet = self.sheet_combo.currentText()
            df = pd.read_excel(self.file_path, sheet_name=sheet,engine='calamine')
            
            mode = self.out_mode_group.checkedId()
            output_dir = ""
            output_file = ""
            
            if mode == 0: # Single File
                output_file, _ = QFileDialog.getSaveFileName(self, "保存结果", "split_result.xlsx", "Excel Files (*.xlsx)")
                if not output_file:
                    return
            else: # Multi Files
                output_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
                if not output_dir:
                    return

            success_msg = ""

            # Logic for filtering first, then saving
            results = []

            for cond in self.conditions:
                col = cond['col']
                params = cond['params']
                t_id = cond['type_id']
                
                if col not in df.columns:
                    continue 
                
                filtered_df = pd.DataFrame()
                mask = None

                if t_id == 0: # Numeric
                    series = pd.to_numeric(df[col], errors='coerce')
                    if params['op'] == 'range':
                        mn = min(params['v1'], params['v2'])
                        mx = max(params['v1'], params['v2'])
                        mask = (series >= mn) & (series <= mx)
                    elif params['op'] == '>=':
                        mask = series >= params['v1']
                    elif params['op'] == '>':
                        mask = series > params['v1']
                    elif params['op'] == '<=':
                        mask = series <= params['v1']
                    elif params['op'] == '<':
                        mask = series < params['v1']
                    elif params['op'] == '==':
                        mask = series == params['v1']
                    
                elif t_id == 1: # Text
                    series = df[col].astype(str)
                    mask = series.str.contains(params['text'], na=False)

                elif t_id == 2: # Regex
                    series = df[col].astype(str)
                    mask = series.str.match(params['pattern'], na=False) | series.str.contains(params['pattern'], regex=True, na=False)
                
                # Apply negation if needed
                if mask is not None:
                    if cond.get('is_negate', False):
                        mask = ~mask
                    filtered_df = df[mask]

                if not filtered_df.empty:
                    results.append((filtered_df, cond['output_name']))

            if not results:
                QMessageBox.information(self, "提示", "没有数据符合任何条件")
                return

            if mode == 0: # Single File
                with pd.ExcelWriter(output_file) as writer:
                    for res_df, name in results:
                         # Check dedup
                        if name in writer.sheets:
                            name = f"{name}_{len(writer.sheets)}"
                        res_df.to_excel(writer, sheet_name=name, index=False)
                success_msg = f"拆分完成！已保存至 {output_file}"
            else: # Multi Files
                saved_count = 0
                for res_df, name in results:
                    # Sanitize filename
                    safe_name = "".join([c for c in name if c.isalnum() or c in (' ', '-', '_')]).strip()
                    if not safe_name:
                        safe_name = "result"
                    
                    # Handle duplicates in file names?
                    fname = os.path.join(output_dir, f"{safe_name}.xlsx")
                    counter = 1
                    while os.path.exists(fname):
                        fname = os.path.join(output_dir, f"{safe_name}_{counter}.xlsx")
                        counter += 1
                    
                    res_df.to_excel(fname, index=False)
                    saved_count += 1
                success_msg = f"拆分完成！共保存 {saved_count} 个文件至 {output_dir}"

            QMessageBox.information(self, "成功", success_msg)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"拆分过程中发生错误: {str(e)}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelSplitterApp()
    window.show()
    sys.exit(app.exec_())
