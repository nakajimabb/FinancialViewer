import copy
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QComboBox, QHBoxLayout, QLabel, QLineEdit,
    QGridLayout, QCheckBox, QPushButton, QGroupBox
)
from tools import safe_get

class FormatDialog(QDialog):
    def __init__(self, df, config):
        super().__init__()
        self.df = df
        self.config = copy.deepcopy(config)

        self.setWindowTitle("帳票設定")
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 8, 16, 8)
        layout.setSpacing(20)

        vbox = QVBoxLayout()
        vbox.setSpacing(4)
        self.label = QLabel("帳票名称")
        font = QFont()
        font.setPointSize(10)  # 小さく
        self.label.setFont(font)
        self.line_edit = QLineEdit()
        self.line_edit.setText(self.config.get("name"))
        self.line_edit.setPlaceholderText("名前を入力してください")
        self.line_edit.setFixedWidth(200)
        vbox.addWidget(self.label)
        vbox.addWidget(self.line_edit)
        layout.addLayout(vbox)

        # グループボックス
        group_box1 = QGroupBox("キー選択")
        group_layout = QHBoxLayout()
        # キー選択
        self.key_combos = []
        keys = config.get("keys") or []
        for i in range(3):
            combo = QComboBox()
            combo.addItem("")
            combo.addItems(df.columns)
            self.key_combos.append(combo)
            key = safe_get(keys, i)
            index = combo.findText(key)
            if index != -1:
                combo.setCurrentIndex(index)
            group_layout.addWidget(combo)

        group_box1.setLayout(group_layout)
        layout.addWidget(group_box1)
        # 表示項目選択
        columns = config.get("columns") or []
        self.column_checkboxes = []
        group_box2 = QGroupBox("表示項目選択")
        grid = QGridLayout()
        max_columns = 4
        for i, name in enumerate(df.columns):
            checkbox = QCheckBox(name)
            checkbox.setChecked(name in columns)
            row = i // max_columns
            col = i % max_columns
            grid.addWidget(checkbox, row, col)
            self.column_checkboxes.append(checkbox)
        group_box2.setLayout(grid)
        layout.addWidget(group_box2)
        # OKボタン
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        # キャンセルボタン
        self.cancel_button = QPushButton("キャンセル")
        self.cancel_button.clicked.connect(self.reject)
        button_layout = QHBoxLayout()
        button_layout.addStretch()  # 右寄せにしたい場合
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        # レイアウト
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def get_config(self):
        self.config["name"] = self.line_edit.text()
        self.config["keys"] = []
        for combo in self.key_combos:
            key = combo.currentText()
            if key:
                self.config["keys"].append(key)
        self.config["columns"] = []
        for checkbox in self.column_checkboxes:
            if checkbox.isChecked():
                self.config["columns"].append(checkbox.text())
        
        return self.config
    