import sys
import os
import json
import re
import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel
from PySide6.QtGui import QFont, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QListWidget,
    QHBoxLayout, QVBoxLayout, QFileDialog, QTableView
)
from MultiInputDialog import MultiInputDialog

def safe_int(s):
    try:
        return int(s)
    except (ValueError, TypeError):
        return 0

def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def natural_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]

def parse_int_or_str(s):
    try:
        return int(s)
    except ValueError:
        return s

class DataFrameModel(QAbstractTableModel):
    def __init__(self, df):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            value = self._df.iloc[index.row(), index.column()]
            return str(value)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._df.columns[section])
            else:
                return str(self._df.index[section])
        return None


class CSVExplorer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV ファイル一覧表示")
        self.resize(960, 800)
        # フォルダ選択
        hbox1 = QHBoxLayout()
        self.label = QLabel("フォルダを選択してください")
        self.button = QPushButton("フォルダ選択")
        self.button.setFixedWidth(100)  
        self.button_xls = QPushButton("帳票設定")
        self.button_xls.setFixedWidth(80)  
        hbox1.addWidget(self.button)
        hbox1.addWidget(self.label)
        hbox1.addWidget(self.button_xls)
        # ファイルリスト＆プレビュー
        hbox2 = QHBoxLayout()
        self.list_widget = QListWidget()
        self.list_widget.setFixedWidth(360)  
        self.font = QFont()
        self.font.setPointSize(8)
        self.preview = QTableView()
        self.preview.setFont(self.font)
        self.preview.verticalHeader().setDefaultSectionSize(20)
        self.preview.horizontalHeader().setDefaultSectionSize(60)
        self.preview.horizontalHeader().setStretchLastSection(False)
        # 必要なら、自動サイズ調整を無効化して固定サイズにしたい場合は下記を追加
        hbox2.addWidget(self.list_widget)
        hbox2.addWidget(self.preview)
        # レイアウト
        layout = QVBoxLayout()
        layout.addLayout(hbox1)
        layout.addLayout(hbox2)
        self.setLayout(layout)
        # シグナル接続
        self.button.clicked.connect(self.select_directory)
        self.button_xls.clicked.connect(self.set_project)
        self.list_widget.itemClicked.connect(self.update_preview)
        # 設定の読み込み
        self.config = {}    # 全体の設定
        self.load_config()
        current_dir = self.config.get("current_dir")
        if current_dir:
            self.change_directory(current_dir)

    def current_dir(self):
        return self.config.get("current_dir")

    def update_preview(self):
        self.preview.setModel(QStandardItemModel()) # プレビュー画面をクリア
        selected_items = self.list_widget.selectedItems()
        if selected_items:
            item = selected_items[0]
            path = os.path.join(self.current_dir(), item.text())
            df = self.load_dataframe(path)
            model = DataFrameModel(df)
            self.preview.setModel(model)

    def load_config(self):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                self.config = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.config = {}

    def save_config(self, config):
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
            
    def save_project(self, project):
        path = os.path.join(self.current_dir(), "project.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(project, f, indent=2, ensure_ascii=False)

    def select_directory(self):
        initial_dir = self.config["current_dir"] or ""
        dir_path = QFileDialog.getExistingDirectory(self, "ディレクトリを選択", initial_dir)
        if dir_path:
            self.change_directory(dir_path)

    def set_project(self):
        dialog = MultiInputDialog(title="帳票設定",
                                  labels=["対象シート", "ヘッダ行"],
                                  values=[str(self.project.get("sheet_name") or 0),
                                          str(self.project.get("header") or 0)])
        if dialog.exec():
            values = dialog.get_data()
            self.project = {"sheet_name": parse_int_or_str(values[0]),
                            "header": safe_int(values[1])}
            self.save_project(self.project)

    def load_dataframe(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".csv":
            return pd.read_csv(file_path, encoding="utf-8-sig")
        elif ext in [".xls", ".xlsx"]:
            sheet_name = self.project.get("sheet_name") or 0
            if is_int(sheet_name):
                sheet_name = int(sheet_name)
            header = self.project.get("header") or 0
            return pd.read_excel(file_path, sheet_name=sheet_name, header=header)
        else:
            raise ValueError(f"未対応のファイル形式です: {ext}")

    def change_directory(self, dir_path):
        self.config["current_dir"] = dir_path
        self.label.setText(dir_path)
        self.save_config(self.config)
        self.load_project(dir_path)

    def load_project(self, dir_path):
        # ディレクトリの設定の読込み
        self.project = {}
        path = os.path.join(dir_path, "project.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                self.project = json.load(f)
        # CSV・EXCELファイル一覧の取得
        csv_files = self.get_csv_files(dir_path)
        self.list_widget.clear()
        for file_path in csv_files:
            self.list_widget.addItem(file_path)


    def get_csv_files(self, directory, sort_by='name', reverse=False):
        csv_files = []
        for root, _, files in os.walk(directory):
            for file in files:
                if file.startswith("~$"):
                    continue
                if file.lower().endswith((".csv", ".xlsx", ".xls")):
                    full_path = os.path.join(root, file)
                    relative_path = os.path.relpath(full_path, directory)
                    mtime = os.path.getmtime(full_path)
                    csv_files.append((relative_path, mtime))

        if sort_by == 'name':
            csv_files.sort(key=lambda x: natural_key(os.path.basename(x[0])), reverse=reverse)
        elif sort_by == 'path':
            csv_files.sort(key=lambda x: natural_key(x[0]), reverse=reverse)
        elif sort_by == 'date':
            csv_files.sort(key=lambda x: x[1], reverse=reverse)

        return [path for path, _ in csv_files]

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CSVExplorer()
    window.show()
    sys.exit(app.exec())
