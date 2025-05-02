import sys
import os
import json
import re
import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QListWidget,
    QHBoxLayout, QVBoxLayout, QFileDialog, QAbstractItemView, QTextEdit,
    QTableView, QMessageBox
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
        self.button_xls = QPushButton("Excel設定")
        self.button_xls.setFixedWidth(80)  
        hbox1.addWidget(self.button)
        hbox1.addWidget(self.label)
        hbox1.addWidget(self.button_xls)

        # ファイルリスト＆プレビュー
        hbox2 = QHBoxLayout()
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        self.list_widget.setFixedWidth(360)  
        self.font = QFont()
        self.font.setPointSize(8)
        table_view = QTableView()
        table_view.setFont(self.font)
        table_view.verticalHeader().setDefaultSectionSize(20)
        table_view.horizontalHeader().setDefaultSectionSize(60)
        table_view.horizontalHeader().setStretchLastSection(False)
        self.preview = QVBoxLayout()
        self.preview.addWidget(table_view)
        # 必要なら、自動サイズ調整を無効化して固定サイズにしたい場合は下記を追加
        hbox2.addWidget(self.list_widget)
        hbox2.addLayout(self.preview)

        # ツール
        hbox3 = QHBoxLayout()
        self.reset_btn = QPushButton("リセット")
        # self.concat_btn = QPushButton("結合")
        self.sum_btn = QPushButton("合計")
        hbox3.addWidget(self.reset_btn)
        # hbox3.addWidget(self.concat_btn)
        hbox3.addWidget(self.sum_btn)

        # レイアウト
        layout = QVBoxLayout()
        layout.addLayout(hbox1)
        layout.addLayout(hbox2)
        self.setLayout(layout)
        layout.addLayout(hbox3)

        # シグナル接続
        self.button.clicked.connect(self.select_directory)
        self.button_xls.clicked.connect(self.set_xls_config)
        self.reset_btn.clicked.connect(self.reset_data)
        # self.concat_btn.clicked.connect(self.concat_data)
        self.sum_btn.clicked.connect(self.sum_data)
        self.list_widget.itemClicked.connect(self.update_preview)

        # 設定ファイル読み込み
        self.config = {"header": 0}
        self.load_config()
        current_dir = self.config.get("current_dir")
        if current_dir:
            self.change_directory(current_dir)

    def reset_data(self):
        self.list_widget.clearSelection()
        self.clear_preview()
        table_view = QTableView()
        table_view.setFont(self.font)
        table_view.verticalHeader().setDefaultSectionSize(20)
        table_view.horizontalHeader().setDefaultSectionSize(60)
        table_view.horizontalHeader().setStretchLastSection(False)
        self.preview.addWidget(table_view)

    def sum_data(self):
        current_dir = self.config.get("current_dir")
        dfs = []
        for item in self.list_widget.selectedItems():
            path = os.path.join(current_dir, item.text())
            df = self.load_dataframe(path)
            dfs.append(df)

        df = pd.concat(dfs, ignore_index=True)
        cols = df["残高"].replace(",", "", regex=True).astype(int)
        total = cols.sum()
        QMessageBox.warning(self, "残高合計", f"{total:,.0f}")

    def concat_data(self):
        self.clear_preview()
        current_dir = self.config.get("current_dir")
        dfs = []
        for item in self.list_widget.selectedItems():
            path = os.path.join(current_dir, item.text())
            df = self.load_dataframe(path)
            dfs.append(df)
            
        self.df = pd.concat(dfs, ignore_index=True)
        model = DataFrameModel(self.df.head(10))
        table_view = QTableView()
        table_view.setModel(model)
        table_view.setFont(self.font)
        table_view.verticalHeader().setDefaultSectionSize(20)
        table_view.horizontalHeader().setDefaultSectionSize(60)
        table_view.horizontalHeader().setStretchLastSection(False)
        self.preview.addWidget(table_view)

    def clear_preview(self):
        while self.preview.count():
            item = self.preview.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def update_preview(self):
        self.clear_preview()
        current_dir = self.config.get("current_dir")
        size = len(self.list_widget.selectedItems())
        if size > 0:
            for item in self.list_widget.selectedItems():
                path = os.path.join(current_dir, item.text())
                df = self.load_dataframe(path)
                model = DataFrameModel(df.head().copy())
                del df
                table_view = QTableView()
                table_view.setModel(model)
                table_view.setFont(self.font)
                table_view.verticalHeader().setDefaultSectionSize(20)
                table_view.horizontalHeader().setDefaultSectionSize(60)
                table_view.horizontalHeader().setStretchLastSection(False)
                self.preview.addWidget(table_view)
        else:
            self.preview.addWidget(QTableView())


    def load_config(self):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                self.config = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.config = {}

    def save_config(self, config):
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
            
    def select_directory(self):
        initial_dir = self.config["current_dir"] or ""
        dir_path = QFileDialog.getExistingDirectory(self, "ディレクトリを選択", initial_dir)

        if dir_path:
            self.change_directory(dir_path)

    def set_xls_config(self):
        excel_config = self.config.get("excel_config") or {}
        dialog = MultiInputDialog(title="EXEL読込み設定",
                                  labels=["対象シート", "ヘッダ行"],
                                  values=[excel_config.get("sheet_name"),
                                          str(excel_config.get("header") or 0)])
        if dialog.exec():
            values = dialog.get_data()
            self.config["excel_config"] = {"sheet_name": values[0], "header": safe_int(values[1])}
            self.save_config(self.config)

    def load_dataframe(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".csv":
            return pd.read_csv(file_path, encoding="utf-8-sig")
        elif ext in [".xls", ".xlsx"]:
            excel_config = self.config.get("excel_config") or {}
            sheet_name = excel_config.get("sheet_name") or 0
            if is_int(sheet_name):
                sheet_name = int(sheet_name)
            header = excel_config.get("header") or 0
            return pd.read_excel(file_path, sheet_name=sheet_name, header=header)
        else:
            raise ValueError(f"未対応のファイル形式です: {ext}")

    def change_directory(self, dir_path):
        self.config["current_dir"] = dir_path
        self.label.setText(dir_path)
        csv_files = self.get_csv_files(dir_path)
        self.list_widget.clear()
        for file_path in csv_files:
            self.list_widget.addItem(file_path)
        self.save_config(self.config)

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
