import sys
import os
import json
import pathlib
import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel
from PySide6.QtGui import QFont, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QListWidget,
    QHBoxLayout, QVBoxLayout, QFileDialog, QTableView, QMessageBox,
    QComboBox
)
from MultiInputDialog import MultiInputDialog
from FormatDialog import FormatDialog
from tools import safe_int, is_int, natural_key, parse_int_or_str, safe_get

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
        self.df = None
        # フォルダ選択
        hbox1 = QHBoxLayout()
        self.label = QLabel("フォルダを選択してください")
        self.button = QPushButton("フォルダ選択")
        self.button.setFixedWidth(100)
        hbox1.addWidget(self.button)
        hbox1.addWidget(self.label)
        # 帳票設定
        hbox2 = QHBoxLayout()
        self.button_xls = QPushButton("ヘッダ設定")
        self.button_xls.setFixedWidth(100)
        self.format_cmb = QComboBox()
        self.format_cmb.addItem("")
        self.format_cmb.setFixedWidth(200)
        self.button_edit = QPushButton("帳票編集")
        self.button_edit.setFixedWidth(80)
        self.button_add = QPushButton("帳票追加")
        self.button_add.setFixedWidth(80)
        self.button_output = QPushButton("集計")
        self.button_output.setFixedWidth(80)
        self.button_save = QPushButton("保存")
        self.button_save.setFixedWidth(80)
        hbox2.addWidget(self.button_xls)
        hbox2.addWidget(self.format_cmb)
        hbox2.addWidget(self.button_edit)
        hbox2.addWidget(self.button_add)
        hbox2.addWidget(self.button_output)
        hbox2.addWidget(self.button_save)
        hbox2.addStretch()
        # ファイルリスト＆プレビュー
        hbox3 = QHBoxLayout()
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
        hbox3.addWidget(self.list_widget)
        hbox3.addWidget(self.preview)
        # レイアウト
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.addLayout(hbox1)
        layout.addLayout(hbox2)
        layout.addLayout(hbox3)
        self.setLayout(layout)
        # シグナル接続
        self.button.clicked.connect(self.select_directory)
        self.button_xls.clicked.connect(self.set_project)
        self.button_edit.clicked.connect(self.edit_format)
        self.button_add.clicked.connect(self.add_format)
        self.button_output.clicked.connect(self.output)
        self.button_save.clicked.connect(self.save_result)
        self.list_widget.itemClicked.connect(self.update_preview)
        # 設定の読み込み
        self.config = {}    # 全体の設定
        self.load_config()
        current_dir = self.current_dir()
        if current_dir:
            self.change_directory(current_dir)

    def current_dir(self):
        return self.config.get("current_dir")

    def current_format(self):
        index = self.format_cmb.currentIndex()
        if index >= 0:
            format = safe_get(self.project["formats"], index) or {}
            return format
        else:
            QMessageBox.information(None, "", "帳票が選択されていません。")
            return None

    def update_preview(self):
        self.preview.setModel(QStandardItemModel()) # プレビュー画面をクリア
        selected_items = self.list_widget.selectedItems()
        if selected_items:
            item = selected_items[0]
            path = os.path.join(self.current_dir(), item.text())
            self.df = self.load_dataframe(path)
            model = DataFrameModel(self.df)
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
        dialog = MultiInputDialog(title="ヘッダ設定",
                                  labels=["対象シート", "ヘッダ行"],
                                  values=[str(self.project.get("sheet_name") or 0),
                                        str(self.project.get("header") or 0)])
        if dialog.exec():
            values = dialog.get_data()
            self.project["sheet_name"] = parse_int_or_str(values[0])
            self.project["header"] = safe_int(values[1])
            self.save_project(self.projects)

    def edit_format(self):
        if self.df is not None and not self.df.empty:
            index = self.format_cmb.currentIndex()
            if index >= 0:
                format = safe_get(self.project["formats"], index) or {}
                dialog = FormatDialog(df=self.df, config=format)
                if dialog.exec():
                    self.project["formats"][index] = dialog.get_config()
                    self.update_format_cmb()
                    self.format_cmb.setCurrentIndex(index)
                    self.save_project(self.project)
        else:
            QMessageBox.information(None, "", "シートを選択してください。")

    def add_format(self):
        if self.df is not None and not self.df.empty:
            dialog = FormatDialog(df=self.df, config={})
            if dialog.exec():
                self.project["formats"].append(dialog.get_config())
                self.update_format_cmb()
                self.format_cmb.setCurrentIndex(self.format_cmb.count() -1)
                self.save_project(self.project)
        else:
            QMessageBox.information(None, "", "シートを選択してください。")

    def df_agg(self, df, keys, columns):
        dtypes_before = df.dtypes.to_dict()
        ordered_columns = [col for col in df.columns if col in columns]
        df = df[ordered_columns]
        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        numeric_cols = [col for col in numeric_cols if col in ordered_columns and col not in keys]
        other_cols = [col for col in ordered_columns if col not in numeric_cols]
        if len(keys) > 0:
            df = df.groupby(keys).agg(
                {col: 'sum' for col in numeric_cols} |
                {col: 'first' for col in other_cols}
            )
            df = self.set_df_int_types(df, dtypes_before)
            df = df[ordered_columns]
        else:
            df = df.sum(numeric_only=True).to_frame().T

        return df

    def set_df_int_types(self, df, dtypes):
        for col in df.columns:
            if col in dtypes and pd.api.types.is_integer_dtype(dtypes[col]):
                df[col] = df[col].astype('Int64')   # Int64: nullable integer(for IntCastingNaNError)
        return df


    def output(self):
        errors = []
        format = self.current_format()
        if format:
            current_dir = self.current_dir()
            files = self.get_csv_xls_files(current_dir)
            keys = format["keys"]
            columns = list(set(keys + format["columns"]))
            dfs = []
            if len(files) > 0:
                for file in files:
                    try:
                        path = os.path.join(current_dir, file)
                        df = self.load_dataframe(path)
                        df = self.df_agg(df, keys, columns)
                        dfs.append(df)
                    except Exception as e:
                        errors.append(f"{file}: {str(e)}")

                try:
                    dtypes_before = dfs[0].dtypes.to_dict()
                    merged_df = pd.concat(dfs, ignore_index=True)
                    merged_df = self.set_df_int_types(merged_df, dtypes_before)
                    self.df = self.df_agg(merged_df, keys, columns)
                    model = DataFrameModel(self.df)
                    self.preview.setModel(model)
                except Exception as e:
                    errors.append(f"result: {str(e)}")

                if len(errors) > 0:
                    max_len = 256
                    file_name = "error.log"
                    path = os.path.join(current_dir, file_name)
                    text = "\n".join(errors)
                    if len(text) > max_len:
                        text = text[:max_len - 3] + "..."
                    QMessageBox.information(None, "エラー情報", text + "\n" + f"{file_name} 参照")
                    self.save_errors(path, errors)
        else:
            QMessageBox.information(None, "", "帳票の種別を選択してください。")

    def save_errors(self, path, errors):
        with open(path, 'w', encoding='utf-8') as f:
            text = "\n".join(errors)
            f.write(text)


    def save_result(self):
        format = self.current_format()
        if format and format.get("name"):
            out_path = os.path.join(self.current_dir(), format.get("name"))
            out_path = pathlib.Path(out_path).with_suffix(".csv")
            self.df.to_csv(out_path, index=False, encoding="utf-8-sig")

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
        # デフォルト値の設定
        self.project.setdefault("formats", [])
        # CSV・EXCELファイル一覧の取得
        files = self.get_csv_xls_files(dir_path)
        self.list_widget.clear()
        for file_path in files:
            self.list_widget.addItem(file_path)
        self.update_format_cmb()

    # コンボボックス更新
    def update_format_cmb(self):
        formats = self.project.get("formats")
        names = map(lambda f: f.get("name"), formats)
        self.format_cmb.clear()
        self.format_cmb.addItems(names)


    def get_csv_xls_files(self, directory, sort_by='name', reverse=False):
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
