from PySide6.QtWidgets import (
    QDialog, QVBoxLayout,
    QLabel, QLineEdit, QPushButton, QHBoxLayout
)

# input key/value
class MultiInputDialog(QDialog):
    def __init__(self, title, labels, values=None):
        super().__init__()
        values = values or []
        self.setWindowTitle(title)

        self.labels = []
        self.values = []
        layout = QVBoxLayout()

        # 名前入力
        for i, name in enumerate(labels):
            hbox = QHBoxLayout()
            label = QLabel(name)
            label.setFixedWidth(80)  
            self.labels.append(label)
            line_edit = QLineEdit()
            value = values[i] if i < len(values) else ""
            line_edit.setText(value)
            self.values.append(line_edit)
            hbox.addWidget(label)
            hbox.addWidget(line_edit)
            layout.addLayout(hbox)

        # OKボタン
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)

        # レイアウト
        layout.addWidget(self.ok_button)
        self.setLayout(layout)

    def get_data(self):
        return list(map(lambda x: x.text(), self.values))
    