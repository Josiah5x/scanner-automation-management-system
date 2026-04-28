import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QTabWidget, QListView, QFileDialog
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QIcon
from PyQt5.QtCore import QSize


class ScannerUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Scanner UI with Preview")
        self.resize(800, 500)

        layout = QVBoxLayout()

        # ===== Tabs =====
        self.tabs = QTabWidget()

        # ===== Tab 1: Preview =====
        self.preview_tab = QWidget()
        self.preview_layout = QVBoxLayout()

        self.list_view = QListView()
        self.list_view.setViewMode(QListView.IconMode)
        self.list_view.setIconSize(QSize(500, 600))
        self.list_view.setSpacing(10)

        # Model
        self.model = QStandardItemModel()
        self.list_view.setModel(self.model)

        self.preview_layout.addWidget(self.list_view)
        self.preview_tab.setLayout(self.preview_layout)

        # ===== Tab 2: Scan =====
        self.scan_tab = QWidget()
        self.scan_layout = QVBoxLayout()

        self.scan_btn = QPushButton("Scan / Add Image")
        self.scan_btn.clicked.connect(self.add_image)

        self.scan_layout.addWidget(self.scan_btn)
        self.scan_tab.setLayout(self.scan_layout)

        # Add tabs
        self.tabs.addTab(self.preview_tab, "Preview")
        self.tabs.addTab(self.scan_tab, "Scan")

        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def add_image(self):
        # Simulate scan by selecting file
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Images (*.png *.jpg *.jpeg)"
        )

        if file_path:
            item = QStandardItem()
            icon = QIcon(file_path)

            item.setIcon(icon)
            item.setText("Scanned Image")
            item.setEditable(False)

            self.model.appendRow(item)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScannerUI()
    window.show()
    sys.exit(app.exec_())
