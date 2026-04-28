import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QComboBox, QMessageBox
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

import win32com.client
from PIL import Image


class ScannerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Scanner App (Canon WIA)")
        self.setGeometry(200, 100, 500, 500)

        self.image_path = ""

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Preview
        self.preview = QLabel("Preview will appear here")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setFixedHeight(250)
        self.preview.setStyleSheet("border:1px solid gray;")
        layout.addWidget(self.preview)

        # DPI selector
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(QLabel("DPI:"))
        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(["200", "300", "600"])
        self.dpi_combo.setCurrentText("300")
        dpi_layout.addWidget(self.dpi_combo)
        layout.addLayout(dpi_layout)

        # Color mode
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel("Color Mode:"))
        self.color_combo = QComboBox()
        self.color_combo.addItems(["Color", "Grayscale", "Black & White"])
        color_layout.addWidget(self.color_combo)
        layout.addLayout(color_layout)

        # Buttons
        btn_layout = QHBoxLayout()

        self.scan_btn = QPushButton("Scan")
        self.scan_btn.clicked.connect(self.scan_document)
        btn_layout.addWidget(self.scan_btn)

        self.save_btn = QPushButton("Save As")
        self.save_btn.clicked.connect(self.save_file)
        btn_layout.addWidget(self.save_btn)

        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def scan_document(self):
        try:
            dpi = int(self.dpi_combo.currentText())

            color_map = {
                "Color": 1,
                "Grayscale": 2,
                "Black & White": 4
            }
            color_mode = color_map[self.color_combo.currentText()]

            wia = win32com.client.Dispatch("WIA.CommonDialog")
            device = wia.ShowSelectDevice()

            if device is None:
                return

            item = device.Items[0]
            props = item.Properties

            # Set DPI (Canon-safe)
            try:
                props[6147].Value = dpi
                props[6148].Value = dpi
            except:
                pass

            # Set color mode
            try:
                props[6146].Value = color_mode
            except:
                pass

            # Scan
            img = item.Transfer()

            temp_bmp = "temp_scan.bmp"
            output_jpg = "preview.jpg"

            img.SaveFile(temp_bmp)

            Image.open(temp_bmp).convert("RGB").save(output_jpg)
            os.remove(temp_bmp)

            self.image_path = output_jpg

            # Show preview
            pixmap = QPixmap(output_jpg)
            self.preview.setPixmap(pixmap.scaled(
                self.preview.width(),
                self.preview.height(),
                Qt.KeepAspectRatio
            ))

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def save_file(self):
        if not self.image_path:
            QMessageBox.warning(self, "Warning", "No scanned image")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save File", "", "Images (*.jpg *.png)"
        )

        if file_path:
            Image.open(self.image_path).save(file_path)
            QMessageBox.information(self, "Saved", f"Saved to {file_path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScannerApp()
    window.show()
    sys.exit(app.exec_())