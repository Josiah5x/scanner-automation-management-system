import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QComboBox, QSlider, QSpinBox, QFrame, QListWidget,
    QTabWidget, QLineEdit,QFileDialog, QComboBox, QMessageBox,QListWidget, QListWidgetItem,QFormLayout
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from win32com.client import Dispatch
from PIL import Image
from pathlib import Path

class ScannerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Scanner App (Canon WIA)")
        self.setStyleSheet("background-color: #2b2b2b; color: white;")
        self.image_path = ""
        self.insideScans = "scans"
        self.path = Path.cwd() / self.insideScans
        self.currentTabName = ""
        self.init_ui()

    def init_ui(self):
        # ===== MAIN HORIZONTAL LAYOUT =====
        main_layout = QHBoxLayout()

        # ===== LEFT SIDE (CONTROLS) =====
        left_layout = QVBoxLayout()

        # left_layout.addWidget(QLabel("SITE:"))

        self.site_combo = QComboBox()
        self.site_combo.setFixedHeight(35)
        self.site_combo.setFixedWidth(350)

        # Ensure folder exists
        self.path = Path.cwd() / "scans"
        self.path.mkdir(exist_ok=True)

        self.load_folders()
        

        left_layout.addWidget(self.site_combo)
         # --- FORM INPUTS ---
        form_layout = QFormLayout()

        self.supplier_input = QLineEdit()
        self.invoice_input = QLineEdit()
        self.amount_input = QLineEdit()

        self.supplier_input.setPlaceholderText("Enter supplier name")
        self.invoice_input.setPlaceholderText("Enter invoice number")
        self.amount_input.setPlaceholderText("Enter amount")

        form_layout.addRow("Supplier:", self.supplier_input)
        form_layout.addRow("Invoice No:", self.invoice_input)
        form_layout.addRow("Amount:", self.amount_input)

        left_layout.addLayout(form_layout)

        # Buttons
        self.add_dir_btn = QPushButton("+ Add Folder")
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_file)

        left_layout.addWidget(self.add_dir_btn)
        left_layout.addWidget(self.save_btn)

        left_layout.addStretch()  # pushes everything up

        # ===== RIGHT SIDE (PREVIEW BIG AREA) =====
        right_layout = QVBoxLayout()

        self.preview = QLabel("Preview will appear here")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setStyleSheet("""
            border: 2px solid gray;
            background-color: #1e1e1e;
        """)

        # Make preview expand fully
        self.preview.setMinimumSize(400, 600)

        right_layout.addWidget(self.preview)

        # ===== ADD BOTH SIDES =====
        main_layout.addLayout(left_layout, 1)   # small width
        main_layout.addLayout(right_layout, 3)  # bigger width

        self.setLayout(main_layout)

    def load_folders(self):
        self.site_combo.clear()
        subdirs = [f.name for f in self.path.iterdir() if f.is_dir()]
        self.site_combo.addItems(subdirs)

    def save_file(self):
        print(self.path/self.site_combo.currentText())
        # if not self.image_path:
        #     QMessageBox.warning(self, "Warning", "No scanned image")
        #     return
        print(self.path/self.image_path)
        # # Define the file to be moved
        # source_file = Path(self.image_path)

        # # Define the location to put the file
        # destination = Path("data/new/location")

        # # Create the directories if they don't exist
        # destination.mkdir(parents=True)

        # # Move the file
        # source_file.replace(destination / source_file.name)

# /////////////////////////


class ScannerUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Scanner UI")
        self.setGeometry(100, 100, 1200, 700)
        self.setStyleSheet("background-color: #eee; color: #333;")

        main_layout = QHBoxLayout(self)

        # ===== LEFT PANEL =====
        left_panel = QFrame()
        left_panel.setFixedWidth(300)
        left_panel.setStyleSheet("background-color: #1e1e1e; color:#fff")
        left_layout = QVBoxLayout()
        def create_dropdown(label, items):
            layout = QVBoxLayout()
            lbl = QLabel(label)
            combo = QComboBox()
            combo.addItems(items)
            combo.setStyleSheet("background:#2b2b2b; padding:5px;")
            layout.addWidget(lbl)
            layout.addWidget(combo)
            return layout
        def list_scanners():
            """List all available WIA scanners"""
            manager = Dispatch("WIA.DeviceManager")
            devices = manager.DeviceInfos
            # print("Available scanners:")
            for i in range(1, devices.Count + 1):
                device = devices.Item(i)
                # Check if the device is a scanner (Type = 1)
                if device.Type == 1:
                    # print(f"  Name: {device.Properties['Name'].Value}")
                    # print(f"  ID: {device.DeviceID}")
                    # print(f"  Description: {device.Properties['Description'].Value}")
                    # print("  ----------------")
                    pass
                left_layout.addLayout(create_dropdown("Select Scanner", [f"{device.Properties['Name'].Value}"]))
                left_layout.addLayout(create_dropdown("Paper Source", ["Glass"]))
                left_layout.addLayout(create_dropdown("Color Mode", ["Color", "Color"]))
                left_layout.addLayout(create_dropdown("Resolution", ["300 DPI", "600 DPI"]))
                left_layout.addLayout(create_dropdown("Scan Area", ["A4 (210 x 297 mm)"]))
        list_scanners()


        def create_slider(label):
            layout = QVBoxLayout()
            lbl = QLabel(label)

            h_layout = QHBoxLayout()
            spin = QSpinBox()
            spin.setRange(-100, 100)
            slider = QSlider(Qt.Horizontal)
            slider.setRange(-100, 100)

            slider.valueChanged.connect(spin.setValue)
            spin.valueChanged.connect(slider.setValue)

            h_layout.addWidget(spin)
            h_layout.addWidget(slider)

            layout.addWidget(lbl)
            layout.addLayout(h_layout)
            return layout

        left_layout.addLayout(create_slider("Brightness"))
        left_layout.addLayout(create_slider("Contrast"))
        left_layout.addLayout(create_slider("Saturation"))

        scan_btn = QPushButton("Scan")
        scan_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
        """)
        scan_btn.clicked.connect(self.scan_document)
        left_layout.addWidget(scan_btn)

        left_layout.addStretch()
        left_panel.setLayout(left_layout)

        # ===== RIGHT PANEL =====
        right_panel = QVBoxLayout()

        # Top bar
        top_bar = QHBoxLayout()
        pages_label = QLabel("All Pages 0 / 1 Selected")

        save_btn2 = QPushButton("Save")
        save_btn2.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                padding: 8px 15px;
                border-radius: 5px;
            }
        """)
        save_btn2.clicked.connect(self.openlog)
        top_bar.addWidget(pages_label)
        top_bar.addStretch()
        top_bar.addWidget(save_btn2)

        # ===== TABS =====
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab {
                background: #eee;
                color:#2d2d2d;
                padding: 10px;
                width: 100px;
                font-weight:700;
                font-family: 'Roboto', sans-serif;
            }
            QTabBar::tab:selected {
                background: blue;
                color:#fff;
                width: 100px;
                font-weight:700;
                font-family: 'Roboto', sans-serif; 
            }
        """)
        # Create tabs
        self.insideScans = "scans"
        self.path = Path.cwd() / self.insideScans

        # Ensure folder exists
        self.path.mkdir(exist_ok=True)

        # Get only directories (sorted)
        self.subdirs = sorted([f.name for f in self.path.iterdir() if f.is_dir()])

        # Add tabs
        for folder_name in self.subdirs:
            self.tabs.addTab(self.create_tab(folder_name), folder_name)

        # Tab click function
        # self.tabs.setCornerWidget(self.tabs.findChild(self.create_tab(0)))
        self.tabs.currentChanged.connect(self.tab_changed)

        right_panel.addLayout(top_bar)
        right_panel.addWidget(self.tabs)

        # ===== ADD TO MAIN =====
        main_layout.addWidget(left_panel)
        main_layout.addLayout(right_panel)

        # Specify the directory path you want to start from
        # directory_path = './scans'
        # self.list_files_walk(directory_path)

    # ===== CREATE TAB CONTENT =====
    def create_tab2(self, folder_name):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        list_widget = QListWidget()
        list_widget.setIconSize(QSize(100, 100))  # preview size
        self.currentTabName = folder_name
        # Loop through files in folder
        folder_paths = self.path/folder_name
        for file in folder_paths.iterdir():
            if file.is_file():
                item = QListWidgetItem(file.name)

                # Optional: show image preview
                if file.suffix.lower() in [".png", ".jpg", ".jpeg"]:
                    item.setIcon(QIcon(str(file)))

                list_widget.addItem(item)

        layout.addWidget(list_widget)
        return tab
    def create_tab(self, folder_name):
        layout = QVBoxLayout()
        tab = QWidget()

         # Search
        search_layout = QHBoxLayout()
        search_input = QLineEdit()
        search_input.setPlaceholderText("Search...")
        
        btn = QPushButton("Search")
        btn.clicked.connect(lambda: print(f"{folder_name} search:", search_input.text()))

        # ===== PREVIEW LIST =====
        list_widget = QListWidget()
        list_widget.setIconSize(QSize(100, 100))  # preview size
       # Loop through files in folder
        folder_paths = self.path/folder_name
        for file in folder_paths.iterdir():
            if file.is_file():
                item = QListWidgetItem(file.name)

                # Optional: show image preview
                if file.suffix.lower() in [".png", ".jpg", ".jpeg"]:
                    item.setIcon(QIcon(str(file)))

                list_widget.addItem(item)

        # btn.clicked.connect(handle_search)

        search_layout.addWidget(search_input)
        search_layout.addWidget(btn)
        layout.addLayout(search_layout)
        layout.addWidget(list_widget)
        # layout.addStretch()

        tab.setLayout(layout)
        return tab

    # ===== TAB CLICK FUNCTION =====
    def tab_changed(self, index):
        name = self.tabs.tabText(index)
        # self.create_tab(Path.cwd()/name)
        print(f"Switched to Tab {name},{index}")
        return name
    
    def doc_view(self, start_path='.'):
        print(start_path)
        for root, dirs, files in os.walk(start_path):
            for file in files:
                # Show doc_view
                # Find all text files inside a directory
                pixmap = QPixmap(os.path.join(root,file))
                self.preview.setPixmap(pixmap.scaled(
                    self.preview.width(),
                    self.preview.height(),
                    Qt.IgnoreAspectRatio,
                    Qt.SmoothTransformation
                ))

    # ===== SEARCH BUTTON FUNCTION =====
    def search_action(self, tab_name, text):
        print(f"Searching '{text}' in {tab_name}")
    
    def scan_document(self):
        scanner_name=None
        """
        Scan a document using WIA and save as image (Canon-safe)
        """
        wia = Dispatch("WIA.CommonDialog")
        selected_device = None
        if scanner_name:
            manager = Dispatch("WIA.DeviceManager")
            devices = manager.DeviceInfos
            for i in range(1, devices.Count + 1):
                device = devices.Item(i)
                if device.Type == 1 and device.Properties['Name'].Value == scanner_name:
                    selected_device = device.Connect()
                    break
            if not selected_device:
                print(f"Scanner '{scanner_name}' not found!")
                return None
        print("Scanning...")
        if selected_device is None:
            img = wia.ShowAcquireImage()
        else:
            img = wia.ShowTransfer(selected_device.Items[1])
        if img is None:
            print("Scanning cancelled or failed")
            return None
        # Ensure output directory exists
        output_dir = os.path.dirname("image.jpg")
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        # 🔥 KEY FIX: always save as BMP first
        temp_bmp = os.path.splitext("image.jpg")[0] + "_temp.bmp"
        try:
            img.SaveFile(temp_bmp)  # Canon-safe
            # Convert to final format (jpg/png/etc.)
            Image.open(temp_bmp).convert("RGB").save("image.jpg")
            os.remove(temp_bmp)
            print(f"✅ Scan saved to: {"image.jpg"}")
            self.window = ScannerApp()
            self.window.image_path = "image.jpg"
            pixmap = QPixmap(self.window.image_path)
            self.window.show()
            self.window.preview.setPixmap(pixmap.scaled(
                self.window.preview.width(),
                self.window.preview.height(),
                Qt.KeepAspectRatio
            ))
            
            return "image.jpg"
        except Exception as e:
            print("❌ Error:", e)
            return None
    def openlog(self):
        print("hello")
        self.window = ScannerApp()
        # self.window.image_path = "image.jpg"
        # pixmap = QPixmap(self.window.image_path)
        self.window.show()
        # self.window.preview.setPixmap(pixmap.scaled(
        #     self.window.preview.width(),
        #     self.window.preview.height(),
        #     Qt.KeepAspectRatio
        # ))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion') 
    window = ScannerApp()
    window.show()
    sys.exit(app.exec_())