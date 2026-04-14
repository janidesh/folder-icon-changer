import sys
import os
import ctypes
import winreg
from pathlib import Path
from PySide6.QtGui import QPixmap, QFont, QColor
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                               QLabel, QPushButton, QFileDialog, QFrame, QMessageBox, 
                               QTabWidget, QCheckBox, QComboBox, QMenu,
                               QGraphicsDropShadowEffect)
from PySide6.QtCore import Qt, QThread, Signal, QPoint, QTimer

try:
    import win32com.client
    import win32api
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# --- CONSTANTS & GUIDS FOR SYSTEM ICONS ---
SYSTEM_ICONS = {
    "This PC": "{20D04FE0-3AEA-1069-A2D8-08002B30309D}",
    "Recycle Bin (Empty)": "RecycleBinEmpty", 
    "Recycle Bin (Full)": "RecycleBinFull",   
    "Network": "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}",
    "User Folder": "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
}


# --- 3D GLASS SPLASH SCREEN ---
class GlassySplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        # Frameless and completely transparent base
        self.setWindowFlags(Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(540, 340)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20) # Space for the 3D shadow

        # The glassy container
        self.bg_frame = QFrame(self)
        self.bg_frame.setStyleSheet("""
            QFrame {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                    stop:0 rgba(80, 30, 110, 180), 
                    stop:0.5 rgba(45, 15, 65, 200), 
                    stop:1 rgba(20, 5, 30, 220));
                border-radius: 20px;
                border: 2px solid rgba(255, 255, 255, 0.2);
                border-top: 2px solid rgba(255, 255, 255, 0.4); /* Highlights top edge */
                border-left: 2px solid rgba(255, 255, 255, 0.3);
            }
        """)

        # 3D Drop Shadow
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(0, 0, 0, 180))
        shadow.setOffset(0, 10)
        self.bg_frame.setGraphicsEffect(shadow)

        frame_layout = QVBoxLayout(self.bg_frame)
        frame_layout.setAlignment(Qt.AlignCenter)

        # Logo Logic
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setStyleSheet("border: none; background: transparent;")
        
        logo_path = os.path.join("assets", "logo.png")
        pixmap = QPixmap(logo_path)
        
        if not pixmap.isNull():
            # Scale logo to fit nicely
            self.logo_label.setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            # Fallback text if logo is missing
            self.logo_label.setText("🔮")
            self.logo_label.setStyleSheet("font-size: 72px; border: none; background: transparent;")

        # Branding Text
        self.title_label = QLabel("Ultimate OS Icon Manager")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: white; font-size: 20px; font-weight: bold; border: none; background: transparent; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI';")
        
        self.subtitle_label = QLabel("Janith Rathnayake Creations")
        self.subtitle_label.setAlignment(Qt.AlignCenter)
        self.subtitle_label.setStyleSheet("color: rgba(255,255,255,0.7); font-size: 14px; border: none; background: transparent; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI';")

        frame_layout.addStretch()
        frame_layout.addWidget(self.logo_label)
        frame_layout.addSpacing(15)
        frame_layout.addWidget(self.title_label)
        frame_layout.addWidget(self.subtitle_label)
        frame_layout.addStretch()

        layout.addWidget(self.bg_frame)


# --- CUSTOM TITLE BAR (iOS TRAFFIC LIGHTS) ---
class CustomTitleBar(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setFixedHeight(40)
        self.setStyleSheet("background: transparent;")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 15, 0)
        layout.setSpacing(8)

        # iOS Traffic Light Buttons
        self.btn_close = QPushButton()
        self.btn_minimize = QPushButton()
        self.btn_maximize = QPushButton()

        self.setup_traffic_light(self.btn_close, "#FF5F56", "#FF5F56")
        self.setup_traffic_light(self.btn_minimize, "#FFBD2E", "#FFBD2E")
        self.setup_traffic_light(self.btn_maximize, "#27C93F", "#27C93F")

        self.btn_close.clicked.connect(self.parent.close)
        self.btn_minimize.clicked.connect(self.parent.showMinimized)
        self.btn_maximize.clicked.connect(self.toggle_maximize)

        layout.addWidget(self.btn_close)
        layout.addWidget(self.btn_minimize)
        layout.addWidget(self.btn_maximize)
        
        # Spacer
        layout.addStretch()

        # Title
        title_label = QLabel("🔮 Ultimate OS Icon Manager")
        title_label.setStyleSheet("color: rgba(255, 255, 255, 0.8); font-size: 14px; font-weight: bold; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI';")
        layout.addWidget(title_label)
        
        layout.addStretch()
        
        # Variables for moving the window
        self.start_pos = None

    def setup_traffic_light(self, btn, color, hover_color):
        btn.setFixedSize(12, 12)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                border-radius: 6px;
                border: 1px solid rgba(0,0,0,0.2);
            }}
            QPushButton:hover {{ background-color: {hover_color}; opacity: 0.8; }}
        """)

    def toggle_maximize(self):
        if self.parent.isMaximized():
            self.parent.showNormal()
        else:
            self.parent.showMaximized()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.start_pos:
            delta = event.globalPosition().toPoint() - self.start_pos
            self.parent.move(self.parent.pos() + delta)
            self.start_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        self.start_pos = None


class DropZone(QFrame):
    def __init__(self, title, accepted_exts=None, is_folder_allowed=True, allow_multiple=False):
        super().__init__()
        self.accepted_exts = accepted_exts or []
        self.is_folder_allowed = is_folder_allowed
        self.allow_multiple = allow_multiple
        self.current_paths = []

        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor) 
        
        self.setStyleSheet("""
            QFrame { 
                border: 2px dashed rgba(255, 255, 255, 0.2); 
                border-radius: 15px; 
                background-color: rgba(255, 255, 255, 0.05); 
            }
            QFrame:hover { 
                background-color: rgba(255, 255, 255, 0.1); 
                border: 2px dashed rgba(255, 255, 255, 0.4); 
            }
        """)
        self.setMinimumSize(250, 120)

        layout = QVBoxLayout(self)
        self.label = QLabel(title)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color: rgba(255,255,255,0.9); font-weight: bold; font-size: 14px; border: none; background: transparent;")
        
        self.path_label = QLabel("Drag & Drop here\nor click to browse")
        self.path_label.setAlignment(Qt.AlignCenter)
        self.path_label.setWordWrap(True)
        self.path_label.setStyleSheet("color: rgba(255,255,255,0.5); font-size: 12px; border: none; background: transparent;")

        layout.addStretch()
        layout.addWidget(self.label)
        layout.addWidget(self.path_label)
        layout.addStretch()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.is_folder_allowed and self.accepted_exts:
                menu = QMenu(self)
                menu.setStyleSheet("""
                    QMenu { 
                        background-color: rgba(30, 10, 50, 240); 
                        color: white; 
                        border: 1px solid rgba(255,255,255,0.1); 
                        border-radius: 8px;
                        font-size: 14px; 
                        padding: 5px; 
                    }
                    QMenu::item { padding: 8px 20px; border-radius: 4px; }
                    QMenu::item:selected { background-color: rgba(255,255,255,0.15); }
                """)
                action_files = menu.addAction("📄 Select Files")
                action_folder = menu.addAction("📁 Select Folder")
                
                selected_action = menu.exec(event.globalPos())
                
                if selected_action == action_files:
                    self.browse_files(mode="files")
                elif selected_action == action_folder:
                    self.browse_files(mode="folder")
            elif self.is_folder_allowed:
                self.browse_files(mode="folder")
            else:
                self.browse_files(mode="files")

    def browse_files(self, mode="files"):
        dialog = QFileDialog(self)
        dialog.setStyleSheet("QWidget { background-color: #1c0c2e; color: white; }")
        if mode == "folder":
            path = dialog.getExistingDirectory(self, "Select Folder", "", QFileDialog.ShowDirsOnly)
            if path: self.validate_and_set_paths([path])
        else:
            filter_str = "All Files (*.*)"
            if self.accepted_exts:
                exts = " ".join([f"*{ext}" for ext in self.accepted_exts])
                filter_str = f"Supported Files ({exts});;All Files (*.*)"

            if self.allow_multiple:
                paths, _ = dialog.getOpenFileNames(self, "Select Files", "", filter_str)
                if paths: self.validate_and_set_paths(paths)
            else:
                path, _ = dialog.getOpenFileName(self, "Select File", "", filter_str)
                if path: self.validate_and_set_paths([path])

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls]
        if not self.allow_multiple and len(paths) > 1:
            QMessageBox.warning(self, "Warning", "Please drop only ONE item here.")
            paths = [paths[0]]
        self.validate_and_set_paths(paths)

    def validate_and_set_paths(self, paths):
        valid_paths = []
        for path in paths:
            if os.path.isdir(path) and self.is_folder_allowed: valid_paths.append(path)
            elif os.path.isfile(path):
                ext = os.path.splitext(path)[1].lower()
                if not self.accepted_exts or ext in self.accepted_exts: valid_paths.append(path)

        if valid_paths: self.set_paths(valid_paths)
        else: QMessageBox.warning(self, "Invalid", "No valid files or folders were selected.")

    def set_paths(self, paths):
        self.current_paths = [os.path.abspath(p) for p in paths]
        if len(self.current_paths) == 1:
            self.path_label.setText(f"Selected:\n{os.path.basename(self.current_paths[0])}")
        else:
            self.path_label.setText(f"Selected:\n{len(self.current_paths)} items ready")
            
        self.setStyleSheet("""
            QFrame { border: 2px solid rgba(39, 201, 63, 0.6); border-radius: 15px; background-color: rgba(39, 201, 63, 0.1); }
        """)


class WorkerThread(QThread):
    progress = Signal(str)
    finished = Signal(int, list)

    def __init__(self, targets, icon_path, recursive=False, mode="apply"):
        super().__init__()
        self.targets = targets
        self.icon_path = icon_path
        self.recursive = recursive
        self.mode = mode

    def run(self):
        success_count = 0
        errors = []
        for target in self.targets:
            if os.path.isdir(target):
                if self.recursive:
                    for root, dirs, _ in os.walk(target):
                        try:
                            if self.mode == "apply": change_folder_icon(root, self.icon_path)
                            else: remove_folder_icon(root)
                            success_count += 1
                        except Exception as e: errors.append(f"{root}: {str(e)}")
                try:
                    if self.mode == "apply": change_folder_icon(target, self.icon_path)
                    else: remove_folder_icon(target)
                    success_count += 1
                except Exception as e: errors.append(f"{target}: {str(e)}")
            elif target.lower().endswith('.lnk') and self.mode == "apply":
                try:
                    change_shortcut_icon(target, self.icon_path)
                    success_count += 1
                except Exception as e: errors.append(f"{target}: {str(e)}")
        self.finished.emit(success_count, errors)


# --- CORE ICON LOGIC ---

def force_explorer_refresh():
    SHCNE_ASSOCCHANGED = 0x08000000
    SHCNF_IDLIST = 0x0000
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)

def change_folder_icon(folder_path, icon_path):
    ini_path = os.path.join(folder_path, "desktop.ini")
    if os.path.exists(ini_path):
        if WIN32_AVAILABLE: win32api.SetFileAttributes(ini_path, win32con.FILE_ATTRIBUTE_NORMAL)
        try: os.remove(ini_path)
        except: pass
    with open(ini_path, "w", encoding="mbcs") as f:
        f.write("[.ShellClassInfo]\n")
        f.write(f"IconResource={icon_path},0\n")
    if WIN32_AVAILABLE:
        win32api.SetFileAttributes(ini_path, win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM)
        win32api.SetFileAttributes(folder_path, win32con.FILE_ATTRIBUTE_READONLY)

def remove_folder_icon(folder_path):
    ini_path = os.path.join(folder_path, "desktop.ini")
    if os.path.exists(ini_path):
        if WIN32_AVAILABLE: win32api.SetFileAttributes(ini_path, win32con.FILE_ATTRIBUTE_NORMAL)
        os.remove(ini_path)
        if WIN32_AVAILABLE: win32api.SetFileAttributes(folder_path, win32con.FILE_ATTRIBUTE_NORMAL)

def change_shortcut_icon(shortcut_path, icon_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.IconLocation = f"{icon_path}, 0"
    shortcut.Save()

def set_drive_icon(drive_letter, icon_path):
    drive_root = f"{drive_letter}:\\"
    autorun_path = os.path.join(drive_root, "autorun.inf")
    if os.path.exists(autorun_path):
        win32api.SetFileAttributes(autorun_path, win32con.FILE_ATTRIBUTE_NORMAL)
    with open(autorun_path, "w") as f:
        f.write("[autorun]\n")
        f.write(f"ICON={icon_path}\n")
    win32api.SetFileAttributes(autorun_path, win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM)

def remove_drive_icon(drive_letter):
    autorun_path = os.path.join(f"{drive_letter}:\\", "autorun.inf")
    if os.path.exists(autorun_path):
        win32api.SetFileAttributes(autorun_path, win32con.FILE_ATTRIBUTE_NORMAL)
        os.remove(autorun_path)

def set_system_icon(name, icon_path):
    base_key = r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID"
    try:
        if name.startswith("Recycle Bin"):
            guid = "{645FF040-5081-101B-9F08-00AA002F954E}"
            state = "empty" if "Empty" in name else "full"
            key_path = fr"{base_key}\{guid}\DefaultIcon"
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, state, 0, winreg.REG_SZ, f"{icon_path},0")
            winreg.CloseKey(key)
        else:
            guid = SYSTEM_ICONS[name]
            key_path = fr"{base_key}\{guid}\DefaultIcon"
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f"{icon_path},0")
            winreg.CloseKey(key)
    except Exception as e:
        raise Exception(f"Registry Edit Failed: {str(e)}")


# --- MAIN UI ---

class UltimateIconApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.resize(750, 560) 
        
        base_layout = QVBoxLayout(self)
        base_layout.setContentsMargins(10, 10, 10, 10) 

        self.bg_frame = QFrame(self)
        self.bg_frame.setObjectName("AppContainer")
        self.bg_frame.setStyleSheet("""
            #AppContainer {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                    stop:0 rgba(45, 15, 65, 240), 
                    stop:0.5 rgba(25, 5, 40, 245), 
                    stop:1 rgba(15, 0, 25, 250));
                border-radius: 20px;
                border: 1px solid rgba(255, 255, 255, 0.15);
            }
        """)
        
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 150))
        shadow.setOffset(0, 5)
        self.bg_frame.setGraphicsEffect(shadow)

        main_layout = QVBoxLayout(self.bg_frame)
        main_layout.setContentsMargins(0, 0, 0, 15)
        main_layout.setSpacing(0)
        
        self.title_bar = CustomTitleBar(self)
        main_layout.addWidget(self.title_bar)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 10, 20, 10)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { 
                border: none; 
                background: transparent; 
                padding-top: 15px;
            }
            QTabBar::tab { 
                background: rgba(255,255,255, 0.05); 
                color: rgba(255,255,255,0.6); 
                padding: 10px 30px; 
                border-radius: 10px; 
                margin-right: 5px;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI';
                font-weight: bold;
                font-size: 13px;
            }
            QTabBar::tab:selected { 
                background: rgba(255,255,255, 0.15); 
                color: white; 
                border: 1px solid rgba(255,255,255,0.2);
            }
            QTabBar::tab:hover:!selected { 
                background: rgba(255,255,255, 0.1); 
            }
        """)
        content_layout.addWidget(self.tabs)

        self.init_bulk_tab()
        self.init_drive_tab()
        self.init_system_tab()

        footer = QLabel('Developed by <b>Janith Rathnayake</b> | <a href="http://jdr.kesug.com" style="color: #00d4ff; text-decoration: none;">Visit jdr.kesug.com</a>')
        footer.setOpenExternalLinks(True) 
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("margin-top: 10px; font-size: 12px; color: rgba(255,255,255,0.4); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI';")
        content_layout.addWidget(footer)

        main_layout.addWidget(content_widget)
        base_layout.addWidget(self.bg_frame)


    def init_bulk_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        zones_layout = QHBoxLayout()
        self.target_zone = DropZone("🎯 Targets (Folders, Shortcuts)\nSupports Multiple", accepted_exts=['.lnk'], is_folder_allowed=True, allow_multiple=True)
        self.icon_zone = DropZone("🖼 Icon (.ico, .exe, .dll)", accepted_exts=['.ico', '.exe', '.dll'], is_folder_allowed=False, allow_multiple=False)
        zones_layout.addWidget(self.target_zone)
        zones_layout.addWidget(self.icon_zone)
        layout.addLayout(zones_layout)

        options_layout = QHBoxLayout()
        self.chk_recursive = QCheckBox("Apply to ALL subfolders inside selected folders (Recursive)")
        self.chk_recursive.setStyleSheet("color: rgba(255,255,255,0.8); font-size: 13px;")
        options_layout.addWidget(self.chk_recursive)
        layout.addLayout(options_layout)

        btn_layout = QHBoxLayout()
        btn_apply = QPushButton("✨ Apply Icons")
        btn_remove = QPushButton("🗑 Remove Custom Icons")
        
        btn_style = """
            QPushButton { 
                background-color: rgba(184, 41, 255, 0.6); 
                border: 1px solid rgba(255,255,255,0.2); 
                border-radius: 12px; 
                padding: 12px; 
                color: white;
                font-weight: bold; 
                font-size: 14px;
            } 
            QPushButton:hover { background-color: rgba(184, 41, 255, 0.8); }
        """
        btn_remove_style = """
            QPushButton { 
                background-color: rgba(255, 68, 68, 0.3); 
                border: 1px solid rgba(255, 68, 68, 0.5); 
                border-radius: 12px; 
                padding: 12px; 
                color: white;
                font-weight: bold; 
                font-size: 14px;
            } 
            QPushButton:hover { background-color: rgba(255, 68, 68, 0.5); }
        """
        
        btn_apply.setStyleSheet(btn_style)
        btn_remove.setStyleSheet(btn_remove_style)
        
        btn_apply.clicked.connect(lambda: self.process_bulk("apply"))
        btn_remove.clicked.connect(lambda: self.process_bulk("remove"))

        btn_layout.addWidget(btn_apply)
        btn_layout.addWidget(btn_remove)
        layout.addLayout(btn_layout)

        self.tabs.addTab(tab, "📁 Folders")

    def init_drive_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        lbl = QLabel("Select a drive to change its root icon (Requires restarting Explorer):")
        lbl.setStyleSheet("color: rgba(255,255,255,0.8);")
        layout.addWidget(lbl)
        
        self.drive_combo = QComboBox()
        self.drive_combo.setStyleSheet("""
            QComboBox { 
                padding: 10px; font-size: 14px; color: white;
                background: rgba(255,255,255,0.05); 
                border: 1px solid rgba(255,255,255,0.2);
                border-radius: 8px;
            }
            QComboBox::drop-down { border: none; }
        """)
        drives = [f"{chr(x)}:" for x in range(65, 91) if os.path.exists(f"{chr(x)}:")]
        self.drive_combo.addItems(drives)
        layout.addWidget(self.drive_combo)

        self.drive_icon_zone = DropZone("🖼 Drop New Drive Icon (.ico)", accepted_exts=['.ico'], is_folder_allowed=False, allow_multiple=False)
        layout.addWidget(self.drive_icon_zone)

        btn_layout = QHBoxLayout()
        btn_apply = QPushButton("💾 Set Drive Icon")
        btn_reset = QPushButton("🗑 Reset Drive Icon")
        
        btn_apply.setStyleSheet("""
            QPushButton { background: rgba(39, 201, 63, 0.6); border: 1px solid rgba(255,255,255,0.2); padding: 12px; border-radius: 12px; font-weight: bold; color: white;} 
            QPushButton:hover { background: rgba(39, 201, 63, 0.8); }
        """)
        btn_reset.setStyleSheet("""
            QPushButton { background: rgba(255, 68, 68, 0.3); border: 1px solid rgba(255, 68, 68, 0.5); padding: 12px; border-radius: 12px; font-weight: bold; color: white;} 
            QPushButton:hover { background: rgba(255, 68, 68, 0.5); }
        """)

        btn_apply.clicked.connect(self.apply_drive_icon)
        btn_reset.clicked.connect(self.reset_drive_icon)

        btn_layout.addWidget(btn_apply)
        btn_layout.addWidget(btn_reset)
        layout.addLayout(btn_layout)

        self.tabs.addTab(tab, "💾 Drives")

    def init_system_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        lbl = QLabel("Override Core Windows Desktop Icons (Current User Only):")
        lbl.setStyleSheet("color: rgba(255,255,255,0.8);")
        layout.addWidget(lbl)
        
        self.sys_combo = QComboBox()
        self.sys_combo.setStyleSheet("""
            QComboBox { 
                padding: 10px; font-size: 14px; color: white;
                background: rgba(255,255,255,0.05); 
                border: 1px solid rgba(255,255,255,0.2);
                border-radius: 8px;
            }
            QComboBox::drop-down { border: none; }
        """)
        self.sys_combo.addItems(list(SYSTEM_ICONS.keys()))
        layout.addWidget(self.sys_combo)

        self.sys_icon_zone = DropZone("🖼 Drop System Icon (.ico, .dll)", accepted_exts=['.ico', '.dll'], is_folder_allowed=False, allow_multiple=False)
        layout.addWidget(self.sys_icon_zone)

        btn_apply = QPushButton("🖥 Apply System Icon")
        btn_apply.setStyleSheet("""
            QPushButton { background: rgba(0, 120, 215, 0.6); border: 1px solid rgba(255,255,255,0.2); padding: 12px; border-radius: 12px; font-weight: bold; font-size: 14px; color: white;} 
            QPushButton:hover { background: rgba(0, 120, 215, 0.8); }
        """)
        btn_apply.clicked.connect(self.apply_system_icon)
        layout.addWidget(btn_apply)

        self.tabs.addTab(tab, "💻 System")

    def process_bulk(self, mode):
        targets = self.target_zone.current_paths
        icon_paths = self.icon_zone.current_paths
        
        if not targets:
            QMessageBox.warning(self, "Error", "Drop or select target folders/files first!")
            return
        if mode == "apply" and not icon_paths:
            QMessageBox.warning(self, "Error", "Drop or select an icon file first!")
            return

        icon = icon_paths[0] if icon_paths else None
        recursive = self.chk_recursive.isChecked()

        self.setEnabled(False)
        self.thread = WorkerThread(targets, icon, recursive, mode)
        self.thread.finished.connect(self.on_bulk_finished)
        self.thread.start()

    def on_bulk_finished(self, success_count, errors):
        self.setEnabled(True)
        force_explorer_refresh()
        if errors:
            err_txt = "\n".join(errors[:5]) + ("\n..." if len(errors) > 5 else "")
            QMessageBox.warning(self, "Warning", f"Processed {success_count} items.\nErrors:\n{err_txt}")
        else:
            QMessageBox.information(self, "Success", f"✨ Magic successful! Modified {success_count} items.")

    def apply_drive_icon(self):
        drive = self.drive_combo.currentText()
        icon_paths = self.drive_icon_zone.current_paths
        if not icon_paths: return QMessageBox.warning(self, "Error", "Select an icon first!")
        try:
            set_drive_icon(drive, icon_paths[0])
            force_explorer_refresh()
            QMessageBox.information(self, "Success", f"Drive icon set for {drive}.\nRestart Explorer or replug drive to see changes.")
        except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def reset_drive_icon(self):
        try:
            remove_drive_icon(self.drive_combo.currentText())
            force_explorer_refresh()
            QMessageBox.information(self, "Success", "Drive icon reset.")
        except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def apply_system_icon(self):
        sys_name = self.sys_combo.currentText()
        icon_paths = self.sys_icon_zone.current_paths
        if not icon_paths: return QMessageBox.warning(self, "Error", "Select an icon first!")
        try:
            set_system_icon(sys_name, icon_paths[0])
            force_explorer_refresh()
            QMessageBox.information(self, "Success", f"{sys_name} icon updated!\nRight click desktop -> Refresh.")
        except Exception as e: QMessageBox.critical(self, "Error", f"Registry Error: {str(e)}")


if __name__ == "__main__":
    if not WIN32_AVAILABLE:
        print("ERROR: Please run: pip install pywin32")
        sys.exit(1)
        
    app = QApplication(sys.argv)

    # 1. Show our completely custom 3D Glass Splash Screen
    splash = GlassySplashScreen()
    splash.show()
    
    # Initialize the main app in the background
    window = UltimateIconApp()

    # 2. Use Qt's built-in timers to close the splash and open the app smoothly
    # Waits 2000 milliseconds (2 seconds) before transitioning
    QTimer.singleShot(2000, splash.close)
    QTimer.singleShot(2000, window.show)
    
    sys.exit(app.exec())