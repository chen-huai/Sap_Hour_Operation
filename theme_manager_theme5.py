from PyQt5.QtWidgets import QApplication

class ThemeManager:
    def __init__(self, app):
        self.app = app
        self.current_theme = "blue"

    def set_theme(self, theme):
        if theme == "light":
            self.set_light_theme()
        elif theme == "dark":
            self.set_dark_theme()
        elif theme == "blue":
            self.set_blue_theme()
        elif theme == "green":
            self.set_green_theme()
        elif theme == "purple":
            self.set_purple_theme()

    def set_light_theme(self):
        self.app.setStyleSheet("""
            QWidget {
                background-color: #f5f5f7;
                color: #333333;
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding: 8px 16px;
                color: #333333;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
                border-color: #d0d0d0;
            }
            QPushButton:pressed {
                background-color: #e8e8e8;
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding: 6px 10px;
                selection-background-color: #0078d4;
                selection-color: white;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #0078d4;
            }
            QComboBox {
                padding-right: 20px;
            }
            QComboBox::drop-down {
                border: none;
                border-left: 1px solid #e0e0e0;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(down_arrow.png);
                width: 12px;
                height: 12px;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
        """)
        self.current_theme = "light"

    def set_dark_theme(self):
        self.app.setStyleSheet("""
            QWidget {
                background-color: #1a1a1a;
                color: #ffffff;
            }
            QPushButton {
                background-color: #2d2d2d;
                border: 1px solid #404040;
                border-radius: 6px;
                padding: 8px 16px;
                color: #ffffff;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #353535;
                border-color: #505050;
            }
            QPushButton:pressed {
                background-color: #404040;
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: #2d2d2d;
                border: 1px solid #404040;
                border-radius: 6px;
                padding: 6px 10px;
                color: #ffffff;
                selection-background-color: #0078d4;
                selection-color: white;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #0078d4;
            }
            QComboBox {
                padding-right: 20px;
            }
            QComboBox::drop-down {
                border: none;
                border-left: 1px solid #404040;
                width: 20px;
            }
            QComboBox::down-arrow {theme_manager_theme.py
                image: url(down_arrow_dark.png);
                width: 12px;
                height: 12px;
            }
            QScrollBar:vertical {
                border: none;
                background: #2d2d2d;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #404040;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #505050;
            }
        """)
        self.current_theme = "dark"

    def set_blue_theme(self):
        self.app.setStyleSheet("""
            QWidget {
                background-color: #f0f4f8;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #3498db;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                color: white;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #2473a6;
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: white;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                padding: 6px 10px;
                selection-background-color: #3498db;
                selection-color: white;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #3498db;
            }
            QComboBox {
                padding-right: 20px;
            }
            QComboBox::drop-down {
                border: none;
                border-left: 1px solid #bdc3c7;
                width: 20px;
            }
            QScrollBar:vertical {
                border: none;
                background: #ecf0f1;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #bdc3c7;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #95a5a6;
            }
        """)
        self.current_theme = "blue"

    def set_green_theme(self):
        self.app.setStyleSheet("""
            QWidget {
                background-color: #f0f8f1;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #2ecc71;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                color: white;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:pressed {
                background-color: #219a52;
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: white;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                padding: 6px 10px;
                selection-background-color: #2ecc71;
                selection-color: white;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #2ecc71;
            }
            QComboBox {
                padding-right: 20px;
            }
            QComboBox::drop-down {
                border: none;
                border-left: 1px solid #bdc3c7;
                width: 20px;
            }
            QScrollBar:vertical {
                border: none;
                background: #ecf0f1;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #bdc3c7;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #95a5a6;
            }
        """)
        self.current_theme = "green"

    def set_purple_theme(self):
        self.app.setStyleSheet("""
            QWidget {
                background-color: #f8f0f8;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #9b59b6;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                color: white;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
            QPushButton:pressed {
                background-color: #803d9f;
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: white;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                padding: 6px 10px;
                selection-background-color: #9b59b6;
                selection-color: white;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #9b59b6;
            }
            QComboBox {
                padding-right: 20px;
            }
            QComboBox::drop-down {
                border: none;
                border-left: 1px solid #bdc3c7;
                width: 20px;
            }
            QScrollBar:vertical {
                border: none;
                background: #ecf0f1;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #bdc3c7;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #95a5a6;
            }
        """)
        self.current_theme = "purple"

    def toggle_theme(self):
        themes = ["light", "dark", "blue", "green", "purple"]
        current_index = themes.index(self.current_theme)
        next_index = (current_index + 1) % len(themes)
        self.set_theme(themes[next_index])
