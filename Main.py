import sys
import os
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, 
                            QPushButton, QVBoxLayout, QHBoxLayout, QWidget,
                            QFileDialog, QMessageBox, QSpinBox, QScrollArea,
                            QSlider)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QImage, QPixmap
import fitz  # PyMuPDF
from pdf2docx import Converter
from docx2pdf import convert

class PDFPageViewer(QLabel):
    def __init__(self, parent=None):
        super(PDFPageViewer, self).__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.scale_factor = 1.0
        self.original_pixmap = None

    def load_page(self, image_data):
        height, width, channel = image_data.shape
        bytes_per_line = channel * width
        q_image = QImage(image_data.data, width, height, bytes_per_line, QImage.Format_RGB888)
        self.original_pixmap = QPixmap.fromImage(q_image)
        self.update_scale()

    def update_scale(self):
        if self.original_pixmap:
            # Convert float values to integers for scaling
            new_width = int(self.original_pixmap.width() * self.scale_factor)
            new_height = int(self.original_pixmap.height() * self.scale_factor)
            
            scaled_pixmap = self.original_pixmap.scaled(
                new_width,
                new_height,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.setPixmap(scaled_pixmap)


class PDFViewerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_pdf = None
        self.current_page = 0
        self.total_pages = 0
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Enhanced PDF Viewer & Converter')
        self.setGeometry(100, 100, 1200, 800)
        self.showMaximized()

        # Create central widget and layouts
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Create scroll area and viewer
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # Create viewer
        self.viewer = PDFPageViewer()
        scroll_area.setWidget(self.viewer)
        layout.addWidget(scroll_area)

        # Create zoom controls
        zoom_layout = QHBoxLayout()
        zoom_out_btn = QPushButton("Zoom Out")
        zoom_in_btn = QPushButton("Zoom In")
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(50)
        self.zoom_slider.setMaximum(300)
        self.zoom_slider.setValue(100)
        self.zoom_label = QLabel("100%")

        zoom_out_btn.clicked.connect(self.zoom_out)
        zoom_in_btn.clicked.connect(self.zoom_in)
        self.zoom_slider.valueChanged.connect(self.zoom_changed)

        zoom_layout.addWidget(zoom_out_btn)
        zoom_layout.addWidget(self.zoom_slider)
        zoom_layout.addWidget(self.zoom_label)
        zoom_layout.addWidget(zoom_in_btn)
        layout.addLayout(zoom_layout)

        # Create navigation controls
        nav_layout = QHBoxLayout()
        
        # Page navigation
        self.page_label = QLabel("Page: ")
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setValue(1)
        self.page_spin.valueChanged.connect(self.go_to_page)
        
        self.total_pages_label = QLabel("/ 0")
        
        # Create buttons
        open_btn = QPushButton("Open PDF")
        pdf_to_word_btn = QPushButton("Convert PDF to Word")
        word_to_pdf_btn = QPushButton("Convert Word to PDF")
        prev_btn = QPushButton("Previous Page")
        next_btn = QPushButton("Next Page")

        # Style buttons
        for btn in [open_btn, pdf_to_word_btn, word_to_pdf_btn, prev_btn, next_btn,
                   zoom_in_btn, zoom_out_btn]:
            btn.setMinimumWidth(120)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    padding: 8px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #1976D2;
                }
            """)

        # Connect buttons
        open_btn.clicked.connect(self.open_pdf)
        pdf_to_word_btn.clicked.connect(self.pdf_to_word)
        word_to_pdf_btn.clicked.connect(self.word_to_pdf)
        prev_btn.clicked.connect(self.prev_page)
        next_btn.clicked.connect(self.next_page)

        # Add navigation controls
        nav_layout.addWidget(open_btn)
        nav_layout.addWidget(pdf_to_word_btn)
        nav_layout.addWidget(word_to_pdf_btn)
        nav_layout.addStretch()
        nav_layout.addWidget(prev_btn)
        nav_layout.addWidget(self.page_label)
        nav_layout.addWidget(self.page_spin)
        nav_layout.addWidget(self.total_pages_label)
        nav_layout.addWidget(next_btn)

        layout.addLayout(nav_layout)

    def zoom_changed(self, value):
        self.zoom_label.setText(f"{value}%")
        self.viewer.scale_factor = value / 100
        self.viewer.update_scale()

    def zoom_in(self):
        self.zoom_slider.setValue(self.zoom_slider.value() + 10)

    def zoom_out(self):
        self.zoom_slider.setValue(self.zoom_slider.value() - 10)

    def go_to_page(self, page_num):
        if self.current_pdf and 1 <= page_num <= self.total_pages:
            self.current_page = page_num - 1
            self.load_current_page()

    def open_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF", "", "PDF Files (*.pdf)"
        )
        if file_path:
            try:
                self.current_pdf = fitz.open(file_path)
                self.total_pages = len(self.current_pdf)
                self.total_pages_label.setText(f"/ {self.total_pages}")
                self.page_spin.setMaximum(self.total_pages)
                self.current_page = 0
                self.page_spin.setValue(1)
                self.load_current_page()
                self.setWindowTitle(f'PDF Viewer - {os.path.basename(file_path)}')
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open PDF: {str(e)}")

    def load_current_page(self):
        if self.current_pdf is None:
            return

        page = self.current_pdf[self.current_page]
        # Increased zoom factor for clearer text
        zoom_matrix = fitz.Matrix(2.5, 2.5)
        pix = page.get_pixmap(matrix=zoom_matrix)
        
        img_data = np.frombuffer(pix.samples, dtype=np.uint8)
        img_data = img_data.reshape(pix.height, pix.width, pix.n)
        
        # Convert RGBA to RGB if necessary
        if img_data.shape[2] == 4:
            img_data = img_data[:, :, :3]
            
        self.viewer.load_page(img_data)
        self.page_spin.setValue(self.current_page + 1)

    def prev_page(self):
        if self.current_pdf and self.current_page > 0:
            self.current_page -= 1
            self.load_current_page()

    def next_page(self):
        if self.current_pdf and self.current_page < len(self.current_pdf) - 1:
            self.current_page += 1
            self.load_current_page()

    def pdf_to_word(self):
        input_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF", "", "PDF Files (*.pdf)"
        )
        if not input_path:
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save Word Document", "", "Word Files (*.docx)"
        )
        if not output_path:
            return

        try:
            QMessageBox.information(self, "Processing", 
                                  "Converting PDF to Word. Please wait...")
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()
            QMessageBox.information(self, "Success", 
                                  "PDF converted to Word successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Conversion failed: {str(e)}")

    def word_to_pdf(self):
        input_path, _ = QFileDialog.getOpenFileName(
            self, "Select Word Document", "", "Word Files (*.docx *.doc)"
        )
        if not input_path:
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF", "", "PDF Files (*.pdf)"
        )
        if not output_path:
            return

        try:
            QMessageBox.information(self, "Processing", 
                                  "Converting Word to PDF. Please wait...")
            convert(input_path, output_path)
            QMessageBox.information(self, "Success", 
                                  "Word converted to PDF successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Conversion failed: {str(e)}")



def main():
    app = QApplication(sys.argv)
    viewer = PDFViewerApp()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
