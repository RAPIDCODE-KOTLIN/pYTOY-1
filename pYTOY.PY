import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QInputDialog, QLabel)
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt
from pdf2image import convert_from_path
import img2pdf
from PIL import Image
import cv2
import os
import comtypes.client  # For Word to PDF conversion
from PyPDF2 import PdfWriter, PdfReader

class MultiUtilityApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.setWindowTitle('Multi-Utility App')
        self.setGeometry(100, 100, 400, 400)

        # Set dark theme
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(25, 25, 25))
        palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ToolTipBase, Qt.white)
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ButtonText, Qt.black)
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Highlight, QColor(142, 45, 197).lighter())
        palette.setColor(QPalette.HighlightedText, Qt.black)
        self.setPalette(palette)

        # Set up the UI components
        layout = QVBoxLayout()

        # Buttons for various utilities
        pdf_to_img_btn = QPushButton("PDF to Image", self)
        pdf_to_img_btn.setStyleSheet("color: black;")
        img_to_pdf_btn = QPushButton("Image to PDF", self)
        img_to_pdf_btn.setStyleSheet("color: black;")
        word_to_pdf_btn = QPushButton("Word to PDF", self)
        word_to_pdf_btn.setStyleSheet("color: black;")
        compress_img_btn = QPushButton("Compress Image", self)
        compress_img_btn.setStyleSheet("color: black;")
        lock_pdf_btn = QPushButton("Lock PDF with Password", self)
        lock_pdf_btn.setStyleSheet("color: black;")
        passport_size_btn = QPushButton("Make Passport Size Photo", self)
        passport_size_btn.setStyleSheet("color: black;")

        layout.addWidget(pdf_to_img_btn)
        layout.addWidget(img_to_pdf_btn)
        layout.addWidget(word_to_pdf_btn)
        layout.addWidget(compress_img_btn)
        layout.addWidget(lock_pdf_btn)
        layout.addWidget(passport_size_btn)

        # Assign functionality to buttons
        pdf_to_img_btn.clicked.connect(self.pdf_to_image)
        img_to_pdf_btn.clicked.connect(self.image_to_pdf)
        word_to_pdf_btn.clicked.connect(self.word_to_pdf)
        compress_img_btn.clicked.connect(self.compress_image)
        lock_pdf_btn.clicked.connect(self.lock_pdf)
        passport_size_btn.clicked.connect(self.make_passport_size)

        # Set layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def pdf_to_image(self):
        pdf_path, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF Files (*.pdf)")
        if pdf_path:
            images = convert_from_path(pdf_path)
            for i, img in enumerate(images):
                img.save(f"output_page_{i}.png", "PNG")
            self.show_message("PDF Converted to Images")

    def image_to_pdf(self):
        img_paths, _ = QFileDialog.getOpenFileNames(self, "Select Images", "", "Image Files (*.png *.jpg *.jpeg)")
        if img_paths:
            pdf_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
            with open(pdf_path, "wb") as f:
                f.write(img2pdf.convert([Image.open(img).filename for img in img_paths]))
            self.show_message("Images Converted to PDF")

    def word_to_pdf(self):
        word_path, _ = QFileDialog.getOpenFileName(self, "Select Word Document", "", "Word Files (*.docx)")
        if word_path:
            pdf_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(word_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            word.Quit()
            self.show_message("Word Document Converted to PDF")

    def compress_image(self):
        img_path, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Image Files (*.png *.jpg *.jpeg)")
        if img_path:
            img = Image.open(img_path)
            compressed_img_path, _ = QFileDialog.getSaveFileName(self, "Save Compressed Image", "", "JPEG Files (*.jpg)")
            img.save(compressed_img_path, "JPEG", quality=30)
            self.show_message("Image Compressed")

    def lock_pdf(self):
        pdf_path, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF Files (*.pdf)")
        if pdf_path:
            password, ok = QInputDialog.getText(self, 'Password', 'Enter password for PDF:')
            if ok and password:
                reader = PdfReader(pdf_path)
                writer = PdfWriter()

                for page in reader.pages:
                    writer.add_page(page)

                writer.encrypt(password)

                encrypted_pdf_path, _ = QFileDialog.getSaveFileName(self, "Save Encrypted PDF", "", "PDF Files (*.pdf)")
                with open(encrypted_pdf_path, "wb") as f:
                    writer.write(f)
                self.show_message("PDF Locked with Password")

    def make_passport_size(self):
        img_path, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Image Files (*.png *.jpg *.jpeg)")
        if img_path:
            self.show_message("Crop the image to the appropriate size.")

            image = cv2.imread(img_path)
            r = cv2.selectROI("Select Area", image)
            cropped_img = image[int(r[1]):int(r[1] + r[3]), int(r[0]):int(r[0] + r[2])]

            num_photos, ok = QInputDialog.getInt(self, 'Number of Photos', 'Enter the number of passport size photos you need:', min=1, max=50)
            if not ok:
                return

            size_options = ['20x25mm', '35x45mm']
            size, ok = QInputDialog.getItem(self, "Select Size", "Choose passport size:", size_options, 0, False)
            if not ok:
                return

            size_in_pixels = (236, 295) if size == '20x25mm' else (413, 531)

            resized_img = cv2.resize(cropped_img, size_in_pixels)
            resized_image_path = "resized_passport_image.png"
            cv2.imwrite(resized_image_path, resized_img)

            canvas = self.create_passport_photo_page(resized_image_path, num_photos, size_in_pixels)

            save_path, _ = QFileDialog.getSaveFileName(self, "Save Passport Photos", "", "JPEG Files (*.jpg);;PDF Files (*.pdf)")

            if save_path:
                if save_path.endswith(".jpg"):
                    canvas.save(save_path, "JPEG")
                else:
                    canvas.save("temp_passport_photos.jpg", "JPEG")
                    with open(save_path, "wb") as f:
                        f.write(img2pdf.convert("temp_passport_photos.jpg"))
                self.show_message(f"Passport photos saved to {save_path}")
                os.remove("temp_passport_photos.jpg")

    def create_passport_photo_page(self, image_path, num_photos, photo_size):
        a4_size = (2480, 3508)
        canvas = Image.new("RGB", a4_size, "white")

        img = Image.open(image_path)
        img_width, img_height = photo_size

        x_offset, y_offset = 100, 100
        spacing = 50

        for _ in range(num_photos):
            canvas.paste(img, (x_offset, y_offset))
            x_offset += img_width + spacing
            if x_offset + img_width > a4_size[0]:
                x_offset = 100
                y_offset += img_height + spacing

        return canvas

    def show_message(self, message):
        label = QLabel(message, self)
        label.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MultiUtilityApp()
    main_window.show()
    sys.exit(app.exec_())
