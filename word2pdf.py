import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon, QPixmap
from docx2pdf import convert
import win32com.client

class WordToPDFConverter(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Convertitore Word in PDF')
        self.setWindowIcon(QIcon('./imgs/converter-ico.ico'))

        self.layout = QVBoxLayout()

        image_label = QLabel(self)
        image_label.setPixmap(QPixmap('./imgs/convert-word-document-into-pdf.png').scaled(250, 250, Qt.KeepAspectRatio))
        image_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(image_label)

        self.drop_label = QLabel("Trascina e rilascia i file Word qui", self)
        font = QFont("Times New Roman", 12)
        self.drop_label.setFont(font)
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.drop_label)

        self.file_list_label = QLabel(self)
        self.layout.addWidget(self.file_list_label)

        convert_button = QPushButton("Converti", self)
        convert_button.clicked.connect(self.convert_to_pdf)
        convert_button.setMinimumHeight(50)
        self.layout.addWidget(convert_button)

        self.setLayout(self.layout)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        file_paths = [url.toLocalFile() for url in event.mimeData().urls()]
        self.input_files = [self.resolve_lnk(file) if file.lower().endswith('.lnk') else file for file in file_paths]
        self.update_file_list_label()

    def resolve_lnk(self, lnk_path):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)

        # Aggiunta di questa parte per ottenere il percorso del collegamento
        lnk_dir = os.path.dirname(lnk_path)
        lnk_name = os.path.basename(lnk_path)
        resolved_lnk_path = os.path.normpath(os.path.join(lnk_dir, lnk_name))


        return resolved_lnk_path, shortcut.TargetPath


    def update_file_list_label(self):
        if self.input_files:
            file_info_list = []
            for file_info in self.input_files:
                if isinstance(file_info, tuple):
                    # Se è una tupla (collegamento .lnk)
                    lnk_path, original_path = file_info
                    file_info_list.append(f"Collegamento: {os.path.basename(lnk_path)}\nOrigine: {os.path.basename(original_path)}")
                else:
                    # Se è un percorso diretto
                    file_info_list.append(f"{os.path.basename(file_info)}")

            file_names = "\n".join(file_info_list)
            self.file_list_label.setText(f"File selezionati:\n{file_names}")
        else:
            self.file_list_label.clear()

    def convert_to_pdf(self):
        for word_file_path in self.input_files:
            if isinstance(word_file_path, tuple):
                # Se è una tupla (collegamento .lnk)
                resolved_lnk_path, original_path = word_file_path
                resolved_lnk_path = os.path.dirname(resolved_lnk_path)
                pdf_file_path = os.path.join(resolved_lnk_path, os.path.splitext(os.path.basename(original_path))[0] + ".pdf")
            else:
                # Se è un percorso diretto
                original_path = word_file_path
                pdf_file_path = os.path.splitext(original_path)[0] + ".pdf"


            convert(original_path, pdf_file_path)

        self.file_list_label.setText("Conversione completata")
        self.input_files = []  # Resetta la lista dei file dopo la conversione



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WordToPDFConverter()
    ex.setGeometry(300, 300, 400, 200)
    ex.show()
    sys.exit(app.exec_())
