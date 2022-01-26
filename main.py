from PyQt5.QtWidgets import QApplication,QMainWindow, QTextEdit, QMenuBar, QMenu,\
    QAction, QFileDialog, QMessageBox
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

class notepad(QMainWindow):
    def __init__(self):
        super(notepad, self).__init__()
        self.setWindowTitle("NotePad")
        self.resize(500,500)
        self.editor = QTextEdit()
        self.setCentralWidget(self.editor)
        self.create_menu_bar()
        self.path = ""

    def create_menu_bar(self):
        menubar = QMenuBar()
        file_menu = QMenu("File", self)
        menubar.addMenu(file_menu)
        open_action = QAction('Open', self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        save_action = QAction("Save",self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        self.setMenuBar(menubar)

    def open_file(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "",
                                                   "Text documents (*.text);Text documents (*.txt);All files (*.*)")
        extension = self.path.split(".")[-1]

        text_file_formats = ['txt','py','java','c']

        docx_file_format = ['docx']

        if extension in text_file_formats:
            with open(self.path,'r') as file:
                text_in_file = file.readlines()
                self.editor.setText("".join(text_in_file))
        elif extension in docx_file_format:
            with ZipFile(self.path,'r') as file:
                text = ""
                list_of_files = file.namelist()
                text+=self.read_docx(file.read('word/document.xml'))
                self.editor.setText(text)

    def read_docx(self,xml):
        text = ""
        root = ET.fromstring(xml)
        for child in root.iter():
            if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t":
                t_text = child.text
                text += t_text if t_text is not None else ''
            elif child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab":
                text += '\t'
            elif child.tag in ("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br", "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cr"):
                text += '\n'
            elif child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p":
                text += '\n\n'
        return text


    def save_file(self):
        if not self.path:
            self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "text documents (*.text);Text documents (*.txt);All files (*.*)")
        elif self.path != "" and ".docx" not in self.path:
            with open(self.path,'w') as file:
                file.write(self.editor.toPlainText())
        elif ".docx" in self.path:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(".docx format")
            msg.setInformativeText('The person who made this program is too lazy to write code for writing docx files. You wish to set a new path?')
            msg.setWindowTitle("Error")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            response = msg.buttonClicked.connect(self.button_clicked)
            if "Yes" in response:
                self.path = ""
                self.save_file()
            msg.exec_()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Something went really wrong')
            msg.setWindowTitle("Error")
            msg.exec_()

    def button_clicked(self,signal):
        print(signal.text)
        return signal.text

    def set_font(self):
        pass

    def set_font_size(self):
        pass

    def set_font_style(self):
        pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = notepad()
    window.show()
    app.exec()


