from PyQt5.QtWidgets import QApplication,QMainWindow, QTextEdit, QMenuBar, QMenu,\
    QAction, QFileDialog, QMessageBox, QToolBar, QComboBox, QSpinBox, QColorDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

class notepad(QMainWindow):
    def __init__(self):
        super(notepad, self).__init__()
        self.setWindowTitle("NotePad")
        self.resize(800,500)
        self.editor = QTextEdit()
        self.setCentralWidget(self.editor)
        self.create_menu_bar()
        self.create_tool_bar()
        self.create_status_bar()
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

    def create_status_bar(self):
        self.statusbar = self.statusBar()

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

        self.statusbar.showMessage(f"opening {self.path.split('/')[-1]}", 5)

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
            self.statusbar.showMessage(f"saved {self.path.split('/')[-1]}", 5)
        elif self.path != "" and ".docx" not in self.path:
            with open(self.path,'w') as file:
                file.write(self.editor.toPlainText())
            self.statusbar.showMessage(f"saved {self.path.split('/')[-1]}", 5)
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
        return signal.text

    def create_tool_bar(self):
        toolbar = QToolBar()
        align_left = QAction("left",self)
        align_left.triggered.connect(lambda : self.set_alignment(Qt.AlignLeft))
        align_right = QAction("right", self)
        align_right.triggered.connect(lambda: self.set_alignment(Qt.AlignRight))
        align_center = QAction("center", self)
        align_center.triggered.connect(lambda: self.set_alignment(Qt.AlignCenter))
        align_justify = QAction("justify", self)
        align_justify.triggered.connect(lambda: self.set_alignment(Qt.AlignJustify))
        toolbar.addAction(align_left)
        toolbar.addAction(align_right)
        toolbar.addAction(align_center)
        toolbar.addAction(align_justify)

        self.fonts = QComboBox()
        self.fonts.addItems(["Courier Std", "Hellentic Typewriter Regular", "Helvetica", "Arial", "SansSerif", "Helvetica", "Times", "Monospace"])
        self.fonts.activated.connect(self.set_font)
        toolbar.addWidget(self.fonts)

        self.font_size = QSpinBox()
        self.font_size.setValue(12)
        self.font_size.valueChanged.connect(self.set_font_size)
        toolbar.addWidget(self.font_size)

        self.text_bold = QAction("Bold")
        self.text_bold.triggered.connect(self.bold_text)
        toolbar.addAction(self.text_bold)

        self.text_italics = QAction("Italics")
        self.text_italics.triggered.connect(self.italic_text)
        toolbar.addAction(self.text_italics)

        self.text_underline = QAction("Underline")
        self.text_underline.triggered.connect(self.underline_text)
        toolbar.addAction(self.text_underline)

        self.color_text = QAction("Colour")
        self.color_text.triggered.connect(self.text_colour)
        toolbar.addAction(self.color_text)

        self.highlight_text = QAction("Highlight")
        self.highlight_text.triggered.connect(self.text_highlight)
        toolbar.addAction(self.highlight_text)

        self.addToolBar(toolbar)

    def set_font(self):
        self.editor.setCurrentFont(QFont(self.fonts.currentText()))
        self.statusbar.showMessage(f"font is set to {self.fonts.currentText()}", 5)

    def set_alignment(self,alignment):
        self.editor.setAlignment(alignment)
        self.statusbar.showMessage(f"alignment is set to {alignment}", 5)

    def set_font_size(self):
        self.editor.setFontPointSize(self.font_size.value())
        self.statusbar.showMessage(f"font size is set to {self.font_size.value()}", 5)

    def bold_text(self):
        if self.editor.fontWeight() != QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            self.statusbar.showMessage(f"Text is bold now", 5)
        else:
            self.editor.setFontWeight(QFont.Normal)
            self.statusbar.showMessage(f"Text is normal now", 5)

    def underline_text(self):
        if not self.editor.fontUnderline():
            self.editor.setFontUnderline(True)
            self.statusbar.showMessage(f"Text is underlined now", 5)
        else:
            self.editor.setFontUnderline(False)
            self.statusbar.showMessage(f"Text is not underlined now", 5)

    def italic_text(self):
        if not self.editor.fontItalic():
            self.editor.setFontItalic(True)
            self.statusbar.showMessage(f"Text is italic now", 5)
        else:
            self.editor.setFontItalic(False)
            self.statusbar.showMessage(f"Text is normal now", 5)

    def text_colour(self):
        color = QColorDialog.getColor()
        if color:
            self.editor.setTextColor(color)
            self.statusbar.showMessage(f"Text color is set to {color.name()} now", 5)

    def text_highlight(self):
        color = QColorDialog.getColor()
        if color:
            self.editor.setTextBackgroundColor(color)
            self.statusbar.showMessage(f"Text highlight is set to {color.name()} now", 5)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = notepad()
    window.show()
    app.exec()


