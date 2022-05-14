import sys
import re
from pathlib import Path
from docx import Document
from PyQt6.QtCore import QObject, pyqtSignal 
from PyQt6.QtWidgets import QMainWindow, QPushButton, QApplication, QTextEdit, QFileDialog
from PyQt6.QtGui import QTextCursor

 
class Stream(QObject):
    """Redirects console output to text widget."""
    newText = pyqtSignal(str)
 
    def write(self, text):
        self.newText.emit(str(text))
 
 
class GenMast(QMainWindow):
    """Main application window."""
    def __init__(self):
        super().__init__()
 
        self.initUI()
 
        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)
 
    def onUpdateText(self, text):
        """Write console output to text widget."""
        cursor = self.process.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(text)
        self.process.setTextCursor(cursor)
        self.process.ensureCursorVisible()
 
    def closeEvent(self, event):
        """Shuts down application on close."""
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__
        super().closeEvent(event)
 
    def initUI(self):
        """Creates UI window on launch."""
        # Button for generating the master list.
        btnGenMast = QPushButton('Open', self)
        btnGenMast.move(550, 50)
        btnGenMast.resize(100, 200)
        btnGenMast.clicked.connect(self.showDialog)
 
        # Create the text output widget.
        self.process = QTextEdit(self, readOnly=True)
        self.process.ensureCursorVisible()
        self.process.setLineWrapColumnOrWidth(500)
        self.process.setLineWrapMode(QTextEdit.LineWrapMode.FixedPixelWidth)
        self.process.setFixedWidth(500)
        self.process.setFixedHeight(300)
        self.process.move(30, 50)
        self.process.setText("请点击Open选择以docx为后缀的文件进行查重")

        # Set window size and title, then show the window.
        self.setGeometry(400, 400, 700, 400)
        self.setWindowTitle('题目查重器')
        self.show()
 
    def checkDupli(self, filepath):
    
        doc = Document(filepath)
        issue_pattern = re.compile(r'[0-9]{1,4}[\.。、]')
        issues = []
        paras = doc.paragraphs
        for _,p in enumerate(paras):
            issue_match = re.match(issue_pattern, p.text)
            if issue_match:
                issue = p.text.replace(issue_match.group(0), '').strip()
                if issue not in issues:
                    issues.append(issue)
                else:
                    print('重复：{}'.format(p.text))
 
    def showDialog(self):

        home_dir = str(Path.home())
        path,_ = QFileDialog.getOpenFileName(self, 'Open file', home_dir, "docx files (*.docx)")
        self.process.setText("")

        if path:
            print("### 文件路径是: {}  \n\n".format(path))
            self.checkDupli(path)
 
 
if __name__ == '__main__':
    # Run the application.
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)
    gui = GenMast()
    sys.exit(app.exec())
