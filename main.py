from PyQt5 import QtWidgets
from ui.main_window import MainWindow

def main():
    app = QtWidgets.QApplication([])
    mainWindow = MainWindow()
    mainWindow.show()
    app.exec_()

if __name__ == "__main__":
    main()
