import sys
from PyQt5.QtWidgets import QApplication
from .gui import MainWindow


def main() -> None:
	app = QApplication(sys.argv)
	win = MainWindow()
	win.resize(1000, 700)
	win.show()
	sys.exit(app.exec_())


if __name__ == "__main__":
	main()
