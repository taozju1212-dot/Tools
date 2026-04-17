import sys
from PyQt5.QtWidgets import QApplication
from balance_app import BalanceMeasureApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')   # 统一跨平台外观
    window = BalanceMeasureApp()
    window.show()
    sys.exit(app.exec_())
