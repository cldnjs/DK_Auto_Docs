import sys

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from ExcelFunction import *


class Form(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.Widget)

        # 기본 저장 경로 존재여부 확인 후 설정
        if self.check_default_path() is False:
            QMessageBox.about(self, 'Message', '기본 저장 경로가 설정되지 않았습니다.')
            self.set_default_path()

        # 화면 및 레이아웃 선언
        self.setWindowTitle('DK Auto Docs')
        self.setFixedSize(500, 300)
        self.setWindowIcon(QIcon('./images/DK_logo.png'))
        layout = QBoxLayout(QBoxLayout.TopToBottom)
        self.setLayout(layout)

    @staticmethod
    def check_default_path():
        """
        기본 저장 경로가 설정되어있는지 확인
        :return: Bool
        """
        with open('./settings/path.txt', encoding='utf-8') as f:
            contents = f.readline()
            if len(contents) <= 0:
                return False
            else:
                return True

    def set_default_path(self):
        """
        기본 저장 경로를 설정
        :return:
        """
        file = str(QFileDialog.getExistingDirectory(self, '저장 경로 선택'))

        if len(file) <= 0:
            QMessageBox.about(self, 'Message', '저장 경로를 설정해주세요.')
            self.set_default_path()
        else:
            with open('./settings/path.txt', 'w', encoding='utf-8') as f:
                f.writelines(str(file))

    @staticmethod
    def get_save_path():
        """
        기본 저장 경로 반환
        :return: String
        """
        with open('./settings/path.txt', 'r', encoding='utf-8') as f:
            path = str(f.readline())

        return path


if __name__ == '__main__':
    # 엑셀 파일 로드
    wb, sheet = load_excel(
        filename='제품목록.xlsx',
        sheet_name='Sheet1',
        read_only=False,
        data_only=False
    )

    product_name = load_column_data(sheet, 2, 4)

    for name in product_name:
        if name is None:
            pass
        else:
            print(name)

    app = QApplication(sys.argv)
    form = Form()
    form.show()
    exit(app.exec_())
