import sys
from fileinput import filename
import os
import openpyxl as op
import pandas as pd
from datetime import datetime

from PyQt5.QtWidgets import *
from PyQt5 import uic


# ui 파일 절대 경로 처리
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

form = resource_path("bga_lib_maker_ui.ui")
form_class = uic.loadUiType(form)[0]


# Qtwidgets의 QMainWindow, ui파일의 form_class 상속
class WindowClass(QDialog, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # UI Setup

        # fileSelect 버튼 클릭시 selectFunction 메서드 동작
        self.fileSelect.clicked.connect(self.selectFunction)

    # selectFunction 메서드 정의
    def selectFunction(self):
        # filePath 출력하는 부분 초기화
        self.filePath.clear()
        # comboBox 출력하는 부분 초기화
        # 선택한 엑셀 파일 경로를 받아옴 : 튜플 타입으로 받아오며 0번재 요소가 주소값 string이다.
        path = QFileDialog.getOpenFileName(self, 'Open File', '', 'All File(*);; xlsx File(*.xlsx)')
        # filePath에 현재 읽어온 엑셀 파일 경로를 입력한다.(절대경로)
        self.filePath.setText(path[0])
        self.file_path = path[0]

        # 위 절대 경로 활용해 openpyxl workbook 객체 생성
        wb = op.load_workbook(path[0])
        # 설정한 workbook의 시트리스트를 읽어온다.
        self.shtlist = wb.sheetnames

        
        bga_org_df = self.read_xlsx()               # df 읽어 오기
        bga_list_df = self.bga_df_maker(bga_org_df) # df data 변환
        self.save_xlsx(bga_list_df)                 # df to xlsx 저장

    def read_xlsx(self):
        file_name = self.file_path
        df = pd.read_excel(file_name)
        return df

    def bga_df_maker(self, df):

        # columns list 생성 (1, 2, 3 ...) / str 타입 변경
        pin_int = df.columns.to_list()
        pin_int = list(map(str, pin_int))

        # df columns str로 통일
        for i in range(len(pin_int)):
            df = df.rename(columns={df.columns[i]: pin_int[i]})

        # rows list 생성 (A, B, C ...)
        pin_cha = df["Unnamed: 0"].to_list()

        # 첫 row의 빈 column 삭제
        pin_int.pop(0)

        # pin number list 생성 (A1, A2, A3 ...)
        pin_number = []  # 핀넘버
        for c in pin_cha:
            for i in pin_int:
                i = str(i)
                pin_number.append(c + i)

        # 각 column별 pin name list 생성
        pin_name_all = []
        for i in range(0, len(df["1"])):
            temp_list = df.loc[i].to_list()
            temp_list.pop(0)
            pin_name_all.append(temp_list)
        pin_name = []

        # pin name list 통합
        for i in pin_name_all:
            pin_name.extend(i)

        # 변경 df 생성
        list_df = pd.DataFrame({"pin number": pin_number, "pin name": pin_name})

        # empty 값 삭제
        d_empty_df = list_df.dropna(axis=0)

        # 중복값 column 생성
        d_empty_df["dupl num"] = d_empty_df.groupby("pin name").cumcount() + 1

        # 중복값 개수용 pin name list 생성
        pin_dupl_count_list = d_empty_df["pin name"].values.tolist()

        dupl_count_list = []

        # 중복값 개수 column 생성
        for c in pin_dupl_count_list:
            dupl = 0
            for d in pin_dupl_count_list:
                if c == d:
                    dupl += 1
            dupl_count_list.append(dupl)

        d_empty_df["dupl count"] = dupl_count_list

        # rows 갯수
        rows_count = d_empty_df.shape[0]

        # dupl count > 1 이상이면 pin_name에 dupl num numbering
        for i in range(rows_count):
            if d_empty_df["dupl count"][i] > 1:
                d_empty_df.loc[i, "pin name"] = d_empty_df.loc[i, "pin name"] + "_" + str(d_empty_df.loc[i, "dupl num"])

        out_df = d_empty_df.drop("dupl num", axis=1).drop("dupl count", axis=1)

        return out_df

    def save_xlsx(self, df):
        hour = str(datetime.now().hour)
        minute = str(datetime.now().minute)
        second = str(datetime.now().second)
        now_time = hour + minute + second
        df.to_excel(f"bga_pin_list_{now_time}.xlsx", index=False)


# GUI 출력 부분
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()