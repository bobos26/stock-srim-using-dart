import time

import OpenDartReader
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import requests as requests
import openpyxl


class Stock:
    API_KEY = 'a5c976581917ad4df2ac3509b69d8717f4d15d56'
    COLUMNS = ['회사명', '지배주주지분', 'ROE_2019', 'ROE_2020', 'ROE_2021', 'ROE_2022', '발행주식수', '자사주', '할인율']
    YEAR = 2020
    df = None
    dart = None
    data = {}

    def __init__(self, path):
        self.path = path

        self.df = self.read_xlsx()
        self.dart = self.init_dart()

    def read_xlsx(self):
        df = pd.read_excel(self.path, sheet_name='Sheet1', dtype=str)
        data = np.full((len(df), len(self.COLUMNS)), 0)
        df_add = pd.DataFrame(data=data, columns=self.COLUMNS)
        df = pd.concat([df, df_add], axis=1)
        return df

    def write_xlsx(self):
        with pd.ExcelWriter(self.path, engine='openpyxl', mode='a') as writer:
            writer.book = openpyxl.load_workbook(self.path)
            self.df.to_excel(writer, sheet_name='Result', index=False)

    def init_dart(self):
        dart = OpenDartReader(self.API_KEY)
        return dart

    @staticmethod
    def get_data_of_param(sub_soup, dic, param):
        def str_to_float(dataOfParam):
            out = []
            for data in dataOfParam:
                try:
                    out.append(float(data.get_text().strip()))
                except Exception as e:
                    out.append(0)
            return out

        sub_tbody = sub_soup.find("table", attrs={"class": "tb_type1 tb_num tb_type1_ifrs"}).find("tbody")
        sub_title = sub_tbody.find("th", attrs={"class": param}).get_text().strip()

        dataOfParam = sub_tbody.find("th", attrs={"class": param}).parent.find_all("td")
        value_param = str_to_float(dataOfParam)
        dic[sub_title] = value_param

    def get_roe(self):
        dict = {}
        for i, code in enumerate(self.df['code']):
            link = 'https://finance.naver.com/item/main.nhn?code={0}'.format(code)
            sub_res = requests.get(link)  # 링크를 통해 우리가 원하는 기업별 데이터 페이지 데이터 크롤링
            sub_soup = BeautifulSoup(sub_res.text, 'html.parser')
            ParamList = ['매출액', '영업이익', '당기순이익', 'ROE(지배주주)']

            for idx, pText in enumerate(ParamList):
                param = " ".join(sub_soup.find('strong', text=pText).parent['class'])
                self.get_data_of_param(sub_soup, dict, param)
            self.df.loc[i, ['ROE_2019', 'ROE_2020', 'ROE_2021', 'ROE_2022']] = dict['ROE(지배주주)'][0:4]

    def get_dart(self):
        def int_validate(series):
            str = series.iloc[0] if len(series) == 1 else '0'
            if str is '-':
                return 0
            else:
                return int(str.replace(',', ''))

        for i, code in enumerate(self.df['code']):
            df2 = self.dart.finstate(code, self.YEAR)
            self.data['자산총계'] = int_validate(df2.loc[(df2['fs_nm'] == '연결재무제표') & (df2['sj_nm'] == '재무상태표') & (
                        df2['account_nm'] == '자산총계'), 'thstrm_amount'])
            self.data['부채총계'] = int_validate(df2.loc[(df2['fs_nm'] == '연결재무제표') & (df2['sj_nm'] == '재무상태표') & (
                        df2['account_nm'] == '부채총계'), 'thstrm_amount'])
            self.data['지배주주지분'] = self.data['자산총계'] - self.data['부채총계']
            df2 = self.dart.report(code, '소액주주', self.YEAR)  # 발행주식수
            if df2 is not None:
                self.data['발행주식수'] = int_validate(df2.loc[df2['se'] == '소액주주', 'stock_tot_co'])
            else:
                self.data['발행주식수'] = 0
            df2 = self.dart.report(code, '자기주식', self.YEAR)  # 자사주
            if df2 is not None:
                self.data['자사주'] = int_validate(df2.loc[(df2['stock_knd'] == '보통주') & (df2['acqs_mth2'] == '직접취득') & (
                            df2['acqs_mth3'] == '소계'), 'bsis_qy'])
            else:
                self.data['자사주'] = 0
            # 회사명
            self.data['회사명'] = self.dart.company(code)['stock_name']
            self.df.loc[i, ['회사명', '지배주주지분', '발행주식수', '자사주']] = \
                [self.data['회사명'], self.data['지배주주지분'], self.data['발행주식수'], self.data['자사주']]


def main():
    stock = Stock()
    stock.get_roe()
    stock.get_dart()
    stock.write_xlsx()


if __name__ == '__main__':
    main()