import openpyxl as xl
import pandas as pd


def run(quma, rechnung, speicherort):
    wb_quma = xl.load_workbook(quma)
    wb_rechnung = xl.load_workbook(rechnung)

    ws_quma = wb_quma.active
    ws_rechnung = wb_rechnung.active

