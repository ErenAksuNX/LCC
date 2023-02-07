import openpyxl as xl
import datetime


def run(quma, rechnung, speicherort):
    wb_quma = xl.load_workbook(quma)
    wb_rechnung = xl.load_workbook(rechnung)

    ws_quma = wb_quma.active
    ws_rechnung = wb_rechnung.active

    for r_zeile in ws_rechnung.iter_rows(values_only=True):
        if r_zeile[18] == "x" or r_zeile[18] == "X" or r_zeile[0] == "Geschäftsjahr":
            continue

        datum: datetime.datetime = r_zeile[2]
        r_monat = datum.month
        werk = r_zeile[19]
        frz = r_zeile[20]

        print(r_monat)

        for q_zeile in ws_quma.iter_rows(values_only=True):

            q_monat = q_zeile[32]
            if q_monat == "Monat" or not werk in q_zeile[13] or not frz in q_zeile[0]:
                continue

            if not (r_monat == 12 or r_monat == 1) and int(r_monat) <= int(q_monat) <= (int(r_monat) + 3) % 12:
                break
