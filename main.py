from openpyxl.utils import get_column_letter as getColLetters
from openpyxl.styles import Alignment
from openpyxl.worksheet.table import Table
from openpyxl import load_workbook

# Validate range params and pick max if not given


def getRange(ws, r0=None, r1=None, c0=None, c1=None):
    r0 = r0 if r0 is not None else ws.min_row
    r1 = r1 if r1 is not None else ws.max_row
    c0 = c0 if c0 is not None else ws.min_col
    c1 = c1 if c1 is not None else ws.max_col

    return {
        'r0': r0, 'r1': r1, 'c0': c0, 'c1': c1
    }


def AltA(ws):
    r1 = ws.max_row
    c1 = ws.max_column

    return f"A1:{getColLetters(c1)}{r1}"


def AltHOI(ws):
    colWidths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                # Calculate the maximum length of each col
                colWidths[cell.column_letter] = max(
                    (colWidths.get(cell.column_letter, 0)),
                    len(str(cell.value))
                )
    # Add a little extra width
    for col, width in colWidths.items():
        ws.column_dimensions[col].width = width + 3


def AltHK(ws, rg=None):
    # if list of range (rg) is missing, use max ranges
    if rg is None:
        rg = getRange(ws)

    for col in ws.iter_cols(
            min_row=rg['r0'], max_row=rg['r1'],
            min_col=rg['c0'], max_col=rg['c1']):
        for cell in col:
            cell.number_format = '#,##0;(#,##0);0'


def tableFormatWB(wb):
    for sh in wb.sheetnames:
        ws = wb[sh]
        ws.freeze_panes = 'A2'
        for col in ws.iter_cols(max_row=1, max_col=ws.max_column):
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
        ws.sheet_view.showGridLines = False
        wsData = AltA(ws)
        ws.add_table(Table(displayName=f'{sh}', ref=wsData))
        ws = AltHOI(ws)


def tableFormatWS(ws):
    ws.freeze_panes = 'A2'
    for col in ws.iter_cols(max_row=1, max_col=ws.max_column):
        for cell in col:
            cell.alignment = Alignment(horizontal='left')
    ws.sheet_view.showGridLines = False
    wsData = AltA(ws)
    ws.add_table(Table(displayName=f'{ws.title}', ref=wsData))
    ws = AltHOI(ws)


def formatWS(wb, shz=None):
    # if list of sheets not provided, pick all sheets in wb
    if shz is None:
        shz = wb.sheetnames

    # loop through all sheets in the given list
    for sh in shz:
        ws = wb[sh]
        tableFormatWS(ws)


def formatWB(wbNin, wbNout=None, ty=None, shz=None):
    wb = load_workbook(wbNin)

    # Format all sheets in wb with sheetname loop
    if ty == 'wb':
        tableFormatWB(wb)

    # Format all sheets in wb with additional control
    if ty == 'ws':
        formatWS(wb, shz)

    # To replace the formatted wb with original
    if wbNout is None:
        wbNout = wbNin

    wb.save(wbNout)


def main():
    wbNin = 'fileInput.xlsx'
    wbNout = 'fileOutput.xlsx'

    # Variable to control formatting function choice
    ty = 'wb'

    # Creating the list of desired sheets
    shz = ['Sheet1', 'Sheet3']

    # Formate sheets accordingly
    formatWB(wbNin, wbNout, ty, shz)


if __name__ == '__main__':
    main()
