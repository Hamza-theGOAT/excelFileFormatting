from openpyxl.utils import get_column_letter as getColLetters
from openpyxl.styles import Alignment
from openpyxl.worksheet.table import Table
from openpyxl import load_workbook


def getRange(ws, rg=None):
    """
    Get the usable row and column range of a worksheet.

    Args:
        ws (Worksheet): The openpyxl worksheet object.
        rg (dict, optional): A dictionary with keys 'r0', 'r1', 'c0', 'c1'.
            - 'r0': Minimum row
            - 'r1': Maximum row
            - 'c0': Minimum column
            - 'c1': Maximum column
            If any value is None or missing, the worksheet's min/max values are used.

    Returns:
        dict: A dictionary with normalized values for the range:
              { 'r0': int, 'r1': int, 'c0': int, 'c1': int }
    """
    r0 = rg.get('r0') if rg.get('r0') is not None else ws.min_row
    r1 = rg.get('r1') if rg.get('r1') is not None else ws.max_row
    c0 = rg.get('c0') if rg.get('c0') is not None else ws.min_col
    c1 = rg.get('c1') if rg.get('c1') is not None else ws.max_col

    return {
        'r0': r0, 'r1': r1, 'c0': c0, 'c1': c1
    }


def AltA(ws):
    """
    Get the full data range of a worksheet as an Excel reference string.

    Args:
        ws (Worksheet): The openpyxl worksheet object.

    Returns:
        str: Excel-style range reference (e.g., "A1:D20")
    """
    r1 = ws.max_row
    c1 = ws.max_column

    return f"A1:{getColLetters(c1)}{r1}"


def AltHOI(ws):
    """
    Auto-adjust column widths based on cell contents.

    Args:
        ws (Worksheet): The openpyxl worksheet object.

    Notes:
        - Width is based on the longest string length in each column.
        - Adds +3 padding to avoid truncated values.
    """
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
    rg = getRange(ws)

    for col in ws.iter_cols(
            min_row=rg['r0'], max_row=rg['r1'],
            min_col=rg['c0'], max_col=rg['c1']):
        for cell in col:
            cell.number_format = '#,##0.00;(#,##0.00);-'


def AltHN_Cus(ws, temp, rg=None):
    # if list of range (rg) is missing, use max ranges
    rg = getRange(ws)

    for col in ws.iter_cols(
            min_row=rg['r0'], max_row=rg['r1'],
            min_col=rg['c0'], max_col=rg['c1']):
        for cell in col:
            cell.number_format = f'{temp}'


def AltHNS(ws, rg=None):
    # if list of range (rg) is missing, use max ranges
    rg = getRange(ws)

    for col in ws.iter_cols(
            min_row=rg['r0'], max_row=rg['r1'],
            min_col=rg['c0'], max_col=rg['c1']):
        for cell in col:
            # Check for cell value being a date or datetime
            if cell.is_date:
                # Apply date format MM/DD/YYYY
                cell.number_format = 'mm/dd/yyyy'


def AltHNS_Cus(ws, temp, rg=None):
    # if list of range (rg) is missing, use max ranges
    rg = getRange(ws)

    for col in ws.iter_cols(
            min_row=rg['r0'], max_row=rg['r1'],
            min_col=rg['c0'], max_col=rg['c1']):
        for cell in col:
            # Check for cell value being a date or datetime
            if cell.is_date:
                # Apply date format MM/DD/YYYY
                cell.number_format = f'{temp}'


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
    if ty is None:
        tableFormatWB(wb)

    # Format all sheets in wb with additional control
    if ty == 'ws':
        formatWS(wb, shz)

    # Specific column based formatting
    if ty == 'sp':
        AltHK(wb['common'], {'r0': 2, 'r1': 4, 'c0': 3, 'c1': 3})
        AltHNS(wb['added'], {'r0': 2, 'r1': 3, 'c0': 2, 'c1': 2})

    # To replace the formatted wb with original
    if wbNout is None:
        wbNout = wbNin

    wb.save(wbNout)


def main():
    wbNin = 'fileInput.xlsx'
    wbNout = 'fileOutput.xlsx'

    # Variable to control formatting function choice
    ty = 'ws'

    # Creating the list of desired sheets
    shz = ['added', 'common']

    # Formate sheets accordingly
    formatWB(wbNin, wbNout)
    formatWB(wbNin, wbNout, ty, shz)


if __name__ == '__main__':
    main()
