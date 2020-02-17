# -*- coding: utf-8 -*-

##--------------------------------------------------------##
## Excelワークブックを開く(*.xlsx版)
##--------------------------------------------------------##
def readExcel_xlsx(T_WB):
    for T_WS in T_WB:
        print(T_WS.Name)


##--------------------------------------------------------##
## Excelワークブックを開く(*.xls版)
##--------------------------------------------------------##
def readExcel_xls(T_WB):
    """
    PythonでExcelファイルを読み込み・書き込みするxlrd, xlwtの使い方
    https://note.nkmk.me/python-xlrd-xlwt-usage/
    """
    hostname_pattern = ["-c", "-PN-c"]

    WorkSheets = T_WB.sheets()
    """
    for T_WS in WorkSheets:
        print(T_WS.name)
    """

    if len(WorkSheets) > 2:
        printCenterSW(WorkSheets[1], hostname_pattern, 40)
    if len(WorkSheets) > 3:
        printCenterSW(WorkSheets[2], hostname_pattern, 40)


##--------------------------------------------------------##
## センターSWの設定ポートの行を出力
##--------------------------------------------------------##
def printCenterSW(T_WS, hostname_pattern, LineCountMAX):
    isFound = False

    for i in range(0, LineCountMAX):
        try:
            row, col = cellAddressToNumber("E18")
            portnum = str(T_WS.cell(row + i - 1, col - 1).value)
            hostname = str(T_WS.cell(row + i - 1, col + 3 - 1).value)
            vrf = str(T_WS.cell(row + i - 1, col + 6 - 1).value)

            if len(portnum) == 0 and isFound:
                readCells(T_WS, "E18", i, 12)
                continue

            isFound = False

            if vrf != "10":
                # VRFが"10"でない場合は次の行へ
                continue

            for targetStr in hostname_pattern:
                if targetStr in hostname:
                    # ホスト名に "-c" または "-PN-c" が含まれている場合
                    isFound = True
                    readCells(T_WS, "E18", i, 12)
                    break
        except:
            break

##--------------------------------------------------------##
## 指定セルの内容を表示
##--------------------------------------------------------##
def readCells(T_WS, T_Address, offset_row, count):
#   resultStr = T_WS.Range(T_Address).cells(row, 0).address() + ":"
    row, col = cellAddressToNumber(T_Address)
    resultStr = convertToAddress(row + offset_row, col) + ":"

    for j in range(count):
        if type(T_WS.cell(row + offset_row - 1, col + j - 1).value) is float:
            val = str(int(T_WS.cell(row + offset_row - 1, col + j - 1).value))
        else:
            val = str(T_WS.cell(row + offset_row - 1, col + j - 1).value)
        val = val.replace("\n", "")

        if len(val) < 15:
            resultStr += "{0:<15s}".format(val)
        else:
            resultStr += "{0:<20s}".format(val)
    
    print(resultStr)

##--------------------------------------------------------##
## セル番地から(行番号, 列番号)に変換する
##--------------------------------------------------------##
def cellAddressToNumber(address):
    for i in range(len(address)):
        if address[i].isnumeric():
            break
    col_address = address[:i]
    row_address = address[i:]
    number = 0
    for i in range(0, len(col_address)):
        number *= 26
        number += ord(col_address[i]) - ord('A') + 1
    return int(row_address), number

##--------------------------------------------------------##
## 行番号, 列番号からセル番地に変換する
##--------------------------------------------------------##
def convertToAddress(row, col):
    if col < 1:
        return ""
    temp = col - 1
    target = []
    while True:
        mod = temp % 26
        target.append(mod)
        temp -= mod
        if temp >= 26:
            temp = int(temp / 26) - 1
            continue
        else:
            break
        
    result = ""
    for i in range(len(target)):
        result = chr(ord('A') + target[i]) + result

    return result + str(row)
