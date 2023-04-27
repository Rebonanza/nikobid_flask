lrow1 = sht1.Cells(Rows.count, 1).End(xlUp).Row
sht3.Cells.Clear()

sht1.Rows(1).Copy()
sht3.Range("A1").PasteSpecial()

k = 2
for i in range(2, lrow1 + 1):
    if sht1.Cells(i, 26) == "":
        sht1.Cells(i, 26) = sht1.Cells(i, 25)
    if sht1.Cells(i, 34) > sht4.Range("H11").Value:
        if sht1.Cells(i, 42).Value < (sht4.Range("J17").Value * 0.01) or sht1.Cells(i, 42).Value > (sht4.Range("K17").Value * 0.01):
            if UCase(sht4.Range("J19").Value) == "NO" or (UCase(sht4.Range("J19").Value) == "YES" and sht1.Cells(i, 35) > 0):
                sht1.Range("BA" + str(i)).Formula = "=MATCH(H" + str(i) + ",'60 Days'!H:H,0)"
                if IsError(sht1.Range("BA" + str(i)).Value):
                    pass
                else:
                    if sht2.Cells(sht1.Range("BA" + str(i)).Value, 42) < (sht4.Range("J18").Value * 0.01) or sht2.Cells(sht1.Range("BA" + str(i)).Value, 42) > (sht4.Range("K18").Value * 0.01):
                        if sht2.Cells(sht1.Range("BA" + str(i)).Value, 35) > sht4.Range("J16").Value:
                            sht1.Rows(i).Copy()
                            sht3.Range("A" + str(k)).PasteSpecial()
                            sht3.Cells(k, 52) = sht2.Cells(sht1.Range("BA" + str(i)).Value, 42)
                            impres = sht3.Range("AH" + str(k)).Value
                            if sht4.Range("E22").Value == "BY Percentage":
                                Ans = WorksheetFunction.Lookup(impres, sht4.Range("C11:C20"), sht4.Range("E11:E20"))
                                sht3.Range("BC" + str(k)).Value = sht3.Range("Z" + str(k)).Value * (100 - Ans) / 100
                            else:
                                Ans = WorksheetFunction.Lookup(impres, sht4.Range("C11:C20"), sht4.Range("F11:F20"))
                                sht3.Range("BC" + str(k)).Value = sht3.Range("Z" + str(k)).Value - Ans
                            if (sht3.Range("Z" + str(k)).Value - sht3.Range("BC" + str(k)).Value) < 0.01:
                                sht3.Range("BC" + str(k)).Value = sht3.Range("Z" + str(k)).Value - 0.01
                            if sht3.Range("BC" + str(k)).Value < sht4.Range("J11"):
                                sht3.Range("BC" + str(k)).Value = sht4.Range("J11")
                            k += 1
sht3.range(f"AQ2:AQ{k}").value = sht3.range(f"BA2:BA{k}").value

sht3.range(f"AQ1:AQ{k-1}").interior.color = 65535  # vbYellow equivalent in Excel color code
sht3.range(f"C2:C{k-1}").value = "Update"
sht3.range("AQ1").value = "Acos 60 Days"

sht3.columns("AA:AA").insert(shift="right")
sht3.range(f"AA2:AA{k}").value = sht3.range(f"BE2:BE{k}").value

sht3.range(f"Z1:Z{k-1}").interior.color = 65535  # vbYellow equivalent in Excel color code
sht3.range("AA1").value = "Bid"
sht3.range("Z1").value = "Old Bid"

sht1.columns(53).clear()

sht3.columns(54).clear()
sht3.columns(55).clear()
sht3.columns(57).clear()

sht3.activate()
sht3.range("A1").select()
sht3.range("A1:AR5000").sort(key1=sht3.range("AI1"), order1=2, header=-1)  # xlDescending = 2, xlYes = -1

def copy_man:
    for k in range(1, 3):
    if k == 1:
        sht1 = ThisWorkbook.Sheets("Yesterday")
        sht_name = "Y_SP"
    else:
        sht1 = ThisWorkbook.Sheets("60 Days")
        sht_name = "60_SP"
    
    sht1.Cells.Clear()
    count = WorksheetsInStr(sht_name)
    
    for i in range(1, count + 1):
        sht2 = ThisWorkbook.Worksheets(sht_name + (" " + str(count) if count > 1 else ""))
        
        if sht2.AutoFilterMode == True:
            pass
        else:
            sht2.Range("A1").AutoFilter()
        
        lr = sht1.Cells(sht1.Rows.Count, "A").End(xlUp).Row
        lr = lr + (1 if lr > 1 else 0)
        
        lrow = sht2.Cells(Rows.Count, 1).End(xlUp).Row
        
        sht2.Range("$A$1:$AR$" + str(lrow)).AutoFilter(Field=2, Criteria1="Keyword")
        
        sht2.Rows("1:" + str(lrow)).Copy()
        
        sht1.Range("A" + str(lr)).PasteSpecial()
        
        Application.CutCopyMode = False
        
        try:
            sht2.ShowAllData()
        except:
            pass

def opn:
    import win32com.client as win32
import os

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True
xl.DisplayAlerts = False

sht3 = xl.ActiveWorkbook.Sheets("Final")
sht4 = xl.Sheets("Dashboard")

fd = xl.FileDialog(3)

fd.Filters.Clear()
fd.Filters.Add("Excel Files", "*.xlsx?", 1)
fd.Title = "Choose an Excel file"
fd.AllowMultiSelect = False

if fd.Show() == -1:
    strFile = fd.SelectedItems(0)
else:
    strFile = ""

if strFile == "":
    xl.ScreenUpdating = True
    xl.Quit()
    print("No file is selected please select a valid file")
    exit()

if xl.Caller == "Rectangle 3":
    sht4.Range("C3").Value = strFile
    sht4.Range("C4").Value = "Yesterday file Import Completed"
    sht1 = xl.ActiveWorkbook.Sheets("Yesterday")
else:
    sht4.Range("C6").Value = strFile
    sht4.Range("C7").Value = "60 Days file Import Completed"
    sht1 = xl.ActiveWorkbook.Sheets("60 Days")

sht1.Cells.Clear()

wb1 = xl.ThisWorkbook
wb2 = xl.Workbooks.Open(strFile)
wb2.Sheets("Sponsored Products Campaigns").Select()
xl.Selection.AutoFilter()

lrow = wb2.Sheets("Sponsored Products Campaigns").Cells(xl.Rows.count, 1).End(-4162).Row
xl.ActiveSheet.Range("$A$1:$AR$" + str(lrow)).AutoFilter(2, "Keyword")
xl.Rows("1:" + str(lrow)).Copy()
wb1.Activate()
sht1.Activate()
sht1.Range("A1").PasteSpecial(-4163)
xl.CutCopyMode = False
wb2.Close(False)

sht4.Activate()
xl.ScreenUpdating = True
xl.Quit()


def copy_video:
    for k in range(1, 3):
    if k == 1:
        sht1 = ThisWorkbook.Sheets("Yesterday")
        sht_name = "Y_SB"
    else:
        sht1 = ThisWorkbook.Sheets("60 Days")
        sht_name = "60_SB"
    
    sht1.Cells.Clear()
    count = WorksheetsInStr(sht_name)
    
    for i in range(1, count+1):
        sht2 = ThisWorkbook.Worksheets(sht_name + (" " + str(count) if count > 1 else ""))
        
        if sht2.AutoFilterMode == True:
            pass
        else:
            sht2.Range("A1").AutoFilter()
        
        lr = sht1.Cells(sht1.Rows.Count, "A").End(xlUp).Row
        lr = lr + (1 if lr > 1 else 0)
        
        lrow = sht2.Cells(Rows.Count, 1).End(xlUp).Row
        
        sht2.Range("$A$1:$AX$" + str(lrow)).AutoFilter(Field=2, Criteria1="Product Targeting")
        sht2.Range("$A$1:$AX$" + str(lrow)).AutoFilter(Field=25, Criteria1="=*category*", Operator=xlOr, Criteria2="=*asin*")
        sht2.Range("$A$1:$AX$" + str(lrow)).AutoFilter(Field=11, Criteria1="=*video*")
        sht2.Range("$A$1:$AX$" + str(lrow)).AutoFilter(Field=22, Criteria1="<>")
        
        sht2.Rows("1:" + str(lrow)).Copy()
        
        sht1.Range("A" + str(lr)).PasteSpecial()
        
        Application.CutCopyMode = False
        
        On Error Resume Next
        sht2.ShowAllData()
        
        On Error GoTo 0
