Set objExcel = CreateObject("Excel.Application")
objExcel.Visible=true
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\lxdie\A&B corp\LE DODO - Documents\MacroDODO.xlsm")
WScript.Sleep 1000 * 10* 1
objExcel.Run "MacroDODO.xlsm!ThisWorkbook.UpdateCoucher"
objWorkbook.Save
objWorkbook.close
objExcel.DisplayAlerts = True
objExcel.Application.Quit