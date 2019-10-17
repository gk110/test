Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'C:\Users\gauravkumar4\Desktop\exce\sample.xlsm'!Module1.SUM"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing
