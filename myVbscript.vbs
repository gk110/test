Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'C:\Users\gauravkumar4\Downloads\Copy of sample (002).xlsm'!Module1.SUM"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing