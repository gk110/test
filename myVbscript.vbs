Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'sample.xlsm'!Module1.SUM"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing