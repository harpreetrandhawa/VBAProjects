Option Explicit
 
ExcelMacroExample
 
Sub ExcelMacroExample()
 
  Dim xlApp
  Dim xlBook
 
  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = False
  xlApp.DisplayAlerts = False
  Set xlBook = xlApp.Workbooks.Open("E:\Proc_Reports\Reports\Handover sheet generation.xlsm")
  xlApp.Run "run_process"
  xlBook.Close False
  xlApp.Quit
 
  Set xlBook = Nothing
  Set xlApp = Nothing
 
End Sub
