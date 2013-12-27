' The following script uses the Excel Find method to locate all the cells in a spreadsheet that have a Value equal to :

Const xlValues = -4163

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("Test.xlsx")
Set objWorksheet = objWorkbook.Worksheets("Sheet1")

Set objRange = objWorksheet.UsedRange

Set objTarget = objRange.Find(4)

If Not objTarget Is Nothing Then
    Wscript.Echo objTarget.AddressLocal(False,False)
    strFirstAddress = objTarget.AddressLocal(False,False)
End If

Do Until (objTarget Is Nothing)
    Set objTarget = objRange.FindNext(objTarget)

    strHolder = objTarget.AddressLocal(False,False)
    If strHolder = strFirstAddress Then
        Exit Do
    End If

    Wscript.Echo objTarget.AddressLocal(False,False)
Loop