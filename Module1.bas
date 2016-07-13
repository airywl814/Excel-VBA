Attribute VB_Name = "Module1"

Sub Button1_Click()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("Sheet1")

    'Dim myprinter As String
    'myprinter = "Microsoft XPS Document Writer"

    Dim datarange As Range
    Set datarange = Range("H1:H10005")

    '  For RowIndex = 1 To datarange.Rows.Count
    ' step is optional if increment by 1

    For RowIndex = 1 To datarange.Rows.Count Step 2

        If WS1.Cells(RowIndex, 8) <> "" And WS1.Cells(RowIndex + 1, 8) <> "" Then

            WS1.Cells(1, 5) = WS1.Cells(RowIndex, 8)
            WS1.Cells(1, 6) = WS1.Cells(RowIndex + 1, 8)
            
            'This shows the printer dialog and allows user to choose the printer
            'Application.Dialogs(xlDialogPrinterSetup).Show
          '  ActiveSheet.PrintPreview
            'ActiveSheet.PrintOut

            Application.ActivePrinter = "Microsoft XPS Document Writer on Ne00:"

            'if user cancels printing, don't empty the cells
            On Error GoTo printcanceled
                'ActiveSheet.PrintOut From:=2, To:=3, Copies:=3, PrintToFile:= True, PrToFileName:="er.xps"
                ActiveSheet.PrintOut
                WS1.Cells(RowIndex, 8) = ""
                WS1.Cells(RowIndex + 1, 8) = ""
printcanceled:
            Exit For
        End If
    Next RowIndex

End Sub


'Sub Button1_Click()
'
'
'    Dim WS1 As Worksheet
'    Set WS1 = ThisWorkbook.Worksheets("Sheet1")
'
'    Dim myprinter As String
'    myprinter = "Microsoft XPS Document Writer"
'
'
'    Dim datarange As Range
'
'    Set datarange = Range("H1:H10005")
'
'    '  For RowIndex = 1 To datarange.Rows.Count
'    ' step is optional if increment by 1
'
'    For RowIndex = 1 To datarange.Rows.Count Step 2
'
'        If WS1.Cells(RowIndex, 8) <> "" And WS1.Cells(RowIndex + 1, 8) <> "" Then
'
'            WS1.Cells(1, 5) = WS1.Cells(RowIndex, 8)
'            WS1.Cells(1, 6) = WS1.Cells(RowIndex + 1, 8)
'
'
'
'            'Application.Dialogs(xlDialogPrinterSetup).Show
'          '  ActiveSheet.PrintPreview
'            'ActiveSheet.PrintOut
'
'            'myprinter = ActivePrinter
'            Application.ActivePrinter = "Microsoft XPS Document Writer on Ne00:"
'
'
'
'            ''''''''' if you know the name of the printer but not NeXX: then uncommemnt the follwoing
'            For i = 1 To 9
'                On Error GoTo WrongPrinter
'                Application.ActivePrinter = "Microsoft XPS Document Writer on Ne0" & i & ":"
'
'            On Error GoTo printcanceled
'                ActiveSheet.PrintOut
'
'                WS1.Cells(RowIndex, 8) = ""
'                WS1.Cells(RowIndex + 1, 8) = ""
'
'
'                Exit For
'CheckNextPrinter:
'            Next i
'
'            Exit Sub
'
'WrongPrinter:
'                MsgBox ("Microsoft XPS Document Writer on Ne0" & i & ":")
'                Resume CheckNextPrinter
'
'printcanceled:
'            Exit For
'
'        End If
'
'
'
'    Next RowIndex
'
'
'
'End Sub






