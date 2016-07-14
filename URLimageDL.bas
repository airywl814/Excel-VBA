Attribute VB_Name = "URLimageDL"
'''''''''''''''''''''''''''''''DOWNLOAD URL Barcode IMAGE TO LOCAL HARD DRIVE
Option Explicit
    #If VBA7 Then
    '64-Bit Windows OS
        Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    #Else
        Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    #End If
    Dim Check As Long
    Const FolderName As String = "F:\Desktop\temp\"
    
Sub DownloadBarcode()
    Dim ws As Worksheet
    Set ws = Sheets("Sheet2")
        
    Dim b As Range
    Set b = Range("B1:B9999")
    Dim i As Integer
    i = 1
    
    Dim strPath As String
    Dim url As String
    url = "https://www.barcodesinc.com/generator/image.php?code=" & Worksheets("Sheet2").Cells(i, 2) & "&style=197&type=C128B&width=128&height=50&xres=1&font=3"

    For i = 2 To b.Rows.Count
        If Cells(i, 2) <> "" Then
            strPath = FolderName & ws.Range("B" & i).Value & ".jpg"
            Check = URLDownloadToFile(0, url, strPath, 0, 0)
            If Check = 0 Then
                ws.Range("C" & i).Value = "Successful"
            Else
                ws.Range("C" & i).Value = "Failed to download"
            End If
        End If
    Next i
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

'''''''''''INSERT BARCODE IMAGE TO EXCEL
Sub Barcode_Click()
       
    Dim url As String
    'url = "https://www.barcodesinc.com/generator/image.php?code=121121&style=197&type=C128B&width=128&height=50&xres=1&font=3"
    
    Dim b As Range
    Set b = Range("B1:B9999")
    Dim i As Integer
    i = 1
    
    url = "https://www.barcodesinc.com/generator/image.php?code=" & Worksheets("Sheet2").Cells(i, 2) & "&style=197&type=C128B&width=128&height=50&xres=1&font=3"
    
    'With ActiveSheet.Pictures.Insert(Filename:=url, LinkToFile:=False, SaveWithDocument:=Ture)

    With ActiveSheet.Pictures.Insert(url)
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Width = 75
            .Height = 100
        End With
        .Left = ActiveSheet.Cells(i, 3).Left
        .Top = ActiveSheet.Cells(i, 3).Top
        .Placement = 1
        .PrintObject = True
    End With
    
    
    

End Sub







