Attribute VB_Name = "Module1"
Public stopButtonClick As Boolean
Public autoSaveEn As Boolean

Sub ReadSerialDataContinuously()
    
    Dim rowIndex As Long
    Dim Input_Buffer As String
    Dim COM_Byte As Byte
    Dim wb As Workbook
    Dim comNum As String
    Dim comSetup As String
    
    Worksheets("Sheet1").Unprotect
    Set wb = ActiveWorkbook
    Range("A2", Cells(Rows.Count, Range("Setup").Column - 1)).ClearContents
    
    comNum = wb.Worksheets("Sheet1").Shapes("COMportNumber").OLEFormat.Object.Object.Value
    stopButtonClick = False
    Range("Status").Value = "Initiated"
    ' Set up COM port
    comSetup = ":" & Range("BaudRate").Value & ","
    Select Case Range("Parity").Value
        Case 2
            comSetup = comSetup & "e"
        Case 3
            comSetup = comSetup & "o"
        Case Else
            comSetup = comSetup & "n"
    End Select
    comSetup = comSetup & "," & Range("dataLength").Value & "," & Range("stopBits").Value
    Shell "mode.com com" & comNum & comSetup
    Range("Status").Value = "waiting"
    Worksheets("Sheet1").Protect
    Application.Wait (1) 'wait for COM port to be set up

    On Error Resume Next
    Open "COM" & comNum & ":" For Random As #1 Len = 1 ' open file in random access mode byte-by-byte
    If Err.Number <> 0 Then
        Worksheets("Sheet1").Unprotect
        Range("Status").Value = "COM error"
        Worksheets("Sheet1").Protect
        GoTo endsub
    End If
    On Error GoTo 0
    
    Input_Buffer = ""
    CharsRemaining = 0
    rowIndex = 2
    Worksheets("Sheet1").Unprotect
    Range("Status").Value = "Active"
    
    Do
        Get #1, , COM_Byte
        If COM_Byte Then
            If COM_Byte = 10 Then               ' look for \n
                Dim columnArray() As String
                columnArray = Split(Input_Buffer, ",")
                For col = 0 To UBound(columnArray)
                    If columnArray(col) <> "" Then Cells(rowIndex, col + 1).Value = columnArray(col)
                Next col
                rowIndex = rowIndex + 1
                Input_Buffer = ""
            Else
                Input_Buffer = Input_Buffer & Chr(COM_Byte)
            End If
        End If
        DoEvents
        ' Check if it's time to save data
        If (rowIndex - 1) Mod Range("AutoSaveLines").Value = 0 Then
            wb.Save
            DoEvents
        End If
    Loop Until stopButtonClick = True
    Close
    
    'now fix first row if misaligned
    Dim lastCol As Integer
    lastCol = 1
    While (Cells(3, lastCol + 1).Value <> "")
        lastCol = lastCol + 1
    Wend
    Dim shiftCols As Integer
    shiftCols = lastCol - Cells(2, 1).End(xlToRight).Column
    If shiftCols > 0 Then
        ' Cut the top data row cells and paste them to the right
        Range(Cells(2, 1), Cells(2, lastCol)).Cut Destination:=Cells(2, shiftCols + 1)
    End If
    
    Range("Status").Value = "Stopped"
    Worksheets("Sheet1").Protect
    
endsub:
End Sub

Sub StopButton_Click()
    stopButtonClick = True
End Sub

Sub ExportData()
    Dim SetupColumn As Integer
    Dim LastRow As Integer
    Dim ExportRange As Range
    Dim SaveFileDialog As FileDialog
    Dim FilePath As String

    ' Find the column number of the named cell "Setup"
    SetupColumn = Range("Setup").Column

    ' Find the last column in the worksheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Define the range to export
    Set ExportRange = Range(Cells(1, 1), Cells(LastRow, SetupColumn - 1))

    ' Create a FileDialog object to save the file
    Set SaveFileDialog = Application.FileDialog(msoFileDialogSaveAs)

    ' Show the Save As dialog
    If SaveFileDialog.Show = -1 Then
        FilePath = SaveFileDialog.SelectedItems(1)
        
        ' Export the range to a CSV file
        ExportRange.Copy
        Workbooks.Add(1).Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues
        ActiveWorkbook.SaveAs FilePath, Local:=True
        ActiveWorkbook.Close SaveChanges:=False
    End If
End Sub

