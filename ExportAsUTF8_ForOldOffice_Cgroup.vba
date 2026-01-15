Sub ExportAsUTF8_ForOldOffice_Cgroup()
    Dim textStream As Object
    Dim rowRange As Range
    Dim cellRange As Range
    Dim strLine As String
    Dim strPath As Variant
    Dim rowCount As Long
    
    ' Stop screen flickering and pause calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler ' Catch unexpected errors
    
    ' Open the Save As Dialog
    strPath = Application.GetSaveAsFilename( _
        InitialFileName:="Export_Data_" & Format(Now, "yyyy-mm-dd"), _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Select Destination")
    
    If strPath = False Then Exit Sub ' User cancelled
    
    ' Initialize UTF-8 Engine
    Set textStream = CreateObject("ADODB.Stream")
    textStream.Type = 2
    textStream.Charset = "utf-8"
    textStream.Open
    
    ' Process Data
    rowCount = 0
    ' Only loop through the used range to keep it clean
    For Each rowRange In ActiveSheet.UsedRange.Rows
        strLine = ""
        For Each cellRange In rowRange.Cells
            strLine = strLine & Replace(cellRange.Value, ",", " ") & ","
        Next cellRange
        
        If Len(strLine) > 0 Then
            strLine = Left(strLine, Len(strLine) - 1) ' Remove last comma
            textStream.WriteText strLine, 1
            rowCount = rowCount + 1
        End If
    Next rowRange
    
    ' Save and Cleanup
    textStream.SaveToFile strPath, 2
    textStream.Close
    
    MsgBox "Success!" & vbCrLf & _
           "Rows Exported: " & rowCount & vbCrLf & _
           "Location: " & strPath, vbInformation, "Export Complete"

CleanExit:
    ' Turn optimizations back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: x_x " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

