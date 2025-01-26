Attribute VB_Name = "Update_WB_links"
Sub Update_WB_link()
    '=============================================
    'Update All Equations to match new sheet name
    '=============================================
    Dim old_v As String, new_v As String
    Dim wsData As Worksheet
    Dim StartTime As Double
    Dim i As Long, LR As Long
    Dim edit As String


    StartTime = Timer
    
    Set wsData = ThisWorkbook.Sheets("Data")

    ' Reset application settings in case of error
    On Error GoTo errHandler
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Set dynamic main folder with date by getting last update + 1 day
    wsData.Range("W2").Value = Int(FileDateTime(ThisWorkbook.FullName) + 1)

    ' Retrieve all files
    Call Get_rest_file

    ' Get current link source
    wsData.Range("W3").Value = ThisWorkbook.LinkSources
    old_v = wsData.Range("W3").Value
    
    Set My_Rng = wsData.Range("A10:A1048576")

    ' Start processing each file from the NewFiles array
    For i = 1 To FileCount ' FileCount keeps track of how many files in NewFiles
        
        new_v = NewFiles(i) ' Get file path from the array
        edit = new_v & Format(FileDateTime(new_v), "yyyymmdd_hhnnss")
        
        
        ' Validate new file path
        If Len(new_v) > 0 And Dir(new_v) <> "" And WorksheetFunction.CountIf(My_Rng, edit) = 0 Then
            ' Update "Data" sheet with primary data
            wsData.Range("B9:C9").Value = Array(new_v, Format(FileDateTime(new_v), "dd/mm/yyyy hh:mm:ss ampm"))
            wsData.Range("A9").Value = new_v & Format(FileDateTime(new_v), "yyyymmdd_hhnnss")

            ' Update workbook links
            ThisWorkbook.ChangeLink old_v, new_v, xlExcelLinks
            old_v = new_v ' Update old link for the next iteration

            ' Copy updated row to the next available row
            LR = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row + 1
            wsData.Rows(LR).Value = wsData.Rows(9).Value
        End If
    Next i

    ' Perform find/replace for "N/A"
    wsData.Cells.Replace what:="N/A", Replacement:="", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False

    ' Save workbook and log execution time
    ActiveWorkbook.Save
    wsData.Range("C1").Value = Format((Timer - StartTime) / 86400, "hh:mm:ss")

errHandler:
    ' Error handling and clean-up
    On Error Resume Next
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Set wsData = Nothing
    Set wsData = Nothing
    
End Sub


