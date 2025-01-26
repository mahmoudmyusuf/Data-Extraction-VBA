Attribute VB_Name = "Get_rest_files"
Public NewFiles() As String ' Declare the array globally to store new files
Public FileCount As Long ' To keep track of the number of files

Sub Get_rest_file()
    '======================================================
    '  Process all files in specified folder/Sub folders  '
    '======================================================
    Dim FileSystem As Object
    Dim HostFolder As String
    
    ' Clear the array before starting new file processing
    Erase NewFiles ' This clears the array and resets its size
    
    ' Get host folder path from cell D3
    HostFolder = ThisWorkbook.Sheets("Data").Range("T2").Value
    
    ' Validate host folder path
    If Dir(HostFolder, vbDirectory) = "" Then
        MsgBox "Invalid folder path in Data sheet Cell T2!", vbCritical
        Exit Sub
    End If
    
    ' Force recalculation of the workbook (if needed)
    Application.Calculate ' Ensure all calculations are updated before processing
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Process the folder
    DoFolder FileSystem.GetFolder(HostFolder)
    
    ' Reset status bar
    Application.StatusBar = False
End Sub

Sub DoFolder(Folder)
    Dim wsData As Worksheet
    Dim My_Rng As Range
    Dim SubFolder As Object, File As Object
    Dim LastRow  As Long, edit As String

    ' Set worksheet and range
    Set wsData = ThisWorkbook.Sheets("Data")
    Set My_Rng = wsData.Range("A10:A1048576")
    
    ' Loop through subfolders recursively
    For Each SubFolder In Folder.SubFolders
        Call DoFolder(SubFolder) ' Recursive call to process subfolders

        ' Check if folder matches criteria
        If SubFolder.Name Like wsData.Range("T3").Value Then
            ' Loop through files in folder
            For Each File In SubFolder.Files
                If File.Name Like wsData.Range("T4").Value And Not File.Name Like "*~*" Then
                
                    ' Generate unique identifier for file (name + timestamp)
                    edit = File.Path & Format(FileDateTime(File.Path), "yyyymmdd_hhnnss")
                    
                    ' Avoid duplicates using CountIf on the range
                    If WorksheetFunction.CountIf(My_Rng, edit) = 0 Then
                        
                        ' Add file path to the temporary array
                        FileCount = FileCount + 1
                        ReDim Preserve NewFiles(1 To FileCount) ' Resize array to add new file
                        NewFiles(FileCount) = File.Path ' Store the file path in the array
                        
                    End If
                End If
            Next File
        End If
    Next SubFolder
End Sub
