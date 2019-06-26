Attribute VB_Name = "ExcelTools"
' Create a new folder for every selected cell in excel
Public Sub CreateFolders()
    Dim path As String
    path = InputBox("Paste in the location where you want folders to be created")
    
    For Each cell In Selection
        Dim location As String
        location = path & "\" & cell
        MkDir location
    Next
    
    MsgBox "Done"
End Sub
