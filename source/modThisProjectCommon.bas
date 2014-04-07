Attribute VB_Name = "modThisProjectCommon"
Option Explicit


Public Function GetSubFolder(ByVal ParentFolder As MAPIFolder, ByVal SubFolderName As String) As MAPIFolder
    On Error Resume Next
    Dim oReturnFolder As MAPIFolder
    Set oReturnFolder = ParentFolder.Folders(SubFolderName)
    If Err.Number <> 0 Then
        Set oReturnFolder = Nothing
    End If
    Set GetSubFolder = oReturnFolder
End Function


