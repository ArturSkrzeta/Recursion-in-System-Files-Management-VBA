Option Explicit

Sub Main()

    Dim directory As String:    directory = ThisWorkbook.Sheets("Main").Range("B9").Value
    
    Call WalkFolders(directory)
    
    MsgBox "Done!", vbInformation

End Sub

Private Sub WalkFolders(directory As String)

    Dim fso As Object:              Set fso = CreateObject("Scripting.FileSystemObject")
    Dim startFolder As Object:      Set startFolder = fso.GetFolder(directory)
    Dim subFolder As Object
    Dim file As Object

    For Each file In startFolder.files
    
        Dim extension As String:    extension = fso.GetExtensionName(file)
        
        If extension = "docx" Or extension = "doc" Then
        
            Dim objWord As Object:      Set objWord = CreateObject("Word.Application")
            Dim doc As Object:          Set doc = objWord.Documents.Open(Filename:=directory & file.Name)
         
            doc.ExportAsFixedFormat OutputFileName:=Replace(doc.FullName, "." & extension, ".pdf"), _
                ExportFormat:=17, OpenAfterExport:=False
            doc.Close
           
            Set doc = Nothing
           
        End If
        
    Next
    
    For Each subFolder In startFolder.SubFolders
        Call WalkFolders(subFolder.Path & "\")
    Next subFolder

    Set startFolder = Nothing
    Set fso = Nothing
    
End Sub
