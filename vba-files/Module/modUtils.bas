Attribute VB_Name = "modUtils"
Public Function SelectFolderOrNull( _
    Optional ByVal initialFolderPath As String _
) As Variant

    If Not StringHasContent(initialFolderPath) Then initialFolderPath = CheckLastBackSlash(CreateObject("WScript.Shell").SpecialFolders("MyDocuments"))

    With Application.FileDialog(msoFileDialogFolderPicker)
        
        .InitialFileName = initialFolderPath
        .Title = modValues.folderSelectionTitle
        .Show

        If .SelectedItems.Count <> 0 Then
            SelectFolderOrNull = CheckLastBackSlash(.SelectedItems(1))
        Else
            SelectFolderOrNull = Null
        End If

    End With

End Function
Public Function CheckLastBackSlash( _
    ByVal str As String _
) As String

    If VBA.Right(str, 1) <> "\" Then
        CheckLastBackSlash = str & "\"
    Else
        CheckLastBackSlash = str
    End If

End Function
Public Function StringHasContent( _
    ByVal str As String _
) As Boolean

    StringHasContent = (VBA.Len(VBA.Trim(str)) > 0)

End Function
Public Function IsStringInArray( _
    ByVal str As String, _
    ByVal arr As Variant _
) As Boolean

    Dim filterArray As Variant

    filterArray = VBA.Filter(arr, str)
    
    If UBound(filterArray) <> -1 Then IsStringInArray = True
    
End Function
