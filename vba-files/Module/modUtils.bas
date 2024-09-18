Attribute VB_Name = "modUtils"
Public Function SelectFolderOrNull( _
    Optional ByVal initialFolderPath As String _
) As Variant
    If Not StringHasContent(initialFolderPath) Then
        initialFolderPath = CheckLastBackSlash(CreateObject("WScript.Shell").SpecialFolders("MyDocuments"))
    End If
    
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
    Dim filterArray As Variant: filterArray = VBA.Filter(arr, str)
    
    If UBound(filterArray) <> -1 Then IsStringInArray = True
End Function

Public Sub SetCellSizeInInches( _
    ByVal cell As Range, _
    Optional ByVal height As Single, _
    Optional ByVal width As Single _
)
    Dim i As Integer
    
    If height > 0 Then cell.RowHeight = Application.InchesToPoints(height)
    
    If width > 0 Then
        If cell.width = 0 Then cell.ColumnWidth = 72
        
        For i = 1 To 2
            cell.ColumnWidth = Application.InchesToPoints(width) * (cell.ColumnWidth / cell.width)
        Next
    End If
End Sub

Public Sub SetRowHeightInInches( _
    ByVal row As Range, _
    Optional ByVal height As Single _
)
    If height > 0 Then row.RowHeight = Application.InchesToPoints(height)
End Sub

Public Sub SetColumnWidthInInches( _
    ByVal column As Range, _
    Optional ByVal width As Single _
)
    Dim i As Integer
    
    If width > 0 Then
        If column.width = 0 Then column.ColumnWidth = 72
        
        For i = 1 To 2
            column.ColumnWidth = Application.InchesToPoints(width) * (column.ColumnWidth / column.width)
        Next
    End If
End Sub

Public Function CheckShapeExist( _
    ByVal name As String, _
    ByVal sheet As Worksheet _
) As Boolean
    Dim shp As shape
    
    For Each shp In sheet.Shapes
        If shp.name = name Then
            CheckShapeExist = True
        Else
            CheckShapeExist = False
        End If
    Next shp
End Function
