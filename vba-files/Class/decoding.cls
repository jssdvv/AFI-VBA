VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "decoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ListImagesWithinAFolder()

    Dim image As String
    Dim folder As Variant
    Dim imageFormats As Variant
    Dim currentFormat As Variant
    Dim i As Long
    
    folder = modUtils.SelectFolderOrNull(Me.Cells(2, 3).Value)
    
    If IsNull(folder) Then Exit Sub
    
    Me.Cells(2, 3).Value = folder
    imageFormats = Array("tiff", "jpeg", "jpg", "png", "gif")
    image = Dir(folder & "*.*")
    i = 1
    
    Do While image <> ""
        currentFormat = VBA.Split(image, ".")
        If modUtils.IsStringInArray(currentFormat(UBound(currentFormat)), imageFormats) Then
            Me.Cells(4 + i, 2).Value = image
            i = i + 1
        End If
        image = Dir
    Loop

End Sub
Sub DecodeBarcodesFromImagesList()

    Dim image As String
    Dim folder As Variant
    Dim lastListRow As Long
    Dim firstListRow As Long
    Dim decodedBarcode As String
    Dim currentFormat As Variant
    Dim i As Long
    
    lastListRow = Me.Cells(Rows.Count, 2).End(xlUp).Row
    firstListRow = Me.Cells(lastListRow, 2).End(xlUp).Offset(1, 0).Row
    
    If (lastListRow - firstListRow) <= 0 Then Exit Sub
    
    folder = Me.Cells(2, 3).Value
    
    If Not modUtils.StringHasContent(folder) Then
        folder = modUtils.SelectFolderOrNull(Me.Cells(2, 3).Value)
        Me.Cells(2, 3).Value = folder
    End If
    
    If IsNull(folder) Then
        MsgBox modValues.folderMustNotBeEmpty
        Exit Sub
    End If
    
    image = Dir(folder & "*.*")
    i = 1
    
    For i = firstListRow To lastListRow
        image = folder & Me.Cells(i, 2).Value
        currentFormat = VBA.Split(image, ".")
        decodedBarcode = modBarcodeReader.decodeBarcode128(image)
        
        If modUtils.StringHasContent(decodedBarcode) Then
            Me.Cells(i, 3).Value = decodedBarcode & "." & currentFormat(UBound(currentFormat))
        End If
    Next i
    
End Sub
Sub RenameImagesWithDecodedBarcode()

    Dim image As String
    Dim folder As Variant
    Dim lastListRow As Long
    Dim firstListRow As Long
    
    Dim oldName As String
    Dim newName As String
    Dim i As Long
    
    lastListRow = Me.Cells(Rows.Count, 2).End(xlUp).Row
    firstListRow = Me.Cells(lastListRow, 2).End(xlUp).Offset(1, 0).Row
    
    If (lastListRow - firstListRow) <= 0 Then Exit Sub
    
    folder = Me.Cells(2, 3).Value
    
    If Not modUtils.StringHasContent(folder) Then
        folder = modUtils.SelectFolderOrNull(Me.Cells(2, 3).Value)
        Me.Cells(2, 3).Value = folder
    End If
    
    If IsNull(folder) Then
        MsgBox modValues.folderMustNotBeEmpty
        Exit Sub
    End If
    
    image = Dir(folder & "*.*")
    i = 1
    
    For i = firstListRow To lastListRow
        oldName = folder & Me.Cells(i, 2).Value
        newName = folder & Me.Cells(i, 3).Value
        
        If modUtils.StringHasContent(Me.Cells(i, 3).Value) Then
            Name oldName As newName
        End If
    Next i


End Sub


