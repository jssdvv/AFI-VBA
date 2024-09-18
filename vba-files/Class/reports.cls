VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Local format dimensions sizes in inches for re-formatting the sheet
' spaces, dimensions and scale. Some of the Getters are located inside the
' global file "ThisWorkbook" to use them globaly around more objects.

Private Property Get offsetCellSideSize() As Single: offsetCellSideSize = ThisWorkbook.defaultOffsetCellSize: End Property

Private Property Get columnWidthB() As Single: columnWidthB = ThisWorkbook.defaultFormatMarginSize: End Property
Private Property Get columnWidthC() As Single: columnWidthC = 0.8: End Property
Private Property Get columnWidthD() As Single: columnWidthD = 1.8: End Property
Private Property Get columnWidthE() As Single: columnWidthE = 1.2: End Property
Private Property Get columnWidthF() As Single: columnWidthF = 1: End Property
Private Property Get columnWidthG() As Single: columnWidthG = 1.2: End Property
Private Property Get columnWidthH() As Single: columnWidthH = 1.2: End Property
Private Property Get columnWidthI() As Single: columnWidthI = 1.2: End Property
Private Property Get columnWidthJ() As Single: columnWidthJ = ThisWorkbook.defaultFormatMarginSize: End Property

Private Property Get rowHeight1() As Single: rowHeight1 = ThisWorkbook.defaultFormatMarginSize: End Property
Private Property Get rowHeight2() As Single: rowHeight2 = ThisWorkbook.defaultFormatHeaderRowHeightSize: End Property
Private Property Get rowHeight3() As Single: rowHeight3 = ThisWorkbook.defaultFormatMarginSize: End Property
Private Property Get rowHeight4() As Single: rowHeight4 = ThisWorkbook.defaultFormatTitleRowHeightSize - ThisWorkbook.defaultFormatMarginSize: End Property
Private Property Get rowHeight5() As Single: rowHeight5 = ThisWorkbook.defaultFormatInfoRowHeightSize: End Property
Private Property Get rowHeight6() As Single: rowHeight6 = ThisWorkbook.defaultFormatInfoRowHeightSize: End Property
Private Property Get rowHeight7() As Single: rowHeight7 = ThisWorkbook.defaultFormatTitleRowHeightSize: End Property
Private Property Get rowHeight8() As Single: rowHeight8 = ThisWorkbook.defaultListHeaderRowHeightSize: End Property
Private Property Get contentListRowHeight() As Single: rowHeight1 = 1.2: End Property

Sub FixFormatDimensions()
    
    Dim offsetCell As Range: Set offsetCell = Me.Cells(1, 1)
    
    Call modUtils.SetCellSizeInInches(offsetCell, offsetCellSideSize, offsetCellSideSize): Set offsetCell = Nothing
    
    Dim columnB As Range: Set columnB = Me.Columns(2)
    Dim columnC As Range: Set columnC = Me.Columns(3)
    Dim columnD As Range: Set columnD = Me.Columns(4)
    Dim columnE As Range: Set columnE = Me.Columns(5)
    Dim columnF As Range: Set columnF = Me.Columns(6)
    Dim columnG As Range: Set columnG = Me.Columns(7)
    Dim columnH As Range: Set columnH = Me.Columns(8)
    Dim columnI As Range: Set columnI = Me.Columns(9)
    Dim columnJ As Range: Set columnJ = Me.Columns(10)
    
    Call modUtils.SetColumnWidthInInches(columnB, columnWidthB): Set columnB = Nothing
    Call modUtils.SetColumnWidthInInches(columnC, columnWidthC): Set columnC = Nothing
    Call modUtils.SetColumnWidthInInches(columnD, columnWidthD): Set columnD = Nothing
    Call modUtils.SetColumnWidthInInches(columnE, columnWidthE): Set columnE = Nothing
    Call modUtils.SetColumnWidthInInches(columnF, columnWidthF): Set columnF = Nothing
    Call modUtils.SetColumnWidthInInches(columnG, columnWidthG): Set columnG = Nothing
    Call modUtils.SetColumnWidthInInches(columnH, columnWidthH): Set columnH = Nothing
    Call modUtils.SetColumnWidthInInches(columnI, columnWidthI): Set columnI = Nothing
    Call modUtils.SetColumnWidthInInches(columnJ, columnWidthJ): Set columnJ = Nothing
    
    Dim row1 As Range: Set row1 = Me.Rows(2)
    Dim row2 As Range: Set row2 = Me.Rows(3)
    Dim row3 As Range: Set row3 = Me.Rows(4)
    Dim row4 As Range: Set row4 = Me.Rows(5)
    Dim row5 As Range: Set row5 = Me.Rows(6)
    Dim row6 As Range: Set row6 = Me.Rows(7)
    Dim row7 As Range: Set row7 = Me.Rows(8)
    Dim row8 As Range: Set row8 = Me.Rows(9)
    
    Call modUtils.SetRowHeightInInches(row1, rowHeight1): Set row1 = Nothing
    Call modUtils.SetRowHeightInInches(row2, rowHeight2): Set row2 = Nothing
    Call modUtils.SetRowHeightInInches(row3, rowHeight3): Set row3 = Nothing
    Call modUtils.SetRowHeightInInches(row4, rowHeight4): Set row4 = Nothing
    Call modUtils.SetRowHeightInInches(row5, rowHeight5): Set row5 = Nothing
    Call modUtils.SetRowHeightInInches(row6, rowHeight6): Set row6 = Nothing
    Call modUtils.SetRowHeightInInches(row7, rowHeight7): Set row7 = Nothing
    Call modUtils.SetRowHeightInInches(row8, rowHeight8): Set row8 = Nothing
        
    Dim logo As shape: Set logo = Me.Shapes("logo")
    Dim rect1 As shape: Set rect1 = Me.Shapes("rect1")
    Dim rect2 As shape: Set rect2 = Me.Shapes("rect2")
    Dim rect3 As shape: Set rect3 = Me.Shapes("rect3")
    Dim rect4 As shape: Set rect4 = Me.Shapes("rect4")
    Dim rect5 As shape: Set rect5 = Me.Shapes("rect5")
    Dim rect6 As shape: Set rect6 = Me.Shapes("rect6")
    Dim header As shape
    Dim headerRange As Range: Set headerRange = Me.Range(Cells(3, 3), Cells(3, 9))
    
    With logo
        .height = headerRange.height
        .width = ThisWorkbook.logoAspectRatio * headerRange.height
        .Top = headerRange.Top
        .Left = headerRange.Left
    End With
    
    With rect1
        .height = logo.height
        .width = logo.width
        .Top = logo.Top
        .Left = logo.Left
    End With
        
    With rect2
        .height = headerRange.height / 2
        .width = headerRange.width - 2 * rect1.width
        .Top = headerRange.Top
        .Left = rect1.Left + rect1.width
    End With
        
    With rect3
        .height = headerRange.height / 2
        .width = headerRange.width - 2 * rect1.width
        .Top = rect2.Top + rect2.height
        .Left = rect1.Left + rect1.width
    End With
        
    With rect4
        .height = headerRange.height / 3
        .width = headerRange.width - rect1.width - rect2.width
        .Top = headerRange.Top
        .Left = rect2.Left + rect2.width
    End With
        
    With rect5
        .height = headerRange.height / 3
        .width = headerRange.width - rect1.width - rect2.width
        .Top = rect4.Top + rect4.height
        .Left = rect2.Left + rect2.width
    End With
        
    With rect6
        .height = headerRange.height - rect4.height - rect5.height
        .width = headerRange.width - rect1.width - rect2.width
        .Top = rect5.Top + rect5.height
        .Left = rect2.Left + rect2.width
    End With
    
    Set logo = Nothing
    Set rect1 = Nothing
    Set rect2 = Nothing
    Set rect3 = Nothing
    Set rect4 = Nothing
    Set rect5 = Nothing
    Set rect6 = Nothing
    Set header = Nothing
    Set headerRange = Nothing

End Sub
Sub UploadPictures()
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim folder As Variant: folder = modUtils.SelectFolderOrNull(ThisWorkbook.CustomDocumentProperties("previousFolder").Value)
    Dim imageFormats As Variant: imageFormats = Array("tiff", "jpeg", "jpg", "png", "gif")
    Dim totalRows As Long: totalRows = Me.Cells(Cells(Rows.Count, 3).End(xlUp).row, 3).row
    Dim row As Long
    Dim invNum As String
    Dim currentFormat As Variant
    Dim invNumRoute As String
    Dim i As Long
    Dim inv As String
    
    ' Handles the folder selection to pick up where are the pictures storaged
    ' saving the previous folder route into a custom property within the file and
    ' if the folder is NULL it shows a message box, then it cancels the script.
    
    If IsNull(folder) Then
        MsgBox modValues.folderNotSelected
        Exit Sub
    Else
        ThisWorkbook.CustomDocumentProperties("previousFolder").Value = folder
    End If
    
    For row = 10 To totalRows Step 1
        If Me.Cells(row, 3).Value = "" Then
            invNum = ""
            currentFormat = Nothing
        Else
        inv = Me.Cells(row, 3).Value
            invNum = Dir(folder & Me.Cells(row, 3).Value & "*.*")
            currentFormat = VBA.Split(invNum, ".")
            
            If modUtils.StringHasContent(invNum) Then
                If modUtils.IsStringInArray(currentFormat(UBound(currentFormat)), imageFormats) Then
                    invNumRoute = folder & invNum
                Else
                    invNumRoute = ""
                End If
            
            
                If Dir(invNumRoute) = "" Then
                Else
                    With Me.Shapes.AddPicture( _
                        Filename:=invNumRoute, _
                        linktofile:=msoFalse, _
                        savewithdocument:=msoTrue, _
                        Left:=Me.Cells(row, 8).Left + 1, _
                        Top:=Me.Cells(row, 8).Top + 1, _
                        width:=Me.Cells(row, 8).width - 1, _
                        height:=Me.Cells(row, 8).height - 1 _
                    )
                        .Rotation = 0
                        .LockAspectRatio = 0
                        .name = invNum
                    End With
                End If
            
            End If
        End If
    Next row

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub















