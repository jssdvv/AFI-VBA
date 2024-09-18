VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Default global values

Public Property Get logoAspectRatio() As Single: logoAspectRatio = 39 / 22: End Property

' Default global size dimensions in points

Public Property Get defaultNormalStrokeThickness() As Single: defaultNormalStrokeThickness = 0.5: End Property

' Default global size dimensions in inches

Public Property Get defaultFormatMarginSize() As Single: defaultFormatMarginSize = 0.05: End Property
Public Property Get defaultOffsetCellSize() As Single: defaultOffsetCellSize = 1: End Property
Public Property Get defaultFormatHeaderRowHeightSize() As Single: defaultFormatHeaderRowHeightSize = 1: End Property
Public Property Get defaultFormatTitleRowHeightSize() As Single: defaultFormatTitleRowHeightSize = 1 / 2: End Property
Public Property Get defaultFormatInfoRowHeightSize() As Single: defaultFormatInfoRowHeightSize = 1 / 3: End Property
Public Property Get defaultListHeaderRowHeightSize() As Single: defaultListHeaderRowHeightSize = 1 / 2: End Property
