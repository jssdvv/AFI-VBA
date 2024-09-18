Attribute VB_Name = "modBarcodeReader"
Public Function decodeBarcode128( _
    ByVal imagePath As String, _
    Optional ByVal maxLength As Integer = 6 _
) As String

    Dim barcodeReader As IBarcodeReader: Set barcodeReader = CreateObject("ZXing.Interop.Decoding.BarcodeReader")
    Dim decodedBarcode As String
    
        barcodeReader.Options.TryHarder = True
        barcodeReader.Options.PureBarcode = True
        barcodeReader.Options.PossibleFormats.Add BarcodeFormat_CODE_128
    
    decodedBarcode = barcodeReader.DecodeImageFile(imagePath).Text
    
    If Len(decodedBarcode) <= maxLength And IsNumeric(decodedBarcode) Then
        decodeBarcode128 = "S" & decodedBarcode
    End If
    
    Set barcodeReader = Nothing
    
End Function
