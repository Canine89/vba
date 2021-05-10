Sub CaptionInlinePictures()
    Dim pic As InlineShape
    Dim vData() As Byte
    Dim i As Long
    Dim lWritePos As Long
    Dim strOutFileName As String
    
    For Each pic In ActiveDocument.InlineShapes
        pic.Select
        Selection.InsertCaption "Figure", "", "", wdCaptionPositionBelow, 0
        'strOutFileName = ".\_img\" & "Figure" & CStr(i) & ".png"
        'Open strOutFileName For Binary Access Write As #1
        'i = i + 1
        'vData = pic.Range.EnhMetaFileBits
        'lWritePos = 1

        'Put #1, lWritePos, vData

        'Close #1
    Next
End Sub

Sub CaptionFloatingPictures()
    Dim pic As Shape
    Dim vData() As Byte
    Dim i As Long
    Dim lWritePos As Long
    Dim strOutFileName As String
    
    For Each pic In ActiveDocument.Shapes
        If pic.Type = msoPicture Then
            pic.Select
            Selection.InsertCaption "Figure", "", "", wdCaptionPositionBelow, 0
            strOutFileName = ".\_img\" & "Figure" & CStr(i) & ".png"
            Open strOutFileName For Binary Access Write As #1
            i = i + 1
            vData = pic.Range.EnhMetaFileBits
            lWritePos = 1

            Put #1, lWritePos, vData

            Close #1
        End If
    Next
End Sub
