' Удаляет объекты из документа word, полезно при чистке файла.

Sub RemoveObjects()
    Dim obj As Object
    
    For Each obj In ActiveDocument.Content.InlineShapes
        If obj.Type = wdInlineShapePicture Or _
           obj.Type = wdInlineShapeChart Then
            obj.Delete
        End If
    Next obj
    
    For Each obj In ActiveDocument.InlineShapes
        If obj.Type = wdInlineShapeLinkedPicture Or _
           obj.Type = wdInlineShapeLinkedChart Then
            obj.Delete
        End If
    Next obj
    
    For Each obj In ActiveDocument.Shapes
        If obj.Type = msoPicture Or _
           obj.Type = msoChart Then
            obj.Delete
        End If
    Next obj
End Sub
