
Sub dictEx()

Dim sDict As New Scripting.Dictionary
Dim key As Variant

With ActiveSheet
'create dictionary with key and associated 2 item array
    For i = 1 To 3
        key = .Cells(i, 1).Value2
        If sDict.Exists(key) Then
            sDict(key)(0) = sDict(key)(0) + .Cells(i, 2).Value2
            sDict(key)(1) = sDict(key)(1) + .Cells(i, 2).Value2
        Else
            sDict.Add key, Array(.Cells(i, 2).Value2, .Cells(i, 3).Value2)
        End If
    Next i
'display dictionary
    i = 5
    For Each key In sDict.Keys
        .Cells(i, 1).Value2 = key
        .Cells(i, 2).Value2 = sDict(key)(0)
        .Cells(i, 3).Value2 = sDict(key)(1)
        Debug.Print key, sDict(key)(0), sDict(key)(1)
        
        i = i + 1
    Next key
End With
Set sDict = Nothing

End Sub
