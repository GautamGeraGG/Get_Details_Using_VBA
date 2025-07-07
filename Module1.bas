Attribute VB_Name = "Module1"
Sub match()
Dim a As Object
Set a = CreateObject("scripting.dictionary")
Dim xl As Long
xl = ThisWorkbook.Worksheets("sheet1").Range("a1048576").End(xlUp).Row
Set am = ThisWorkbook.Worksheets("sheet1")

For b = 2 To xl
If am.Cells(b, 1).Value <> "" Then
    x = LCase(am.Cells(b, 1).Value)
    If Not a.exists(b) Then
        a.Add x, b
    End If
    End If
Next b

Dim qdict As Object
Set qdict = CreateObject("scripting.dictionary")
Dim rm As Range

For x = 0 To a.Count - 1
Key = a.keys()(x)
vl = a.items()(x)
Set rm = am.Range(am.Cells(vl + 1, 4), am.Cells(vl + 3, 6))

If Not qdict.exists(Key) Then
    qdict.Add Key, rm
    End If
Next x
a.RemoveAll

If am.Cells(2, 12).Value <> "" Then
    Dim found As Boolean
    found = False
    
    For x = 0 To qdict.Count - 1
        If LCase(am.Cells(2, 12).Value) = qdict.keys()(x) Then
            Set aw = qdict.items()(x)
            aw.Copy
            am.Range(am.Cells(8, 11), am.Cells(11, 13)).PasteSpecial xlPasteValues
            am.Cells(6, 11).Value = UCase(am.Cells(2, 12).Value)
            found = True
            Exit For
        End If
    Next x
    
    If Not found Then
        MsgBox "The name you entered is not found"
        am.Range(am.Cells(8, 11), am.Cells(11, 13)) = none
    End If
End If

If am.Cells(2, 12).Value = "" Then
    MsgBox "Please Enter The Name First"
End If



am.Cells(2, 12).Value = none
Application.CutCopyMode = False
End Sub

