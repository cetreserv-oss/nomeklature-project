Option Explicit

Public Sub FormatZakaz()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim colOrder As Range
    Dim lastRow As Long
    Dim i As Long
    Dim val As String

    Set ws = GetFsmRequestSheet
    Set colOrder = ws.Rows(1).Find(What:="Заказ", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If colOrder Is Nothing Then
        Err.Raise vbObjectError + 701, , "Столбец 'Заказ' не найден на листе '" & ws.Name & "'."
    End If

    lastRow = ws.Cells(ws.Rows.Count, colOrder.Column).End(xlUp).Row

    For i = 2 To lastRow
        val = Trim(CStr(ws.Cells(i, colOrder.Column).Value))
        If Len(val) > 0 Then
            val = ReplaceCyrillicWithLatin(val)
            ws.Cells(i, colOrder.Column).Value = UCase$(val)
        End If
    Next i

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "FormatZakaz", Err.Description
End Sub

Private Function ReplaceCyrillicWithLatin(ByVal txt As String) As String
    Dim mapCyr As Object
    Dim ch As String
    Dim res As String
    Dim j As Long

    Set mapCyr = CreateObject("Scripting.Dictionary")

    mapCyr.Add "А", "A": mapCyr.Add "В", "B": mapCyr.Add "Е", "E": mapCyr.Add "К", "K": mapCyr.Add "М", "M"
    mapCyr.Add "Н", "H": mapCyr.Add "О", "O": mapCyr.Add "Р", "P": mapCyr.Add "С", "C": mapCyr.Add "Т", "T"
    mapCyr.Add "У", "Y": mapCyr.Add "Х", "X": mapCyr.Add "а", "A": mapCyr.Add "в", "B": mapCyr.Add "е", "E"
    mapCyr.Add "к", "K": mapCyr.Add "м", "M": mapCyr.Add "н", "H": mapCyr.Add "о", "O": mapCyr.Add "р", "P"
    mapCyr.Add "с", "C": mapCyr.Add "т", "T": mapCyr.Add "у", "Y": mapCyr.Add "х", "X"

    For j = 1 To Len(txt)
        ch = Mid$(txt, j, 1)
        If mapCyr.Exists(ch) Then
            res = res & mapCyr(ch)
        Else
            res = res & ch
        End If
    Next j

    ReplaceCyrillicWithLatin = res
End Function
