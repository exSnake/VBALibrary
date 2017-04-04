Attribute VB_Name = "Tab_Controller"
Option Explicit

'Tab Controller che serve ad implementare tutte le funzioni non presenti
'nella listobject. Possiamo definirla come un interfaccia per le tabelle

'Ritorna la ListObject di una tabella presenta all'interno del foglio, accetta in ingresso
'il nome della ListObject
Public Function GetLobj(name As String) As ListObject
    Set GetLobj = Range(name).ListObject
End Function

'Ritorna una singola cella che e' l'intersezione di una data intestazione di colonna e una riga
Public Function GetCell(lobj As ListObject, riga As Range, col As String) As Range
    Set GetCell = Intersect(riga.EntireRow, lobj.ListColumns(col).DataBodyRange)
End Function

'Come sopra a differenza che invece della riga come range, accetta in ingresso il numero di riga della listobject (non assoluto)
Public Function GetCellByRow(lobj As ListObject, riga As Long, col As String) As Range
    Set GetCellByRow = Intersect(lobj.ListRows(riga).Range.EntireRow, lobj.ListColumns(col).DataBodyRange)
End Function

'Prende in ingresso la tabella e il nome di intestazione di una colonna e ne ritorna il range dei dati
Public Function GetColumnData(lobj As ListObject, colName As String) as Range
    If myTbl.ListRows.Count = 0 Then
        Set GetColumnData = GetHeaderRange(myTbl).Find(col).offset(1)
    Else
        Set GetColumnData = myTbl.ListColumns(col).DataBodyRange
    End If
End Function
                
'Ritorna il numero della prima riga vuota della tabella, relativa alla tabella stessa
Public Function GetFirstEmptyRow(myTbl As ListObject) As Long
    
    If myTbl.ListRows.Count = 0 Then
        GetFirstEmptyRow = 1
    Else
        Dim Row As Integer
        Dim i As Integer
        Row = myTbl.ListRows.Count
        For i = 1 To myTbl.ListColumns.Count
            If LTrim(myTbl.DataBodyRange(Row, i)) <> vbNullString Then
                GetFirstEmptyRow = Row + 1
                Exit Function
            End If
        Next i
        GetFirstEmptyRow = Row
    End If
    
    
End Function
                
                
'Ritorna  il range della prima riga vuota della tabella
Public Function GetFirstEmptyRowRange(myTbl As ListObject) As Range
    If GetFirstEmptyRow(myTbl) > myTbl.ListRows.Count Then
        Set GetFirstEmptyRowRange = myTbl.ListRows(myTbl.ListRows.Count).Range.offset(1)
    Else
        Set GetFirstEmptyRowRange = myTbl.ListRows(myTbl.ListRows.Count).Range
    End If
End Function

'Cancella tutti i dati presenti all'interno della tabella e ne crea una riga vuota
Public Sub Reset(lobj As ListObject)
    lobj.DataBodyRange.Delete
    lobj.ListRows.Add
End Sub

'Accetta in ingresso la listobject, il numero di righe e il numero di colonne,
'ridimensiona la tabella eliminando i dati superflui che dopo il ridimensionamento
'finiranno all'esterno della tabella
Public Sub Resize(lobj As ListObject, Row As Long, colNumber As Long)
    Dim rng As Range, val As Variant, rngs As Range
    Set rng = Range(lobj.name & "[#All]").Resize(Row + 1, colNumber)
    Set rngs = Nothing
    For Each val In lobj.ListRows
        If Intersect(val.Range, rng) Is Nothing Then
            If Not rngs Is Nothing Then
                Set rngs = Union(rngs, val.Range)
            Else
                Set rngs = val.Range
            End If
        End If
    Next val
    lobj.Resize rng
    rng.Interior.Pattern = xlNone
    If rngs Is Nothing Then Exit Sub
    With rngs
        .ClearContents
        .Interior.Pattern = xlSolid: .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark2
    End With
End Sub

'Ordina la tabella in ordine ascendente in base alla colonna scelta
Public Sub Sort(lobj As ListObject, col As String, Optional ascending As Boolean)
    Dim field As SortField
    lobj.Sort.SortFields.Clear
    Set field = lobj.Sort.SortFields.Add(Range(lobj.name & "[[#All],[" & col & "]]"))
    With field
        .SortOn = xlSortOnValues
        .order = IIf(ascending, xlAscending, xlDescending)
        .DataOption = xlSortNormal
    End With
    
    lobj.Sort.Apply
End Sub


