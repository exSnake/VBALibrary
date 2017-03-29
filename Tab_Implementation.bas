Attribute VB_Name = "insertName"
Option Explicit

Private Const name As String = "tableName" //Insert the name of the table, (you choose it when you create the table if you click on the table property)
Private Const col As Integer = 19
Private lobj As ListObject

Public Sub DoubleClick(Target As Range, cancel As Boolean)
    '' TODO
End Sub

Public Function GetLobj() As ListObject
    On Error Resume Next
    If lobj Is Nothing Then Set lobj = Tab_Controller.GetLobj(name)
    Set GetLobj = lobj
    On Error GoTo 0
End Function

Public Function GetCell(riga As Range, col As String) As Range
    Set GetCell = Tab_Controller.GetCell(GetLobj, riga, col)
End Function

Public Function GetCellByRow(riga As Long, col As String) As Range
    Set GetCellByRow = Tab_Controller.GetCellByRow(GetLobj, riga, col)
End Function

Public Function GetColRng(col As String) As Range
    Set GetColRng = Tab_Controller.GetColumnData(GetLobj, col)
End Function

Public Sub Reset()
    Tab_Controller.Reset GetLobj
End Sub

Public Sub Resize(Row As Long)
    Tab_Controller.Resize GetLobj, Row, col
End Sub

Public Sub SortByColHeader(col As String)
    Tab_Controller.Sort GetLobj, col
End Sub

Public Sub Update()

End Sub
