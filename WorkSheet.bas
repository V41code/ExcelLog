
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
Dim lRow, lCol, lCol0, lCellCnt, lr, lWHId As Long, sColName As String
Dim lChngRowCnt, lChngColCnt As Long ' neieuei y?aae eciaiaii
Dim fNeedRefresh As Boolean
Dim msg As String
Dim c  As Variant, chCell As Range
Dim sv, CellVal As Variant
' caiiiiei ia?aua cia?aiey, aaeuoa eo aoaai ia?aeniieuciaaou
Dim sVal As String
    If fModCell Then Exit Sub
    If fRefreshing Then
        'Debug.Print "In refreshing"
        Exit Sub
    End If
    If fInForm Then Exit Sub
Debug.Print "WorkSheet_Change begin", Time()
    'Call CheckGlobals
    lCol0 = 0: fNeedRefresh = False: msg = "": lCellCnt = 0
    With Target
    Debug.Print "Areas.Count=", .Areas.Count, .Areas(1).Columns.Count, .Areas(1).Rows.Count
    For Each c In Target
        With c
            lCol = .Column
            lRow = .Row
            CellVal = Cells(lRow, lCol)
            If IsError(CellVal) Then
                sVal = "IsError"
            ElseIf IsEmpty(CellVal) Then
                sVal = "IsError"
            ElseIf IsNull(CellVal) Then
                sVal = "IsNull"
            Else
                sVal = CellVal
            End If
        End With
        glCommand.Parameters("@RowIndex") = lRow
        glCommand.Parameters("@ColIndex") = lCol
        'glCommand.Parameters("@SheetId") = glSheet1Id
        glCommand.Parameters("@NewValue") = sVal
        glCommand.Execute
        Debug.Print lRow, lCol, "Command result=", glCommand.Parameters("@res").Value
    Next c
    End With

exit_sub:
Debug.Print "WorkSheet_Change end", Time()
End Sub
