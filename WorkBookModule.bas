Option Explicit


 
Dim glConnection As ADODB.Connection
Dim glConnString As String
Dim glCommand As ADODB.Command
Dim glUserName As String
Dim glFileName As String
Dim glSheetList As String
Dim glFileId As Long
Dim glSheet1Id As Long

Sub WorkBook_Open()
Dim idx As Long
    glUserName = Environ("UserName")
    glFileName = Application.ActiveWorkbook.FullName
    For idx = 1 To Application.ActiveWorkbook.Worksheets.Count
        glSheetList = glSheetList + Application.ActiveWorkbook.Worksheets.Item(idx).Name + ","
    Next
    glSheetList = Left(glSheetList, Len(glSheetList) - 1)
    If glUserName = "EEE" Then
        glConnString = "ODBC;SERVER=LLL\SQLEXPRESS;DATABASE=SPData;DSN=SPDataSQLExpress;Trusted_Connection=Yes;App=ProtExl"
    Else
        glConnString = "ODBC;SERVER=MT-SV-TS4\SQLEXPRESS;DATABASE=SPData;DSN=SPData_User;App=ProtExl"
    End If
    Set glConnection = New ADODB.Connection
    glConnection.ConnectionString = glConnString
    glConnection.Open
    Set glCommand = New ADODB.Command
    glCommand.CommandText = "uspExlLogGetFileSheetIds"
    glCommand.CommandType = adCmdStoredProc
    glCommand.ActiveConnection = glConnection
    Dim Param1 As ADODB.Parameter
    Set Param1 = New ADODB.Parameter
    Set Param1 = glCommand.CreateParameter("@FileName", adVarChar, adParamInput, LenB(glFileName), glFileName)
    glCommand.Parameters.Append Param1
    Dim Param2 As ADODB.Parameter
    Set Param2 = glCommand.CreateParameter("@SheetNameList", adVarChar, adParamInput, LenB(glSheetList), glSheetList)
    glCommand.Parameters.Append Param2
    Dim rs As ADODB.Recordset
    Set rs = glCommand.Execute
    With rs
     Do While Not .EOF()
        glFileId = rs.Fields("FK_ExlLogFileId")
        If Trim(rs.Fields("SheetName")) = "Eeno1" Then
            glSheet1Id = rs.Fields("SheetId")
        End If
        rs.MoveNext
     Loop
    End With
    rs.Close
    glCommand.Parameters.Delete (1)
    glCommand.Parameters.Delete (0)
    ' iiaaioiaea eiiaiau e eiae?iaaie?
    Dim ParamRes As ADODB.Parameter
    Set ParamRes = glCommand.CreateParameter("@res", adBigInt, adParamOutput)
    glCommand.Parameters.Append ParamRes
    Set Param1 = glCommand.CreateParameter("@SheetId", adBigInt, adParamInput)
    glCommand.Parameters.Append Param1
    Set Param2 = glCommand.CreateParameter("@RowIndex", adBigInt, adParamInput)
    glCommand.Parameters.Append Param2
    Dim Param3 As ADODB.Parameter
    Set Param3 = glCommand.CreateParameter("@ColIndex", adBigInt, adParamInput)
    glCommand.Parameters.Append Param3
    Dim Param4 As ADODB.Parameter
    Set Param4 = glCommand.CreateParameter("@NewValue", adVarChar, adParamInput, 512)
    glCommand.Parameters.Append Param4
    ' set parameters
    glCommand.CommandText = "uspExlLogDataInsert"
    glCommand.Parameters("@RowIndex") = 0
    glCommand.Parameters("@ColIndex") = 0
    glCommand.Parameters("@SheetId") = glSheet1Id
    glCommand.Parameters("@NewValue") = "Test text for log file " + glFileName
    glCommand.Execute
    Debug.Print "Command result=", glCommand.Parameters("@res").Value
End Sub
