Option Compare Database
Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(filename)
    oFile.WriteLine txt
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Function
Public Function SHA256(sIn As String) As String
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    SHA256 = ConvToHexString(bytes)
    Set oT = Nothing
    Set oSHA256 = Nothing
End Function
Public Function ConvToHexString(vIn As Variant) As Variant
    Dim oD As Object
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    Set oD = Nothing
End Function
Public Function LoginUserName() As String
  Dim er, lSize As Long
  Dim sBuffer As String
  On Error GoTo Err_LoginUserName
  lSize = 255
  sBuffer = Space$(lSize)
  er = GetUserNameA(sBuffer, lSize)
  If lSize > 0 And lSize < 100 Then
    LoginUserName = Left$(sBuffer, lSize - 1)
    On Error GoTo 0
    Exit Function
  End If
Err_LoginUserName:
  LoginUserName = "user"
End Function
Public Function DictCheck(searchStr As String, searchColumn As String, tableName As String, idName As String) As Long
    Dim searchRS As Recordset
    Set searchDB = CurrentDb
    Dim searchSQL As String
    searchSQL = "select " & idName & " from " & tableName & " where " & searchColumn & " = '" & searchStr & "';"
    Set searchRS = searchDB.OpenRecordset(searchSQL)
    DictCheck = searchRS.Fields.Item(0).Value
    Set searchRS = Nothing
    Set searchDB = Nothing
End Function
Public Function PrepareInsertSQL(sqlString As String) As String
    sqlString = Replace(sqlString, "{$}{$}", "null")
    sqlString = Replace(sqlString, "'", "''")
    sqlString = Replace(sqlString, "{$}", "'")
    PrepareInsertSQL = sqlString
End Function
Public Function ReserveID (tblName As String, keyField As String) As String
    Dim result As String
    strSQL = "insert into " & tblName & "(reserved) values ('+');"
    Set reserveDB = CurrentDb
    reserveDB.Execute strSQL
    strSQL = "select " & keyField & " from " & tblName & " where reserved is not null;"
    Set rs = reserveDB.OpenRecordset (strSQL)
    result = CStr(rs.Fields(keyField).Value)
    strSQL = "update " & tblName & " set reserved = null where reserved is not null;"
    reserveDB.Execute strSQL
    Set rs = Nothing
    Set insertDB = Nothing
    ReserveID = result
End Function