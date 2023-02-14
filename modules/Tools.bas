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
Public Function ReserveID(tblName As String, keyField As String, Optional sha256hash As String) As String
    Dim result As String
    Randomize
    Dim guid As String
    guid = SHA256(CStr(Rnd) + Now + CStr(Rnd))
    If sha256hash = "" Then
        strSQL = "insert into " & tblName & "(reserved) values ('" & guid & "');"
    Else
        strSQL = "insert into " & tblName & "(reserved, sha256hash) values ('" & guid & "','" & sha256hash & "');"
    End If
    Set reserveDB = CurrentDb
    reserveDB.Execute strSQL
    strSQL = "select " & keyField & " from " & tblName & " where reserved ='" & guid & "';"
    Set rs = reserveDB.OpenRecordset(strSQL)
    result = CStr(rs.Fields(keyField).Value)
    strSQL = "update " & tblName & " set reserved = null where reserved ='" & guid & "';"
    reserveDB.Execute strSQL
    Set rs = Nothing
    Set insertDB = Nothing
    ReserveID = result
End Function
'Public Function ReserveID(tblName As String, keyField As String) As String
'    Dim result As String
'    Randomize
'    Dim timestamp As String
'    strSQL = "insert into " & tblName & "(reserved) values ('+');"
'    Set reserveDB = CurrentDb
'    reserveDB.Execute strSQL
'    strSQL = "select " & keyField & " from " & tblName & " where reserved is not null;"
'    Set rs = reserveDB.OpenRecordset(strSQL)
'    result = CStr(rs.Fields(keyField).Value)
'    strSQL = "update " & tblName & " set reserved = null where reserved is not null;"
'    reserveDB.Execute strSQL
'    Set rs = Nothing
'    Set insertDB = Nothing
'    ReserveID = result
'End Function
Public Function ClearObjTable(objType As String, xmlf_id As Long) As Boolean
    'Список таблиц
    Dim tbl(14, 4) As String
    tbl(0, 0) = "public_import_t_buildings_cult_docs"
    tbl(1, 0) = "public_import_t_buildings_cult"
    tbl(2, 0) = "public_import_t_buildings_facl"
    tbl(3, 0) = "public_import_t_buildings_subb_enbr_docs"
    tbl(4, 0) = "public_import_t_buildings_subb_enbr"
    tbl(5, 0) = "public_import_t_buildings_subb"
    tbl(6, 0) = "public_import_t_buildings_cost_docs"
    tbl(7, 0) = "public_import_t_buildings_cost"
    tbl(8, 0) = "public_import_t_buildings_perm"
    tbl(9, 0) = "public_import_t_buildings_unit"
    tbl(10, 0) = "public_import_t_buildings_cars"
    tbl(11, 0) = "public_import_t_buildings_flat"
    tbl(12, 0) = "public_import_t_buildings_prev"
    tbl(13, 0) = "public_import_t_buildings_prnt"
    tbl(14, 0) = "public_import_t_buildings"
    tbl(0, 1) = "public_import_t_flats_cult_docs"
    tbl(1, 1) = "public_import_t_flats_cult"
    tbl(2, 1) = "public_import_t_flats_facl"
    tbl(3, 1) = "public_import_t_flats_subf_enbr_docs"
    tbl(4, 1) = "public_import_t_flats_subf_enbr"
    tbl(5, 1) = "public_import_t_flats_subf"
    tbl(6, 1) = "public_import_t_flats_cost_docs"
    tbl(7, 1) = "public_import_t_flats_cost"
    tbl(8, 1) = "public_import_t_flats_perm"
    tbl(9, 1) = "public_import_t_flats_unit"
    tbl(10, 1) = "public_import_t_flats_pstn"
    tbl(11, 1) = "public_import_t_flats_asgn"
    tbl(12, 1) = "public_import_t_flats_prev"
    tbl(13, 1) = "public_import_t_flats_poks"
    tbl(14, 1) = "public_import_t_flats"
    tbl(0, 2) = "public_import_t_constructions_cult_docs"
    tbl(1, 2) = "public_import_t_constructions_cult"
    tbl(2, 2) = "public_import_t_constructions_facl"
    tbl(3, 2) = "public_import_t_constructions_subc_enbr_docs"
    tbl(4, 2) = "public_import_t_constructions_subc_enbr"
    tbl(5, 2) = "public_import_t_constructions_subc"
    tbl(6, 2) = "public_import_t_constructions_cost_docs"
    tbl(7, 2) = "public_import_t_constructions_cost"
    tbl(8, 2) = "public_import_t_constructions_perm"
    tbl(9, 2) = "public_import_t_constructions_unit"
    tbl(10, 2) = "public_import_t_constructions_flat"
    tbl(11, 2) = "public_import_t_constructions_prev"
    tbl(12, 2) = "public_import_t_constructions_prnt"
    tbl(13, 2) = "public_import_t_constructions_keyp"
    tbl(14, 2) = "public_import_t_constructions"
    tbl(0, 3) = "public_import_t_uncompleted_facl"
    tbl(1, 3) = "public_import_t_uncompleted_cost_docs"
    tbl(2, 3) = "public_import_t_uncompleted_cost"
    tbl(3, 3) = "public_import_t_uncompleted_prev"
    tbl(4, 3) = "public_import_t_uncompleted_prnt"
    tbl(5, 3) = "public_import_t_uncompleted_keyp"
    tbl(6, 3) = ""
    tbl(7, 3) = ""
    tbl(8, 3) = ""
    tbl(9, 3) = ""
    tbl(10, 3) = ""
    tbl(11, 3) = ""
    tbl(12, 3) = ""
    tbl(13, 3) = ""
    tbl(14, 3) = "public_import_t_uncompleted"
    tbl(0, 4) = "public_import_t_cars_facl"
    tbl(1, 4) = "public_import_t_cars_cost_docs"
    tbl(2, 4) = "public_import_t_cars_cost"
    tbl(3, 4) = "public_import_t_cars_unit"
    tbl(4, 4) = "public_import_t_cars_pstn"
    tbl(5, 4) = "public_import_t_cars_prev"
    tbl(6, 4) = "public_import_t_cars_poks"
    tbl(7, 4) = ""
    tbl(8, 4) = ""
    tbl(9, 4) = ""
    tbl(10, 4) = ""
    tbl(11, 4) = ""
    tbl(12, 4) = ""
    tbl(13, 4) = ""
    tbl(14, 4) = "public_import_t_cars"
    
    Dim dRs As DAO.Recordset
    Dim c As Integer
    Dim i As Integer
    Dim sqlStr As String
    Dim flag As Boolean
    flag = False
    
    If objType = "002001002000" Then c = 0
    If objType = "002001003000" Then c = 1
    If objType = "002001004000" Then c = 2
    If objType = "002001005000" Then c = 3
    If objType = "002001009000" Then c = 4
    
    'Чистим таблицы
    sqlStr = "select cadastralnumber from " & tbl(14, c) & " where xmlf_id=" & CStr(xmlf_id) & ";"
    Set delDB = CurrentDb
    Set dRs = delDB.OpenRecordset(sqlStr)
    If dRs.RecordCount > 0 Then
        dRs.MoveFirst
        Do While Not dRs.EOF
            For i = 0 To 14
                If tbl(i, c) <> "" Then
                    sqlStr = "delete from " & tbl(i, c) & " where cadastralnumber='" & dRs.Fields("cadastralnumber").Value & "';"
                    delDB.Execute sqlStr
                End If
            Next i
            dRs.MoveNext
        Loop
        flag = True
    End If
    Set delDB = Nothing
    Set dRs = Nothing
    ClearObjTable = flag
End Function
