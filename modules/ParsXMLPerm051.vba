Option Compare Database
Public Function ParsXMLPerm051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal permNode As Object) As String
    '��������
    '   tblName - �������� �������� ������� �������
    '   tblKeyName - �������� �������������� �������
    '   tblKeyValue - ������������� �������
    '   cadNum - ����������� ����� �������
    '   ������ �� ���� XML
    ' ------------------------
    ' ----- ������������ -----
    ' ------------------------
    Dim pDBFields(3) As String
        pDBFields(0) = "ObjectPermittedUses"
        pDBFields(1) = tblKeyName
        pDBFields(2) = "CadastralNumber"
        pDBFields(3) = "Reserved"
    Dim sqlStr As String
    ' ---------------------------------
    ' ----- ������� � ������ � �� -----
    ' ---------------------------------
    Set permChild = permNode.FirstChild
    Set insertDB = CurrentDb
    While (Not permChild Is Nothing)
        '��� ������ ������. ������� ������ ������ ����� ����� � �� � ��������� � ��������� �������
        sqlStr = "insert into " & tblName & "(" & pDBFields(0) & "," & pDBFields(1) & "," & pDBFields(2)
        sqlStr = sqlStr & ") values ("
        sqlStr = sqlStr & "'" & permChild.Text & "'," & tblKeyValue & ",'" & cadNum & "');"
        insertDB.Execute sqlStr
        Set permChild = permChild.NextSibling
    Wend
    Set insertDB = Nothing
    ParsXMLPerm051 = "+"
End Function