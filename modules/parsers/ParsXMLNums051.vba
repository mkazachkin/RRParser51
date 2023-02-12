Option Compare Database
Public Function ParsXMLNums051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal numsNode As Object) As String
    '��������
    '   tblName - �������� �������� ������� �������
    '   tblKeyName - �������� �������������� �������
    '   tblKeyValue - ������������� �������
    '   cadNum - ����������� ����� �������
    '   ������ �� ���� XML
    ' ------------------------
    ' ----- ������������ -----
    ' ------------------------
    Dim nDBFields(3) As String
        nDBFields(0) = "ChildCadastralNumber"
        nDBFields(1) = tblKeyName
        nDBFields(2) = "CadastralNumber"
        nDBFields(3) = "Reserved"
    Dim sqlStr As String
    ' ---------------------------------
    ' ----- ������� � ������ � �� -----
    ' ---------------------------------
    Set numsChild = numsNode.FirstChild
    Set insertDB = CurrentDb
    While (Not numsChild Is Nothing)
        '��� ������ ������. ������� ������ ������ ����� ����� � �� � ��������� � ��������� ����������
        sqlStr = "insert into " & tblName & "(" & nDBFields(0) & "," & nDBFields(1) & "," & nDBFields(2)
        sqlStr = sqlStr & ") values ("
        sqlStr = sqlStr & "'" & numsChild.Text & "'," & tblKeyValue & ",'" & cadNum & "');"
        insertDB.Execute sqlStr
        Set numsChild = numsChild.NextSibling
    Wend
    Set insertDB = Nothing
    ParsXMLNums051 = "+"
End Function