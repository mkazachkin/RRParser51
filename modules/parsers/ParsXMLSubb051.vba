Option Compare Database
Public Function ParsXMLSubb051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal subbNode As Object) As String
    '��������
    '   tblName - �������� �������� ������� �������
    '   tblKeyName - �������� �������������� �������
    '   tblKeyValue - ������������� �������
    '   cadNum - ����������� ����� �������
    '   ������ �� ���� XML Documents
    ' ------------------------
    ' ----- ������������ -----
    ' ------------------------
    '�������� ����
    Dim subbXMLTags(6) As String
        subbXMLTags = GetSubbConfig051(true)
    '�������� ���� ��
    Dim subbDBFields(6) As String
        subbXMLTags = GetSubbConfig051(false)
        subbDBFields(4) = tblKeyName
    Dim subbDBValues(6) As String
    '�������� ���� ������
    Dim subbDBTypes(6) As Boolean
        subbDBTypes = GetSubbTypes051()
    '���������
    Dim i As Integer
    Dim subb_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- ������� -----
    ' -------------------
    '��� �������������� ���� �������� �������
    subbDBValues(4) = tblKeyValue
    subbDBValues(5) = cadNum
    Set builChild = subbNode.FirstChild
    Set insertDB = CurrentDb
    While (Not builChild Is Nothing)
        '������������� � ������� id ������� ������
        subb_id = ReserveID(tblName, "subb_id")
        subbDBValues(6) = "null"
        If builChild.getAttribute("NumberRecord") <> nill Then subbDBValues(0) = builChild.getAttribute("NumberRecord")
        If builChild.getAttribute("DateCreated") <> nill Then subbDBValues(1) = builChild.getAttribute("DateCreated")
        '������
        Set subbChild = builChild.FirstChild
        While (Not subbChild Is Nothing)
            '������ ��������
            If (subbChild.NodeName = subbXMLTags(2)) Then subbDBValues(2) = Replace(subbChild.Text, ".", ",")
            '������ ���� ���
            If (subbChild.NodeName = subbXMLTags(3)) Then subbDBValues(3) = ParsXMLEnbr051(tblName & "_enbr", "subb_id", subb_id, cadNum, subbChild)
            Set subbChild = subbChild.NextSibling
        Wend
        ' -----------------------
        ' ----- ������ � �� -----
        ' -----------------------
        '������������ ������ � ������
        For i = 0 To 5
            If subbDBTypes(i) Then subbDBValues(i) = "{$}" & subbDBValues(i) & "{$}"
        Next i
        '��������� �������
        For i = 0 To 4
            subbDBValues(i) = subbDBValues(i) & ","
        Next i
        '������� ������ �� ���������� ������
        sqlStr = "update " & tblName & " set "
        For i = 0 To 5
            sqlStr = sqlStr & subbDBFields(i) & "=" & subbDBValues(i)
        Next i
        sqlStr = sqlStr & " where subb_id = " & subb_id & ";"
        sqlStr = PrepareInsertSQL(sqlStr)
        insertDB.Execute sqlStr
        Set builChild = builChild.NextSibling
    Wend
    Set insertDB = Nothing
    ParsXMLSubb051 = "+"
End Function
