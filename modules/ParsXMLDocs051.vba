Option Compare Database
Public Function ParsXMLDocs051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal docsNode As Object) As String
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
    Dim docsXMLTags(9) As String
        docsXMLTags = GetDocsConfig051(true)
    '�������� ���� ��
    Dim docsDBFields(9) As String
        docsDBFields =  GetDocsConfig051(false)
        docsDBFields(7) = tblKeyName
    Dim docsDBValues(9) As String
    '�������� ���� ������
    Dim docsDBTypes(9) As Boolean
        docsDBTypes = GetDocsTypes051()
    '���������
    Dim i As Integer
    Dim docs_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- ������� -----
    ' -------------------
    '��� �������������� ���� �������� �������
    docsDBValues(7) = tblKeyValue
    docsDBValues(8) = cadNum
    '������������� � ������� id ������� ������
    docs_id = ReserveID(tblName, "docs_id")
    docsDBValues(9) = "null"
    '������
    Set docsChild = docsNode.FirstChild
    While (Not docsChild Is Nothing)
        '������ ��������
        For i = 0 To 6
            If (docsChild.NodeName = docsXMLTags(i)) Then docsDBValues(i) = docsChild.Text
        Next i
        '����� ���, �� ������� �� ����
        Set docsChild = docsChild.NextSibling
    Wend
    ' -----------------------
    ' ----- ������ � �� -----
    ' -----------------------
    '������������ ������ � ������
    For i = 0 To 8
        If docsDBTypes(i) Then docsDBValues(i) = "{$}" & docsDBValues(i) & "{$}"
    Next i
    '��������� �������
    For i = 0 To 7
        docsDBValues(i) = docsDBValues(i) & ","
    Next i
    '������� ������ �� ���������� ������
    sqlStr = "update " & tblName & " set "
    For i = 0 To 8
        sqlStr = sqlStr & docsDBFields(i) & "=" & docsDBValues(i)
    Next i
    sqlStr = sqlStr & " where docs_id = " & docs_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    parsXMLDocuments051 = "+"
End Function