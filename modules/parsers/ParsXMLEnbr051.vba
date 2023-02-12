Option Compare Database
Public Function ParsXMLEnbr051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal enbrNode As Object) As String
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
    Dim enbrXMLTags(7) As String
        enbrXMLTags = GetEnbrConfig051(true)
    '�������� ���� ��
    Dim enbrDBFields(7) As String
        enbrDBFields = GetEnbrConfig051(false)
        enbrDBFields(5) = tblKeyName
    Dim enbrDBValues(7) As String
    '�������� ���� ������
    Dim enbrDBTypes(7) As Boolean
    '���������
    Dim i As Integer
    Dim enbr_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- ������� -----
    ' -------------------
    '��� �������������� ���� �������� �������
    enbrDBValues(5) = tblKeyValue
    enbrDBValues(6) = cadNum
    Set enbrNode = enbrNode.FirstChild
    While (Not enbrNode Is Nothing)
        '������������� � ������� id ������� ������
        enbr_id = ReserveID(tblName, "enbr_id")
        enbrDBValues(7) = "null"
        '������
        Set enbrChild = enbrNode.FirstChild
        While (Not enbrChild Is Nothing)
            '������ ��������
            If (enbrChild.NodeName = enbrXMLTags(0)) Then enbrDBValues(0) = enbrChild.Text
            If (enbrChild.NodeName = enbrXMLTags(1)) Then enbrDBValues(1) = enbrChild.Text
            If (enbrChild.NodeName = enbrXMLTags(2)) Then
                Set subb = enbrChild.FirstChild
                While (Not subb Is Nothing)
                    If (subb.NodeName = "RightNumber") Then enbrDBValues(2) = subb.Text
                    If (subb.NodeName = "RegistrationDate") Then enbrDBValues(3) = subb.Text
                    Set subb = subb.NextSibling
                Wend
            End If
            '������ ���� ���
            If (enbrChild.NodeName = enbrXMLTags(4)) Then enbrDBValues(4) = ParsXMLDocs051(tblName & "_docs", "enbr_id", enbr_id, cadNum, enbrChild)
            Set enbrChild = enbrChild.NextSibling
        Wend
        '������������ ������ � ������
        For i = 0 To 6
            If enbrDBTypes(i) Then enbrDBValues(i) = "{$}" & enbrDBValues(i) & "{$}"
        Next i
        '��������� �������
        For i = 0 To 5
            enbrDBValues(i) = enbrDBValues(i) & ","
        Next i
        '������� ������ �� ���������� ������
        sqlStr = "update " & tblName & " set "
        For i = 0 To 6
            sqlStr = sqlStr & enbrDBFields(i) & "=" & enbrDBValues(i)
        Next i
        sqlStr = sqlStr & " where enbr_id = " & enbr_id & ";"
        sqlStr = PrepareInsertSQL(sqlStr)
        Set insertDB = CurrentDb
        insertDB.Execute sqlStr
        Set insertDB = Nothing
        Set enbrNode = enbrNode.NextSibling
    Wend
    ParsXMLEnbr051 = "+"
End Function