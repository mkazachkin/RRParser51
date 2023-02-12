Option Compare Database
Public Function ParsXMLCost051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal costNode As Object) As String
    '��������
    '   tblName - ������� ������ XML
    '   tblKeyName - �������� �������������� XML
    '   tblKeyValue - ������������� XML
    '   cadNum - ����������� ����� �������
    '   ������ �� ���� XML CadastralNumber
    ' ------------------------
    ' ----- ������������ -----
    ' ------------------------
    '�������� ����
    Dim cdcsXMLTags(7) As String
        cdcsXMLTags = GetCostConfig051(true)
    '�������� ���� ��
    Dim cdcsDBFields(10) As String
        cdcsDBFields = GetCostConfig051(false)
        cdcsDBFields(8) = tblKeyName
    Dim cdcsDBValues(10) As String
    '�������� ���� ������
    Dim cdcsDBTypes(10) As Boolean
        cdcsDBTypes = GetCostTypes051()
    '���������
    Dim i As Integer
    Dim cdcs_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- ������� -----
    ' -------------------
    '��� �������������� ���� �������� �������
    cdcsDBValues(8) = tblKeyValue
    cdcsDBValues(9) = cadNum
    '������������� � ������� id ������� ������
    cdcs_id = ReserveID(tblName, "cdcs_id")
    cdcsDBValues(10) = "null"
    '����������� ��������� ���� �������� "�������"
    If costNode.getAttribute("Value") <> nill Then
        cdcsDBValues(0) = Replace(costNode.getAttribute("Value"), ".", ",")
    End If
    Set costChild = costNode.FirstChild
    While (Not costChild Is Nothing)
        '������ ��������
        For i = 1 To 6
            If (costChild.NodeName = cdcsXMLTags(i)) Then cdcsDBValues(i) = costChild.Text
        Next i
        '������ ����. �� ��� � ��� ����
        If (costChild.NodeName = cdcsXMLTags(7)) Then
            cdcsDBValues(7) = ParsXMLDocs051(tblName & "_docs", "cdcs_id", cdcs_id, cadNum, costChild)
        End If
        Set costChild = costChild.NextSibling
    Wend
    ' -----------------------
    ' ----- ������ � �� -----
    ' -----------------------
    '������������ ������ � ������
    For i = 0 To 9
        If cdcsDBTypes(i) Then cdcsDBValues(i) = "{$}" & cdcsDBValues(i) & "{$}"
    Next i
    '��������� �������
    For i = 0 To 8
        cdcsDBValues(i) = cdcsDBValues(i) & ","
    Next i
    '������� ������ �� ���������� ������
    sqlStr = "update " & tblName & " set "
    For i = 0 To 9
        sqlStr = sqlStr & cdcsDBFields(i) & "=" & cdcsDBValues(i)
    Next i
    sqlStr = sqlStr & " where cdcs_id = " & cdcs_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLCost051 = "+"
End Function