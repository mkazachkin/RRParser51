Option Compare Database
Public Function ParsXMLBuil051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal builNode As Object) As String
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
    Dim builXMLTags(26) As String
        builXMLTags = GetBuilConfig051 (true)
    '�������� ���� ��
    Dim builDBFields(26)
        builDBFields = GetBuilConfig051 (false)
        builDBFields(25) = tblKeyName
    '�������������� ��������
    Dim builDBValues(26)
    '�������� ���� ������
    Dim builDBTypes(26) As Boolean
        builDBTypes = GetBuilTypes051()
    '���������
    Dim i As Integer
    Dim buil_id As String
    Dim cadNum As String
    Dim sqlStr As String
    ' -------------------
    ' ----- ������� -----
    ' -------------------
    '���� ���� �������� �������
    builDBValues(25) = tblKeyValue
    '������������� � ������� id ������� ������
    buil_id = ReserveID(tblName, "buil_id")
    builDBValues(26) = "null"
    '� �������� ��������� ���� �������� ��� ��������� �����
    If builNode.getAttribute(builXMLTags(0)) <> nill Then
        builDBValues(0) = builNode.getAttribute(builXMLTags(0))
        cadNum = builDBValues(0)
    End If
    If builNode.getAttribute(builXMLTags(1)) <> nill Then
        builDBValues(1) = builNode.getAttribute(builXMLTags(1))
    End If
    If builNode.getAttribute(builXMLTags(2)) <> nill Then
        builDBValues(2) = builNode.getAttribute(builXMLTags(2))
    End If
    Set builChild = builNode.FirstChild
    While (Not builChild Is Nothing)
        '������ ��������
        For i = 3 To 8
            If (builChild.NodeName = builXMLTags(i)) Then builDBValues(i) = builChild.Text
        Next i
        '������ ��������
        If (builChild.NodeName = builXMLTags(9)) Then
            Set child = builChild.FirstChild
            If child.getAttribute("Wall") <> nill Then builDBValues(9) = child.getAttribute("Wall")
            Set child = Nothing
        End If
        If (builChild.NodeName = builXMLTags(10)) Then
            If builChild.getAttribute("YearBuilt") <> nill Then builDBValues(10) = builChild.getAttribute("YearBuilt")
            If builChild.getAttribute("YearUsed") <> nill Then builDBValues(11) = builChild.getAttribute("YearUsed")
        End If
        If (builChild.NodeName = builXMLTags(12)) Then
            If builChild.getAttribute("Floors") <> nill Then builDBValues(12) = builChild.getAttribute("Floors")
            If builChild.getAttribute("UndergroundFloors") <> nill Then builDBValues(13) = builChild.getAttribute("UndergroundFloors")
        End If
        '������ ����
        If (builChild.NodeName = builXMLTags(14)) Then builDBValues(14) = ParsXMLNums051(tblName & "_prnt", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(15)) Then builDBValues(15) = ParsXMLNums051(tblName & "_prev", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(16)) Then builDBValues(16) = ParsXMLNums051(tblName & "_flat", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(17)) Then builDBValues(17) = ParsXMLNums051(tblName & "_cars", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(18)) Then builDBValues(18) = ParsXMLNums051(tblName & "_unit", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(19)) Then builDBValues(19) = ParsXMLAddr051(builChild)
        If (builChild.NodeName = builXMLTags(20)) Then builDBValues(20) = ParsXMLPerm051(tblName & "_perm", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(21)) Then builDBValues(21) = ParsXMLCost051(tblName & "_cost", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(22)) Then builDBValues(22) = ParsXMLSubb051(tblName & "_subb", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(23)) Then builDBValues(23) = ParsXMLFacl051(tblName & "_facl", "buil_id", buil_id, cadNum, builChild)
        If (builChild.NodeName = builXMLTags(24)) Then builDBValues(24) = ParsXMLCult051(tblName & "_cult", "buil_id", buil_id, cadNum, builChild)

        Set builChild = builChild.NextSibling
    Wend
    ' -----------------------
    ' ----- ������ � �� -----
    ' -----------------------
    'Area ����� ���������� ��������
    builDBValues(8) = Replace(builDBValues(8), ".", ",")
    '������������ ������ � ������
    For i = 0 To 25
        If builDBTypes(i) Then builDBValues(i) = "{$}" & builDBValues(i) & "{$}"
    Next i
    '��������� �������
    For i = 0 To 24
        builDBValues(i) = builDBValues(i) & ","
    Next i
    '������� ������ �� ���������� ������
    sqlStr = "update " & tblName & " set "
    For i = 0 To 25
        sqlStr = sqlStr & builDBFields(i) & "=" & builDBValues(i)
    Next i
    sqlStr = sqlStr & " where buil_id = " & buil_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
End Function