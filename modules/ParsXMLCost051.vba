Option Compare Database
Public Function ParsXMLCost051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal costNode As Object) As String
    '��������
    '   tblName - �������� �������� ������� �������
    '   tblKeyName - �������� �������������� �������
    '   tblKeyValue - ������������� �������
    '   cadNum - ����������� ����� �������
    '   ������ �� ���� XML CadastralNumber

    ' -----------------------------------------------------
    ' ----- ������������ ������ ����������� ��������� -----
    ' -----------------------------------------------------
    '�������� ����� ������� � XML ����������
    Dim cdcsXMLTags(7) As String
        cdcsXMLTags(0) = "CadastralCost"
        cdcsXMLTags(1) = "DateValuation"
        cdcsXMLTags(2) = "DateEntering"
        cdcsXMLTags(3) = "DateApproval"
        cdcsXMLTags(4) = "ApplicationDate"
        cdcsXMLTags(5) = "RevisalStatementDate"
        cdcsXMLTags(6) = "ApplicationLastDate"
        cdcsXMLTags(7) = "ApprovalDocument"

    '���� � ������� ����������� ���������� � ��
    Dim cdcsDBFields(10) As String
        cdcsDBFields(0) = "CadastralCost"
        cdcsDBFields(1) = "DatesValuation"
        cdcsDBFields(2) = "DatesEntering"
        cdcsDBFields(3) = "DatesApproval"
        cdcsDBFields(4) = "ApplicationDates"
        cdcsDBFields(5) = "RevisalStatementDates"
        cdcsDBFields(6) = "ApplicationLastDates"
        cdcsDBFields(7) = "ApprovalDocument"
        cdcsDBFields(8) = tblKeyName                        '������������� � ������� ��������, ��� ������� �������� ���������� ���������
        cdcsDBFields(9) = "CadastralNumber"                '����������� ����� �������, ��� �������� �������� ����������� ���������
        cdcsDBFields(10) = "Reserved"                       '����������������� ��������� ����
    Dim cdcsDBValues(10) As String

    '���� ������ � �� ��������� (s) ��� ��������� (d)
    Dim cdcsDBTypes(10) As String
        cdcsDBTypes(0) = "s"
        cdcsDBTypes(1) = "s"
        cdcsDBTypes(2) = "s"
        cdcsDBTypes(3) = "s"
        cdcsDBTypes(4) = "s"
        cdcsDBTypes(5) = "s"
        cdcsDBTypes(6) = "s"
        cdcsDBTypes(7) = "s"
        cdcsDBTypes(8) = "d"
        cdcsDBTypes(9) = "s"
        cdcsDBTypes(10) = "d"

    '��������� ���������� � ���� ������
    Dim i As Integer
    Dim cdcs_id As String
    Dim insertSQL As String

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

    '������
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

    '������������ ������ � ������
    For i = 0 To 9
        If cdcsDBTypes(i) = "s" Then cdcsDBValues(i) = "{$}" & cdcsDBValues(i) & "{$}"
    Next i

    '��������� �������
    For i = 0 To 8
        cdcsDBValues(i) = cdcsDBValues(i) & ","
    Next i

    '������� ������ �� ���������� ������
    insertSQL = "update " & tblName & " set "
    For i = 0 To 9
        insertSQL = insertSQL & cdcsDBFields(i) & "=" & cdcsDBValues(i)
    Next i
    insertSQL = insertSQL & " where cdcs_id = " & cdcs_id & ";"
    insertSQL = PrepareInsertSQL(insertSQL)
    Set insertDB = CurrentDb
    insertDB.Execute insertSQL
    Set insertDB = Nothing

    ParsXMLCost051 = "+"
End Function