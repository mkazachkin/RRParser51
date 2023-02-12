Option Compare Database
Public Function ParsXMLCult051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal cltrNode As Object) As String
    Dim cultXMLTags(7) As String
        cultXMLTags(0) = "InclusionEGROKN"
        cultXMLTags(1) = ""                                 'Не используется
        cultXMLTags(2) = ""                                 'Не используется
        cultXMLTags(3) = "AssignmentEGROKN"
        cultXMLTags(4) = ""                                 'Не используется
        cultXMLTags(5) = ""                                 'Не используется
        cultXMLTags(6) = "RequirementsEnsure"
        cultXMLTags(7) = "Document"


    Dim cultDBFields(10) As String
        cultDBFields(0) = "EGROKNRegNum"
        cultDBFields(1) = "EGROKNObjCultural"
        cultDBFields(2) = "EGROKNNameCultural"
        cultDBFields(3) = "AssignEGROKNRegNum"
        cultDBFields(4) = "AssignEGROKNObjCultural"
        cultDBFields(5) = "AssignAssignEGROKNRegNum"
        cultDBFields(6) = "RequirementsEnsure"
        cultDBFields(7) = "Document"
        cultDBFields(8) = tblKeyName
        cultDBFields(9) = "CadastralNumber"
        cultDBFields(10) = "Reserved"
    Dim cultDBValues(10) As String

    Dim cultDBTypes(10) As String
        cultDBTypes(0) = "s"
        cultDBTypes(1) = "s"
        cultDBTypes(2) = "s"
        cultDBTypes(3) = "s"
        cultDBTypes(4) = "s"
        cultDBTypes(5) = "s"
        cultDBTypes(6) = "s"
        cultDBTypes(7) = "s"
        cultDBTypes(8) = "d"
        cultDBTypes(9) = "s"
        cultDBTypes(10) = "d"

    Dim i As Integer
    Dim insertSQL As String
    Dim cult_id As String

    cultDBValues(8) = tblKeyValue
    cultDBValues(9) = cadCode

    'Зарезервируем и получим id будущей записи
    cult_id = ReserveID(tblName, "cult_id")
    cultDBValues(10) = "null"

    Set cultNode = cltrNode.FirstChild
    While (Not cultNode Is Nothing)
        If cultNode.NodeName = "InclusionEGROKN" Then
            Set egrkonNode = cultNode.FirstChild
            While (Not egrkonNode Is Nothing)
                If egrkonNode.NodeName = "RegNum" Then cultDBValues(0) = egrkonNode.Text
                If egrkonNode.NodeName = "ObjCultural" Then cultDBValues(1) = egrkonNode.Text
                If egrkonNode.NodeName = "NameCultural" Then cultDBValues(2) = egrkonNode.Text
                Set egrkonNode = egrkonNode.NextSibling
            Wend
        End If
        If cultNode.NodeName = "AssignmentEGROKN" Then
            Set egrkonNode = cultNode.FirstChild
            While (Not egrkonNode Is Nothing)
                If egrkonNode.NodeName = "RegNum" Then cultDBValues(3) = egrkonNode.Text
                If egrkonNode.NodeName = "ObjCultural" Then cultDBValues(4) = egrkonNode.Text
                If egrkonNode.NodeName = "NameCultural" Then cultDBValues(5) = egrkonNode.Text
                Set egrkonNode = egrkonNode.NextSibling
            Wend
        End If
        If cultNode.NodeName = "RequirementsEnsure" Then cultDBValues(6) = cultNode.Text
        If cultNode.NodeName = "Document" Then cultDBValues(7) = ParsXMLDocs051(tblName & "_docs", "cult_id", cult_id, cadNum, cultNode)
        Set cultNode = cultNode.NextSibling
    Wend
    'Обрабатываем строки в данных
    For i = 0 To 9
        If cultDBTypes(i) = "s" Then cultDBValues(i) = "{$}" & cultDBValues(i) & "{$}"
    Next i
    'Добавляем запятые
    For i = 0 To 8
        cultDBValues(i) = cultDBValues(i) & ","
    Next i
    'Готовим запрос на добавление данных
    insertSQL = "update " & tblName & " set "
    For i = 0 To 9
        insertSQL = insertSQL & cultDBFields(i) & "=" & cultDBValues(i)
    Next i
    insertSQL = insertSQL & " where cult_id = " & cult_id & ";"
    insertSQL = PrepareInsertSQL(insertSQL)
    Set insertDB = CurrentDb
    tmp = SaveTXTfile("C:\Users\Kaz_MYu\Downloads\sql.txt", insertSQL)
    insertDB.Execute insertSQL
    Set insertDB = Nothing
    ParsXMLCultur051 = "+"
End Function