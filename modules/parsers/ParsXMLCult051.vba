Option Compare Database
Public Function ParsXMLCult051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal cultNode As Object) As String
    'Получаем
    '   tblName - префикс таблиц XML
    '   tblKeyName - название идентификатора XML
    '   tblKeyValue - идентификатор XML
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    'Получаем теги
    Dim cultXMLTags() As Variant
        cultXMLTags = GetCultConfig051(True)
    'Получаем поля БД
    Dim cultDBFields() As Variant
        cultDBFields = GetCultConfig051(False)
        cultDBFields(8) = tblKeyName
    Dim cultDBValues(10) As String
    'Получаем типы данных
    Dim cultDBTypes() As Variant
        cultDBTypes = GetCultTypes051()
    'Служебное
    Dim i As Integer
    Dim sqlStr As String
    Dim cult_id As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    cultDBValues(8) = tblKeyValue
    cultDBValues(9) = cadNum
    'Зарезервируем и получим id будущей записи
    cult_id = ReserveID(tblName, "cult_id")
    cultDBValues(10) = "null"
    Set cultNode = cultNode.FirstChild
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
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Обрабатываем строки в данных
    For i = 0 To 9
        If cultDBTypes(i) Then cultDBValues(i) = "{$}" & cultDBValues(i) & "{$}"
    Next i
    'Добавляем запятые
    For i = 0 To 8
        cultDBValues(i) = cultDBValues(i) & ","
    Next i
    'Готовим запрос на добавление данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 9
        sqlStr = sqlStr & cultDBFields(i) & "=" & cultDBValues(i)
    Next i
    sqlStr = sqlStr & " where cult_id = " & cult_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLCult051 = "+"
End Function
