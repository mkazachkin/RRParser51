Option Compare Database
Public Function ParsXMLDocs051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal docsNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    'Получаем теги
    Dim docsXMLTags() As Variant
        docsXMLTags = GetDocsConfig051(True)
    'Получаем поля БД
    Dim docsDBFields() As Variant
        docsDBFields = GetDocsConfig051(False)
        docsDBFields(7) = tblKeyName
    Dim docsDBValues(9) As String
    'Получаем типы данных
    Dim docsDBTypes() As Variant
        docsDBTypes = GetDocsTypes051()
    'Служебное
    Dim i As Integer
    Dim docs_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Два дополнительных поля приходят снаружи
    docsDBValues(7) = tblKeyValue
    docsDBValues(8) = cadNum
    'Зарезервируем и получим id будущей записи
    docs_id = ReserveID(tblName, "docs_id")
    'Парсим
    Set docsChild = docsNode.FirstChild
    While (Not docsChild Is Nothing)
        'Парсим значения
        For i = 0 To 6
            If (docsChild.NodeName = docsXMLTags(i)) Then docsDBValues(i) = docsChild.Text
        Next i
        'Типов нет, их парсить не надо
        Set docsChild = docsChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Готовим запрос на добавление данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 8
        If docsDBTypes(i) Then docsDBValues(i) = "{$}" & docsDBValues(i) & "{$}"
        If (i < 8) Then docsDBValues(i) = docsDBValues(i) & ","
        sqlStr = sqlStr & docsDBFields(i) & "=" & docsDBValues(i)
    Next i
    sqlStr = sqlStr & " where docs_id = " & docs_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLDocs051 = "+"
End Function
