Option Compare Database
Public Function ParsXMLSubb051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal subbNode As Object) As String
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
    Dim subbXMLTags() As Variant
        subbXMLTags = GetSubbConfig051(True)
    'Получаем поля БД
    Dim subbDBFields() As Variant
        subbDBFields = GetSubbConfig051(False)
        subbDBFields(4) = tblKeyName
    Dim subbDBValues(6) As String
    'Получаем типы данных
    Dim subbDBTypes() As Variant
        subbDBTypes = GetSubbTypes051()
    'Служебное
    Dim i As Integer
    Dim subb_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Два дополнительных поля приходят снаружи
    Set builChild = subbNode.FirstChild
    Set insertDB = CurrentDb
    While (Not builChild Is Nothing)
        'Зарезервируем и получим id будущей записи
        subb_id = ReserveID(tblName, "subb_id")
        subbDBValues(0) = ""
        subbDBValues(1) = ""
        subbDBValues(2) = ""
        subbDBValues(3) = ""
        subbDBValues(4) = tblKeyValue
        subbDBValues(5) = cadNum
        subbDBValues(6) = "null"
        If builChild.getAttribute("NumberRecord") <> nill Then subbDBValues(0) = builChild.getAttribute("NumberRecord")
        If builChild.getAttribute("DateCreated") <> nill Then subbDBValues(1) = builChild.getAttribute("DateCreated")
        'Парсим
        Set subbChild = builChild.FirstChild
        While (Not subbChild Is Nothing)
            'Парсим значения
            If (subbChild.NodeName = subbXMLTags(2)) Then subbDBValues(2) = Replace(subbChild.Text, ".", ",")
            'Парсим один тип
            If (subbChild.NodeName = subbXMLTags(3)) Then subbDBValues(3) = ParsXMLEnbr051(tblName & "_enbr", "subb_id", subb_id, cadNum, subbChild)
            Set subbChild = subbChild.NextSibling
        Wend
        ' -----------------------
        ' ----- Запись в БД -----
        ' -----------------------
        'Обрабатываем строки в данных
        For i = 0 To 5
            If subbDBTypes(i) Then subbDBValues(i) = "{$}" & subbDBValues(i) & "{$}"
        Next i
        'Добавляем запятые
        For i = 0 To 4
            subbDBValues(i) = subbDBValues(i) & ","
        Next i
        'Готовим запрос на добавление данных
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
