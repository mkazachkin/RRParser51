Option Compare Database
Public Function ParsXMLUcmp051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal ucmpNode As Object) As String
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
    Dim ucmpXMLTags() As Variant
        ucmpXMLTags = GetUcmpConfig051(True)
    'Получаем поля БД
    Dim ucmpDBFields() As Variant
        ucmpDBFields = GetUcmpConfig051(False)
        ucmpDBFields(14) = tblKeyName
    'Инициализируем значения
    Dim ucmpDBValues(15) As String
    'Получаем типы данных
    Dim ucmpDBTypes() As Variant
        ucmpDBTypes = GetUcmpTypes051()
    'Служебное
    Dim i As Integer
    Dim ucmp_id As String
    Dim cadNum As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Одно поле приходит снаружи
    ucmpDBValues(14) = tblKeyValue
    'Зарезервируем и получим id будущей записи
    ucmp_id = ReserveID(tblName, "ucmp_id")
    ucmpDBValues(15) = "null"
    'В качестве атрибутов узла приходят еще несколько полей
    If ucmpNode.getAttribute(ucmpXMLTags(0)) <> nill Then
        ucmpDBValues(0) = ucmpNode.getAttribute(ucmpXMLTags(0))
        cadNum = ucmpDBValues(0)
    End If
    If ucmpNode.getAttribute(ucmpXMLTags(1)) <> nill Then
        ucmpDBValues(1) = ucmpNode.getAttribute(ucmpXMLTags(1))
    End If
    If ucmpNode.getAttribute(ucmpXMLTags(2)) <> nill Then
        ucmpDBValues(2) = ucmpNode.getAttribute(ucmpXMLTags(2))
    End If
    Set ucmpChild = ucmpNode.FirstChild
    While (Not ucmpChild Is Nothing)
        'Парсим значения
        If (ucmpChild.NodeName = ucmpXMLTags(3)) Then ucmpDBValues(3) = ucmpChild.Text
        If (ucmpChild.NodeName = ucmpXMLTags(4)) Then ucmpDBValues(4) = ucmpChild.Text
        If (ucmpChild.NodeName = ucmpXMLTags(5)) Then ucmpDBValues(5) = ucmpChild.Text
        If (ucmpChild.NodeName = ucmpXMLTags(6)) Then ucmpDBValues(6) = ucmpChild.Text
        If (ucmpChild.NodeName = ucmpXMLTags(7)) Then ucmpDBValues(7) = ucmpChild.Text
        'Парсим типы
        If (ucmpChild.NodeName = ucmpXMLTags(8)) Then ucmpDBValues(8) = ParsXMLKeyp051(tblName & "_keyp", "ucmp_id", ucmp_id, cadNum, ucmpChild)
        If (consChild.NodeName = ucmpXMLTags(9)) Then ucmpDBValues(9) = ParsXMLNums051(tblName & "_prnt", "ucmp_id", ucmp_id, cadNum, consChild)
        If (consChild.NodeName = ucmpXMLTags(10)) Then ucmpDBValues(10) = ParsXMLNums051(tblName & "_prev", "ucmp_id", ucmp_id, cadNum, consChild)
        If (builChild.NodeName = builXMLTags(11)) Then builDBValues(11) = ParsXMLAddr051(ucmpChild)
        If (ucmpChild.NodeName = ucmpXMLTags(12)) Then ucmpDBValues(12) = ParsXMLCost051(tblName & "_cost", "ucmp_id", ucmp_id, cadNum, ucmpChild)
        If (ucmpChild.NodeName = ucmpXMLTags(13)) Then ucmpDBValues(13) = ParsXMLFacl051(tblName & "_facl", "ucmp_id", ucmp_id, cadNum, ucmpChild)

        Set ucmpChild = ucmpChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Обрабатываем строки в данных
    For i = 0 To 14
        If ucmpDBTypes(i) Then ucmpDBValues(i) = "{$}" & ucmpDBValues(i) & "{$}"
    Next i
    'Добавляем запятые
    For i = 0 To 13
        ucmpDBValues(i) = ucmpDBValues(i) & ","
    Next i
    'Готовим запрос на добавление данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 14
        sqlStr = sqlStr & ucmpDBFields(i) & "=" & ucmpDBValues(i)
    Next i
    sqlStr = sqlStr & " where ucmp_id = " & ucmp_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLUcmp051 = "+"
End Function
