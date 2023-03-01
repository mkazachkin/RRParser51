Option Compare Database
Public Function ParsXMLCars051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal carsNode As Object) As String
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
    Dim carsXMLTags() As Variant
        carsXMLTags = GetCarsConfig051(True)
    'Получаем поля БД
    Dim carsDBFields() As Variant
        carsDBFields = GetCarsConfig051(False)
        carsDBFields(15) = tblKeyName
    'Инициализируем значения
    Dim carsDBValues(16) As String
    'Получаем типы данных
    Dim carsDBTypes() As Variant
        carsDBTypes = GetCarsTypes051()
    'Служебное
    Dim i As Integer
    Dim cars_id As String
    Dim cadNum As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Одно поле приходит снаружи
    carsDBValues(15) = tblKeyValue
    'Зарезервируем и получим id будущей записи
    cars_id = ReserveID(tblName, "cars_id")
    'В качестве атрибутов узла приходят еще несколько полей
    If carsNode.getAttribute(carsXMLTags(0)) <> nill Then
        carsDBValues(0) = carsNode.getAttribute(carsXMLTags(0))
        cadNum = carsDBValues(0)
    End If
    If carsNode.getAttribute(carsXMLTags(1)) <> nill Then
        carsDBValues(1) = carsNode.getAttribute(carsXMLTags(1))
    End If
    If carsNode.getAttribute(carsXMLTags(2)) <> nill Then
        carsDBValues(2) = carsNode.getAttribute(carsXMLTags(2))
    End If
    Set carsChild = carsNode.FirstChild
    While (Not carsChild Is Nothing)
        'Парсим значения
        If (carsChild.NodeName = carsXMLTags(3)) Then carsDBValues(3) = carsChild.Text
        If (carsChild.NodeName = carsXMLTags(4)) Then carsDBValues(4) = carsChild.Text
        If (carsChild.NodeName = carsXMLTags(5)) Then carsDBValues(5) = carsChild.Text
        If (carsChild.NodeName = carsXMLTags(6)) Then carsDBValues(6) = carsChild.Text
        If (carsChild.NodeName = carsXMLTags(7)) Then carsDBValues(7) = carsChild.Text
        'Парсим типы
        If (carsChild.NodeName = carsXMLTags(8)) Then carsDBValues(8) = ParsXMLPoks051(tblName & "_poks", "cars_id", cars_id, cadNum, carsChild)
        If (carsChild.NodeName = carsXMLTags(9)) Then carsDBValues(9) = ParsXMLNums051(tblName & "_prev", "cars_id", cars_id, cadNum, carsChild)
        If (carsChild.NodeName = carsXMLTags(10)) Then carsDBValues(10) = ParsXMLPstn051(tblName & "_pstn", "cars_id", cars_id, cadNum, carsChild)
        If (carsChild.NodeName = carsXMLTags(11)) Then carsDBValues(11) = ParsXMLFacl051(tblName & "_unit", "cars_id", cars_id, cadNum, carsChild)
        If (carsChild.NodeName = carsXMLTags(12)) Then carsDBValues(12) = ParsXMLAddr051(carsChild)
        If (carsChild.NodeName = carsXMLTags(13)) Then carsDBValues(13) = ParsXMLCost051(tblName & "_cost", "cars_id", cars_id, cadNum, carsChild)
        If (carsChild.NodeName = carsXMLTags(14)) Then carsDBValues(14) = ParsXMLFacl051(tblName & "_facl", "cars_id", cars_id, cadNum, carsChild)

        Set carsChild = carsChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Area нужно отработать отдельно
    carsDBValues(7) = Replace(carsDBValues(7), ".", ",")
    'Готовим запрос на добавление данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 15
        If carsDBTypes(i) Then carsDBValues(i) = "{$}" & carsDBValues(i) & "{$}"
        If (i < 15) Then carsDBValues(i) = carsDBValues(i) & ","
        sqlStr = sqlStr & carsDBFields(i) & "=" & carsDBValues(i)
    Next i
    sqlStr = sqlStr & " where cars_id = " & cars_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    ParsXMLCars051 = sqlStr
End Function

