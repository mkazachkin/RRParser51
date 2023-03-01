Option Compare Database
Public Function ParsXMLFlat051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal flatNode As Object) As String
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
    Dim flatXMLTags() As Variant
        flatXMLTags = GetFlatConfig051(True)
    'Получаем поля БД
    Dim flatDBFields() As Variant
        flatDBFields = GetFlatConfig051(False)
        flatDBFields(22) = tblKeyName
    'Инициализируем значения
    Dim flatDBValues(23) As String
    'Получаем типы данных
    Dim flatDBTypes() As Variant
        flatDBTypes = GetFlatTypes051()
    'Служебное
    Dim i As Integer
    Dim flat_id As String
    Dim cadNum As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Одно поле приходит снаружи
    flatDBValues(22) = tblKeyValue
    'Зарезервируем и получим id будущей записи
    flat_id = ReserveID(tblName, "flat_id")
    'В качестве атрибутов узла приходят еще несколько полей
    If flatNode.getAttribute(flatXMLTags(0)) <> nill Then
        flatDBValues(0) = flatNode.getAttribute(flatXMLTags(0))
        cadNum = flatDBValues(0)
    End If
    If flatNode.getAttribute(flatXMLTags(1)) <> nill Then
        flatDBValues(1) = flatNode.getAttribute(flatXMLTags(1))
    End If
    If flatNode.getAttribute(flatXMLTags(2)) <> nill Then
        flatDBValues(2) = flatNode.getAttribute(flatXMLTags(2))
    End If
    Set flatChild = flatNode.FirstChild
    While (Not flatChild Is Nothing)
        'Парсим значения
        For i = 3 To 9
            If (flatChild.NodeName = flatXMLTags(i)) Then flatDBValues(i) = flatChild.Text
        Next i
        'Парсим типы
        If (flatChild.NodeName = flatXMLTags(10)) Then flatDBValues(10) = ParsXMLPoks051(tblName & "_poks", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(11)) Then flatDBValues(11) = ParsXMLNums051(tblName & "_prev", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(12)) Then flatDBValues(12) = ParsXMLAsgn051(tblName & "_asgn", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(13)) Then flatDBValues(13) = ParsXMLPstn051(tblName & "_pstn", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(14)) Then flatDBValues(14) = ParsXMLPerm051(tblName & "_perm", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(15)) Then flatDBValues(15) = ParsXMLCost051(tblName & "_cost", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(16)) Then flatDBValues(16) = ParsXMLSubb051(tblName & "_subf", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(17)) Then flatDBValues(17) = ParsXMLFacl051(tblName & "_unit", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(18)) Then flatDBValues(18) = ParsXMLFacl051(tblName & "_facl", "flat_id", flat_id, cadNum, flatChild)
        If (flatChild.NodeName = flatXMLTags(19)) Then flatDBValues(19) = ParsXMLCult051(tblName & "_cult", "flat_id", flat_id, cadNum, flatChild)
        'Адрес тут особый
        If (flatChild.NodeName = flatXMLTags(20)) Then
            flatDBValues(20) = ParsXMLAddr051(flatChild)
            Set roomNode = flatChild.FirstChild
            While (Not roomNode Is Nothing)
                If (roomNode.NodeName = flatXMLTags(21)) Then flatDBValues(21) = roomNode.Text
                Set roomNode = roomNode.NextSibling
            Wend
        End If

        Set flatChild = flatChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Area нужно отработать отдельно
    flatDBValues(9) = Replace(flatDBValues(9), ".", ",")
    'Готовим запрос на добавление данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 22
        If flatDBTypes(i) Then flatDBValues(i) = "{$}" & flatDBValues(i) & "{$}"
        If (i < 22) Then flatDBValues(i) = flatDBValues(i) & ","
        sqlStr = sqlStr & flatDBFields(i) & "=" & flatDBValues(i)
    Next i
    sqlStr = sqlStr & " where flat_id = " & flat_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    ParsXMLFlat051 = sqlStr
End Function
