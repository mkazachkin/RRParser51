Option Compare Database
Public Function ParsXMLCons051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal consNode As Object) As String
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
    Dim consXMLTags() As Variant
        consXMLTags = GetConsConfig051(True)
    'Получаем поля БД
    Dim consDBFields() As Variant
        consDBFields = GetConsConfig051(False)
        consDBFields(24) = tblKeyName
    'Инициализируем значения
    Dim consDBValues(25) As String
    'Получаем типы данных
    Dim consDBTypes() As Variant
        consDBTypes = GetConsTypes051()
    'Служебное
    Dim i As Integer
    Dim cons_id As String
    Dim cadNum As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Одно поле приходит снаружи
    consDBValues(24) = tblKeyValue
    'Зарезервируем и получим id будущей записи
    cons_id = ReserveID(tblName, "cons_id")
    consDBValues(25) = "null"
    'В качестве атрибутов узла приходят еще несколько полей
    If consNode.getAttribute(consXMLTags(0)) <> nill Then
        consDBValues(0) = consNode.getAttribute(consXMLTags(0))
        cadNum = consDBValues(0)
    End If
    If consNode.getAttribute(consXMLTags(1)) <> nill Then
        consDBValues(1) = consNode.getAttribute(consXMLTags(1))
    End If
    If consNode.getAttribute(consXMLTags(2)) <> nill Then
        consDBValues(2) = consNode.getAttribute(consXMLTags(2))
    End If
    'Парсим
    Set consChild = consNode.FirstChild
    While (Not consChild Is Nothing)
        'Парсим значения
        For i = 3 To 7
            If (consChild.NodeName = consXMLTags(i)) Then consDBValues(i) = consChild.Text
        Next i
        'Парсим атрибуты
        If (consChild.NodeName = consXMLTags(8)) Then
            If consChild.getAttribute("YearBuilt") <> nill Then consDBValues(8) = consChild.getAttribute("YearBuilt")
            If consChild.getAttribute("YearUsed") <> nill Then consDBValues(9) = consChild.getAttribute("YearUsed")
        End If
        If (consChild.NodeName = consXMLTags(10)) Then
            If consChild.getAttribute("Floors") <> nill Then consDBValues(10) = consChild.getAttribute("Floors")
            If consChild.getAttribute("UndergroundFloors") <> nill Then consDBValues(11) = consChild.getAttribute("UndergroundFloors")
        End If
        'Парсим типы
        If (consChild.NodeName = consXMLTags(12)) Then consDBValues(12) = ParsXMLKeyp051(tblName & "_keyp", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(13)) Then consDBValues(13) = ParsXMLNums051(tblName & "_prnt", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(14)) Then consDBValues(14) = ParsXMLNums051(tblName & "_prev", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(15)) Then consDBValues(15) = ParsXMLNums051(tblName & "_flat", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(16)) Then consDBValues(16) = ParsXMLNums051(tblName & "_cars", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(17)) Then consDBValues(17) = ParsXMLNums051(tblName & "_unit", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(18)) Then consDBValues(18) = ParsXMLPerm051(tblName & "_perm", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(19)) Then consDBValues(19) = ParsXMLCost051(tblName & "_cost", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(20)) Then consDBValues(20) = ParsXMLSubc051(tblName & "_subc", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(21)) Then consDBValues(21) = ParsXMLFacl051(tblName & "_facl", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(22)) Then consDBValues(22) = ParsXMLCult051(tblName & "_cult", "cons_id", cons_id, cadNum, consChild)
        'Адреса парсятся особо
        If (consChild.NodeName = consXMLTags(23)) Then consDBValues(23) = ParsXMLAddr051(consChild)

        Set consChild = consChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Обрабатываем строки в данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 24
        If consDBTypes(i) Then consDBValues(i) = "{$}" & consDBValues(i) & "{$}"
        If i < 24 Then consDBValues(i) = consDBValues(i) & ","
        sqlStr = sqlStr & consDBFields(i) & "=" & consDBValues(i)
    Next i
    sqlStr = sqlStr & " where cons_id = " & cons_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLCons051 = "+"
End Function
