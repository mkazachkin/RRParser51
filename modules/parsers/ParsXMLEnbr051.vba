Option Compare Database
Public Function ParsXMLEnbr051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal enbrNode As Object) As String
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
    Dim enbrXMLTags() As Variant
        enbrXMLTags = GetEnbrConfig051(True)
    'Получаем поля БД
    Dim enbrDBFields() As Variant
        enbrDBFields = GetEnbrConfig051(False)
        enbrDBFields(5) = tblKeyName
    Dim enbrDBValues(7) As String
    'Получаем типы данных
    Dim enbrDBTypes() As Variant
        enbrDBTypes = GetEnbrTypes051()
    'Служебное
    Dim i As Integer
    Dim enbr_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    Set enbrNode = enbrNode.FirstChild
    While (Not enbrNode Is Nothing)
        'Зарезервируем и получим id будущей записи
        enbr_id = ReserveID(tblName, "enbr_id")
        enbrDBValues(0) = ""
        enbrDBValues(1) = ""
        enbrDBValues(2) = ""
        enbrDBValues(3) = ""
        enbrDBValues(4) = ""
        enbrDBValues(5) = tblKeyValue
        enbrDBValues(6) = cadNum
        enbrDBValues(7) = "null"
        'Парсим
        Set enbrChild = enbrNode.FirstChild
        While (Not enbrChild Is Nothing)
            'Парсим значения
            If (enbrChild.NodeName = enbrXMLTags(0)) Then enbrDBValues(0) = enbrChild.Text
            If (enbrChild.NodeName = enbrXMLTags(1)) Then enbrDBValues(1) = enbrChild.Text
            If (enbrChild.NodeName = enbrXMLTags(2)) Then
                Set subb = enbrChild.FirstChild
                While (Not subb Is Nothing)
                    If (subb.NodeName = "RightNumber") Then enbrDBValues(2) = subb.Text
                    If (subb.NodeName = "RegistrationDate") Then enbrDBValues(3) = subb.Text
                    Set subb = subb.NextSibling
                Wend
            End If
            'Парсим один тип
            If (enbrChild.NodeName = enbrXMLTags(4)) Then enbrDBValues(4) = ParsXMLDocs051(tblName & "_docs", "enbr_id", enbr_id, cadNum, enbrChild)
            Set enbrChild = enbrChild.NextSibling
        Wend
        'Обрабатываем строки в данных
        For i = 0 To 6
            If enbrDBTypes(i) Then enbrDBValues(i) = "{$}" & enbrDBValues(i) & "{$}"
        Next i
        'Добавляем запятые
        For i = 0 To 5
            enbrDBValues(i) = enbrDBValues(i) & ","
        Next i
        'Готовим запрос на добавление данных
        sqlStr = "update " & tblName & " set "
        For i = 0 To 6
            sqlStr = sqlStr & enbrDBFields(i) & "=" & enbrDBValues(i)
        Next i
        sqlStr = sqlStr & " where enbr_id = " & enbr_id & ";"
        sqlStr = PrepareInsertSQL(sqlStr)
        Set insertDB = CurrentDb
        insertDB.Execute sqlStr
        Set insertDB = Nothing
        Set enbrNode = enbrNode.NextSibling
    Wend
    ParsXMLEnbr051 = "+"
End Function
