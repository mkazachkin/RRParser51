Option Compare Database
Public Function ParsXMLCost051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal costNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML CadastralNumber

    ' -----------------------------------------------------
    ' ----- Конфигурация таблиц кадастровой стоимости -----
    ' -----------------------------------------------------
    'Названия тегов адресов в XML Росреестра
    Dim cdcsXMLTags(7) As String
        cdcsXMLTags(0) = "CadastralCost"
        cdcsXMLTags(1) = "DateValuation"
        cdcsXMLTags(2) = "DateEntering"
        cdcsXMLTags(3) = "DateApproval"
        cdcsXMLTags(4) = "ApplicationDate"
        cdcsXMLTags(5) = "RevisalStatementDate"
        cdcsXMLTags(6) = "ApplicationLastDate"
        cdcsXMLTags(7) = "ApprovalDocument"

    'Поля в таблице кадастровых стоимостей в БД
    Dim cdcsDBFields(10) As String
        cdcsDBFields(0) = "CadastralCost"
        cdcsDBFields(1) = "DatesValuation"
        cdcsDBFields(2) = "DatesEntering"
        cdcsDBFields(3) = "DatesApproval"
        cdcsDBFields(4) = "ApplicationDates"
        cdcsDBFields(5) = "RevisalStatementDates"
        cdcsDBFields(6) = "ApplicationLastDates"
        cdcsDBFields(7) = "ApprovalDocument"
        cdcsDBFields(8) = tblKeyName                        'Идентификатор в таблице объектов, для которой парсится кадастовая стоимость
        cdcsDBFields(9) = "CadastralNumber"                'Кадастровый номер объекта, для которого парсится кадастровая стоимость
        cdcsDBFields(10) = "Reserved"                       'Зарезервированное служебное поле
    Dim cdcsDBValues(10) As String

    'Типы данных в БД строковые (s) или численные (d)
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

    'Служебные переменные и база данных
    Dim i As Integer
    Dim cdcs_id As String
    Dim insertSQL As String

    'Два дополнительных поля приходят снаружи
    cdcsDBValues(8) = tblKeyValue
    cdcsDBValues(9) = cadNum

    'Зарезервируем и получим id будущей записи
    cdcs_id = ReserveID(tblName, "cdcs_id")
    cdcsDBValues(10) = "null"

    'Кадастровая стоимость тоже приходит "снаружи"
    If costNode.getAttribute("Value") <> nill Then
        cdcsDBValues(0) = Replace(costNode.getAttribute("Value"), ".", ",")
    End If

    'Парсим
    Set costChild = costNode.FirstChild
    While (Not costChild Is Nothing)
        'Парсим значения
        For i = 1 To 6
            If (costChild.NodeName = cdcsXMLTags(i)) Then cdcsDBValues(i) = costChild.Text
        Next i
        'Парсим типы. Он тут у нас один
        If (costChild.NodeName = cdcsXMLTags(7)) Then
            cdcsDBValues(7) = ParsXMLDocs051(tblName & "_docs", "cdcs_id", cdcs_id, cadNum, costChild)
        End If
        Set costChild = costChild.NextSibling
    Wend

    'Обрабатываем строки в данных
    For i = 0 To 9
        If cdcsDBTypes(i) = "s" Then cdcsDBValues(i) = "{$}" & cdcsDBValues(i) & "{$}"
    Next i

    'Добавляем запятые
    For i = 0 To 8
        cdcsDBValues(i) = cdcsDBValues(i) & ","
    Next i

    'Готовим запрос на добавление данных
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