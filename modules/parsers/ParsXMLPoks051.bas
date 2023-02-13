Option Compare Database
Public Function ParsXMLPoks051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal poksNode As Object) As String
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
    Dim poksXMLTags() As Variant
        poksXMLTags = GetPoksConfig051(True)
    'Получаем поля БД
    Dim poksDBFields() As Variant
        poksDBFields = GetPoksConfig051(False)
        poksDBFields(9) = tblKeyName
    Dim poksDBValues(11) As String
    'Получаем типы данных
    Dim poksDBTypes() As Variant
        poksDBTypes = GetPoksTypes051()
    'Служебное
    Dim i As Integer
    Dim prnt_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Часть данных получаем извне
    poksDBValues(9) = tblKeyValue
    poksDBValues(10) = cadNum
    'Зарезервируем и получим id будущей записи
    prnt_id = ReserveID(tblName, "prnt_id")
    poksDBValues(11) = "null"
    'Парсим
    Set poksChild = poksNode.FirstChild
    While (Not poksChild Is Nothing)
        If poksChild.NodeName = poksXMLTags(0) Then poksDBValues(0) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(1) Then poksDBValues(1) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(2) Then poksDBValues(2) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(3) Then poksDBValues(3) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(4) Then If poksChild.FirstChild.getAttribute("Wall") <> nill Then poksDBValues(4) = poksChild.FirstChild.getAttribute("Wall")
        If poksChild.NodeName = poksXMLTags(5) Then If poksChild.getAttribute("YearBuilt") <> nill Then poksDBValues(5) = poksChild.getAttribute("YearBuilt")
        If poksChild.NodeName = poksXMLTags(6) Then If poksChild.getAttribute("YearUsed") <> nill Then poksDBValues(6) = poksChild.getAttribute("YearUsed")
        If poksChild.NodeName = poksXMLTags(7) Then If poksChild.getAttribute("Floors") <> nill Then poksDBValues(7) = poksChild.getAttribute("Floors")
        If poksChild.NodeName = poksXMLTags(8) Then If poksChild.getAttribute("UndergroundFloors") <> nill Then poksDBValues(8) = poksChild.getAttribute("UndergroundFloors")
        Set poksChild = poksChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Обрабатываем строки в данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 10
        If poksDBTypes(i) Then poksDBValues(i) = "{$}" & poksDBValues(i) & "{$}"
        If i < 10 Then poksDBValues(i) = poksDBValues(i) & ","
        sqlStr = sqlStr & poksDBFields(i) & "=" & poksDBValues(i)
    Next i
    sqlStr = sqlStr & " where prnt_id = " & prnt_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLPoks051 = "+"
End Function
