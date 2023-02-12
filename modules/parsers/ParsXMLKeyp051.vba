Option Compare Database
Public Function ParsXMLKeyp051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal keypNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML Documents
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    'Получаем поля БД
    Dim numsDBFields(4) As String
        numsDBFields(0) = "KeyType"
        numsDBFields(1) = "KeyValue"
        numsDBFields(2) = tblKeyName
        numsDBFields(3) = "CadastralNumber"
        numsDBFields(4) = "Reserved"
    'Служебное
    Dim insertSQL As String
    Dim keypType, keypValue As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    Set keypChild = keypNode.FirstChild
    Set insertDB = CurrentDb
    While (Not keypChild Is Nothing)
        If keypChild.getAttribute("Type") <> nill Then keypType = keypChild.getAttribute("Type")
        If keypChild.getAttribute("Value") <> nill Then keypValue = Replace(keypChild.getAttribute("Value"), ".", ",")
        insertSQL = "insert into " & tblName & "(" & numsDBFields(0) & "," & numsDBFields(1) & "," & numsDBFields(2) & "," & numsDBFields(3)
        insertSQL = insertSQL & ") values ("
        insertSQL = insertSQL & "'" & keypType & "','" & keypValue & "'," & tblKeyValue & ",'" & cadNum & "');"
        insertDB.Execute insertSQL
        Set keypChild = keypChild.NextSibling
    Wend
    Set insertDB = Nothing
    ParsXMLKeyp051 = "+"
End Function