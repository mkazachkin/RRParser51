Option Compare Database
Public Function ParsXMLFacl051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal faclNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML Documents
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    'Получаем теги
    Dim faclXMLTags(2) As String
        faclXMLTags(0) = "CadastralNumber"
        faclXMLTags(1) = "Purpose"
        faclXMLTags(2) = "Name"
    'Получаем поля БД
    Dim faclDBFields(5) As String
        faclDBFields(0) = "FacilityCadastralNumber"
        faclDBFields(1) = "Purpose"
        faclDBFields(2) = "Names"
        faclDBFields(3) = tblKeyName
        faclDBFields(4) = "CadastralNumber"
        faclDBFields(5) = "Reserved"
    Dim faclDBValues(5)
    'Служебное
    Dim insertSQL As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    faclDBValues(3) = tblKeyValue & ","
    faclDBValues(4) = "{$}" & cadNum & "{$});"
    Set faclChild = faclNode.FirstChild
    While (Not faclChild Is Nothing)
        If faclChild.NodeName = faclXMLTags(0) Then faclDBValues(0) = "({$}" & faclChild.Text & "{$},"
        If faclChild.NodeName = faclXMLTags(1) Then faclDBValues(1) = "{$}" & faclChild.Text & "{$},"
        If faclChild.NodeName = faclXMLTags(2) Then faclDBValues(2) = "{$}" & faclChild.Text & "{$},"
        Set faclChild = faclChild.NextSibling
    Wend
    sqlStr = "insert into " & tblName & "(" & faclDBFields(0) & "," & faclDBFields(1) & "," & faclDBFields(2) & "," & faclDBFields(3) & "," & faclDBFields(4) & ")"
    sqlStr = sqlStr & " values "
    sqlStr = sqlStr & faclDBValues(0) & faclDBValues(1) & faclDBValues(2) & faclDBValues(3) & faclDBValues(4)
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLFacl051 = "+"
End Function