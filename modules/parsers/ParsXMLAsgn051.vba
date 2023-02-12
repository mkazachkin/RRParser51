Option Compare Database
Public Function ParsXMLAsgn051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal asgnNode As Object) As String
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
    Dim asgnXMLTags(7) As String
        asgnXMLTags = GetAsgnConfig051(true)
    'Получаем поля БД
    Dim asgnDBFields(7) As String
        asgnDBFields = GetAsgnConfig051(false)
        asgnDBFields(5) = tblKeyName
    Dim asgnDBValues(7) As String
    'Получаем типы данных
    Dim asgnDBTypes(7) As Boolean
        asgnDBTypes = GetAsgnTypes051()
    'Служебное
    Dim i As Integer
    Dim asgn_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Часть данных получаем извне
    asgnDBValues (5) = tblKeyValue
    asgnDBValues (6) = cadNum
    'Зарезервируем и получим id будущей записи
    asgn_id = ReserveID(tblName, "asgn_id")
    poksDBValues(7) = "null"
    'Парсим
    Set asgnChild = asgnNode.FirstChild
    While (Not asgnChild Is Nothing)
        For i = 0 To 4
            If asgnChild.NodeName = asgnXMLTags (i) Then asgnDBFields (i) = asgnChild.Text
        Next i
        Set asgnChild = asgnChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Обрабатываем строки в данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 7
        If asgnDBTypes(i) Then asgnDBValues(i) = "{$}" & asgnDBValues(i) & "{$}"
        If i < 7 Then asgnDBValues(i) = asgnDBValues(i) & ","
        sqlStr = sqlStr & asgnDBFields(i) & "=" & asgnDBValues(i)
    Next i
    sqlStr = sqlStr & " where asgn_id = " & asgn_id & ";"
    sqlStr = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLPoks051 = "+"
End Function