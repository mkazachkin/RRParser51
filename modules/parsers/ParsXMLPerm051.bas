Option Compare Database
Public Function ParsXMLPerm051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal permNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    Dim pDBFields(3) As String
        pDBFields(0) = "ObjectPermittedUses"
        pDBFields(1) = tblKeyName
        pDBFields(2) = "CadastralNumber"
        pDBFields(3) = "Reserved"
    Dim sqlStr As String
    ' ---------------------------------
    ' ----- Парсинг и запись в БД -----
    ' ---------------------------------
    Set permChild = permNode.FirstChild
    Set insertDB = CurrentDb
    While (Not permChild Is Nothing)
        'Это просто список. Поэтому каждую запись сразу пишем в БД с привязкой к основному объекту
        sqlStr = "insert into " & tblName & "(" & pDBFields(0) & "," & pDBFields(1) & "," & pDBFields(2) & "," & pDBFields(3)
        sqlStr = sqlStr & ") values ("
        sqlStr = sqlStr & "'" & permChild.Text & "'," & tblKeyValue & ",'" & cadNum & "','" & SHA256(CStr(Rnd) + CStr(Now) + CStr(Timer) + CStr(Rnd)) & "');"
        insertDB.Execute sqlStr
        Set permChild = permChild.NextSibling
    Wend
    Set insertDB = Nothing
    ParsXMLPerm051 = "+"
End Function
