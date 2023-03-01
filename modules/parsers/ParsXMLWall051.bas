Option Compare Database
Public Function ParsXMLWall051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal wallNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    Dim wDBFields(3) As String
        wDBFields(0) = "WallsCode"
        wDBFields(1) = tblKeyName
        wDBFields(2) = "CadastralNumber"
        wDBFields(3) = "Reserved"
    Dim sqlStr As String
    ' ---------------------------------
    ' ----- Парсинг и запись в БД -----
    ' ---------------------------------
    Set wallChild = wallNode.FirstChild
    Set insertDB = CurrentDb
    While (Not wallChild Is Nothing)
        If (wallChild.getAttribute("Wall")) <> nill Then
            'Это просто список. Поэтому каждую запись сразу пишем в БД с привязкой к основному объектукту
            sqlStr = "insert into " & tblName & "(" & wDBFields(0) & "," & wDBFields(1) & "," & wDBFields(2) & "," & wDBFields(3)
            sqlStr = sqlStr & ") values ("
            sqlStr = sqlStr & "'" & wallChild.getAttribute("Wall") & "'," & tblKeyValue & ",'" & cadNum & "','" & SHA256(CStr(Rnd) + CStr(Now) + CStr(Timer) + CStr(Rnd)) & "');"
            insertDB.Execute sqlStr
        End If
        Set wallChild = wallChild.NextSibling
    Wend
    Set insertDB = Nothing
    ParsXMLWall051 = "+"
End Function
