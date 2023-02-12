Public Function ParsXMLPstn051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal pstnNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML

    ' ----------------------------------------------
    ' ----- Конфигурация таблиц PositionOnPlan -----
    ' ----------------------------------------------
     'Названия тегов адресов в XML Росреестра

    Dim pstnDBFields(6) As String
        pstnDBFields(0) = "Types"
        pstnDBFields(1) = "Numbers"
        pstnDBFields(2) = "NumberOnPlan"
        pstnDBFields(3) = "Description"
        pstnDBFields(4) = tblKeyName
        pstnDBFields(5) = "CadastralNumber"
        pstnDBFields(6) = "Reserved"
    Dim pstnDBFields(6) As String

    Dim pstnDBTypes(6) As Boolean
        pstnDBTypes(0) = true
        pstnDBTypes(1) = true
        pstnDBTypes(2) = true
        pstnDBTypes(3) = true
        pstnDBTypes(4) = false
        pstnDBTypes(5) = true
        pstnDBTypes(6) = false

    Dim insertSQL As String
    Dim pstn_id As String
    Dim i As Integer

    'Часть данных получаем извне
    pstnDBValues (4) = tblKeyValue
    pstnDBValues (5) = cadNum

    'Зарезервируем и получим id будущей записи
    pstn_id = ReserveID(tblName, "pstn_id")
    poksDBValues(6) = "null"

    'Парсим
    Set pstnChild = pstnNode.FirstChild
    While (Not pstnChild Is Nothing)
        If pstnChild.getAttribute ("Number") <> nill Then pstnDBValues(0) = pstnChild.getAttribute ("Number")
        If pstnChild.getAttribute ("Type") <> nill Then pstnDBValues(1) = pstnChild.getAttribute ("Type")
        If pstnChild.FirstChild.getAttribute ("NumberOnPlan") <> nill Then pstnDBValues(2) = pstnChild.FirstChild.getAttribute ("NumberOnPlan")
        If pstnChild.FirstChild.getAttribute ("Description") <> nill Then pstnDBValues(3) = pstnChild.FirstChild.getAttribute ("Description")
        Set pstnChild = pstnChild.NextSibling
    Wend

    'Обрабатываем строки в данных
    insertSQL = "update " & tblName & " set "
    For i = 0 To 5
        If pstnDBTypes(i) Then pstnDBValues(i) = "{$}" & pstnDBValues(i) & "{$}"
        If i < 5 Then pstnDBValues(i) = pstnDBValues(i) & ","
        insertSQL = insertSQL & pstnDBFields(i) & "=" & pstnDBValues(i)
    Next i
    insertSQL = insertSQL & " where pstn_id = " & pstn_id & ";"
    insertSQL = PrepareInsertSQL(insertSQL)
    Set insertDB = CurrentDb
    insertDB.Execute insertSQL
    Set insertDB = Nothing
    ParsXMLPstn051 = "+"
End Function
