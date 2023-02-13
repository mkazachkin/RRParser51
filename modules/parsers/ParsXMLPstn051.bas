Option Compare Database
Public Function ParsXMLPstn051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal pstnNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    'Получаем поля БД
    Dim pstnDBFields() As Variant
        pstnDBFields = GetPstnConfig051(False)
        pstnDBFields(4) = tblKeyName
    Dim pstnDBValues(6) As String
    'Получаем типы данных
    Dim pstnDBTypes() As Variant
        pstnDBTypes = GetPstnTypes051()
    'Служебное
    Dim i As Integer
    Dim pstn_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Часть данных получаем извне
    pstnDBValues(4) = tblKeyValue
    pstnDBValues(5) = cadNum
    'Зарезервируем и получим id будущей записи
    pstn_id = ReserveID(tblName, "pstn_id")
    pstnDBValues(6) = "null"
    'Парсим
    Set pstnChild = pstnNode.FirstChild
    While (Not pstnChild Is Nothing)
        If pstnChild.getAttribute("Number") <> nill Then pstnDBValues(0) = pstnChild.getAttribute("Number")
        If pstnChild.getAttribute("Type") <> nill Then pstnDBValues(1) = pstnChild.getAttribute("Type")
        If (Not pstnChild.FirstChild Is Nothing) Then If pstnChild.FirstChild.getAttribute("NumberOnPlan") <> nill Then pstnDBValues(2) = pstnChild.FirstChild.getAttribute("NumberOnPlan")
        If (Not pstnChild.FirstChild Is Nothing) Then If pstnChild.FirstChild.getAttribute("Description") <> nill Then pstnDBValues(3) = pstnChild.FirstChild.getAttribute("Description")
        Set pstnChild = pstnChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Обрабатываем строки в данных
    sqlStr = "update " & tblName & " set "
    For i = 0 To 5
        If pstnDBTypes(i) Then pstnDBValues(i) = "{$}" & pstnDBValues(i) & "{$}"
        If (i < 5) Then pstnDBValues(i) = pstnDBValues(i) & ","
        sqlStr = sqlStr & pstnDBFields(i) & "=" & pstnDBValues(i)
    Next i
    sqlStr = sqlStr & " where pstn_id = " & pstn_id & ";"
    insertSQL = PrepareInsertSQL(sqlStr)
    Set insertDB = CurrentDb
    insertDB.Execute sqlStr
    Set insertDB = Nothing
    ParsXMLPstn051 = "+"
End Function
