Option Compare Database
Public Function ParsXMLSubc051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal subcNode As Object) As String
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
    Dim subcXMLTags() As Variant
        subcXMLTags = GetSubcConfig051(True)
    'Получаем поля БД
    Dim subcDBFields() As Variant
        subcDBFields = GetSubcConfig051(False)
        subcDBFields(5) = tblKeyName
    Dim subcDBValues(7) As String
    'Типы данных в БД строковые (s) или численные (d)
    Dim subcDBTypes() As Variant
        subcDBTypes = GetSubcTypes051()
    'Служебное
    Dim i As Integer
    Dim subc_id As String
    Dim sqlStr As String
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Два дополнительных поля приходят снаружи

    Set subcNode = subcNode.FirstChild
    While (Not subcNode Is Nothing)
        'Зарезервируем и получим id будущей записи
        subc_id = ReserveID(tblName, "subc_id")
        subcDBValues(0) = ""
        subcDBValues(1) = ""
        subcDBValues(2) = ""
        subcDBValues(3) = ""
        subcDBValues(4) = ""
        subcDBValues(5) = tblKeyValue
        subcDBValues(6) = cadNum
        If subcNode.getAttribute("NumberRecord") <> nill Then subcDBValues(0) = subcNode.getAttribute("NumberRecord")
        If subcNode.getAttribute("DateCreated") <> nill Then subcDBValues(1) = subcNode.getAttribute("DateCreated")
        'Парсим
        Set subcChild = subcNode.FirstChild
        While (Not subcChild Is Nothing)
            'Парсим значения
            If (subcChild.NodeName = subcXMLTags(2)) Then
                If subcChild.getAttribute("Type") <> nill Then subcDBValues(2) = subcChild.getAttribute("Type")
                If subcChild.getAttribute("Value") <> nill Then subcDBValues(3) = Replace(subcChild.getAttribute("Value"), ".", ",")
            End If
            'Парсим один тип
            If (subcChild.NodeName = subcXMLTags(4)) Then subcDBValues(4) = ParsXMLEnbr051(tblName & "_enbr", "subc_id", subc_id, cadNum, subcChild)
            Set subcChild = subcChild.NextSibling
        Wend
        ' -----------------------
        ' ----- Запись в БД -----
        ' -----------------------
        'Готовим запрос на добавление данных
        sqlStr = "update " & tblName & " set "
        For i = 0 To 6
            If subcDBTypes(i) Then subcDBValues(i) = "{$}" & subcDBValues(i) & "{$}"
            If (i < 6) Then subcDBValues(i) = subcDBValues(i) & ","
            sqlStr = sqlStr & subcDBFields(i) & "=" & subcDBValues(i)
        Next i
        sqlStr = sqlStr & " where subc_id = " & subc_id & ";"
        sqlStr = PrepareInsertSQL(sqlStr)
        Set insertDB = CurrentDb
        insertDB.Execute sqlStr
        Set insertDB = Nothing
        Set subcNode = subcNode.NextSibling
    Wend
    ParsXMLSubc051 = "+"
End Function
