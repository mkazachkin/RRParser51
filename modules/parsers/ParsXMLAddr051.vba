Option Compare Database
Public Function ParsXMLAddr051(ByVal addrNode As Object) As String
    'Получаем
    '   Ссылка на узел XML Address
    ' ------------------------
    ' ----- Конфигурация -----
    ' ------------------------
    'Получаем теги
    Dim addrXMLTag(33) As String
        addrXMLTag = GetAddrConfig051(true)
    'Получаем поля БД
    Dim addrDBFields(33) As String
        addrDBFields = GetAddrConfig051(false)
    'Инициализируем значения
    Dim addrDBValues(33) As String
    'Получаем типы данных
    Dim addrDBTypes(33) As Boolean
        addrDBTypes = GetAddrTypes051()
    'Задаем название таблицы адресов
    Dim tblName As String
        tblName = "public_import_t_address"
    'Задаем параметры словаря регионов
    Dim dictRegi(2) As String
        dictRegi(0) = "RegionCode"
        dictRegi(1) = "public_import_dict_region"
        dictRegi(2) = "regi_id"
    'Служебное
    Dim i As Integer
    Dim rs As Recordset
    Dim shaStr, sqlStr As String
    Dim addr_id As Long
    ' -------------------
    ' ----- Парсинг -----
    ' -------------------
    'Получаем значения адреса из XML
    'Код вида адреса и его название
    If (addrNode.getAttribute(addrXMLTag(30)) <> nill) Then
        addrDBValues(30) = addrNode.getAttribute(addrXMLTag(30))
        If addrNode.getAttribute(addrXMLTag(30))= "0" Then
            addrDBValues(31) = "Местоположение объекта недвижимости"
        Else
            addrDBValues(31) = "Присвоенный в установленном порядке адрес объекта недвижимости"
        End If
    End If
    'Парсим потомков addressChild
    Set addrChild = addrNode.FirstChild
    While (Not addrChild Is Nothing)
        For i = 0 To 8
            If (addrChild.NodeName = addrXMLTag(i)) Then
                addrDBValues(0) = addrChild.Text
            End If
        Next i
        'Region
        If (addrChild.NodeName = addrXMLTag(9)) Then
            addrDBValues(9) = CStr(DictCheck(addrChild.Text, dictRegi(0), dictRegi(1), dictRegi(2)))
        End If
        For i = 10 To 20 Step 2
            If (addrChild.NodeName = addrXMLTag(i)) Then
                If addrChild.getAttribute("Type") <> nill Then
                    addrDBValues(i) = addrChild.getAttribute("Type")
                End If
                If addrChild.getAttribute("Name") <> nill Then
                    addrDBValues(i+1) = addrChild.getAttribute("Name")
                End If
            End If
        Next i
        For i = 22 To 28 Step 2
            If (addrChild.NodeName = addrXMLTag(i)) Then
                If addrChild.getAttribute("Type") <> nill Then
                    addrDBValues(i) = addrChild.getAttribute("Type")
                End If
                If addrChild.getAttribute("Value") <> nill Then
                    addrDBValues(i+1) = addrChild.getAttribute("Value")
                End If
            End If
        Next i

        Set addrChild = addrChild.NextSibling
    Wend
    ' -----------------------
    ' ----- Запись в БД -----
    ' -----------------------
    'Считаем хэш
    shaStr = ""
    For i = 0 To 31
        shaStr = shaStr & "$" & addrDBValues(i)
    Next i
    addrDBValues(32) = SHA256(shaStr)
    'Проверяем, есть ли такой адрес в БД
    sqlStr = "select addr_id from " & tblName & " where " & addrDBFields(32) & "='" & addrDBValues(32) & "';"
    Set insertDB = CurrentDb
    Set rs = insertDB.OpenRecordset(sqlStr)
    If (rs.RecordCount = 0) Then
        'Зарезервируем и получим id будущей записи
        addr_id = ReserveID(tblName, "addr_id")
        addrDBValues(33) = "null"
        'Обрабатываем строки в данных
        For i = 0 To 32
            If addrDBTypes(i) Then addrDBValues(i) = "{$}" & addrDBValues(i) & "{$}"
        Next i
        'Добавляем запятые
        For i = 0 To 31
            addrDBValues(i) = addrDBValues(i) & ","
        Next i
        'Готовим запрос на добавление данных
        sqlStr = "update " & tblName & " set "
        For i = 0 To 32
            sqlStr = sqlStr & addrDBFields(i) & "=" & addrDBValues(i)
        Next i
        sqlStr = sqlStr & " where addr_id = " & addr_id & ";"
        sqlStr = PrepareInsertSQL(sqlStr)
        insertDB.Execute sqlStr
    Else
        addr_id = Cstr(rs.Fields.Item(0).Value)
    End If
    Set insertDB = Nothing
    Set rs = Nothing
    'Возвращаем id
    parsXMLAddr051 = addr_id
End Function



Public Function ParsXMLPoks051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal poksNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML

    ' -----------------------------------------
    ' ----- Конфигурация таблиц ParentOKS -----
    ' -----------------------------------------
     'Названия тегов адресов в XML Росреестра
    Dim poksXMLTags(7) As String
        poksXMLTags(0) = "CadastralNumberOKS"
        poksXMLTags(1) = "ObjectType"
        poksXMLTags(2) = "AssignationBuilding"
        poksXMLTags(3) = "AssignationName"
        poksXMLTags(4) = "ElementsConstruct"
        poksXMLTags(5) = "ExploitationChar"
        poksXMLTags(6) = "ExploitationChar"
        poksXMLTags(7) = "Floors"
        poksXMLTags(8) = "Floors"

    Dim poksDBFields(11) As String
        poksDBFields(0) = "CadastralNumberOKS"
        poksDBFields(1) = "ObjectType"
        poksDBFields(2) = "AssignationBuilding"
        poksDBFields(3) = "AssignationNames"
        poksDBFields(4) = "WallsCode"
        poksDBFields(5) = "YearBuilt"
        poksDBFields(6) = "YearUsed"
        poksDBFields(7) = "Floors"
        poksDBFields(8) = "UndergroundFloors"
        poksDBFields(9) = tblKeyName
        poksDBFields(10) = "CadastralNumber"
        poksDBFields(11) = "Reserved"
    Dim poksDBValues(11) As String

    Dim poksDBTypes(11) As Boolean
        poksDBTypes(0) = True
        poksDBTypes(1) = True
        poksDBTypes(2) = True
        poksDBTypes(3) = True
        poksDBTypes(4) = True
        poksDBTypes(5) = True
        poksDBTypes(6) = True
        poksDBTypes(7) = True
        poksDBTypes(8) = True
        poksDBTypes(9) = False
        poksDBTypes(10) = True
        poksDBTypes(11) = False

    Dim insertSQL As String
    Dim prnt_id As String
    Dim i As Integer

    'Часть данных получаем извне
    poksDBValues(9) = tblKeyValue
    poksDBValues(10) = cadNum

    'Зарезервируем и получим id будущей записи
    prnt_id = ReserveID(tblName, "prnt_id")
    poksDBValues(11) = "null"

    'Парсим
    Set poksChild = poksNode.FirstChild
    While (Not poksChild Is Nothing)
        If poksChild.NodeName = poksXMLTags(0) Then poksDBFields(0) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(1) Then poksDBFields(1) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(2) Then poksDBFields(2) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(3) Then poksDBFields(3) = poksChild.Text
        If poksChild.NodeName = poksXMLTags(4) Then If poksChild.FirstChild.getAttribute("Wall") <> nill Then poksDBFields(4) = poksChild.FirstChild.getAttribute("Wall")
        If poksChild.NodeName = poksXMLTags(5) Then If poksChild.getAttribute("YearBuilt") <> nill Then poksDBFields(5) = poksChild.getAttribute("YearBuilt")
        If poksChild.NodeName = poksXMLTags(6) Then If poksChild.getAttribute("YearUsed") <> nill Then poksDBFields(6) = poksChild.getAttribute("YearUsed")
        If poksChild.NodeName = poksXMLTags(7) Then If poksChild.getAttribute("Floors") <> nill Then poksDBFields(7) = poksChild.getAttribute("Floors")
        If poksChild.NodeName = poksXMLTags(8) Then If poksChild.getAttribute("UndergroundFloors") <> nill Then poksDBFields(8) = poksChild.getAttribute("UndergroundFloors")
        Set poksChild = poksChild.NextSibling
    Wend

    'Обрабатываем строки в данных
    insertSQL = "update " & tblName & " set "
    For i = 0 To 10
        If poksDBTypes(i) Then poksDBValues(i) = "{$}" & poksDBValues(i) & "{$}"
        If i < 10 Then poksDBValues(i) = poksDBValues(i) & ","
        insertSQL = insertSQL & poksDBFields(i) & "=" & poksDBValues(i)
    Next i
    insertSQL = insertSQL & " where prnt_id = " & prnt_id & ";"
    insertSQL = PrepareInsertSQL(insertSQL)
    Set insertDB = CurrentDb
    insertDB.Execute insertSQL
    Set insertDB = Nothing
    ParsXMLPoks051 = "+"
End Function
Public Function ParsXMLAsgn051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal asgnNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML

    ' -------------------------------------------
    ' ----- Конфигурация таблиц Assignation -----
    ' -------------------------------------------
     'Названия тегов адресов в XML Росреестра
    Dim asgnXMLTags(4) As String
        asgnXMLTags(0) = "AssignationCode"
        asgnXMLTags(1) = "AssignationType"
        asgnXMLTags(2) = "SpecialType"
        asgnXMLTags(3) = "TotalAssets"
        asgnXMLTags(4) = "AuxiliaryFlat"

    Dim asgnDBFields(7) As String
        asgnDBFields(0) = "AssignationCode"
        asgnDBFields(1) = "AssignationType"
        asgnDBFields(2) = "SpecialType"
        asgnDBFields(3) = "TotalAssets"
        asgnDBFields(4) = "AuxiliaryFlat"
        asgnDBFields(5) = tblKeyName
        asgnDBFields(6) = "CadastralNumber"
        asgnDBFields(7) = "Reserved"
    Dim asgnDBValues(7) As String

    Dim asgnDBTypes(7) As Boolean
        asgnDBTypes(0) = true
        asgnDBTypes(1) = true
        asgnDBTypes(2) = true
        asgnDBTypes(3) = true
        asgnDBTypes(4) = true
        asgnDBTypes(5) = false
        asgnDBTypes(6) = true
        asgnDBTypes(7) = false

    Dim insertSQL As String
    Dim asgn_id As String
    Dim i As Integer

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

    'Обрабатываем строки в данных
    insertSQL = "update " & tblName & " set "
    For i = 0 To 7
        If asgnDBTypes(i) Then asgnDBValues(i) = "{$}" & asgnDBValues(i) & "{$}"
        If i < 7 Then asgnDBValues(i) = asgnDBValues(i) & ","
        insertSQL = insertSQL & asgnDBFields(i) & "=" & asgnDBValues(i)
    Next i
    insertSQL = insertSQL & " where asgn_id = " & asgn_id & ";"
    insertSQL = PrepareInsertSQL(insertSQL)
    Set insertDB = CurrentDb
    insertDB.Execute insertSQL
    Set insertDB = Nothing
    ParsXMLPoks051 = "+"
End Function
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
