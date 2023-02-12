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


Public Function ParsXMLCons051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal consNode As Object) As String
    'Получаем
    '   tblName - префикс таблиц XML
    '   tblKeyName - название идентификатора XML
    '   tblKeyValue - идентификатор XML
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML CadastralNumber

    ' --------------------------------------
    ' ----- Конфигурация таблиц зданий -----
    ' --------------------------------------
    Dim consXMLTags(23) As String
        'Прямые значения
        consXMLTags(0) = "CadastralNumber"
        consXMLTags(1) = "DateCreated"
        consXMLTags(2) = "FoundationDate"
        consXMLTags(3) = "CadastralBlock"
        consXMLTags(4) = "PreviouslyPosted"
        consXMLTags(5) = "Name"
        consXMLTags(6) = "ObjectType"
        consXMLTags(7) = "AssignationName"
        'Дополнительная обработка
        consXMLTags(8) = "ExploitationChar"
        consXMLTags(9) = ""                             'Не используется
        consXMLTags(10) = "Floors"
        consXMLTags(11) = ""                            'Не используется
        'Парсится отдельно
        consXMLTags(12) = "KeyParameters"
        consXMLTags(13) = "ParentCadastralNumbers"
        consXMLTags(14) = "PrevCadastralNumbers"
        consXMLTags(15) = "FlatsCadastralNumbers"
        consXMLTags(16) = "CarParkingSpacesCadastralNumbers"
        consXMLTags(17) = "UnitedCadastralNumber"
        consXMLTags(18) = "ObjectPermittedUses"
        consXMLTags(19) = "CadastralCost"
        consXMLTags(20) = "SubConstructions"
        consXMLTags(21) = "FacilityCadastralNumber"
        consXMLTags(22) = "CulturalHeritage"
        consXMLTags(23) = "Location"

    Dim consDBFields(25)
        consDBFields(0) = "CadastralNumber"
        consDBFields(1) = "DatesCreated"
        consDBFields(2) = "FoundationDates"
        consDBFields(3) = "CadastralBlock"
        consDBFields(4) = "PreviouslyPosted"
        consDBFields(5) = "Names"
        consDBFields(6) = "ObjectType"
        consDBFields(7) = "AssignationNames"
        consDBFields(8) = "YearBuilt"
        consDBFields(9) = "YearUsed"
        consDBFields(10) = "Floors"
        consDBFields(11) = "UndergroundFloors"
        consDBFields(12) = "KeyParameters"
        consDBFields(13) = "ParentCadastralNumbers"
        consDBFields(14) = "PrevCadastralNumbers"
        consDBFields(15) = "FlatsCadastralNumbers"
        consDBFields(16) = "CarParkingSpacesCadastralNumbers"
        consDBFields(17) = "UnitedCadastralNumber"
        consDBFields(18) = "ObjectPermittedUses"
        consDBFields(19) = "CadastralCost"
        consDBFields(20) = "SubConstructions"
        consDBFields(21) = "FacilityCadastralNumber"
        consDBFields(22) = "CulturalHeritage"
        consDBFields(23) = "addr_id"
        consDBFields(24) = tblKeyName
        consDBFields(25) = "Reserved"
    Dim consDBValues(25)

    'Типы данных в БД строковые (true) или численные (false)
    Dim consDBTypes(25) As Boolean
        consDBTypes(0) = True
        consDBTypes(1) = True
        consDBTypes(2) = True
        consDBTypes(3) = True
        consDBTypes(4) = True
        consDBTypes(5) = True
        consDBTypes(6) = True
        consDBTypes(7) = True
        consDBTypes(8) = True
        consDBTypes(9) = True
        consDBTypes(10) = True
        consDBTypes(11) = True
        consDBTypes(12) = True
        consDBTypes(13) = True
        consDBTypes(14) = True
        consDBTypes(15) = True
        consDBTypes(16) = True
        consDBTypes(17) = True
        consDBTypes(18) = True
        consDBTypes(19) = True
        consDBTypes(20) = True
        consDBTypes(21) = True
        consDBTypes(22) = True
        consDBTypes(23) = False
        consDBTypes(24) = False
        consDBTypes(25) = False

    'Служебные переменные и база данных
    Dim i As Integer
    Dim cons_id As String
    Dim cadNum As String
    Dim insertSQL As String

    'Одно поле приходит снаружи
    consDBValues(24) = tblKeyValue
    'Зарезервируем и получим id будущей записи
    cons_id = ReserveID(tblName, "cons_id")
    consDBValues(25) = "null"
    'В качестве атрибутов узла приходят еще несколько полей
    If consNode.getAttribute(consXMLTags(0)) <> nill Then
        consDBValues(0) = consNode.getAttribute(consXMLTags(0))
        cadNum = consDBValues(0)
    End If
    If consNode.getAttribute(consXMLTags(1)) <> nill Then
        consDBValues(1) = consNode.getAttribute(consXMLTags(1))
    End If
    If consNode.getAttribute(consXMLTags(2)) <> nill Then
        consDBValues(2) = consNode.getAttribute(consXMLTags(2))
    End If

    'Парсим
    Set consChild = consNode.FirstChild
    While (Not consChild Is Nothing)
        'Парсим значения
        For i = 3 To 7
            If (consChild.NodeName = consXMLTags(i)) Then consDBValues(i) = consChild.Text
        Next i
        'Парсим атрибуты
        If (consChild.NodeName = consXMLTags(8)) Then
            If consChild.getAttribute("YearBuilt") <> nill Then consDBValues(8) = consChild.getAttribute("YearBuilt")
            If consChild.getAttribute("YearUsed") <> nill Then consDBValues(9) = consChild.getAttribute("YearUsed")
        End If
        If (consChild.NodeName = consXMLTags(10)) Then
            If consChild.getAttribute("Floors") <> nill Then consDBValues(10) = consChild.getAttribute("Floors")
            If consChild.getAttribute("UndergroundFloors") <> nill Then consDBValues(11) = consChild.getAttribute("UndergroundFloors")
        End If
        'Парсим типы
        If (consChild.NodeName = consXMLTags(12)) Then consDBValues(12) = ParsXMLKeyp051(tblName & "_keyp", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(13)) Then consDBValues(13) = ParsXMLNums051(tblName & "_prnt", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(14)) Then consDBValues(14) = ParsXMLNums051(tblName & "_prev", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(15)) Then consDBValues(15) = ParsXMLNums051(tblName & "_flat", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(16)) Then consDBValues(16) = ParsXMLNums051(tblName & "_cars", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(17)) Then consDBValues(17) = ParsXMLNums051(tblName & "_unit", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(18)) Then consDBValues(18) = ParsXMLPrUs051(tblName & "_perm", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(19)) Then consDBValues(19) = ParsXMLCost051(tblName & "_cost", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(20)) Then consDBValues(20) = ParsXMLSubc051(tblName & "_subc", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(21)) Then consDBValues(21) = ParsXMLFacl051(tblName & "_facl", "cons_id", cons_id, cadNum, consChild)
        If (consChild.NodeName = consXMLTags(22)) Then consDBValues(22) = ParsXMLCult051(tblName & "_cult", "cons_id", cons_id, cadNum, consChild)
        'Адреса парсятся особо
        If (consChild.NodeName = consXMLTags(23)) Then consDBValues(23) = CStr(parsXMLAddress051(consChild.FirstChild))

        Set consChild = consChild.NextSibling
    Wend
    'Обрабатываем строки в данных
    insertSQL = "update " & tblName & " set "
    For i = 0 To 24
        If consDBTypes(i) Then consDBValues(i) = "{$}" & consDBValues(i) & "{$}"
        If i < 24 Then consDBValues(i) = consDBValues(i) & ","
        insertSQL = insertSQL & consDBFields(i) & "=" & consDBValues(i)
    Next i
    insertSQL = insertSQL & " where cons_id = " & cons_id & ";"
    insertSQL = PrepareInsertSQL(insertSQL)
    Set insertDB = CurrentDb
    tmp = SaveTXTfile("C:\Users\Kaz_MYu\Downloads\sql.txt", insertSQL)
    insertDB.Execute insertSQL
    Set insertDB = Nothing
End Function
Public Function ParsXMLSubc051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal subcNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML Documents

    ' -------------------------------------------------
    ' ----- Конфигурация таблиц SubConstructions  -----
    ' -------------------------------------------------
    'Названия тегов адресов в XML Росреестра
    Dim subcXMLTags(4) As String
        subcXMLTags(0) = "NumberRecord"
        subcXMLTags(1) = "DateCreated"
        subcXMLTags(2) = "KeyParameter"
        subcXMLTags(3) = ""
        subcXMLTags(4) = "Encumbrances"

    'Поля в таблице кадастровых стоимостей в БД
    Dim subcDBFields(7) As String
        subcDBFields(0) = "NumberRecord"
        subcDBFields(1) = "DatesCreated"
        subcDBFields(2) = "Types"
        subcDBFields(3) = "Values"
        subcDBFields(4) = "Encumbrances"
        subcDBFields(5) = tblKeyName
        subcDBFields(6) = "CadastralNumber"
        subcDBFields(7) = "Reserved"
    Dim subcDBValues(7) As String

    'Типы данных в БД строковые (s) или численные (d)
    Dim subcDBTypes(7) As String
        subcDBTypes(0) = "s"
        subcDBTypes(1) = "s"
        subcDBTypes(2) = "s"
        subcDBTypes(3) = "s"
        subcDBTypes(4) = "s"
        subcDBTypes(5) = "d"
        subcDBTypes(6) = "s"
        subcDBTypes(7) = "d"

    'Служебные переменные и база данных
    Dim i As Integer
    Dim subc_id As String
    Dim insertSQL As String

    'Два дополнительных поля приходят снаружи
    subcDBValues(5) = tblKeyValue
    subcDBValues(6) = cadNum

    Set subcNode = subcNode.FirstChild
    While (Not subcNode Is Nothing)
        'Зарезервируем и получим id будущей записи
        subc_id = ReserveID(tblName, "subc_id")
        subcDBValues(7) = "null"
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
            If (subcChild.NodeName = subcXMLTags(4)) Then subcDBValues(3) = ParsXMLEnbr051(tblName & "_enbr", "subc_id", subc_id, cadNum, subcChild)
            Set subcChild = subcChild.NextSibling
        Wend
        'Обрабатываем строки в данных
        For i = 0 To 6
            If subcDBTypes(i) = "s" Then subcDBValues(i) = "{$}" & subcDBValues(i) & "{$}"
        Next i
        'Добавляем запятые
        For i = 0 To 5
            subcDBValues(i) = subcDBValues(i) & ","
        Next i
        'Готовим запрос на добавление данных
        insertSQL = "update " & tblName & " set "
        For i = 0 To 6
            insertSQL = insertSQL & subcDBFields(i) & "=" & subcDBValues(i)
        Next i
        insertSQL = insertSQL & " where subc_id = " & subc_id & ";"
        insertSQL = PrepareInsertSQL(insertSQL)
        Set insertDB = CurrentDb
        insertDB.Execute insertSQL
        Set insertDB = Nothing
        Set subcNode = subcNode.NextSibling
    Wend
    ParsXMLSubc051 = "+"
End Function
Public Function ParsXMLKeyp051(ByVal tblName As String, ByVal tblKeyName As String, ByVal tblKeyValue As String, ByVal cadNum As String, ByVal keypNode As Object) As String
    'Получаем
    '   tblName - название основной таблицы объекта
    '   tblKeyName - название идентификатора объекта
    '   tblKeyValue - идентификатор объекта
    '   cadNum - кадастрвоый номер объекта
    '   Ссылка на узел XML

    ' ---------------------------------------------
    ' ----- Конфигурация таблиц KeyParameters -----
    ' ---------------------------------------------
 
    Dim numsDBFields(4) As String
        numsDBFields(0) = "KeyType"
        numsDBFields(1) = "KeyValue"
        numsDBFields(2) = tblKeyName
        numsDBFields(3) = "CadastralNumber"
        numsDBFields(4) = "Reserved"

    Dim insertSQL As String
    Dim keypType, keypValue As String

    'Парсим
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
