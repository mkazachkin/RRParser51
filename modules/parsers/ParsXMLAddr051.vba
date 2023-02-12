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