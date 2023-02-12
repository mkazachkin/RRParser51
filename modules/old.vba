Public Function parsXMLAddress051(ByVal addressNode As Object) As Long
    ' ---------------------------------------
    ' ----- Конфигурация таблиц адресов -----
    ' ---------------------------------------
    'Названия тегов адресов в XML Росреестра
    Dim addressXMLValues(20) As String
    addressXMLValues(0) = "FIAS"
    addressXMLValues(1) = "OKATO"
    addressXMLValues(2) = "KLADR"
    addressXMLValues(3) = "OKTMO"
    addressXMLValues(4) = "PostalCode"
    addressXMLValues(5) = "RussianFederation"
    addressXMLValues(6) = "Region"
    addressXMLValues(7) = "District"
    addressXMLValues(8) = "City"
    addressXMLValues(9) = "UrbanDistrict"
    addressXMLValues(10) = "SovietVillage"
    addressXMLValues(11) = "Locality"
    addressXMLValues(12) = "Street"
    addressXMLValues(13) = "Level1"
    addressXMLValues(14) = "Level2"
    addressXMLValues(15) = "Level3"
    addressXMLValues(16) = "Apartment"
    addressXMLValues(17) = "Other"
    addressXMLValues(18) = "Note"
    addressXMLValues(19) = "ReadableAddress"
    addressXMLValues(20) = "AddressOrLocation"

    'Название таблицы адресов в БД
    Dim addressDBTable  As String
    addressDBTable = "public_import_t_address"
    'Поля в таблице адресов в БД
    Dim addressDBTableKey As String
    addressDBTableKey = "addr_id"
    Dim addressDBTableFields(32) As String
    addressDBTableFields(0) = "FIAS"
    addressDBTableFields(1) = "OKATO"
    addressDBTableFields(2) = "KLADR"
    addressDBTableFields(3) = "OKTMO"
    addressDBTableFields(4) = "PostalCode"
    addressDBTableFields(5) = "RussianFederation"
    addressDBTableFields(6) = "regi_id"
    addressDBTableFields(7) = "DistrictType"
    addressDBTableFields(8) = "DistrictName"
    addressDBTableFields(9) = "CityType"
    addressDBTableFields(10) = "CityName"
    addressDBTableFields(11) = "UrbanDistrictType"
    addressDBTableFields(12) = "UrbanDistrictName"
    addressDBTableFields(13) = "SovietVillageType"
    addressDBTableFields(14) = "SovietVillageName"
    addressDBTableFields(15) = "LocalityType"
    addressDBTableFields(16) = "LocalityName"
    addressDBTableFields(17) = "StreetType"
    addressDBTableFields(18) = "StreetName"
    addressDBTableFields(19) = "Level1Type"
    addressDBTableFields(20) = "Level1Name"
    addressDBTableFields(21) = "Level2Type"
    addressDBTableFields(22) = "Level2Name"
    addressDBTableFields(23) = "Level3Type"
    addressDBTableFields(24) = "Level3Name"
    addressDBTableFields(25) = "ApartmentType"
    addressDBTableFields(26) = "ApartmentName"
    addressDBTableFields(27) = "Other"
    addressDBTableFields(28) = "Notes"
    addressDBTableFields(29) = "ReadableAddress"
    addressDBTableFields(30) = "AddressOrLocationCode"
    addressDBTableFields(31) = "AddressOrLocation"
    addressDBTableFields(32) = "sha256hash"
    
    Dim hashstring As String
    
    'Словарь регионов
    Dim dictRegi(2) As String
    dictRegi(0) = "RegionCode"
    dictRegi(1) = "public_import_dict_region"
    dictRegi(2) = "regi_id"


    'Парсим и добавляем информацию по адресам в БД. Возвращаем id адреса в таблице. Поддерживаемая версия 5.1
    'Получаем ссылку на узел в XML файле
    Set addressObject = addressNode
    
    'Задаем массив значений
    Dim addressDBTableValues(32) As String
    
    'Служебное
    Dim counter As Integer
    Dim rs As Recordset

    'Получаем значения адреса из XML
    'Код вида адреса и его название
    If (addressObject.getAttribute(addressXMLValues(20)) <> nill) Then
        addressDBTableValues(30) = addressObject.getAttribute(addressXMLValues(20))
        If addressDBTableValues(30) = "0" Then
            addressDBTableValues(31) = "Местоположение объекта недвижимости"
        Else
            addressDBTableValues(31) = "Присвоенный в установленном порядке адрес объекта недвижимости"
        End If
    Else
        addressDBTableValues(30) = ""
        addressDBTableValues(31) = ""
    End If
    'Парсим потомков addressChild
    Set addressChild = addressObject
    While (Not addressChild Is Nothing)
        'FIAS
        If (addressChild.NodeName = addressXMLValues(0)) Then
            addressDBTableValues(0) = addressChild.Text
        End If
        'OKATO
        If (addressChild.NodeName = addressXMLValues(1)) Then
            addressDBTableValues(1) = addressChild.Text
        End If
        'KLADR
        If (addressChild.NodeName = addressXMLValues(2)) Then
            addressDBTableValues(2) = addressChild.Text
        End If
        'Код OKTMO
        If (addressChild.NodeName = addressXMLValues(3)) Then
            addressDBTableValues(3) = addressChild.Text
        End If
        'PostalCode
        If (addressChild.NodeName = addressXMLValues(4)) Then
            addressDBTableValues(4) = addressChild.Text
        End If
        'RussianFederation
        If (addressChild.NodeName = addressXMLValues(5)) Then
            addressDBTableValues(5) = addressChild.Text
        End If
        'Region
        If (addressChild.NodeName = addressXMLValues(6)) Then
            addressDBTableValues(6) = CStr(DictCheck(addressChild.Text, dictRegi(0), dictRegi(1), dictRegi(2)))
        End If
        'District
        If (addressChild.NodeName = addressXMLValues(7)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(7) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Name") <> nill Then
                addressDBTableValues(8) = addressChild.getAttribute("Name")
            End If
        End If
        'City
        If (addressChild.NodeName = addressXMLValues(8)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(9) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Name") <> nill Then
                addressDBTableValues(10) = addressChild.getAttribute("Name")
            End If
        End If
        'UrbanDistrict
        If (addressChild.NodeName = addressXMLValues(9)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(11) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Name") <> nill Then
                addressDBTableValues(12) = addressChild.getAttribute("Name")
            End If
        End If
        'SovietVillage
        If (addressChild.NodeName = addressXMLValues(10)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(13) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Name") <> nill Then
                addressDBTableValues(14) = addressChild.getAttribute("Name")
            End If
        End If
        'Locality
        If (addressChild.NodeName = addressXMLValues(11)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(15) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Name") <> nill Then
                addressDBTableValues(16) = addressChild.getAttribute("Name")
            End If
        End If
        'Street
        If (addressChild.NodeName = addressXMLValues(12)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(17) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Name") <> nill Then
                addressDBTableValues(18) = addressChild.getAttribute("Name")
            End If
        End If
        'Level1
        If (addressChild.NodeName = addressXMLValues(13)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(19) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Value") <> nill Then
                addressDBTableValues(20) = addressChild.getAttribute("Value")
            End If
        End If
        'Level2
        If (addressChild.NodeName = addressXMLValues(14)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(21) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Value") <> nill Then
                addressDBTableValues(22) = addressChild.getAttribute("Value")
            End If
        End If
        'Level3
        If (addressChild.NodeName = addressXMLValues(15)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(23) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Value") <> nill Then
                addressDBTableValues(24) = addressChild.getAttribute("Value")
            End If
        End If
        'Apartment
        If (addressChild.NodeName = addressXMLValues(16)) Then
            If addressChild.getAttribute("Type") <> nill Then
                addressDBTableValues(25) = addressChild.getAttribute("Type")
            End If
            If addressChild.getAttribute("Value") <> nill Then
                addressDBTableValues(26) = addressChild.getAttribute("Value")
            End If
        End If
        'Other
        If (addressChild.NodeName = addressXMLValues(17)) Then
            addressDBTableValues(27) = addressChild.Text
        End If
        'Note
        If (addressChild.NodeName = addressXMLValues(18)) Then
            addressDBTableValues(28) = addressChild.Text
        End If
        'Адрес в текстовом виде
        If (addressChild.NodeName = addressXMLValues(19)) Then
            addressDBTableValues(29) = addressChild.Text
        End If
        Set addressChild = addressChild.NextSibling
    Wend
    
    'Считаем хэш
    hashstring = ""
    For counter = 0 To 31
        hashstring = hashstring & "$" & addressDBTableValues(counter)
    Next counter
    addressDBTableValues(32) = SHA256(hashstring)
    
    'Проверяем, есть ли такой адрес в БД
    selectSQL = "select " & addressDBTableKey & " from " & addressDBTable & " where " & addressDBTableFields(32) & "='" & addressDBTableValues(32) & "';"
    Set cadastrDB = CurrentDb
    Set rs = cadastrDB.OpenRecordset(selectSQL)
    If (rs.RecordCount = 0) Then
        'Добавляем новый адрес
        insertSQL = "insert into " & addressDBTable & " ("
        For counter = 0 To 31
            insertSQL = insertSQL & addressDBTableFields(counter) & ","
        Next counter
        insertSQL = insertSQL & addressDBTableFields(32) & ") values ("
        For counter = 0 To 31
            If counter <> 6 Then insertSQL = insertSQL & "{$}"
            insertSQL = insertSQL & addressDBTableValues(counter)
            If counter <> 6 Then insertSQL = insertSQL & "{$}"
            insertSQL = insertSQL & ", "
        Next counter
        insertSQL = insertSQL & "{$}" & addressDBTableValues(32) & "{$});"
        insertSQL = Replace(insertSQL, "{$}{$}", "null")
        insertSQL = Replace(insertSQL, "'", "''")
        insertSQL = Replace(insertSQL, "{$}", "'")
        cadastrDB.Execute insertSQL
        Set rs = cadastrDB.OpenRecordset(selectSQL)
    End If
    'Возвращаем данные
    rs.MoveLast
    rs.MoveFirst
    parsXMLAddress051 = rs.Fields.Item(0).Value
End Function