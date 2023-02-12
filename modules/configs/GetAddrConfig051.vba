Option Compare Database
Public Function GetAddrConfig051 (xmlOrdb As Boolean) As String()
    Dim conf0 (33), conf1 (33) As String
    'XML теги                         'Поля в БД
    conf0(0) = "FIAS"               : conf1(0) = "FIAS"
    conf0(1) = "OKATO"              : conf1(1) = "OKATO"
    conf0(2) = "KLADR"              : conf1(2) = "KLADR"
    conf0(3) = "OKTMO"              : conf1(3) = "OKTMO"
    conf0(4) = "PostalCode"         : conf1(4) = "PostalCode"
    conf0(5) = "RussianFederation"  : conf1(5) = "RussianFederation"
    conf0(6) = "Other"              : conf1(6) = "Other"
    conf0(7) = "Note"               : conf1(7) = "Notes"
    conf0(8) = "ReadableAddress"    : conf1(8) = "ReadableAddress"
    conf0(9) = "Region"             : conf1(9) = "regi_id"
    conf0(10) = "District"          : conf1(10) = "DistrictType"
    conf0(11) = ""                  : conf1(11) = "DistrictName"
    conf0(12) = "City"              : conf1(12) = "CityType"
    conf0(13) = ""                  : conf1(13) = "CityName"
    conf0(14) = "UrbanDistrict"     : conf1(14) = "UrbanDistrictType"
    conf0(15) = ""                  : conf1(15) = "UrbanDistrictName"
    conf0(16) = "SovietVillage"     : conf1(16) = "SovietVillageType"
    conf0(17) = ""                  : conf1(17) = "SovietVillageName"
    conf0(18) = "Locality"          : conf1(18) = "LocalityType"
    conf0(19) = ""                  : conf1(19) = "LocalityName"
    conf0(20) = "Street"            : conf1(20) = "StreetType"
    conf0(21) = ""                  : conf1(21) = "StreetName"
    conf0(22) = "Level1"            : conf1(22) = "Level1Type"
    conf0(23) = ""                  : conf1(23) = "Level1Name"
    conf0(24) = "Level2"            : conf1(24) = "Level2Type"
    conf0(25) = ""                  : conf1(25) = "Level2Name"
    conf0(26) = "Level3"            : conf1(26) = "Level3Type"
    conf0(27) = ""                  : conf1(27) = "Level3Name"
    conf0(28) = "Apartment"         : conf1(28) = "ApartmentType"
    conf0(29) = ""                  : conf1(29) = "ApartmentName"
    conf0(30) = "AddressOrLocation" : conf1(30) = "AddressOrLocationCode"
    conf0(31) = ""                  : conf1(31) = "AddressOrLocation"
    conf0(32) = ""                  : conf1(32) = "sha256hash"
    conf0(33) = ""                  : conf1(33) = "Reserved"
    If xmlOrdb GetAddrConfig051 = conf0 Else GetAddrConfig051 = conf1
End Function

Public Function GetAddrTypes051 () As Boolean()
    Dim conf (33) As Boolean
    Dim i As Integer;
    'Все строки
    For i = 0 To 33
        conf (i) = true
    Next i
    'Исключая regi_id
    conf (6) = false
    'Исключая Reserved
    conf (33) = false
    GetAddrTypes051 = conf
End Function