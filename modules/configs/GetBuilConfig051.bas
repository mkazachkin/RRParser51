Option Compare Database
Public Function GetBuilConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(26) As Variant
    Dim conf1(26) As Variant
    'XML теги                                         'Поля в БД
    conf0(0) = "CadastralNumber":                   conf1(0) = "CadastralNumber"
    conf0(1) = "DateCreated":                       conf1(1) = "DatesCreated"
    conf0(2) = "FoundationDate":                    conf1(2) = "FoundationDates"
    conf0(3) = "CadastralBlock":                    conf1(3) = "CadastralBlock"
    conf0(4) = "PreviouslyPosted":                  conf1(4) = "PreviouslyPosted"
    conf0(5) = "Name":                              conf1(5) = "Names"
    conf0(6) = "ObjectType":                        conf1(6) = "ObjectType"
    conf0(7) = "AssignationBuilding":               conf1(7) = "AssignationBuilding"
    conf0(8) = "Area":                              conf1(8) = "Area"
    conf0(9) = "ExploitationChar":                  conf1(9) = "YearBuilt"
    conf0(10) = "":                                 conf1(10) = "YearUsed"
    conf0(11) = "Floors":                           conf1(11) = "Floors"
    conf0(12) = "":                                 conf1(12) = "UndergroundFloors"
    conf0(13) = "ElementsConstruct":                conf1(13) = "WallsCode"
    conf0(14) = "ParentCadastralNumbers":           conf1(14) = "ParentCadastralNumbers"
    conf0(15) = "PrevCadastralNumbers":             conf1(15) = "PrevCadastralNumbers"
    conf0(16) = "FlatsCadastralNumbers":            conf1(16) = "FlatsCadastralNumbers"
    conf0(17) = "CarParkingSpacesCadastralNumbers": conf1(17) = "CarParkingSpacesCadastralNumbers"
    conf0(18) = "UnitedCadastralNumber":            conf1(18) = "UnitedCadastralNumber"
    conf0(19) = "Location":                         conf1(19) = "addr_id"
    conf0(20) = "ObjectPermittedUses":              conf1(20) = "ObjectPermittedUses"
    conf0(21) = "CadastralCost":                    conf1(21) = "CadastralCost"
    conf0(22) = "SubBuildings":                     conf1(22) = "SubBuildings"
    conf0(23) = "FacilityCadastralNumber":          conf1(23) = "FacilityCadastralNumber"
    conf0(24) = "CulturalHeritage":                 conf1(24) = "CulturalHeritage"
    conf0(25) = "":                                 conf1(25) = ""
    conf0(26) = "":                                 conf1(26) = "Reserved"
    If xmlOrdb Then GetBuilConfig051 = conf0 Else GetBuilConfig051 = conf1
End Function
Public Function GetBuilTypes051() As Variant
    Dim conf(26) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 26
        conf(i) = True
    Next i
    'Исключая addr_id
    conf(19) = False
    'Исключая id
    conf(25) = False
    'Исключая Reserved
    conf(26) = False
    GetBuilTypes051 = conf
End Function
