Option Compare Database
Public Function GetConsConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(25) As Variant
    Dim conf1(25) As Variant
    'XML теги                                         'Поля в БД
    conf0(0) = "CadastralNumber":                       conf1(0) = "CadastralNumber"
    conf0(1) = "DateCreated":                           conf1(1) = "DatesCreated"
    conf0(2) = "FoundationDate":                        conf1(2) = "FoundationDates"
    conf0(3) = "CadastralBlock":                        conf1(3) = "CadastralBlock"
    conf0(4) = "PreviouslyPosted":                      conf1(4) = "PreviouslyPosted"
    conf0(5) = "Name":                                  conf1(5) = "Names"
    conf0(6) = "ObjectType":                            conf1(6) = "ObjectType"
    conf0(7) = "AssignationName":                       conf1(7) = "AssignationNames"
    conf0(8) = "ExploitationChar":                      conf1(8) = "YearBuilt"
    conf0(9) = "":                                      conf1(9) = "YearUsed"
    conf0(10) = "Floors":                               conf1(10) = "Floors"
    conf0(11) = "":                                     conf1(11) = "UndergroundFloors"
    conf0(12) = "KeyParameters":                        conf1(12) = "KeyParameters"
    conf0(13) = "ParentCadastralNumbers":               conf1(13) = "ParentCadastralNumbers"
    conf0(14) = "PrevCadastralNumbers":                 conf1(14) = "PrevCadastralNumbers"
    conf0(15) = "FlatsCadastralNumbers":                conf1(15) = "FlatsCadastralNumbers"
    conf0(16) = "CarParkingSpacesCadastralNumbers":     conf1(16) = "CarParkingSpacesCadastralNumbers"
    conf0(17) = "UnitedCadastralNumber":                conf1(17) = "UnitedCadastralNumber"
    conf0(18) = "ObjectPermittedUses":                  conf1(18) = "ObjectPermittedUses"
    conf0(19) = "CadastralCost":                        conf1(19) = "CadastralCost"
    conf0(20) = "SubConstructions":                     conf1(20) = "SubConstructions"
    conf0(21) = "FacilityCadastralNumber":              conf1(21) = "FacilityCadastralNumber"
    conf0(22) = "CulturalHeritage":                     conf1(22) = "CulturalHeritage"
    conf0(23) = "Location":                             conf1(23) = "addr_id"
    conf0(24) = "":                                     conf1(24) = ""
    conf0(25) = "":                                     conf1(25) = "Reserved"
    If xmlOrdb Then GetConsConfig051 = conf0 Else GetConsConfig051 = conf1
End Function
Public Function GetConsTypes051() As Variant
    Dim conf(25) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 25
        conf(i) = True
    Next i
    'Исключая addr_id
    conf(23) = False
    'Исключая id
    conf(24) = False
    'Исключая Reserved
    conf(25) = False
    GetConsTypes051 = conf
End Function
