Option Compare Database
Public Function GetFlatConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(23) As Variant
    Dim conf1(23) As Variant
    'XML теги                                         'Поля в БД
    conf0(0) = "CadastralNumber": conf1(0) = "CadastralNumber"
    conf0(1) = "DateCreated": conf1(1) = "DatesCreated"
    conf0(2) = "FoundationDate": conf1(2) = "FoundationDates"
    conf0(3) = "CadastralBlock": conf1(3) = "CadastralBlock"
    conf0(4) = "PreviouslyPosted": conf1(4) = "PreviouslyPosted"
    conf0(5) = "Name": conf1(5) = "Names"
    conf0(6) = "ObjectType": conf1(6) = "ObjectType"
    conf0(7) = "CadastralNumberFlat": conf1(7) = "CadastralNumberFlat"
    conf0(8) = "CadastralNumberOKS": conf1(8) = "CadastralNumberOKS"
    conf0(9) = "Area": conf1(9) = "Area"
    conf0(10) = "ParentOKS": conf1(10) = "ParentOKS"
    conf0(11) = "PrevCadastralNumbers": conf1(11) = "PrevCadastralNumbers"
    conf0(12) = "Assignation": conf1(12) = "Assignation"
    conf0(13) = "PositionInObject": conf1(13) = "PositionInObject"
    conf0(14) = "ObjectPermittedUses": conf1(14) = "ObjectPermittedUses"
    conf0(15) = "CadastralCost": conf1(15) = "CadastralCost"
    conf0(16) = "SubFlats": conf1(16) = "SubFlats"
    conf0(17) = "UnitedCadastralNumber": conf1(17) = "UnitedCadastralNumber"
    conf0(18) = "FacilityCadastralNumber": conf1(18) = "FacilityCadastralNumber"
    conf0(19) = "CulturalHeritage": conf1(19) = "CulturalHeritage"
    conf0(20) = "Location": conf1(20) = "addr_id"
    conf0(21) = "RoomNumber": conf1(21) = "RoomNumber"
    conf0(22) = "": conf1(22) = ""
    conf0(23) = "": conf1(23) = "Reserved"
    If xmlOrdb Then GetFlatConfig051 = conf0 Else GetFlatConfig051 = conf1
End Function
Public Function GetFlatTypes051() As Variant
    Dim conf(23) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 23
        conf(i) = True
    Next i
    'Исключая addr_id
    conf(20) = False
    'Исключая id
    conf(22) = False
    'Исключая Reserved
    conf(23) = False
    GetFlatTypes051 = conf
End Function
