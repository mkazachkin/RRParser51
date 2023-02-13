Option Compare Database
Public Function GetUcmpConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(15) As Variant
    Dim conf1(15) As Variant
    'XML теги                                           'Поля в БД
    conf0(0) = "CadastralNumber":                       conf1(0) = "CadastralNumber"
    conf0(1) = "DateCreated":                           conf1(1) = "DatesCreated"
    conf0(2) = "FoundationDate":                        conf1(2) = "FoundationDates"
    conf0(3) = "CadastralBlock":                        conf1(3) = "CadastralBlock"
    conf0(4) = "PreviouslyPosted":                      conf1(4) = "PreviouslyPosted"
    conf0(5) = "ObjectType":                            conf1(5) = "ObjectType"
    conf0(6) = "AssignationName":                       conf1(6) = "AssignationNames"
    conf0(7) = "DegreeReadiness":                       conf1(7) = "DegreeReadiness"
    conf0(8) = "KeyParameters":                         conf1(8) = "KeyParameters"
    conf0(9) = "ParentCadastralNumbers":                conf1(9) = "ParentCadastralNumbers"
    conf0(10) = "PrevCadastralNumbers":                 conf1(10) = "PrevCadastralNumbers"
    conf0(11) = "Location":                             conf1(11) = "addr_id"
    conf0(12) = "CadastralCost":                        conf1(12) = "CadastralCost"
    conf0(13) = "FacilityCadastralNumber":              conf1(13) = "FacilityCadastralNumber"
    conf0(14) = "":                                     conf1(14) = ""
    conf0(15) = "":                                     conf1(15) = "Reserved"
    If xmlOrdb Then GetUcmpConfig051 = conf0 Else GetUcmpConfig051 = conf1
End Function
Public Function GetUcmpTypes051() As Variant
    Dim conf(15) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 15
        conf(i) = True
    Next i
    'Исключая addr_id
    conf(11) = False
    'Исключая id
    conf(14) = False
    'Исключая Reserved
    conf(15) = False
    GetUcmpTypes051 = conf
End Function
