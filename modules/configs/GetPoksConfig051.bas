Option Compare Database
Public Function GetPoksConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(11) As Variant
    Dim conf1(11) As Variant
    'XML теги                     'Поля в БД
    conf0(0) = "CadastralNumberOKS":        conf1(0) = "CadastralNumberOKS"
    conf0(1) = "ObjectType":                conf1(1) = "ObjectType"
    conf0(2) = "AssignationBuilding":       conf1(2) = "AssignationBuilding"
    conf0(3) = "AssignationName":           conf1(3) = "AssignationNames"
    conf0(4) = "ElementsConstruct":         conf1(4) = "WallsCode"
    conf0(5) = "ExploitationChar":          conf1(5) = "YearBuilt"
    conf0(6) = "ExploitationChar":          conf1(6) = "YearUsed"
    conf0(7) = "Floors":                    conf1(7) = "Floors"
    conf0(8) = "Floors":                    conf1(8) = "UndergroundFloors"
    conf0(9) = "":                          conf1(9) = ""
    conf0(10) = "":                         conf1(10) = "CadastralNumber"
    conf0(11) = "":                         conf1(11) = "Reserved"
    If xmlOrdb Then GetPoksConfig051 = conf0 Else GetPoksConfig051 = conf1
End Function
Public Function GetPoksTypes051() As Variant
    Dim conf(11) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 11
        conf(i) = True
    Next i
    'Исключая id
    conf(9) = False
    'Исключая Reserved
    conf(11) = False
    GetPoksTypes051 = conf
End Function
