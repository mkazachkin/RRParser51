Option Compare Database
Public Function GetAsgnConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(7) As Variant
    Dim conf1(7) As Variant
    'XML теги                         'Поля в БД
    conf0(0) = "AssignationCode":   conf1(0) = "AssignationCode"
    conf0(1) = "AssignationType":   conf1(1) = "AssignationType"
    conf0(2) = "SpecialType":       conf1(2) = "SpecialType"
    conf0(3) = "TotalAssets":       conf1(3) = "TotalAssets"
    conf0(4) = "AuxiliaryFlat":     conf1(4) = "AuxiliaryFlat"
    conf0(5) = "":                  conf1(5) = ""
    conf0(6) = "":                  conf1(6) = "CadastralNumber"
    conf0(7) = "":                  conf1(7) = "Reserved"
    If xmlOrdb Then GetAsgnConfig051 = conf0 Else GetAsgnConfig051 = conf1
End Function
Public Function GetAsgnTypes051() As Variant
    Dim conf(7) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 7
        conf(i) = True
    Next i
    'Исключая id
    conf(5) = False
    'Исключая Reserved
    conf(7) = False
    GetAsgnTypes051 = conf
End Function
