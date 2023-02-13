Option Compare Database
Public Function GetPstnConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(6) As Variant
    Dim conf1(6) As Variant
    'XML теги         'Поля в БД
    conf0(0) = "": conf1(0) = "Types"
    conf0(1) = "": conf1(1) = "Numbers"
    conf0(2) = "": conf1(2) = "NumberOnPlan"
    conf0(3) = "": conf1(3) = "Description"
    conf0(4) = "": conf1(4) = ""
    conf0(5) = "": conf1(5) = "CadastralNumber"
    conf0(6) = "": conf1(6) = "Reserved"
    If xmlOrdb Then GetPstnConfig051 = conf0 Else GetPstnConfig051 = conf1
End Function
Public Function GetPstnTypes051() As Variant
    Dim conf(6) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 6
        conf(i) = True
    Next i
    'Исключая id
    conf(4) = False
    'Исключая Reserved
    conf(6) = False
    GetPstnTypes051 = conf
End Function
