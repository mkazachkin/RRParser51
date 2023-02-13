Option Compare Database
Public Function GetSubcConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(7) As Variant
    Dim conf1(7) As Variant
    'XML теги                     'Поля в БД
    conf0(0) = "NumberRecord": conf1(0) = "NumberRecord"
    conf0(1) = "DateCreated": conf1(1) = "DatesCreated"
    conf0(2) = "KeyParameter": conf1(2) = "Types"
    conf0(3) = "": conf1(3) = "Values"
    conf0(4) = "Encumbrances": conf1(4) = "Encumbrances"
    conf0(5) = "": conf1(5) = ""
    conf0(6) = "": conf1(6) = "CadastralNumber"
    conf0(7) = "": conf1(7) = "Reserved"
    If xmlOrdb Then GetSubcConfig051 = conf0 Else GetSubcConfig051 = conf1
End Function
Public Function GetSubcTypes051() As Variant
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
    GetSubcTypes051 = conf
End Function
