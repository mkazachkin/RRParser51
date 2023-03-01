Option Compare Database
Public Function GetSubbConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(6) As Variant
    Dim conf1(6) As Variant
    'XML теги                     'Поля в БД
    conf0(0) = "NumberRecord":  conf1(0) = "NumberRecord"
    conf0(1) = "DateCreated":   conf1(1) = "DatesCreated"
    conf0(2) = "Area":          conf1(2) = "Area"
    conf0(3) = "Encumbrances":  conf1(3) = "Encumbrances"
    conf0(4) = "":              conf1(4) = ""
    conf0(5) = "":              conf1(5) = "CadastralNumber"
    conf0(6) = "":              conf1(6) = "Reserved"
    If xmlOrdb Then GetSubbConfig051 = conf0 Else GetSubbConfig051 = conf1
End Function
Public Function GetSubbTypes051() As Variant
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
    GetSubbTypes051 = conf
End Function
