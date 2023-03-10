Option Compare Database
Public Function GetEnbrConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(7) As Variant
    Dim conf1(7) As Variant
    'XML теги                     'Поля в БД
    conf0(0) = "Name":          conf1(0) = "Names"
    conf0(1) = "Type":          conf1(1) = "Type"
    conf0(2) = "Registration":  conf1(2) = "RightNumber"
    conf0(3) = "":              conf1(3) = "RegistrationDates"
    conf0(4) = "Document":      conf1(4) = "Document"
    conf0(5) = "":              conf1(5) = ""
    conf0(6) = "":              conf1(6) = "CadastralNumber"
    conf0(7) = "":              conf1(7) = "Reserved"
    If xmlOrdb Then GetEnbrConfig051 = conf0 Else GetEnbrConfig051 = conf1
End Function
Public Function GetEnbrTypes051() As Variant
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
    GetEnbrTypes051 = conf
End Function
