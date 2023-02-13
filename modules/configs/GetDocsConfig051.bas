Option Compare Database
Public Function GetDocsConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(9) As Variant
    Dim conf1(9) As Variant
    'XML теги                     'Поля в БД
    conf0(0) = "CodeDocument": conf1(0) = "CodeDocument"
    conf0(1) = "Name": conf1(1) = "Names"
    conf0(2) = "Series": conf1(2) = "Series"
    conf0(3) = "Number": conf1(3) = "Numbers"
    conf0(4) = "Date": conf1(4) = "Dates"
    conf0(5) = "IssueOrgan": conf1(5) = "IssueOrgan"
    conf0(6) = "Desc": conf1(6) = "Descr"
    conf0(7) = "": conf1(7) = ""
    conf0(8) = "": conf1(8) = "CadastralNumber"
    conf0(9) = "": conf1(9) = "Reserved"
    If xmlOrdb Then GetDocsConfig051 = conf0 Else GetDocsConfig051 = conf1
End Function
Public Function GetDocsTypes051() As Variant
    Dim conf(9) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 9
        conf(i) = True
    Next i
    'Исключая id
    conf(7) = False
    'Исключая Reserved
    conf(9) = False
    GetDocsTypes051 = conf
End Function
