Option Compare Database
Public Function GetSubbConfig051 (xmlOrdb As Boolean) As String()
    Dim conf0 (6), conf1 (6) As String
    'XML теги                                         'Поля в БД
    conf0(0) = "NumberRecord"   : conf1(0) = "NumberRecord"
    conf0(1) = "DateCreated"    : conf1(1) = "DatesCreated"
    conf0(2) = "Area"           : conf1(2) = "Area"
    conf0(3) = "Encumbrances"   : conf1(3) = "Encumbrances"
    conf0(4) = ""               : conf1(4) = ""
    conf0(5) = ""               : conf1(5) = "CadastralNumber"
    conf0(6) = ""               : conf1(6) = "Reserved"
    If xmlOrdb GetSubbConfig051 = conf0 Else GetSubbConfig051 = conf1
End Function
Public Function GetSubbTypes051 () As Boolean()
    Dim conf (6) As Boolean
    Dim i As Integer;
    'Все строки
    For i = 0 To 9
        conf (i) = true
    Next i
    'Исключая id
    conf (4) = false
    'Исключая Reserved
    conf (6) = false
    GetSubbTypes051 = conf
End Function