Option Compare Database
Public Function GetCostConfig051 (xmlOrdb As Boolean) As String()
    Dim conf0 (10), conf1 (10) As String
    'XML теги                                 'Поля в БД
    conf0(0) = "CadastralCost"              : conf1(0) = "CadastralCost"
    conf0(1) = "DateValuation"              : conf1(1) = "DatesValuation"
    conf0(2) = "DateEntering"               : conf1(2) = "DatesEntering"
    conf0(3) = "DateApproval"               : conf1(3) = "DatesApproval"
    conf0(4) = "ApplicationDate"            : conf1(4) = "ApplicationDates"
    conf0(5) = "RevisalStatementDate"       : conf1(5) = "RevisalStatementDates"
    conf0(6) = "ApplicationLastDate"        : conf1(6) = "ApplicationLastDates"
    conf0(7) = "ApprovalDocument"           : conf1(7) = "ApprovalDocument"
    conf0(8) = ""                           : conf1(8) = ""
    conf0(9) = ""                           : conf1(9) = "CadastralNumber"
    conf0(10) = ""                          : conf1(10) = "Reserved"
    If xmlOrdb GetCostConfig051 = conf0 Else GetCostConfig051 = conf1
End Function

Public Function GetCostTypes051 () As Boolean()
    Dim conf (10) As Boolean
    Dim i As Integer;
    'Все строки
    For i = 0 To 10
        conf (i) = true
    Next i
    'Исключая id
    conf (8) = false
    'Исключая Reserved
    conf (10) = false
    GetCostTypes051 = conf
End Function