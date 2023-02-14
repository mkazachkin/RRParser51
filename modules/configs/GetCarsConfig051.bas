Option Compare Database
Public Function GetCarsConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(16) As Variant
    Dim conf1(16) As Variant
    conf0(0) = "CadastralNumber":           conf1(0) = "CadastralNumber"
    conf0(1) = "DateCreated":               conf1(1) = "DatesCreated"
    conf0(2) = "FoundationDate":            conf1(2) = "FoundationDates"
    conf0(3) = "CadastralBlock":            conf1(3) = "CadastralBlock"
    conf0(4) = "PreviouslyPosted":          conf1(4) = "PreviouslyPosted"
    conf0(5) = "ObjectType":                conf1(5) = "ObjectType"
    conf0(6) = "CadastralNumberOKS":        conf1(6) = "CadastralNumberOKS"
    conf0(7) = "Area":                      conf1(7) = "Area"
    conf0(8) = "ParentOKS":                 conf1(8) = "ParentOKS"
    conf0(9) = "PrevCadastralNumbers":      conf1(9) = "PrevCadastralNumbers"
    conf0(10) = "PositionInObject":         conf1(10) = "PositionInObject"
    conf0(11) = "UnitedCadastralNumber":    conf1(11) = "UnitedCadastralNumber"
    conf0(12) = "Location":                 conf1(12) = "addr_id"
    conf0(13) = "CadastralCost":            conf1(13) = "CadastralCost"
    conf0(14) = "FacilityCadastralNumber":  conf1(14) = "FacilityCadastralNumber"
    conf0(15) = "":                         conf1(15) = ""
    conf0(16) = "":                         conf1(16) = "Reserved"
    If xmlOrdb Then GetCarsConfig051 = conf0 Else GetCarsConfig051 = conf1
End Function
Public Function GetCarsTypes051() As Variant
    Dim conf(16) As Variant
    Dim i As Integer
    For i = 0 To 16
        conf(i) = True
    Next i
    conf(12) = False
    conf(15) = False
    conf(16) = False
    GetCarsTypes051 = conf
End Function

