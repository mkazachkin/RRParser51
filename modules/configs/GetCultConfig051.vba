Option Compare Database
Public Function GetCultConfig051(xmlOrdb As Boolean) As Variant
    Dim conf0(10) As Variant
    Dim conf1(10) As Variant
    'XML теги                         'Поля в БД
    conf0(0) = "InclusionEGROKN": conf1(0) = "EGROKNRegNum"
    conf0(1) = "": conf1(1) = "EGROKNObjCultural"
    conf0(2) = "": conf1(2) = "EGROKNNameCultural"
    conf0(3) = "AssignmentEGROKN": conf1(3) = "AssignEGROKNRegNum"
    conf0(4) = "": conf1(4) = "AssignEGROKNObjCultural"
    conf0(5) = "": conf1(5) = "AssignAssignEGROKNRegNum"
    conf0(6) = "RequirementsEnsure": conf1(6) = "RequirementsEnsure"
    conf0(7) = "Document": conf1(7) = "Document"
    conf0(8) = "": conf1(8) = ""
    conf0(9) = "": conf1(9) = "CadastralNumber"
    conf0(10) = "": conf1(10) = "Reserved"
    If xmlOrdb Then GetCultConfig051 = conf0 Else GetCultConfig051 = conf1
End Function
Public Function GetCultTypes051() As Variant
    Dim conf(10) As Variant
    Dim i As Integer
    'Все строки
    For i = 0 To 10
        conf(i) = True
    Next i
    'Исключая id
    conf(8) = False
    'Исключая Reserved
    conf(10) = False
    GetCultTypes051 = conf
End Function
