Option Compare Database
'===================================
'===== Функции работы с формой =====
'===================================
Private Sub changeXMLFolderButton_Click()
    'Нажата кнопка выбора папки с XML
    'Инициируем форму выбора файла
    Dim dlgOpenFile As Object
    'Определяем форму диалога как выбор папки
    Set dlgOpenFile = Application.FileDialog(4) 'msoFileDialogFolderPicker
    With dlgOpenFile
        'Очищаем фильтры файлов
        .Filters.Clear
        'В качестве начального пути выбираем папку текущего проекта
        .InitialFileName = CurrentProject.Path
        'Разрешаем выбрать только одну папку
        .AllowMultiSelect = False
        'Задаем название диалога
        .Title = "Выбор папки"
        'Если папка выбрана, то в объект формы pathToXML возвращаем путь выбранной папки в строку Value
        If (.Show = -1) And (.SelectedItems.Count > 0) Then
            pathToXML.Value = .SelectedItems(1)
        End If
    End With
    Set dlgOpenFile = Nothing
End Sub
Private Sub loadXMLButton_Click()
    'Нажата кнопка Произвести загузку и распарсить XML
    'Инициализация переменных:
    Dim parseFile As String             'имя файла XML, который будем парсить
    Dim parseFilePath As String         'полный путь к XML файлу, который будем парсить
    Dim parseFilder As String           'путь к каталогу, где находятся XML-файлы
    Dim parseFilter As String           'фильтр файлов в каталоге с XML-файлами
    Dim parseActNumber As String        'Номер акта
    Dim parseReassCode As String        'Номер статьи оценки/дооценки
    Dim parseTransLetter As String      'Номер сопроводительного письма
    'Получаем значения из формы
    parseActNumber = actNumber.Value
    parseReassCode = reassCode.Value
    parseTransLetter = transLetter.Value
    'Проверяем введенные данные
    'Проверяем следующие поля формы:
        'actNumber должно содержать номер акта для дооценке
        'reassCode должно содержать номер статьи, по которой производится дооценка
        'transLetter должно содержать номер сопроводительной статьи
    If parseActNumber = "" Then
        MsgBox "Ошибка. Не введен номер акта дооценки."
        Exit Sub
    End If
    If parseReassCode = "" Then
        MsgBox "Ошибка. Не введен номер статьи, по которой производится дооценка."
        Exit Sub
    End If
    If parseTransLetter = "" Then
        MsgBox "Ошибка. Не указан номер сопроводительного письма."
        Exit Sub
    End If
    
    If Me.pathToXML = "" Then
        MsgBox "Ошибка. Не указан каталог с XML файлами"
        Exit Sub
    End If
    'Все проверки пройдены успешно
    'В обработку пойдут все xml файлы, расположенные в указанном каталоге
    'В качестве пути используем выбранный каталог в форме
    parseFilder = pathToXML.Value + "\"
    'В обработку пускаем все XML-файлы в каталоге
    parseFilter = parseFilder & "*.xml"
    'Получаем первый файл в списке для обработки
    parseFile = Dir(parseFilter)
    'Пускаем в обработку каждый XML файл
    While parseFile <> ""
        'Получаем полный путь к XML-файлу
        parseFilePath = parseFilder & parseFile
        'Парсим XML-файл собственной функцией
        If ParseXML_051(parseFilePath, parseActNumber, parseReassCode, parseTransLetter) Then
            'MsgBox ("Данные из файла " & parseFile & " импортированы успешно.")
        End If
        'Переходим к следующему XML файлу
        parseFile = Dir
    Wend
    MsgBox ("Данные импортированы успешно.")
End Sub
Private Function ParseXML_051(ByVal importFilePath As String, Optional ByVal actNumber As String, Optional ByVal reassessmentcode As String, Optional ByVal transmittalletter As String) As Boolean
    'Парсер XML выходных данных Росреестра версии 5.1
    'Входные параметры:
        'importFilePath - путь к XML файлам, которые надо парсить
        'importActNumber - номер акта, по которому поступили данные
        'importReassessmentCode - номер статьи, по которой проводится дооценка
        'importTransmittalLetter- номер сопроводительного письма к данным

    'Таблица импортированных XML
    Const importXMLTable = "public_import_t_xml_files"
    Dim importXMLFilesTable(11) As String
    importXMLFilesTable(0) = "listguid"
    importXMLFilesTable(1) = "xmlversion"
    importXMLFilesTable(2) = "actnumber"
    importXMLFilesTable(3) = "reassessmentcode"
    importXMLFilesTable(4) = "transmittalletter"
    importXMLFilesTable(5) = "importdate"
    importXMLFilesTable(6) = "formdate"
    importXMLFilesTable(7) = "list_id"
    importXMLFilesTable(8) = "regi_id"
    importXMLFilesTable(9) = "real_id"
    importXMLFilesTable(10) = "ctgr_id"
    importXMLFilesTable(11) = "username"
    'Таблицы словарей
    Dim dictRealty(2) As String
    dictRealty(0) = "RealtyCode"
    dictRealty(1) = "public_dict_realty"
    dictRealty(2) = "real_id"
    Dim dictList(2) As String
    dictList(0) = "ListCode"
    dictList(1) = "public_dict_list"
    dictList(2) = "list_id"
    Dim dictRegi(2) As String
    dictRegi(0) = "RegionCode"
    dictRegi(1) = "public_dict_region"
    dictRegi(2) = "regi_id"
    Dim dictCtgr(2) As String
    dictCtgr(0) = "CategoryCode"
    dictCtgr(1) = "public_dict_categories"
    dictCtgr(2) = "ctgr_id"
    'Служебные переменные
    Dim sqlStr As String
    Dim objType As String
    Dim xmlf_id As Long
    Dim rs As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim tmp As Boolean
    'Переменные для хранения полученных данных XML
    Dim listguid As String
    Dim xmlversion As String
    Dim importdate As String
    Dim list_id As Long
    Dim regi_id As Long
    Dim real_id As Long
    Dim ctgr_id As Long
    Dim formdate As String
    Dim importXMLFilesUserName As String
    
    Dim updStr(6000) As String
    
    'Дату сразу поставим текущую
    formdate = Format(Now, "YYYY-mm-dd")
    'Получаем имя пользователя, который импортирует файлы
    importXMLFilesUserName = LoginUserName()

    'Создаем OLE-объект и отключаем асинхронную загрузку
    Set xmlFile = CreateObject("Msxml2.DOMDocument")
    xmlFile.async = False
    'Открываем XML файл, чтобы его распарсить
    xmlFile.Load importFilePath
    'Получаем корневой элемент. Это должен быть ListForRating с версией 5.1
    Set rootNode = xmlFile.DocumentElement
    'Получаем название корневого элемента и версию XML-файла
    xmlversion = rootNode.getAttribute("Version")
    'Если корневой элемент не соответствует типу обрабатываемого XML файла, то выходим
    If rootNode.NodeName <> "ListForRating" Then
        MsgBox ("Ошибка. Не могу обработать XML-файл. Формат не соответствует заданному")
        Exit Function
    End If
    'Считываем информацию из ListInfo
    Set rootNodeChild = rootNode.FirstChild
    'Оформляем дату списка создания списка
    importdate = rootNodeChild.getAttribute("DateForm")                                       'Обязательный элемент
    'Получаем идентификатор списка
    listguid = rootNodeChild.getAttribute("GUID")                                             'Обязательный элемент
    'Парсим потомков ListInfo
    Set listInfoChild = rootNodeChild.FirstChild
    While (Not listInfoChild Is Nothing)
        If (listInfoChild.NodeName = "ListType") Then
            'Получаем вид перечня
            list_id = DictCheck(listInfoChild.Text, dictList(0), dictList(1), dictList(2))
        End If
        If (listInfoChild.NodeName = "Region") Then
            'Получаем код региона
            regi_id = DictCheck(listInfoChild.Text, dictRegi(0), dictRegi(1), dictRegi(2))
        End If
        If (listInfoChild.NodeName = "ObjectsType") Then
            'Парсим выясняем вид объектов недвижимости
            real_id = DictCheck(listInfoChild.FirstChild.Text, dictRealty(0), dictRealty(1), dictRealty(2))
            objType = listInfoChild.FirstChild.Text
        End If
        If (listInfoChild.NodeName = "Categories") Then
            'Парсим категории земель
            ctgr_id = DictCheck(listInfoChild.FirstChild.Text, dictCtgr(0), dictCtgr(1), dictCtgr(2))
        End If
        Set listInfoChild = listInfoChild.NextSibling
    Wend
    
    'Данные по импорту сформированы. Проверяем, был ли переданный на загрузку список ранее загружен
    'Если загрузка ранее не производилась, то заполняем таблицу импортированных списков
    
    'Проверяем были ли загружены данные из указанного XML в БД
    'xmlf_id - ключ в БД к записям.
    Set cadastrDB = CurrentDb
    sqlStr = "select xmlf_id, " & importXMLFilesTable(0) & " from " & importXMLTable & " where listguid='" & listguid & "';"
    Set rs = cadastrDB.OpenRecordset(sqlStr)
    If rs.RecordCount > 0 Then
        'Данные с GUID из файла уже были добавлены в БД. Либо их удаляем, либо выходим
        If MsgBox("Данные из файла с именем " & importFilePath & " уже были добавлены в базу данных. Заменить их?", vbYesNo) = vbYes Then
            'Удаляем данные с этим GUID из всех таблиц
            rs.MoveFirst
            Do While Not rs.EOF
                xmlf_id = rs.Fields("xmlf_id").Value
                tmp = ClearObjTable(objType, xmlf_id)
                'Других данных в наших таблицах быть не должно
                sqlStr = "delete from " & importXMLTable & " where xmlf_id = " & rs.Fields("xmlf_id").Value & ";"
                cadastrDB.Execute sqlStr
                rs.MoveNext
            Loop
        Else
            'Данные не трогаем, файл пропускаем и выходим из функции
            ParseXML_051 = False
            Exit Function
        End If
    End If
    
    'Заполняем таблицу импортированных XML
    sqlStr = "insert into " & importXMLTable & " ("
    For i = 0 To 10
        sqlStr = sqlStr & importXMLFilesTable(i) & ","
    Next i
    sqlStr = sqlStr & importXMLFilesTable(11) & ") values ('"
    sqlStr = sqlStr & listguid & "', '"
    sqlStr = sqlStr & xmlversion & "', '"
    sqlStr = sqlStr & actNumber & "', '"
    sqlStr = sqlStr & reassessmentcode & "', '"
    sqlStr = sqlStr & transmittalletter & "', '"
    sqlStr = sqlStr & importdate & "', '"
    sqlStr = sqlStr & formdate & "', "
    sqlStr = sqlStr & list_id & ", "
    sqlStr = sqlStr & regi_id & ", "
    sqlStr = sqlStr & real_id & ", "
    sqlStr = sqlStr & ctgr_id & ", '"
    sqlStr = sqlStr & importXMLFilesUserName & "');"

    cadastrDB.Execute sqlStr
    'Получаем id в таблице импортируемых XML
    sqlStr = "select xmlf_id from " & importXMLTable & " where listguid='" & listguid & "';"
    Set rs = cadastrDB.OpenRecordset(sqlStr)
    If rs.RecordCount = 1 Then
        xmlf_id = rs.Fields("xmlf_id").Value
    Else
        'В смысле нет GUID? Выводим сообщение об ошибке
        MsgBox ("Ошибка. Не смог добавить информацию о XML файле в таблицу")
        ParseXML_051 = False
        Exit Function
    End If
    Set rs = Nothing
    Set cadastrDB = Nothing
    
    Set rootNodeChild = rootNodeChild.NextSibling
    'Переходим к парсингу объектов Objects

    Set objectsNode = rootNodeChild.FirstChild
    While (Not objectsNode Is Nothing)
        i = 0
        If objectsNode.NodeName = "Constructions" Then
            'Парсим сооружения
            Set constructionsNode = objectsNode.FirstChild
            While (Not constructionsNode Is Nothing)
                'Парсим каждый объект
                updStr(i) = ParsXMLCons051("public_import_t_constructions", "xmlf_id", xmlf_id, constructionsNode)
                i = i + 1
                Set constructionsNode = constructionsNode.NextSibling
            Wend
        End If
        If objectsNode.NodeName = "Flats" Then
            'Парсим помещения
            Set flatsNode = objectsNode.FirstChild
            While (Not flatsNode Is Nothing)
                'Парсим каждый объект
                updStr(i) = ParsXMLFlat051("public_import_t_flats", "xmlf_id", xmlf_id, flatsNode)
                i = i + 1
                Set flatsNode = flatsNode.NextSibling
            Wend
        End If
        If objectsNode.NodeName = "Buildings" Then
            'Парсим здания
            Set buildingsNode = objectsNode.FirstChild
            While (Not buildingsNode Is Nothing)
                'Парсим каждый объект
                updStr(i) = ParsXMLBuil051("public_import_t_buildings", "xmlf_id", xmlf_id, buildingsNode)
                i = i + 1
                Set buildingsNode = buildingsNode.NextSibling
            Wend
        End If
        If objectsNode.NodeName = "Uncompleteds" Then
            'Парсим недострой
            Set ucmpNode = objectsNode.FirstChild
            While (Not ucmpNode Is Nothing)
                'Парсим каждый объект
                updStr(i) = ParsXMLUcmp051("public_import_t_uncompleted", "xmlf_id", xmlf_id, ucmpNode)
                i = i + 1
                Set ucmpNode = ucmpNode.NextSibling
            Wend
        End If
        If objectsNode.NodeName = "CarParkingSpaces" Then
            'Парсим парковки
            Set carsNode = objectsNode.FirstChild
            While (Not carsNode Is Nothing)
                'Парсим каждый объект
                updStr(i) = ParsXMLCars051("public_import_t_cars", "xmlf_id", xmlf_id, carsNode)
                i = i + 1
                Set carsNode = carsNode.NextSibling
            Wend
        End If
        Set updDb = CurrentDb
        For j = 0 To i - 1
            updDb.Execute updStr(j)
        Next j
        Set updDb = Nothing
        'Парсим дальше объекты
        Set objectsNode = objectsNode.NextSibling
    Wend
    ParseXML_051 = True
End Function
