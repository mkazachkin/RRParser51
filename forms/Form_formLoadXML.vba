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

    'Служебные переменные
    Set cadastrDB = CurrentDb
    Dim insertSQL As String
    Dim selectSQL As String
    Dim deleteSQL As String
    Dim xmlf_id As Long
    Dim rs As DAO.Recordset

    Dim importXMLFilesTable(11) As String
    importXMLFilesTable(0) = "listguid"
    Dim listguid As String
    importXMLFilesTable(1) = "xmlversion"
    Dim xmlversion As String
    importXMLFilesTable(2) = "actnumber"
    'Прилетает снаружи actNumber
    importXMLFilesTable(3) = "reassessmentcode"
    'Прилетает снаружи reassessmentcode
    importXMLFilesTable(4) = "transmittalletter"
    'Прилетает снаружи transmittalletter
    importXMLFilesTable(5) = "importdate"
    Dim importdate As String
    importXMLFilesTable(6) = "formdate"
    Dim formdate As String
    formdate = Format(Date, "YYYY-mm-dd")
    importXMLFilesTable(7) = "list_id"
    Dim list_id As Long
    importXMLFilesTable(8) = "regi_id"
    Dim regi_id As Long
    importXMLFilesTable(9) = "real_id"
    Dim real_id As Long
    importXMLFilesTable(10) = "ctgr_id"
    Dim ctgr_id As Long
    'Получаем имя пользователя, который импортирует файлы
    importXMLFilesTable(11) = "username"
    Dim importXMLFilesUserName As String
    importXMLFilesUserName = LoginUserName()

    Const importXMLTable = "public_import_t_xml_files"                    'Название таблицы со списком импортированных XML

    'Таблицы словарей
    Dim dictRealty(2) As String
    dictRealty(0) = "RealtyCode"
    dictRealty(1) = "public_import_dict_realty"
    dictRealty(2) = "real_id"
    
    Dim dictList(2) As String
    dictList(0) = "ListCode"
    dictList(1) = "public_import_dict_list"
    dictList(2) = "list_id"
    
    Dim dictRegi(2) As String
    dictRegi(0) = "RegionCode"
    dictRegi(1) = "public_import_dict_region"
    dictRegi(2) = "regi_id"

    Dim dictCtgr(2) As String
    dictCtgr(0) = "CategoryCode"
    dictCtgr(1) = "public_import_dict_categories"
    dictCtgr(2) = "ctgr_id"


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
    selectSQL = "select xmlf_id, " & importXMLFilesTable(0) & " from " & importXMLTable & " where listguid='" & listguid & "';"
    Set rs = cadastrDB.OpenRecordset(selectSQL)
    If rs.RecordCount > 0 Then
        'Данные с GUID из файла уже были добавлены в БД. Либо их удаляем, либо выходим
        If MsgBox("Данные из файла с именем " & importFilePath & " уже были добавлены в базу данных. Заменить их?", vbYesNo) = vbYes Then
            'Удаляем данные с этим GUID из всех таблиц
            'Получаем ключ, по которому осуществлялась привязка к XML-файлу
            'Передернем recordset на всякий пожарный случай
            rs.MoveLast
            rs.MoveFirst
            For counter = 1 To rs.RecordCount
                'Запись, вообще говоря, должна быть одна, но черт его знает
                'Удаляем из списка импортированных файлов
                deleteSQL = "delete from " & importXMLTable & " where xmlf_id = " & rs.Fields.Item(0).Value & ";"
                cadastrDB.Execute deleteSQL
                rs.MoveNext
            Next counter
        Else
            'Данные не трогаем, файл пропускаем и выходим из функции
            ParseXML_051 = False
            Exit Function
        End If
    End If
    
    'Заполняем таблицу импортированных XML
    insertSQL = "insert into " & importXMLTable & " ("
    For counter = 0 To 10
        insertSQL = insertSQL & importXMLFilesTable(counter) & ","
    Next counter
    insertSQL = insertSQL & importXMLFilesTable(11) & ") values ('"
    
    insertSQL = insertSQL & listguid & "', '"
    insertSQL = insertSQL & xmlversion & "', '"
    insertSQL = insertSQL & actNumber & "', '"
    insertSQL = insertSQL & reassessmentcode & "', '"
    insertSQL = insertSQL & transmittalletter & "', '"
    insertSQL = insertSQL & importdate & "', '"
    insertSQL = insertSQL & formdate & "', "
    insertSQL = insertSQL & list_id & ", "
    insertSQL = insertSQL & regi_id & ", "
    insertSQL = insertSQL & real_id & ", "
    insertSQL = insertSQL & ctgr_id & ", '"
    insertSQL = insertSQL & importXMLFilesUserName & "');"
    
    cadastrDB.Execute insertSQL
    'Получаем id в таблице импортируемых XML
    selectSQL = "select xmlf_id from " & importXMLTable & " where listguid='" & listguid & "';"
    Set rs = cadastrDB.OpenRecordset(selectSQL)
    If rs.RecordCount = 1 Then
        'Данные с GUID из файла уже были добавлены в БД. Либо их удаляем, либо выходим
        'Передернем recordset на всякий пожарный случай
        rs.MoveLast
        rs.MoveFirst
        xmlf_id = rs.Fields.Item(0).Value
    Else
        'В смысле нет GUID? Выводим сообщение об ошибке
        MsgBox ("Ошибка. Не смог добавить информацию о XML файле в таблицу")
        ParseXML_051 = False
        Exit Function
    End If

    Set rootNodeChild = rootNodeChild.NextSibling
    'Переходим к парсингу объектов Objects

    Set objectsNode = rootNodeChild.FirstChild
    While (Not objectsNode Is Nothing)
        If objectsNode.NodeName = "Buildings" Then
            'Парсим здания
            Set buildingsNode = objectsNode.FirstChild
            While (Not buildingsNode Is Nothing)
                'Парсим каждый объект
                a = ParsXMLBuil051("public_import_t_buildings", "xmlf_id", xmlf_id, buildingsNode)
                Set buildingsNode = buildingsNode.NextSibling
            Wend
        End If
        If objectsNode.NodeName = "Constructions" Then
            'Парсим здания
            Set constructionsNode = objectsNode.FirstChild
            While (Not constructionsNode Is Nothing)
                'Парсим каждый объект
                a = ParsXMLCons051("public_import_t_constructions", "xmlf_id", xmlf_id, constructionsNode)
                Set constructionsNode = constructionsNode.NextSibling
            Wend
        End If
        'Парсим дальше объекты
        Set objectsNode = objectsNode.NextSibling
    Wend
    cadastrDB.Close
    ParseXML_051 = True
End Function
