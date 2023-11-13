Attribute VB_Name = "Module1"
Sub SendMassEmails()
    Dim ExcelApp As Object
    Dim ExcelSheet As Object
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim OutApp As Object
    Dim OutMail As Object
    Dim FilePath As String
    Dim WordTemplatePath As String
    Dim i As Integer
    Dim LastRow As Integer

    ' Создаем экземпляры приложений
    Set ExcelApp = CreateObject("Excel.Application")
    Set WordApp = CreateObject("Word.Application")
    Set OutApp = CreateObject("Outlook.Application")

    ' Просим пользователя выбрать файл Excel с адресами и шаблон Word
    With ExcelApp.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Excel file"
        .Filters.Add "Excel Files", "*.xls; *.xlsx"
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        Else
            MsgBox "You did not select an Excel file. Exiting..."
            Exit Sub
        End If
    End With

    With WordApp.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Word template"
        .Filters.Add "Word Files", "*.doc; *.docx"
        If .Show = -1 Then
            WordTemplatePath = .SelectedItems(1)
        Else
            MsgBox "You did not select a Word template. Exiting..."
            Exit Sub
        End If
    End With

    ' Открываем Excel и читаем данные
    Set ExcelSheet = ExcelApp.Workbooks.Open(FilePath).Sheets(1)
    LastRow = ExcelSheet.Cells(ExcelSheet.Rows.Count, "A").End(-4162).Row

    For i = 2 To LastRow
        ' Открываем документ Word
        Set WordDoc = WordApp.Documents.Open(WordTemplatePath, ReadOnly:=True)

        ' Копируем содержимое Word-документа в буфер обмена
        WordDoc.Content.Copy
        WordDoc.Close False
        Set WordDoc = Nothing

        ' Создаем новое письмо в Outlook
        Set OutMail = OutApp.CreateItem(0)
        
        With OutMail
            .Display
            ' Вставляем содержимое из буфера обмена
            .GetInspector.WordEditor.Content.PasteAndFormat (wdFormatOriginalFormatting)
    
            ' Заменяем плейсхолдер на обращение
            With .GetInspector.WordEditor.Content.Find
                .Text = "[Имя]"
                .Replacement.Text = ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "Обращение")).Value
                .Wrap = 1
                .Execute Replace:=2
            End With
    
            ' Добавляем получателя, тему и вложения
            .To = ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "Email адресатов")).Value
            .Subject = ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "Тема письма")).Value
            .Attachments.Add ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "Путь к файлу вложению")).Value
            .Save
            .Close olDiscard
        End With
        Set OutMail = Nothing
    Next i

    ' Закрываем Excel
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelApp = Nothing

    ' Закрываем Outlook (необязательно, если Outlook уже открыт)
    ' Set OutApp = Nothing
End Sub

Function FindColumnByName(sheet As Object, columnName As String) As Integer
    Dim col As Integer
    col = 1
    While sheet.Cells(1, col).Value <> ""
        If StrComp(sheet.Cells(1, col).Value, columnName, vbTextCompare) = 0 Then
            FindColumnByName = col
            Exit Function
        End If
        col = col + 1
    Wend
    FindColumnByName = 0
End Function

