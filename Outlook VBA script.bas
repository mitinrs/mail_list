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

    ' ������� ���������� ����������
    Set ExcelApp = CreateObject("Excel.Application")
    Set WordApp = CreateObject("Word.Application")
    Set OutApp = CreateObject("Outlook.Application")

    ' ������ ������������ ������� ���� Excel � �������� � ������ Word
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

    ' ��������� Excel � ������ ������
    Set ExcelSheet = ExcelApp.Workbooks.Open(FilePath).Sheets(1)
    LastRow = ExcelSheet.Cells(ExcelSheet.Rows.Count, "A").End(-4162).Row

    For i = 2 To LastRow
        ' ��������� �������� Word
        Set WordDoc = WordApp.Documents.Open(WordTemplatePath, ReadOnly:=True)

        ' �������� ���������� Word-��������� � ����� ������
        WordDoc.Content.Copy
        WordDoc.Close False
        Set WordDoc = Nothing

        ' ������� ����� ������ � Outlook
        Set OutMail = OutApp.CreateItem(0)
        
        With OutMail
            .Display
            ' ��������� ���������� �� ������ ������
            .GetInspector.WordEditor.Content.PasteAndFormat (wdFormatOriginalFormatting)
    
            ' �������� ����������� �� ���������
            With .GetInspector.WordEditor.Content.Find
                .Text = "[���]"
                .Replacement.Text = ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "���������")).Value
                .Wrap = 1
                .Execute Replace:=2
            End With
    
            ' ��������� ����������, ���� � ��������
            .To = ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "Email ���������")).Value
            .Subject = ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "���� ������")).Value
            .Attachments.Add ExcelSheet.Cells(i, FindColumnByName(ExcelSheet, "���� � ����� ��������")).Value
            .Save
            .Close olDiscard
        End With
        Set OutMail = Nothing
    Next i

    ' ��������� Excel
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelApp = Nothing

    ' ��������� Outlook (�������������, ���� Outlook ��� ������)
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

