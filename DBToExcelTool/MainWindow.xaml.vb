Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports OfficeOpenXml

Class MainWindow

    Private Sub ClickConvert(sender As Object, e As RoutedEventArgs) Handles button.Click
        Try
            MySettings.Default.Save()
            button.IsEnabled = False

            Dim dataSource = input_source.Text
            Dim dbName = input_db.Text
            Dim username = input_username.Text
            Dim password = input_password.Text
            Dim fileName = input_file.Text
            Dim limit = Integer.Parse(input_limit.Text)

            If (File.Exists(fileName)) Then
                File.Delete(fileName)
            End If

            Dim Thread1 As New Thread(Sub() CreateExcelFile(dataSource, username, password, dbName, fileName, limit))
            Thread1.Start()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub CreateExcelFile(dataSource As String, userName As String, password As String, dbname As String, excelName As String, limit As Integer)
        ' DB Verbindung
        Dim Connection As New SqlConnection
        Connection.ConnectionString = "Data Source=" & dataSource & "; Database=" & dbname & "; User Id=" & userName & "; Password=" & password
        Connection.Open()

        ' Excel-Datei
        Dim excel As New ExcelPackage(New FileInfo(excelName))

        ' Alle Tabellennamen laden
        Dim tableNames = TableSelect("TABLE_SCHEMA, TABLE_NAME", "INFORMATION_SCHEMA.TABLES",
                                    "TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG = '" & dbname & "'", Connection)

        ' Für jede Tabelle alle Daten laden und ins Excel speichern

        For tableIndex = 1 To tableNames.Rows.Count
            ' Fortschritt anzeigen
            Dim row = tableNames.Rows(tableIndex - 1)

            ' Inhalt laden
            Dim tableContent = TableSelect("*", dbname & "." & row(0) & "." & row(1), "", Connection)

            ' Worksheet laden oder erstellen
            Dim worksheet = excel.Workbook.Worksheets.Add(row(0) & "." & row(1))

            Dim rowCount = Math.Min(tableContent.Rows.Count, limit)

            ' Worksheet befüllen
            For rowIndex = 1 To rowCount
                Dispatcher.BeginInvoke(Sub() progress.Content = "Tabelle " & tableIndex & "/" & tableNames.Rows.Count & ", Zeile " & rowIndex & "/" & rowCount)
                For columnIndex = 1 To tableContent.Columns.Count
                    worksheet.Cells(rowIndex, columnIndex).Value = tableContent.Rows(rowIndex - 1)(columnIndex - 1)
                Next
            Next
        Next

        Dispatcher.BeginInvoke(Sub()
                                   progress.Content = "Datenbank erfolgreich In Excel-Datei konvertiert."
                                   button.IsEnabled = True
                               End Sub)
        excel.Save()
        Connection.Close()
    End Sub

    Private Function TableSelect(columns As String, table As String, where As String, connection As SqlConnection) As DataTable
        Dim sqlString = "Select " & columns & " FROM " & table &
            If(String.IsNullOrWhiteSpace(where), "", " WHERE " & where)
        Dim DataAdapter As New SqlDataAdapter(sqlString, connection)
        Dim CommandBuilder As New SqlCommandBuilder(DataAdapter)
        Dim DataTable As New DataTable()
        DataAdapter.Fill(DataTable)

        Return DataTable
    End Function

End Class