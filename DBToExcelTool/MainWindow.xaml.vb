Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports OfficeOpenXml

Class MainWindow

    Private Connection As New SqlConnection
    Private Limit = 10000
    Private trd As Thread

    Private Sub ClickConvert(sender As Object, e As RoutedEventArgs) Handles button.Click
        Try
            MySettings.Default.Save()
            ' DB Verbindung
            If (Connection.State <> ConnectionState.Open) Then
                Connection.ConnectionString = "Data Source=(localdb)\V11.0; Database=" & input_db.Text
                Connection.Open()
            End If

            ' Excel-Datei
            Dim excel As New ExcelPackage(New FileInfo(input_file.Text))

            ' Alle Tabellennamen laden
            Dim tableNames = TableSelect("TABLE_SCHEMA, TABLE_NAME", "INFORMATION_SCHEMA.TABLES",
                                        "TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG = '" & input_db.Text & "'", Connection)

            ' Für jede Tabelle alle Daten laden und ins Excel speichern
            For Each row As DataRow In tableNames.Rows
                Dim tableContent = TableSelect("*", input_db.Text & "." & row(0) & "." & row(1), "", Connection)

                ' Fortschritt anzeigen
                progress.Content = "Verarbeite Tabelle " & row(0) & "." & row(1) & "(" & tableNames.Rows.IndexOf(row) & "/" & tableNames.Rows.Count & ")"

                ' Worksheet erstellen
                Dim worksheet = excel.Workbook.Worksheets.Add(row(0) & "." & row(1))

                ' Worksheet befüllen
                FillCells(tableContent, worksheet)
            Next

            excel.Save()
            progress.Content = "Datenbank erfolgreich In Excel-Datei konvertiert."
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub FillCells(tableContent As DataTable, worksheet As ExcelWorksheet)
        For rowIndex = 1 To Math.Min(tableContent.Rows.Count, Limit)
            For columnIndex = 1 To tableContent.Columns.Count
                worksheet.Cells(rowIndex, columnIndex).Value = tableContent.Rows(rowIndex - 1)(columnIndex - 1)
            Next
        Next
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
