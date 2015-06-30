﻿Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports OfficeOpenXml

Class MainWindow

    Private Connection As New SqlConnection
    Private Limit = 2000

    Private Sub ClickConvert(sender As Object, e As RoutedEventArgs) Handles button.Click
        Try
            MySettings.Default.Save()
            button.IsEnabled = False

            Dim dbName = input_db.Text
            Dim fileName = input_file.Text

            If (File.Exists(fileName)) Then
                File.Delete(fileName)
            End If

            Dim Thread1 As New Thread(Sub() CreateExcelFile(dbName, fileName))
            Thread1.Start()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub CreateExcelFile(dbname As String, excelName As String)
        ' DB Verbindung
        If (Connection.State <> ConnectionState.Open) Then
            Connection.ConnectionString = "Data Source=(localdb)\V11.0; Database=" & dbname
            Connection.Open()
        End If

        ' Excel-Datei
        Dim excel As New ExcelPackage(New FileInfo(excelName))

        ' Alle Tabellennamen laden
        Dim tableNames = TableSelect("TABLE_SCHEMA, TABLE_NAME", "INFORMATION_SCHEMA.TABLES",
                                    "TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG = '" & dbname & "'", Connection)

        ' Zählvariablen
        Dim tableCount = tableNames.Rows.Count
        Dim counter = 1

        ' Für jede Tabelle alle Daten laden und ins Excel speichern
        For Each row As DataRow In tableNames.Rows
            ' Fortschritt anzeigen
            counter += 1

            ' Inhalt laden
            Dim tableContent = TableSelect("*", dbname & "." & row(0) & "." & row(1), "", Connection)

            ' Worksheet laden oder erstellen
            Dim worksheet = Nothing  'excel.Workbook.Worksheets.FirstOrDefault(Name = row(0) & "." & row(1))
            worksheet = If(worksheet, excel.Workbook.Worksheets.Add(row(0) & "." & row(1)))

            Dim rowCount = Math.Min(tableContent.Rows.Count, Limit)
            ' Worksheet befüllen
            For rowIndex = 1 To rowCount
                Dispatcher.BeginInvoke(Sub() progress.Content = "Tabelle " & counter & "/" & tableCount & " Zeile " & rowIndex & "/" & rowCount)
                For columnIndex = 1 To tableContent.Columns.Count
                    worksheet.Cells(rowIndex, columnIndex).Value = tableContent.Rows(rowIndex - 1)(columnIndex - 1)
                Next
            Next
        Next

        Dispatcher.BeginInvoke(Sub() progress.Content = "Datenbank erfolgreich In Excel-Datei konvertiert.")
        Dispatcher.BeginInvoke(Sub() button.IsEnabled = True)

        excel.Save()
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