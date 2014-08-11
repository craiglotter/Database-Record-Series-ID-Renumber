Imports System.IO

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public DatabaseToUse As String = ""
    Public TableToUse As String = ""
    Public ColumnSort As String = ""
    Public ColumnSortOrder As String = ""
    Public ColumnRenumber As String = ""
    Public SeriesPrefix As String = ""
    Public SeriesSuffix As String = ""
    Public SeriesStart As Long = 0

    Private columnsorttype As OleDb.OleDbType
    Private columnrenumbertype As OleDb.OleDbType

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal switch As Integer)



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal exc As Exception)
        Try
            If exc.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Worker Thread encountered the following problem: " & vbCrLf & exc.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & exc.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch ex As Exception
            MsgBox("An error occurred in Database Record Series ID Renumber's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal exc As String)
        Try

            Dim Display_Message1 As New Display_Message("The Worker Thread encountered the following problem: " & vbCrLf & exc.ToString)
            Display_Message1.ShowDialog()
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & exc.ToString)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            MsgBox("An error occurred in Database Record Series ID Renumber's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try
            Dim f As FileInfo = New FileInfo(DatabaseToUse)
            If f.Exists = True Then
                result = "Success"
                Dim Conn As Data.OleDb.OleDbConnection
                Conn = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseToUse & ";")
                Conn.Open()
                Dim Conn2 As Data.OleDb.OleDbConnection
                Conn2 = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseToUse & ";")
                Conn2.Open()



                '************
                Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, TableToUse, Nothing})
                Dim myrow2 As DataRow
                Dim mycol2 As DataColumn
                Dim tempdatatype As OleDb.OleDbType
                Dim tempcolumnname As String
                For Each myRow2 In schemaTable2.Rows
                    For Each myCol2 In schemaTable2.Columns
                        If myCol2.ColumnName = "DATA_TYPE" Then
                            tempdatatype = myrow2(mycol2)
                        End If
                        If myCol2.ColumnName = "COLUMN_NAME" Then
                            tempcolumnname = myrow2(mycol2).ToString()
                        End If
                    Next
                    If tempcolumnname = ColumnSort Then
                        columnsorttype = tempdatatype
                    End If
                    If tempcolumnname = ColumnRenumber Then
                        columnrenumbertype = tempdatatype
                    End If
                Next
                ' MsgBox("columnsort: " & ColumnSort & " - " & columnsorttype.ToString & vbCrLf & "columnrenum: " & ColumnRenumber & " - " & columnrenumbertype.ToString)

                myrow2 = Nothing
                mycol2 = Nothing
                schemaTable2.Dispose()
                '************


                Dim SQL2 As String
                Dim SQL As String = "select * from [" & TableToUse & "] order by [" & ColumnSort & "] " & ColumnSortOrder
                Dim recset As Data.OleDb.OleDbCommand = New Data.OleDb.OleDbCommand(SQL, Conn)
                Dim recread As Data.OleDb.OleDbDataReader
                Dim recset2 As Data.OleDb.OleDbCommand

                recread = recset.ExecuteReader()
                Dim sortquotes As String = ""
                Dim renumquotes As String = ""
                If columnsorttype.ToString.ToLower = "wchar" Then
                    sortquotes = """"
                Else
                    sortquotes = ""
                End If
                If columnrenumbertype.ToString.ToLower = "wchar" Then
                    renumquotes = """"
                Else
                    renumquotes = ""
                    If IsNumeric(SeriesPrefix) = False Then
                        SeriesPrefix = ""
                    End If
                    If IsNumeric(SeriesSuffix) = False Then
                        SeriesSuffix = ""
                    End If
                End If

                If recread.HasRows = True Then

                    While recread.Read() = True
                        Try

                            SQL2 = "update [" & TableToUse & "] set [" & ColumnRenumber & "] = " & renumquotes & SeriesPrefix & SeriesStart & SeriesSuffix & renumquotes & " where [" & ColumnRenumber & "] = " & renumquotes & recread.Item(ColumnRenumber).ToString & renumquotes & " and [" & ColumnSort & "] = " & sortquotes & recread.Item(ColumnSort).ToString & sortquotes
                            SeriesStart = SeriesStart + 1

                            Dim resultt As Integer
                            recset2 = New Data.OleDb.OleDbCommand(SQL2, Conn2)
                            resultt = (recset2.ExecuteNonQuery())
                            If resultt > 0 Then
                                RaiseEvent WorkerProgress(1)
                            End If
                            RaiseEvent WorkerProgress(2)

                        Catch ex As Exception
                            Error_Handler(ex)
                        End Try
                    End While
                End If

                recread.Close()
                recset.Dispose()
                recset2.Dispose()
                Conn.Close()
                Conn.Dispose()
                Conn2.Close()
                Conn2.Dispose()
            Else
                    result = "Failure"
                    Error_Handler("Database file inaccessible")
                End If

                RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            Error_Handler(ex)
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub







End Class
