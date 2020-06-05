Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports ExcelLibrary.SpreadSheet

Namespace ExcelService
    Public Class ExcelObject
        Implements IDisposable

        Private excelObject As String = "Provider=Microsoft.{0}.OLEDB.{1};Data Source={2};Extended Properties=""Excel {3};HDR=YES;READONLY=FALSE\"""
        Private filepath As String = String.Empty
        Private con As OleDbConnection = Nothing

        Public Delegate Sub ProgressWork(ByVal percentage As Single)
        Private Event Reading As ProgressWork
        Private Event Writeing As ProgressWork
        Private Event connectionStringChange As EventHandler

        Public Custom Event ReadProgress As ProgressWork
            AddHandler(ByVal value As ProgressWork)

                AddHandler Reading, value

            End AddHandler

            RemoveHandler(ByVal value As ProgressWork)

                RemoveHandler Reading, value

            End RemoveHandler

            RaiseEvent(ByVal percentage As Single)

                RaiseEvent Reading(percentage)

            End RaiseEvent
        End Event

        Public Custom Event WriteProgress As ProgressWork
            AddHandler(ByVal value As ProgressWork)

                AddHandler Writeing, value

            End AddHandler

            RemoveHandler(ByVal value As ProgressWork)

                RemoveHandler Writeing, value

            End RemoveHandler

            RaiseEvent(ByVal percentage As Single)

                RaiseEvent Writeing(percentage)

            End RaiseEvent
        End Event

        Public Custom Event ConnectionStringChanged As EventHandler
            AddHandler(ByVal value As EventHandler)
                AddHandler connectionStringChange, value
            End AddHandler

            RemoveHandler(ByVal value As EventHandler)
                RemoveHandler connectionStringChange, value
            End RemoveHandler

            RaiseEvent(ByVal sender As Object, ByVal e As System.EventArgs)

                If Me.Connection IsNot Nothing AndAlso
                Not Me.Connection.ConnectionString.Equals(Me.ConnectionString) Then

                    If Me.Connection.State = ConnectionState.Open Then
                        Me.Connection.Close()
                    End If
                    Me.Connection.Dispose()
                    Me.con = Nothing

                End If

                RaiseEvent connectionStringChange(sender, e)

            End RaiseEvent
        End Event

        Public ReadOnly Property ConnectionString() As String
            Get
                If Not Me.filepath = String.Empty Then
                    'Check for File Format
                    Dim fi As New FileInfo(Me.filepath)
                    If fi.Extension.Equals(".xls") Then
                        Return String.Format(Me.excelObject, "Jet", "4.0", Me.filepath, "8.0")
                    ElseIf fi.Extension.Equals(".xlsx") Then
                        Return String.Format(Me.excelObject, "Ace", "12.0", Me.filepath, "12.0")
                    Else : Return String.Empty
                    End If
                Else
                    Return String.Empty
                End If
            End Get
        End Property


        Public ReadOnly Property Connection() As OleDbConnection
            Get
                If con Is Nothing Then
                    Dim _con As New OleDbConnection(Me.ConnectionString)
                    Me.con = _con
                End If
                Return Me.con
            End Get
        End Property

        Sub New(ByVal path As String)

            Me.filepath = path
            RaiseEvent ConnectionStringChanged(Me, New EventArgs())

        End Sub

        Public Function GetSchema() As DataTable
            Dim dtSchema As DataTable = Nothing
            If Me.Connection.State <> ConnectionState.Open Then Me.Connection.Open()
            dtSchema = Me.Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            Return dtSchema
        End Function

        Public Function ReadTable(ByVal tableName As String, Optional ByVal criteria As String = "") As DataTable

            Try
                Dim resultTable As DataTable = Nothing
                If Me.Connection.State <> ConnectionState.Open Then
                    Me.Connection.Open()
                    RaiseEvent ReadProgress(10)
                End If
                Dim cmdText As String = "Select * from [{0}]"
                If criteria <> "" Then
                    cmdText += " Where " + criteria
                End If
                Dim cmd As New OleDbCommand(String.Format(cmdText, tableName))
                cmd.Connection = Me.Connection
                Dim adpt As New OleDbDataAdapter(cmd)
                RaiseEvent ReadProgress(30)
                Dim ds As New DataSet
                RaiseEvent ReadProgress(50)
                adpt.Fill(ds, tableName)
                RaiseEvent ReadProgress(100)
                If ds.Tables.Count = 1 Then
                    Return ds.Tables(0)
                Else
                    Return Nothing
                End If
            Catch
                Console.WriteLine("Table Cannot be read")
                Return Nothing
            End Try
        End Function

        Public Function DropTable(ByVal tablename As String) As Boolean

            Try
                If Me.Connection.State <> ConnectionState.Open Then
                    Me.Connection.Open()
                    RaiseEvent WriteProgress(10)
                End If
                Dim cmdText As String = "Drop Table [{0}]"
                Using cmd As New OleDbCommand(String.Format(cmdText, tablename), Me.Connection)
                    RaiseEvent WriteProgress(30)
                    cmd.ExecuteNonQuery()
                    RaiseEvent WriteProgress(80)
                End Using
                Me.Connection.Close()
                RaiseEvent WriteProgress(100)
                Return True
            Catch ex As Exception
                RaiseEvent WriteProgress(0)
                Console.WriteLine(ex.Message)
                Return False
            End Try
        End Function

        Public Function WriteTable(ByVal tableName As String, ByVal tableDefination As Dictionary(Of String, String)) As Boolean

            Using cmd As New OleDbCommand(Me.GenerateCreateTable(tableName, tableDefination), Me.Connection)

                If Me.Connection.State <> ConnectionState.Open Then Me.Connection.Open()
                cmd.ExecuteNonQuery()

            End Using

        End Function

        Public Function AddNewRow(ByVal dr As DataRow) As Boolean

            Using cmd As New OleDbCommand(Me.GenerateInsertStatement(dr), Me.Connection)

                cmd.ExecuteNonQuery()

            End Using
            Return True
        End Function

        ''' <summary>
        ''' Generates Create Table Script
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="tableDefination"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GenerateCreateTable(ByVal tableName As String, ByVal tableDefination As Dictionary(Of String, String)) As String

            Dim sb As New StringBuilder()
            Dim firstcol As Boolean = True
            sb.AppendFormat("CREATE TABLE [{0}](", tableName)
            firstcol = True
            For Each keyvalue As KeyValuePair(Of String, String) In tableDefination
                If Not firstcol Then
                    sb.Append(",")
                End If
                firstcol = False
                sb.AppendFormat("{0} {1}", keyvalue.Key, keyvalue.Value)
            Next

            sb.Append(")")
            Return sb.ToString()
        End Function

        Private Function GenerateInsertStatement(ByVal dr As DataRow) As String
            Dim sb As New StringBuilder()
            Dim firstcol As Boolean = True
            sb.AppendFormat("INSERT INTO [{0}](", dr.Table.TableName)


            For Each dc As DataColumn In dr.Table.Columns
                If Not firstcol Then
                    sb.Append(",")
                End If
                firstcol = False
                sb.Append(dc.Caption)

            Next

            sb.Append(") VALUES(")
            firstcol = True
            For i As Integer = 0 To dr.Table.Columns.Count - 1
                If dr.Table.Columns(i).DataType IsNot GetType(Integer) Then
                    sb.Append("'")
                    sb.Append(dr(i).ToString().Replace("'", "''"))
                    sb.Append("'")
                Else
                    sb.Append(dr(i).ToString().Replace("'", "''"))
                End If
                If i <> dr.Table.Columns.Count - 1 Then
                    sb.Append(",")
                End If
            Next

            sb.Append(")")
            Return sb.ToString()
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose

            If Me.con IsNot Nothing AndAlso Me.con.State = ConnectionState.Open Then
                Me.con.Close()
            End If
            If Me.con IsNot Nothing Then
                Me.con.Dispose()
            End If
            Me.con = Nothing
        End Sub

#Region "Excel Library Function"
        Public Shared Function WriteXLSFile(ByVal pPath As String, ByVal pDataTable As DataTable) As Boolean
            Try
                Dim _FileFullPath = pPath
                Dim workbook As Workbook = New Workbook()
                Dim worksheet As Worksheet
                Dim iRow As Integer = 0
                Dim iCol As Integer = 0
                Dim sTemp As String = String.Empty
                Dim dTemp As Double = 0
                Dim iTemp As Integer = 0
                Dim dtTemp As DateTime
                Dim count As Integer = 0
                Dim iTotalRows As Integer = 0
                Dim iSheetCount As Integer = 0

                iSheetCount = iSheetCount + 1
                worksheet = New Worksheet("Sheet " & iSheetCount.ToString())

                For Each dc As DataColumn In pDataTable.Columns
                    worksheet.Cells(iRow, iCol) = New Cell(dc.ColumnName)
                    iCol = iCol + 1
                Next

                iRow = 1
                For Each dr As DataRow In pDataTable.Rows
                    iCol = 0
                    For Each dc As DataColumn In pDataTable.Columns
                        sTemp = dr(dc.ColumnName).ToString()
                        If dc.DataType Is GetType(DateTime) Then
                            DateTime.TryParse(sTemp, dtTemp)
                            worksheet.Cells(iRow, iCol) = New Cell(dtTemp, "MM/DD/YYYY")
                        ElseIf dc.DataType Is GetType(Double) Then
                            Double.TryParse(sTemp, dTemp)
                            worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
                        Else
                            worksheet.Cells(iRow, iCol) = New Cell(sTemp)
                        End If
                        iCol = iCol + 1
                    Next
                    iRow = iRow + 1
                Next

                workbook.Worksheets.Add(worksheet)
                iTotalRows = iTotalRows + iRow


                If iTotalRows < 100 Then
                    worksheet = New Worksheet("Sheet X")
                    count = 1
                    Do While count < 100
                        worksheet.Cells(count, 0) = New Cell(" ")
                        count = count + 1
                    Loop
                    workbook.Worksheets.Add(worksheet)
                End If

                If File.Exists(_FileFullPath) Then
                    File.Delete(_FileFullPath)
                End If

                workbook.Save(_FileFullPath)

                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
#End Region
    End Class
End Namespace

