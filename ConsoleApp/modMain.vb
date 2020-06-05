Imports System.IO
Imports System.Net.Mail
Imports System.Text
Imports System.Text.RegularExpressions
Imports ConsoleApp.Common
Imports ConsoleApp.Email
Imports ConsoleApp.ExcelService

Module modMain
#Region "Main Variables & Properties"
    Private ExcelObj As ExcelObject
    Private WithEvents mInteropExcel As New InteropExcel
    Private SkipCount As Integer = 0
    Private FileDirectory As String = ""
    Private WithEvents objMailComponent As New MailComponents
    Private Event onError(ByVal ex As Exception)
#End Region

#Region "Delegate MailComponent"
    Sub MailComponent_Reporting(ByVal pReportType As MailComponents.ReportType,
                                ByVal pReportingStage As MailComponents.ReportingStage,
                                ByVal pMessage As String,
                                ByVal curLeftVal As Nullable(Of Integer),
                                ByVal curTopVal As Nullable(Of Integer),
                                ByVal pDisplayStageInd As Boolean)

        Dim _Left As Integer = Console.CursorLeft
        Dim _Top As Integer = Console.CursorTop
        Dim _CurrentLeft As Integer = 0
        Try
            If curLeftVal.HasValue AndAlso curTopVal.HasValue Then
                Console.WriteLine(Now & vbTab & IIf(pDisplayStageInd, pReportingStage.ToString, vbTab) & vbTab & pMessage.PadRight(_CurrentLeft, Chr(32)))
                _CurrentLeft = Console.CursorLeft
                Console.SetCursorPosition(_Left, _Top)
            Else
                Console.SetCursorPosition(_CurrentLeft, _Top)
                Console.WriteLine(Now & vbTab & IIf(pDisplayStageInd, pReportingStage.ToString, vbTab) & vbTab & pMessage)
            End If
        Catch ex As Exception
            MailComponent_onError(ex)
        End Try
    End Sub

    Function MailComponent_ReadFromFile(ByVal FilePath As String) As Dictionary(Of String, String)
        Dim Result As New Dictionary(Of String, String)
        Try
            Dim _String As String = File.ReadAllText(FilePath)
            Result.Add("Success", _String)
        Catch ex As Exception
            Result.Add("Error", ex.Message)
        End Try
        Return Result
    End Function

    Function MailComponent_GenerateMailBody(ByVal pTemplateFilePath As String, ByVal pShawStaffFilePath As String, ByVal pTemperatureFilePath As String) As Dictionary(Of String, String)
        Dim Result As New Dictionary(Of String, String)
        Try
            Dim _OriginalTemplate As String = File.ReadAllText(pTemplateFilePath)
            Dim _StaffDataTable As Data.DataTable = MailComponent_ReadFromExcelWithInterop(pShawStaffFilePath, "FullList")
            Dim _TemperatureDataTable As Data.DataTable = MailComponent_ReadFromExcelWithInterop(pTemperatureFilePath, "Form Responses 1")

            Dim _DataView As DataView = New DataView(_TemperatureDataTable)
            Dim _DistinctTable As Data.DataTable = _DataView.ToTable(True, "NRIC/FIN (Last 4 Alphanumeric ie.123F)")

            Dim _NRICList As List(Of String) = (From r In _DistinctTable.AsEnumerable() Select r.Field(Of String)(0).Trim).ToList()

            Dim _FilteredTable As Data.DataTable = _StaffDataTable.AsEnumerable.Where(Function(row) Not _NRICList.Contains(row.Field(Of String)("NRIC Last 4").Trim)).CopyToDataTable

            Dim _TemplateString As String = GetTableString(_FilteredTable, _OriginalTemplate)
            Result.Add("Success", _TemplateString)
        Catch ex As Exception
            Result.Add("Error", ex.Message)
        End Try
        Return Result
    End Function

    Function MailComponent_ReadFromExcelWithInterop(ByVal path As String, Optional ByVal pSheetName As String = Nothing) As Data.DataTable
        Dim _TempEventList As New Dictionary(Of String, Object(,))
        Dim _DataTable As Data.DataTable = Nothing
        Dim _DefaultSheetName As String = Nothing

        If pSheetName Is Nothing Then
            Dim _SheetList As List(Of String) = mInteropExcel.GetSheetList(path)
            _DefaultSheetName = _SheetList(0)
        Else
            _DefaultSheetName = pSheetName
        End If

        _TempEventList = mInteropExcel.ExcelAnalysis(path, _DefaultSheetName)

        If _TempEventList(pSheetName).Length > 0 Then
            _DataTable = mInteropExcel.GetDataTableFrom2DArray(_TempEventList(_DefaultSheetName))
        End If

        Return _DataTable
    End Function

    Function MailComponent_ReadFromExcel(ByVal path As String) As Data.DataTable
        Dim dt As Data.DataTable = Nothing
        If ExcelObj Is Nothing Then
            ExcelObj = New ExcelObject(path)
            Dim dtXlsSchema As Data.DataTable = ExcelObj.GetSchema()
            For i As Integer = 0 To dtXlsSchema.Rows.Count - 1
                Debug.WriteLine(dtXlsSchema.Rows(i).Item("Table_Name").ToString)
                Dim strTableName As String = dtXlsSchema.Rows(i).Item("Table_Name").ToString
                If strTableName.Contains("Sheet1$") Then
                    dt = ExcelObj.ReadTable(strTableName, "")
                    Exit For
                End If
            Next
        End If
        ExcelObj = Nothing
        Return dt
    End Function

    Function GetTableString(ByVal pTable As Data.DataTable, pTemplate As String) As String
        Dim _StringBuilder As New StringBuilder

        For Each _DataRow As Data.DataRow In pTable.Rows
            _StringBuilder.AppendLine("<tr>")
            _StringBuilder.AppendFormat("<td>{0}</td>", _DataRow(0)).AppendLine()
            _StringBuilder.AppendFormat("<td>{0}</td>", _DataRow(1)).AppendLine()
            _StringBuilder.AppendFormat("<td>{0}</td>", _DataRow(2)).AppendLine()
            _StringBuilder.AppendFormat("<td>{0}</td>", _DataRow(3)).AppendLine()
            _StringBuilder.AppendFormat("<td>{0}</td>", _DataRow(5)).AppendLine()
            _StringBuilder.AppendLine("</tr>")
        Next
        pTemplate = pTemplate.Replace("{{row}}", _StringBuilder.ToString)

        Return pTemplate
    End Function

#End Region

    Sub Main(ByVal pArgument() As String)
        Try
            Dim _Argument As New Dictionary(Of String, String)
            For Each _String As String In pArgument
                Dim _Key As String = String.Empty
                Dim _Value As String = String.Empty
                _Key = _String.Split("=")(0)
                If _String.Contains("=") Then
                    _Value = _String.Split("=")(1)
                End If
                If _Argument.ContainsKey(_Key) Then
                    _Argument(_Key) = _Key
                Else
                    _Argument.Add(_Key, _Value)
                End If
            Next

            With MailComponents.MailComponentParms
                .SMTPHost = My.Settings.MailServer
                .FromAddress = My.Settings.FromAddress
                .FromName = My.Settings.FromName
                .EnableSSL = My.Settings.EnableSSL

                If _Argument.ContainsKey("-RetryCount") AndAlso IsNumeric(_Argument.Item("-RetryCount")) Then
                    .RetryCount = _Argument.Item("-RetryCount")
                Else
                    .RetryCount = My.Settings.RetryCount
                End If

                If _Argument.ContainsKey("-ToAddress") Then
                    .ToAddress = _Argument.Item("-ToAddress")
                Else
                    .ToAddress = My.Settings.ToAddress
                End If

                If _Argument.ContainsKey("-ToName") Then
                    .ToName = _Argument.Item("-ToName")
                Else
                    .ToName = My.Settings.ToName
                End If

                If _Argument.ContainsKey("-Subject") Then
                    .Subject = _Argument.Item("-Subject")
                Else
                    .Subject = My.Settings.Subject
                End If

                If _Argument.ContainsKey("-Body") Then
                    .Body = _Argument.Item("-Body")
                Else
                    .Body = My.Settings.Body
                End If

                If _Argument.ContainsKey("-CcAddress") Then
                    .CcAddress = _Argument.Item("-CcAddress")
                Else
                    .CcAddress = My.Settings.CcAddress
                End If

                If _Argument.ContainsKey("-BccAddress") Then
                    .BccAddress = _Argument.Item("-BccAddress")
                Else
                    .BccAddress = My.Settings.BccAddress
                End If

                If _Argument.ContainsKey("-MailPriority") AndAlso IsNumeric(_Argument.Item("-MailPriority")) Then
                    .MailPriority = _Argument.Item("-MailPriority")
                Else
                    .MailPriority = My.Settings.MailPriority
                End If

                If _Argument.ContainsKey("-DeliveryNotificationOption") AndAlso IsNumeric(_Argument.Item("-DeliveryNotificationOption")) Then
                    .DeliveryNotificationOptions = _Argument.Item("-DeliveryNotificationOption")
                Else
                    .DeliveryNotificationOptions = My.Settings.DeliveryNotificationOptions
                End If

                If _Argument.ContainsKey("-TemplateFilePath") Then
                    .TemplateFilePath = _Argument.Item("-TemplateFilePath")
                Else
                    .TemplateFilePath = My.Settings.TemplateFilePath
                End If

                If _Argument.ContainsKey("-LogFilePath") Then
                    .LogFilePath = _Argument.Item("-LogFilePath")
                Else
                    .LogFilePath = My.Settings.LogFilePath
                End If

                If _Argument.ContainsKey("-ShawStaffFilePath") Then
                    .ShawStaffFilePath = _Argument.Item("-ShawStaffFilePath")
                Else
                    .ShawStaffFilePath = My.Settings.ShawStaffFilePath
                End If

                If _Argument.ContainsKey("-TemparatureFilePath") Then
                    .TemparatureFilePath = _Argument.Item("-TemparatureFilePath")
                Else
                    .TemparatureFilePath = My.Settings.TemparatureFilePath
                End If

                If _Argument.ContainsKey("-ErrorFlag") Then
                    .ErrorFlag = _Argument.Item("-ErrorFlag")
                Else
                    .ErrorFlag = My.Settings.ErrorFlag
                End If

                Dim _ReturnResult As Dictionary(Of String, String) = Nothing
                _ReturnResult = Utilities.CreateFolder(MailComponents.MailComponentParms.Subject, MailComponents.MailComponentParms.LogFilePath)

                If _ReturnResult.ContainsKey("Success") Then
                    FileDirectory = _ReturnResult.Item("Success")
                    objMailComponent.CallReporting = New MailComponents.DelegateReporting(AddressOf MailComponent_Reporting)
                    objMailComponent.CallReadFromExcel = New MailComponents.DelegateReadFromExcel(AddressOf MailComponent_ReadFromExcel)
                    objMailComponent.CallReadFromFile = New MailComponents.DelegateReadFromFile(AddressOf MailComponent_ReadFromFile)
                    objMailComponent.CallGenerateMailBody = New MailComponents.DelegateGenerateMailBody(AddressOf MailComponent_GenerateMailBody)

                    DisplayInititalizeParameters()

                    DoProcess()
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

#Region "Main Methods"
    Private Sub DisplayInititalizeParameters()
        With MailComponents.MailComponentParms
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "From:" & .FromAddress, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "FromName:" & .FromName, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "To:" & .ToAddress, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "ToName:" & .ToName, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "Cc:" & .CcAddress, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "Bcc:" & .BccAddress, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "Subject:" & .Subject, Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "MailPriority:" & [Enum].GetName(GetType(MailPriority), .MailPriority), Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "DeliveryNotificationOption:" & [Enum].GetName(GetType(DeliveryNotificationOptions), .DeliveryNotificationOptions), Nothing, Nothing, True)
            MailComponent_Reporting(MailComponents.ReportType.Process, MailComponents.ReportingStage.Initialization, "HTMLFileLocation:" & .TemplateFilePath, Nothing, Nothing, True)
        End With
    End Sub
#End Region

#Region "Processing State"
    Private Sub DoProcess()
        With MailComponents.MailComponentParms
            objMailComponent.SendMailMessage()
            If Not .ErrorFlag Then
                MailComponent_Reporting(MailComponents.ReportType.Success, MailComponents.ReportingStage.Finalizing, Nothing, Nothing, Nothing, True)
                Console.WriteLine("Press ENTER to exit...")
                Console.ReadLine()
            Else
                Console.WriteLine("Error occur during processing, Please find the log file.")
                Console.WriteLine("Press ENTER to exit...")
                Console.ReadLine()
            End If
        End With

    End Sub
#End Region

#Region "Error Event Handlers"
    Public Sub MailComponent_onError(Optional ByVal Exception As System.Exception = Nothing,
                                    Optional ByVal Stage As MailComponents.ReportingStage = Nothing,
                                    Optional ByVal ReportParameters As String = Nothing,
                                    Optional ByVal TotalRecord As Integer = 0,
                                    Optional ByVal CurrentRow As Integer = 0) Handles objMailComponent.onError
        Try
            Dim _LocalPath As String = AppDomain.CurrentDomain.BaseDirectory
            Dim _FullFileName As String = String.Empty
            Dim _FileDirectory As String = String.Empty

            _FileDirectory = FileDirectory

            If Not Directory.Exists(_FileDirectory) Then
                Directory.CreateDirectory(_FileDirectory)
            End If

            Dim _NewLogFileName As String = Path.GetFileNameWithoutExtension(My.Settings.ShawStaffFilePath)
            _FullFileName = _FileDirectory & _NewLogFileName & ".log"

            Dim _StreamWriter As New StreamWriter(_FullFileName, True)
            _StreamWriter.WriteLine("Report Parameters : " & vbCrLf)
            _StreamWriter.WriteLine(ReportParameters)
            _StreamWriter.WriteLine("-------------------------------------------------------------------------------------------------------")
            _StreamWriter.WriteLine("Total Record : " & TotalRecord)
            _StreamWriter.WriteLine("-------------------------------------------------------------------------------------------------------")
            _StreamWriter.WriteLine("Current Row Record : " & CurrentRow)
            _StreamWriter.WriteLine("-------------------------------------------------------------------------------------------------------")
            _StreamWriter.WriteLine("Current Stage Message : " & [Enum].GetName(GetType(MailComponents.ReportingStage), Stage))
            _StreamWriter.WriteLine("-------------------------------------------------------------------------------------------------------")
            _StreamWriter.WriteLine("Exception Message : " & Exception.Message)
            _StreamWriter.WriteLine("-------------------------------------------------------------------------------------------------------")
            _StreamWriter.WriteLine("Stack Trace : " & vbCrLf)
            _StreamWriter.WriteLine(Exception.StackTrace)
            _StreamWriter.WriteLine()
            _StreamWriter.Flush()
            _StreamWriter.Close()
            _StreamWriter.Dispose()
            MailComponents.MailComponentParms.ErrorFlag = True
        Catch ex As Exception
            Console.WriteLine(ex.Message & vbNewLine & ex.StackTrace)
            Console.ReadLine()
        End Try
    End Sub
#End Region
End Module
