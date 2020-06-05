Imports System.IO
Imports System.Reflection
Imports System.Data
Imports ConsoleApp.ExcelService

Namespace Common
    Public Class Utilities
#Region "Methods"
        Shared Function ToDictionary(Of T As Class)(ByVal _Object As T, Optional ByVal prefix As String = Nothing) As Dictionary(Of String, String)
            Dim _ReturnVal As New Dictionary(Of String, String)
            If _Object Is Nothing Then
                Return _ReturnVal
            End If

            Dim _PropertyInfo() As PropertyInfo = _Object.GetType().GetProperties()

            For Each _Info As PropertyInfo In _PropertyInfo
                Dim _Value = _Info.GetValue(_Object, Nothing)
                Dim _String = String.Empty

                If _Value IsNot Nothing Then
                    _String = _Value.ToString()
                End If
                _ReturnVal(String.Format("{0}{1}", prefix, _Info.Name)) = _String
            Next
            Return _ReturnVal
        End Function
        Shared Function RegexValidator(ByVal Regex As String(), ByVal Source As String) As String
            Dim _Return As String = Source
            Dim _IsMatch As Boolean
            For Each _String As String In Regex
                Dim _Regex As New System.Text.RegularExpressions.Regex(_String)
                With _Regex
                    _IsMatch = .IsMatch(_Return)
                    If _IsMatch Then
                        _Return = .Match(_Return).Value
                    Else
                        Exit For
                    End If
                End With
            Next
            Return IIf(_IsMatch, _Return, String.Empty)
        End Function
        Shared Function CreateFolder(ByVal Subject As String, ByVal LogFilePath As String) As Dictionary(Of String, String)
            Dim _Result As New Dictionary(Of String, String)
            Dim _LocalPath As String = AppDomain.CurrentDomain.BaseDirectory
            Dim _FolderName As String = Now.ToString("yyyy-MM-dd HHmmss ") & Subject

            For Each _Char As Char In Path.GetInvalidFileNameChars
                _FolderName = _FolderName.Replace(_Char, "")
            Next

            Dim _FolderPath As String = _LocalPath & LogFilePath & _FolderName & "\"
            If Not Directory.Exists(_FolderPath) Then
                Try
                    Directory.CreateDirectory(_FolderPath)
                    _Result.Add("Success", _FolderPath)
                Catch ex As Exception
                    _Result.Add("Error", _FolderPath)
                End Try
            End If

            Return _Result
        End Function
        Shared Function XLSToDataTable(ByVal pExcelFilePath As String) As DataTable
            Dim _ExcelObj As New ExcelObject(pExcelFilePath)
            Dim _DataTable As New DataTable
            If _ExcelObj Is Nothing Then
                Dim dtXlsSchema As DataTable = _ExcelObj.GetSchema()
                For i As Integer = 0 To dtXlsSchema.Rows.Count - 1
                    Debug.WriteLine(dtXlsSchema.Rows(i).Item("Table_Name").ToString)
                    Dim strTableName As String = dtXlsSchema.Rows(i).Item("Table_Name").ToString
                    If strTableName.Contains("Sheet1$") Then
                        _DataTable = _ExcelObj.ReadTable(strTableName, "")
                        Exit For
                    End If
                Next
            End If
            _ExcelObj = Nothing
            Return _DataTable
        End Function
#End Region
    End Class
End Namespace

