Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports SecurityEncryption.Shaw

Namespace ExcelService
    Public Class InteropExcel
        Private Const ExcelPassword As String = "TLsDEKeV8XRmSJmmzr4DrAibvkq1q7HXnTxlTnvFybw="
        Private mEncryption As New Security.Encryption
        Public Function ExcelConnection(ByVal filePath As String) As OleDbCommand

            Dim xConnStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                        "Data Source=" & filePath & ";" &
                        "Extended Properties=Excel 8.0;"

            Dim objXConn As New System.Data.OleDb.OleDbConnection(xConnStr)
            Try
                objXConn.Open()
            Catch ex As OleDbException
                objXConn.Close()
                Console.WriteLine(ex.Message)
            Catch ex As Exception
                objXConn.Close()
                Console.WriteLine(ex.Message)
            End Try

            Dim objCommand As New OleDbCommand("SELECT * FROM [XXXX$]", objXConn)

            Return objCommand

        End Function

        Public Function GetSheetList(ByVal pExcelFileInputPath As String) As List(Of String)
            Dim _SheetList As New List(Of String)

            Dim Proceed As Boolean = False
            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBooks As Excel.Workbooks = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorkSheets As Excel.Sheets = Nothing

            Try
                xlApp = New Excel.Application
                xlApp.DisplayAlerts = False
                xlWorkBooks = xlApp.Workbooks
                xlWorkBook = xlWorkBooks.Open(pExcelFileInputPath, Password:=mEncryption.DecryptData(ExcelPassword))
                xlApp.Visible = False
                xlWorkSheets = xlWorkBook.Sheets

                If IO.File.Exists(pExcelFileInputPath) Then
                    If xlWorkSheets.Count > 0 Then
                        For x As Integer = 1 To xlWorkSheets.Count
                            xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)
                            _SheetList.Add(xlWorkSheet.Name.ToString)
                        Next
                    Else
                        Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
                        xlWorkSheet = Nothing
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                xlWorkBook.Close()
                xlApp.UserControl = True
                xlApp.Quit()
                ReleaseComObject(xlWorkSheets)
                ReleaseComObject(xlWorkSheet)
                ReleaseComObject(xlWorkBook)
                ReleaseComObject(xlWorkBooks)
                ReleaseComObject(xlApp)
            End Try
            Return _SheetList
        End Function

        Public Function ExcelAnalysis(ByVal pExcelFileInputPath As String, Optional ByVal pSheetName As String = Nothing) As Dictionary(Of String, Object(,))
            Dim ExcelSheets As New Dictionary(Of String, Object(,))
            Dim ExcelSheetArray As Object(,) = New Object(,) {}

            Dim Proceed As Boolean = False
            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBooks As Excel.Workbooks = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorkSheets As Excel.Sheets = Nothing
            Dim xlCells As Excel.Range = Nothing

            Try
                xlApp = New Excel.Application
                xlApp.DisplayAlerts = False
                xlWorkBooks = xlApp.Workbooks
                xlWorkBook = xlWorkBooks.Open(pExcelFileInputPath, Password:=mEncryption.DecryptData(ExcelPassword))
                xlApp.Visible = False
                xlWorkSheets = xlWorkBook.Sheets

                If IO.File.Exists(pExcelFileInputPath) Then
                    If pSheetName Is Nothing Then
                        If xlWorkSheets.Count > 0 Then
                            For x As Integer = 1 To xlWorkSheets.Count
                                xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)
                                ExcelSheetArray = Processing(pExcelFileInputPath, xlWorkSheet)
                                ExcelSheets.Add(xlWorkSheet.Name.ToString, ExcelSheetArray)
                            Next
                        Else
                            Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
                            xlWorkSheet = Nothing
                        End If
                    Else
                        For x As Integer = 1 To xlWorkSheets.Count
                            xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)
                            If xlWorkSheet.Name = pSheetName Then
                                Proceed = True
                                Exit For
                            End If
                            Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
                            xlWorkSheet = Nothing
                        Next
                        If Proceed Then
                            ExcelSheetArray = Processing(pExcelFileInputPath, xlWorkSheet)
                            ExcelSheets.Add(xlWorkSheet.Name.ToString, ExcelSheetArray)
                        End If
                    End If
                Else
                    MsgBox("'" & pExcelFileInputPath & "' not located.")
                End If
            Catch ex As Exception
                Throw ex
            Finally
                xlWorkBook.Close()
                xlApp.UserControl = True
                xlApp.Quit()
                ReleaseComObject(xlCells)
                ReleaseComObject(xlWorkSheets)
                ReleaseComObject(xlWorkSheet)
                ReleaseComObject(xlWorkBook)
                ReleaseComObject(xlWorkBooks)
                ReleaseComObject(xlApp)
            End Try
            Return ExcelSheets
        End Function

        Public Function Processing(ByVal pExcelFileInputPath As String, ByRef xlWorkSheet As Excel.Worksheet)
            Dim NewExcelArray(,) As Object = New Object(,) {}
            Dim xlUsedRange As Excel.Range = xlWorkSheet.UsedRange
            Dim _NewRowIndex As Integer = -1
            Dim _NewColumnIndex As Integer = 0

            Try
                Dim ExcelArray(,) As Object = CType(xlUsedRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault), Object(,))
                If ExcelArray IsNot Nothing Then
                    Dim _Rows As Integer = ExcelArray.GetUpperBound(0)
                    Dim _Columns As Integer = ExcelArray.GetUpperBound(1)
                    ReDim Preserve NewExcelArray(_Rows, _Columns)
                    For _Row As Integer = 1 To _Rows
                        Dim _ReadInd As Boolean = False
                        If CStr(ExcelArray(_Row, 1)) & "" <> "" Then
                            Dim _CheckStr As String = CStr(ExcelArray(_Row, 1)).Trim

                            'If _CheckStr.Trim.ToLower.StartsWith("employee") Or _CheckStr.Trim.ToLower.StartsWith("corporation") Then
                            'Else
                            '    _ReadInd = True
                            '    _NewRowIndex = _NewRowIndex + 1
                            'End If
                            _ReadInd = True
                            _NewRowIndex = _NewRowIndex + 1
                        End If
                        If _ReadInd Then
                            _NewColumnIndex = 0
                            For _Column As Integer = 1 To _Columns
                                Dim a As Excel.Range = xlWorkSheet.Cells(_Row, _Column)
                                Dim _CellData As String = CStr(ExcelArray(_Row, _Column)) & ""
                                If _NewRowIndex >= 0 Then
                                    If _NewRowIndex = 0 Then
                                        NewExcelArray(_NewRowIndex, _NewColumnIndex) &= _CellData.Trim
                                    Else
                                        NewExcelArray(_NewRowIndex, _NewColumnIndex) &= _CellData.Trim & Environment.NewLine
                                    End If
                                    _NewColumnIndex = _NewColumnIndex + 1
                                End If
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                Throw ex
            Finally
                ReleaseComObject(xlUsedRange)
            End Try
            Return NewExcelArray
        End Function

        Public Function GetDataTableFrom2DArray(ByVal array As Object(,)) As Data.DataTable
            Dim _DataTable As New Data.DataTable()
            Try
                For i As Integer = 0 To array.GetUpperBound(1) - 1
                    _DataTable.Columns.Add(IIf(array(0, i).ToString.Trim.Length > 0, array(0, i).ToString.Trim, " "), Type.GetType("System.Object"))
                Next

                For j As Integer = 1 To array.GetUpperBound(0) - 1
                    Dim row As Object() = New Object(array.GetUpperBound(1) - 1) {}
                    Dim _RowInd As Boolean = False
                    For k As Integer = 0 To row.Length - 1
                        If Not IsNothing(array(j, k)) Then
                            _RowInd = True
                        End If
                        row(k) = array(j, k)
                    Next
                    If _RowInd Then
                        _DataTable.Rows.Add(row)
                    End If
                Next
            Catch ex As Exception
                Throw ex
            End Try

            Return _DataTable
        End Function
        Public Function GetDataTableFromArray(ByVal array As Object()) As Data.DataTable
            Dim _DataTable As New Data.DataTable()
            _DataTable.LoadDataRow(array, True)
            Return _DataTable
        End Function

        Public Sub ReleaseComObject(ByVal obj As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            End Try
        End Sub
    End Class
End Namespace
