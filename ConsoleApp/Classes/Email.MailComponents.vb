Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports System.Data
Imports ConsoleApp.Common

Namespace Email
    Public Class MailComponents
#Region "Class Constants and Enumerations"
        Public Enum EmailFormat
            Text = 1
            Html = 0
        End Enum
        Public Enum ReportType
            Process
            Success
            Failure
        End Enum
        Public Enum ReportingStage
            Initialization
            Fetching
            Formatting
            Sending
            Finalizing
            Resending
        End Enum
#End Region

#Region "Public Shared Variables"
        Public Shared MailComponentParms As New Email.MailParameters
        Public Shared _DataRow As DataRow
        Public Shared _DataColumnCollection As DataColumnCollection
#End Region

#Region "Class Variables and Objects"
        Public Delegate Sub DelegateReporting(ByVal pReportType As ReportType, ByVal pReportingStage As ReportingStage, ByVal pMessage As String, ByVal curLeftVal As Nullable(Of Integer), ByVal curTopVal As Nullable(Of Integer), ByVal pDisplayStageInd As Boolean)
        Public Delegate Function DelegateReadFromFile(ByVal pFilePath As String) As Dictionary(Of String, String)
        Public Delegate Function DelegateReadFromExcel(ByVal pFilePath As String) As DataTable
        Public Delegate Function DelegateGenerateMailBody(ByVal pTemplateFilePath As String, ByVal pShawStaffFilePath As String, ByVal pTemperatureFilePath As String) As Dictionary(Of String, String)

        Public CallReporting As DelegateReporting
        Public CallReadFromFile As DelegateReadFromFile
        Public CallReadFromExcel As DelegateReadFromExcel
        Public CallGenerateMailBody As DelegateGenerateMailBody

        Private mPassword As String = String.Empty
        Private mSMTPHost As String = String.Empty
        Private mSMTPPort As Int16
        Private mEnableSSL As Boolean
        Private mFrom As MailAddress
        Private mTo As New MailAddressCollection
        Private mCc As New MailAddressCollection
        Private mBcc As New MailAddressCollection
        Private mFromName As String = String.Empty
        Private mToName As String = String.Empty
        Private mSubject As String = String.Empty
        Private mBody As String = String.Empty
        Private mAttachedImages As List(Of AttachedImage)
        Private mMessage As String = String.Empty
        Private mReportParameters As String = String.Empty
        Private mMailPriority As MailPriority
        Private mDeliveryNotificationOptions As DeliveryNotificationOptions
        Private mIsBodyHtml As Boolean
        Private mSendMailInd As Boolean
        Private mCount As Integer
        Private mRound As Integer
        Private mRowCount As Integer
        Private mRetryCount As Integer
        Public intEmailRetryCount As Integer = 0
#End Region

#Region "Class Event Handlers"
        Public Event onSkip(ByVal ex As SmtpException, ByVal dr As DataRow, ByVal dcc As DataColumnCollection)
        Public Event onError(ByVal ex As Exception, ByVal Stage As ReportingStage, ByVal ReportParameters As String, ByVal TotalRecord As Integer, ByVal CurrentRow As Integer)
#End Region

#Region "Class Functions"
        Private Function LinkImages(ByVal Message As String, ByVal IsHtmlInd As Boolean) As MailMessage
            Dim msg = New MailMessage()
            mAttachedImages = New List(Of AttachedImage)
            If Not (IsHtmlInd = True) Then
                msg.IsBodyHtml = False
                'first create the Plain Text part
                Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(
                                                 MailComponentParms.Body,
                                                 Nothing, "text/plain")
                msg.AlternateViews.Add(plainView)
                mBody = MailComponentParms.Body
            Else

                msg.IsBodyHtml = True
                Try

                    Dim matches = Regex.Matches(Message, "<img[^>]*?src\s*=\s*([""']?[^'"">]+?['""])[^>]*?>",
                                                RegexOptions.IgnoreCase Or
                                                RegexOptions.IgnorePatternWhitespace Or
                                                RegexOptions.Multiline)

                    Dim img_list = New List(Of LinkedResource)()
                    Dim cid = 1
                    For Each _Match As Match In matches
                        Dim src As String = _Match.Groups(1).Value
                        src = src.Trim(""""c)
                        src = src.Trim("'")

                        If File.Exists(MailComponentParms.TemplateFilePath) Then
                            'Dim MyFile As String = Path.GetDirectoryName(MailComponentParms.FilePath) + "\" + src
                            Dim MyFile As String = src
                            Dim _AttachedImage As New AttachedImage
                            If File.Exists(MyFile) Then
                                Dim ext = Path.GetExtension(src)
                                If ext.Length > 0 Then
                                    ext = ext.Substring(1)
                                    Dim res As LinkedResource = New LinkedResource(MyFile)
                                    res.ContentId = String.Format("img{0}",
                                    System.Math.Min(System.Threading.Interlocked.Increment(cid), cid - 1), ext)
                                    _AttachedImage.ContentID = res.ContentId
                                    _AttachedImage.ContentLink = MyFile
                                    res.TransferEncoding = System.Net.Mime.TransferEncoding.Base64
                                    res.ContentType.MediaType = String.Format("image/{0}", ext)
                                    res.ContentType.Name = res.ContentId
                                    img_list.Add(res)
                                    src = String.Format("'cid:{0}'", res.ContentId)
                                    Message = Message.Replace(_Match.Groups(1).Value, src)
                                End If
                            End If
                            mBody = Message
                            mAttachedImages.Add(_AttachedImage)
                        Else
                            Console.WriteLine("image source path is missing")
                        End If
                    Next

                    'then we create the Html part
                    Dim HtmlView As AlternateView = AlternateView.CreateAlternateViewFromString(Message,
                                                                                Nothing,
                                                                                System.Net.Mime.MediaTypeNames.Text.Html)

                    For Each img As LinkedResource In img_list
                        HtmlView.LinkedResources.Add(img)
                    Next

                    msg.AlternateViews.Add(HtmlView)

                Catch ex As Exception
                    RaiseEvent onError(ex, ReportingStage.Formatting, mReportParameters, Nothing, Nothing)
                End Try
            End If

            Return msg
        End Function

        Public Sub SendMailMessage()
            Dim _Body As String = String.Empty
            Dim _ReturnResult As New Dictionary(Of String, String)
            Dim _Dictionary As New Dictionary(Of Integer, Dictionary(Of String, String))
            Dim _SmtpException As New SmtpException
            Dim _Count As Integer = 0

            With MailComponentParms
                mPassword = .Password
                mSMTPHost = .SMTPHost
                mSMTPPort = .SMTPPort
                mEnableSSL = .EnableSSL

                mFrom = New MailAddress(.FromAddress, .FromName)
                mFromName = .FromName
                mSubject = .Subject
                mToName = .ToName
                mMailPriority = .MailPriority
                mDeliveryNotificationOptions = .DeliveryNotificationOptions
                mCount = 0
                mRowCount = 0

                ' ''To
                If Not .ToAddress.Trim.Equals("") Then
                    For Each _Address As String In .ToAddress.Split(";")
                        mTo.Add(New MailAddress(_Address))
                    Next
                End If

                ' ''Cc
                If Not .CcAddress.Trim.Equals("") Then
                    For Each _Address As String In .CcAddress.Split(";")
                        mCc.Add(New MailAddress(_Address))
                    Next
                End If
                ' ''Bcc
                If Not .BccAddress.Trim.Equals("") Then
                    For Each _Address As String In .BccAddress.Split(";")
                        mBcc.Add(New MailAddress(_Address))
                    Next
                End If
                ' // Read MessageBody From File
                '_ReturnResult = CallReadFromFile.Invoke(MailComponentParms.TemplateFilePath)
                _ReturnResult = CallGenerateMailBody.Invoke(MailComponentParms.TemplateFilePath, MailComponentParms.ShawStaffFilePath, MailComponentParms.TemparatureFilePath)
                If _ReturnResult.ContainsKey("Error") Then
                    mMessage = _ReturnResult.Item("Error")
                    RaiseEvent onError(New Exception(mMessage), ReportingStage.Fetching, mReportParameters, Nothing, Nothing)
                    Exit Sub
                ElseIf _ReturnResult.ContainsKey("Success") Then
                    _Body = _ReturnResult.Item("Success")
                    Dim _Pattern As String = "</?\w+((\s+\w+(\s*=\s*(?:"".*?""|'.*?'|[^'"">\s]+))?)+\s*|\s*)/?>"    '"(?<=<(\w+)>).*(?=<\/\1>)|(?<=(<\w+[ ]\s?\w*\W([^>]+)>)).*(?=<\/\w+>)"
                    If Utilities.RegexValidator(New String() {_Pattern}, _Body) Is String.Empty Then
                        mIsBodyHtml = False
                    Else
                        mIsBodyHtml = True
                    End If
                End If

                SendEmail(_Count, mFrom.ToString, mFromName, mTo.ToString, mToName, mSubject, mBody, mCc.ToString, mBcc.ToString, EmailFormat.Html, mMailPriority)
            End With
        End Sub

        Public Function SendEmail(ByVal pCount As Integer, ByVal strFromAddress As String, ByVal strFromName As String, ByVal strToAddress As String, ByVal strToName As String, ByVal strSubject As String, ByVal strBody As String, Optional ByVal strCcList As String = "", Optional ByVal strBccList As String = "", Optional ByVal enumFormat As EmailFormat = EmailFormat.Text, Optional ByVal enumPriority As MailPriority = MailPriority.Normal, Optional ByVal enumNotifyOption As DeliveryNotificationOptions = DeliveryNotificationOptions.None, Optional ByVal strEmailServiceCode As String = "", Optional ByVal intMemberID As Integer = -1, Optional ByVal strLogRemarks As String = "") As Integer
            Dim objSMTP As New SmtpClient(My.Settings.MailServer)
            Dim intErrorCode As Integer = -1

            If File.Exists(My.Settings.LogFilePath) Then
            Else
                Directory.CreateDirectory(My.Settings.LogFilePath)
            End If

            Dim sw As New StreamWriter(My.Settings.LogFilePath & Now.ToString("yyyy-MM-dd HHmmssfff") & ".txt", True)
            Try
                If strFromAddress.Trim.Equals("") Then
                    If Not My.Settings.FromAddress.Trim.Equals("") Then
                        strFromAddress = My.Settings.FromAddress.Trim
                    End If
                End If
                If strFromName.Trim.Equals("") Then
                    If Not My.Settings.FromName.Trim.Equals("") Then
                        strFromName = My.Settings.FromName.Trim
                    End If
                End If
                Dim objMailFrom As New MailAddress(strFromAddress, strFromName)
                Dim objMailTo As New MailAddress(strToAddress.Split(New Char() {";"})(0).ToString, strToName.Split(New Char() {";"})(0).ToString)
                Dim objMessage As New MailMessage(objMailFrom, objMailTo)
                objMessage.Subject = strSubject
                objMessage.Body = strBody

                If Not strToAddress.Trim.Equals("") Then
                    Dim arrToList As Array = strToAddress.Split(New Char() {";"})
                    Dim arrToListName As Array = strToName.Split(New Char() {";"})
                    Dim strName As String
                    For intToCount As Integer = 1 To arrToList.Length - 1
                        If arrToListName.Length - 1 >= intToCount Then
                            strName = arrToList.GetValue(intToCount).ToString
                        Else
                            strName = ""
                        End If
                        objMessage.To.Add(New MailAddress(arrToList.GetValue(intToCount).ToString, strName))
                    Next
                End If

                If Not strCcList.Trim.Equals("") Then
                    Dim arrCCList As Array = strCcList.Split(New Char() {";"})

                    For intCCCount As Integer = 0 To arrCCList.Length - 1
                        objMessage.CC.Add(arrCCList.GetValue(intCCCount).ToString)
                    Next
                End If

                If Not strBccList.Trim.Equals("") Then

                    Dim arrBccList As Array = strBccList.Split(New Char() {";"})

                    For intBccCount As Integer = 0 To arrBccList.Length - 1
                        objMessage.Bcc.Add(arrBccList.GetValue(intBccCount).ToString)
                    Next
                End If

                objMessage.DeliveryNotificationOptions = enumNotifyOption
                objMessage.Priority = enumPriority
                If enumFormat = 0 Then
                    objMessage.IsBodyHtml = True
                ElseIf enumFormat = 1 Then
                    objMessage.IsBodyHtml = False
                End If

                ' added by thet lwin (2016 June)
                For Each _AttachImage As AttachedImage In mAttachedImages
                    Dim _AttachementImage As Attachment = New Attachment(_AttachImage.ContentLink)
                    With _AttachementImage
                        .ContentId = _AttachImage.ContentID
                        .ContentDisposition.Inline = True
                        objMessage.Attachments.Add(_AttachementImage)
                    End With
                Next

                objSMTP.Send(objMessage)
                Try
                    If Not IsNothing(My.Settings.LogFilePath) Then
                        Console.Write(pCount & " >> " & vbTab)
                        Console.Write("To:" & strToAddress & vbTab)
                        Console.Write("Timestamp:" & DateTime.Now())
                        Console.WriteLine()

                        sw.Write(pCount & ")" & vbTab & vbTab)
                        sw.Write("To:" & strToAddress & vbTab & vbTab)
                        sw.Write("From:" & strFromAddress & vbTab & vbTab)
                        sw.Write("Timestamp:" & DateTime.Now() & vbTab)
                        sw.WriteLine()
                    End If
                Catch InnerEx As Exception

                End Try
            Catch ex As Exception
                If intEmailRetryCount < My.Settings.RetryCount And enumPriority = MailPriority.High Then
                    intEmailRetryCount += 1
                    Return SendEmail(pCount, strFromAddress, strFromName, strToAddress, strToName, strSubject, strBody, strCcList, strBccList, enumFormat, enumPriority, enumNotifyOption, strEmailServiceCode, intMemberID, strLogRemarks)
                Else
                    intEmailRetryCount = 0
                    Return intErrorCode
                End If
            Finally
                objSMTP = Nothing
                sw.Flush()
                sw.Close()
            End Try
        End Function
#End Region
    End Class

    Public Class AttachedImage
        Private mContentID As String
        Private mContentLink As String

        Public Property ContentID() As String
            Get
                Return mContentID
            End Get
            Set(ByVal value As String)
                mContentID = value
            End Set
        End Property
        Public Property ContentLink() As String
            Get
                Return mContentLink
            End Get
            Set(ByVal value As String)
                mContentLink = value
            End Set
        End Property
    End Class
End Namespace

