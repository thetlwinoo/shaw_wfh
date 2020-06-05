Namespace Email
    <Serializable()>
    Public Class MailParameters

#Region "Local Variable"
        Private mSMTPHost As String
        Private mFromAddress As String
        Private mFromName As String
        Private mPassword As String
        Private mSMTPPort As Integer
        Private mEnableSSL As Boolean
        Private mRetryCount As Integer
        Private mToAddress As String
        Private mToName As String
        Private mSubject As String
        Private mBody As String
        Private mCcAddress As String
        Private mBccAddress As String
        Private mMailPriority As Integer
        Private mDeliveryNotificationOptions As Integer
        Private mLogFilePath As String
        Private mTemplateFilePath As String
        Private mShawStaffFilePath As String
        Private mTemparatureFilePath As String
        Private mErrorFlag As Boolean
#End Region

#Region "Properties"
        Public Property SMTPHost()
            Get
                Return mSMTPHost
            End Get
            Set(ByVal value)
                mSMTPHost = value
            End Set
        End Property
        Public Property FromAddress()
            Get
                Return mFromAddress
            End Get
            Set(ByVal value)
                mFromAddress = value
            End Set
        End Property
        Public Property FromName()
            Get
                Return mFromName
            End Get
            Set(ByVal value)
                mFromName = value
            End Set
        End Property
        Public Property Password()
            Get
                Return mPassword
            End Get
            Set(ByVal value)
                mPassword = value
            End Set
        End Property
        Public Property SMTPPort()
            Get
                Return mSMTPPort
            End Get
            Set(ByVal value)
                mSMTPPort = value
            End Set
        End Property
        Public Property EnableSSL()
            Get
                Return mEnableSSL
            End Get
            Set(ByVal value)
                mEnableSSL = value
            End Set
        End Property
        Public Property RetryCount()
            Get
                Return mRetryCount
            End Get
            Set(ByVal value)
                mRetryCount = value
            End Set
        End Property
        Public Property ToAddress()
            Get
                Return mToAddress
            End Get
            Set(ByVal value)
                mToAddress = value
            End Set
        End Property
        Public Property ToName()
            Get
                Return mToName
            End Get
            Set(ByVal value)
                mToName = value
            End Set
        End Property
        Public Property Subject()
            Get
                Return mSubject
            End Get
            Set(ByVal value)
                mSubject = value
            End Set
        End Property
        Public Property Body()
            Get
                Return mBody
            End Get
            Set(ByVal value)
                mBody = value
            End Set
        End Property
        Public Property CcAddress()
            Get
                Return mCcAddress
            End Get
            Set(ByVal value)
                mCcAddress = value
            End Set
        End Property
        Public Property BccAddress()
            Get
                Return mBccAddress
            End Get
            Set(ByVal value)
                mBccAddress = value
            End Set
        End Property
        Public Property MailPriority()
            Get
                Return mMailPriority
            End Get
            Set(ByVal value)
                mMailPriority = value
            End Set
        End Property
        Public Property DeliveryNotificationOptions()
            Get
                Return mDeliveryNotificationOptions
            End Get
            Set(ByVal value)
                mDeliveryNotificationOptions = value
            End Set
        End Property
        Public Property LogFilePath()
            Get
                Return mLogFilePath
            End Get
            Set(ByVal value)
                mLogFilePath = value
            End Set
        End Property
        Public Property TemplateFilePath()
            Get
                Return mTemplateFilePath
            End Get
            Set(ByVal value)
                mTemplateFilePath = value
            End Set
        End Property
        Public Property TemparatureFilePath()
            Get
                Return mTemparatureFilePath
            End Get
            Set(ByVal value)
                mTemparatureFilePath = value
            End Set
        End Property
        Public Property ShawStaffFilePath()
            Get
                Return mShawStaffFilePath
            End Get
            Set(ByVal value)
                mShawStaffFilePath = value
            End Set
        End Property
        Public Property ErrorFlag()
            Get
                Return mErrorFlag
            End Get
            Set(ByVal value)
                mErrorFlag = value
            End Set
        End Property

#End Region

    End Class
End Namespace
