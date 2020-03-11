Imports log4net

Public Class Logger
    Private Shared logger As ILog = LogManager.GetLogger("FileAppender")

    Public Shared Sub LogInfo(ByVal str As String)
        log4net.ThreadContext.Properties("user") = loggedInUser
        logger.Info(str)
    End Sub

    Public Shared Sub LogError(ByVal str As String)
        log4net.ThreadContext.Properties("user") = loggedInUser
        logger.Error(str)
    End Sub

    Public Shared Sub LogFatal(ByVal str As String, ex As Exception)
        log4net.ThreadContext.Properties("user") = loggedInUser
        logger.Fatal(str, ex)
    End Sub
End Class
