<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

' Environment.asp
' THIS IS THE ONLY FILE THAT SHOULD BE DIFFERENT BETWEEN DEVELOPMENT AND PRODUCTION
' ENVIRONMENTS!
' *********************************************************************************
'
' PRODUCTION
'
Application("IntranetDevelopment") = false

' Database settings
Application("DataConn_ConnectionTimeout") = 15
Application("DataConn_CommandTimeout") = 30
Application("DataConn_CursorLocation") = 3
Application("DataConn_RuntimeUserName") = "IntraProd_User"
Application("DataConn_RuntimePassword") = "intraprod"

'Exchange Server

' File Paths
Application("WebRootDir")         = "d:\inetpub\"
Application("WeatherFile")        = "d:\inetpub\intranet\content\weather.txt"
Application("CNETFile")           = "d:\inetpub\intranet\content\cnet.txt"
Application("AnnouncementsFile")  = "d:\inetpub\intranet\welcome\announcements\announcements1.txt"