<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

' Environment.asp' This file contains settings that are specific to the environment being used' (Development,Testing,Production, etc)' *********************************************************************************
' THIS IS THE ONLY FILE THAT SHOULD BE DIFFERENT BETWEEN DEVELOPMENT AND PRODUCTION
' ENVIRONMENTS!
' *********************************************************************************
'
' PRODUCTION
'' Set development state
Application("IntranetDevelopment") = false

' Database settingsApplication("DataConn_ConnectionString") = "DSN=Sark;User Id=IntraProd_User;PASSWORD=intraprod;UID=IntraProd_User;APP=Microsoft (R) Developer Studio;WSID=WINBOOKXL;DATABASE=IntranetProd"
Application("DataConn_ConnectionTimeout") = 15
Application("DataConn_CommandTimeout") = 30
Application("DataConn_CursorLocation") = 3
Application("DataConn_RuntimeUserName") = "IntraProd_User"
Application("DataConn_RuntimePassword") = "intraprod"

'Exchange ServerApplication("ExchangeServer") = "EDDY" ' EDDY is production Exchange Server

' File Paths
Application("WebRootDir")         = "d:\inetpub\"Application("401KFile")           = "d:\inetpub\intranet\content\stock.txt"
Application("WeatherFile")        = "d:\inetpub\intranet\content\weather.txt"
Application("CNETFile")           = "d:\inetpub\intranet\content\cnet.txt"
Application("AnnouncementsFile")  = "d:\inetpub\intranet\welcome\announcements\announcements1.txt"Application("Announcements2File") = "d:\inetpub\intranet\welcome\announcements\announcements2.txt"</SCRIPT>