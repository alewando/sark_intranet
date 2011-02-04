<!--
Developer:    Barry Behrmann
Date:         12/20/1999
Description:  Submit new news articles
-->



<%
  
' SET THE USER NAME AND EXCHANGE BOX NAME
  strUserName = "pgrimes"
  strServerName = "EDDY"

'DON'T MUCK WITH THIS UNLESS YOU KNOW WHAT YOU'RE IN FOR!
  set objSession = nothing
  Set objSession = Server.CreateObject("mapi.session") 
  
  if not objSession is nothing then
  On Error resume next
'    objSession.Logon "pgrimes/sacincy", "Dartagnan", False, True, 0, True    ' log the user onto the server
    objSession.Logon "", "", False, True, 0, True, strServerName & vbLf & strUserName     ' log the user onto the server

Response.Write(":" & objSession.currentuser & ":")
Response.Write(":" & objSession.name & ":")


    set objInbox = objSession.Inbox
' obtain the Inbox
    if objInbox is nothing then
    else
      set objMessColl = objSession.Inbox.Messages
' message collection = messages in the Inbox
      set objMessFilter = objMessColl.filter
' filter on the collection
      objMessFilter.unread = true
' set the filters to whatever you want
      Response.Write("You have " & objMessColl.count & " new messages.")
' write out how many messages there are
      set objMessFilter = nothing
      set objMessColl = nothing
    end if
      objSession.Logoff
  end if
  set objInbox = nothing
  set objSession = nothing 
 
%>

end of test
