<%
'------------------------
'  Initialize variables  
'------------------------
dim DataConn						' Generic database connection
dim sectionWebMaster, ccWebMaster	' Webmaster email addresses
dim pageTitle						' Page title (in drop down)


'------------------------
'  Initialize constants  
'------------------------
NL = chr(13) & chr(10)
sectionNav = ""
sectionImg = "backgrd2.jpg"
CurrentNavPage = lcase(request.ServerVariables("PATH_INFO"))
if request.ServerVariables("QUERY_STRING") <> "" then CurrentNavPage = CurrentNavPage & "?" & request.ServerVariables("QUERY_STRING")
CurrentNavPage = lcase(CurrentNavPage)
DefaultNavPage = CurrentNavPage

ccWebMaster = ""
sectionWebMaster = getWebMaster()


'-----------------------
'  Set section buttons  
'-----------------------
SectionBtns = "<form><center><hr>" & NL & _
	"<input type=button class=btnSetDefault name=setdefault value='Set Default' onClick=" & chr(34) & "setCookie('Page', curPage); alert('Your default page for the ' + curSection + ' section has been set to the current page.'); return false;" & chr(34) & " onMouseOver=" &chr(34)& " statmsg('Sets this as your default page for this section'); return true;" & chr(34) & "onMouseOut=" & chr(34) & "statmsg('')"& chr(34) & ">" & NL & _
	"<hr><p><input type=button class=button value='Preferences' onClick=" &chr(34)& "window.location='../../preferences/content/default.asp'" &chr(34)& " onMouseOver=" &chr(34)& " statmsg('View and set your preferences'); return true;" & chr(34) & "onMouseOut=" & chr(34) & "statmsg('')"& chr(34) & " id=button1 name=button1>" & NL & _
	"<p><input type=button class=button value='Quick email' onClick=" & chr(34) & "winTech=window.open('../../email.asp?editto=yes&footer=" & Server.URLEncode("INTRANET: Quick Email") & "', 'SendEmail','height=330,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1'); return false;" & chr(34) & " onMouseOver=" &chr(34)& " statmsg('Send quick email to another Sark'); return true;" & chr(34) & "onMouseOut=" & chr(34) & "statmsg('')"& chr(34) & ">" & NL & _
	"<p><input type=button class=button value='Feedback' onClick='SendFeedback()' onMouseOver=" &chr(34)& " statmsg('Send feedback email to webmaster(s).'); return true;" & chr(34) & "onMouseOut=" & chr(34) & "statmsg('')"& chr(34) & " id=button1 name=button1>" & NL & _
	"</center></form>" & NL


'----------------------
'  Debugging code      
'----------------------
if Application("debug") then
	response.write(NL & NL & "<!-- DEBUG:" & NL)
	
	response.write("Variables:" & NL)
	response.write("   Session('ID') = " & Session("ID") & NL)
	response.write("   Session('Username') = '" & Session("Username") & "'" & NL)
	response.write("   CurrentNavPage = '" & CurrentNavPage & "'" & NL)
	
'	response.write("Application:" & NL)
'	for each x in application.contents
'		response.write("   " & lcase(x) & " = " & application(x) & NL)
'		next
'	
'	response.write("Session:" & NL)
'	for each x in session.contents
'		response.write("   " & lcase(x) & " = " & session(x) & NL)
'		next
'	
	response.write("Request.Form:" & NL)
	for each x in request.form
		response.write("   " & x & " = " & request.form(x) & NL)
		next
	
	response.write("Request.Querystring:" & NL)
	for each x in request.querystring
		response.write("   " & x & " = " & request.querystring(x) & NL)
		next
	
	response.write("Cookies:" & NL)
	for each cookie in request.cookies
		if not request.cookies(cookie).HasKeys then
			response.write("   " & cookie & " = " & request.cookies(cookie) & NL)
		else
			for each key in request.cookies(cookie)
				response.write("   " & cookie & " (" & key & ") = " & request.cookies(cookie)(key) & NL)
				next
		end if
		next
	response.write("-->" & NL & NL & NL)
end if


'----------------------
'  Navigation builder  
'----------------------
sub buildNavItem(Item, Url)
	'--------------------------------------
	'  This routine builds the navigation  
	'  select combo box entries            
	'--------------------------------------
	SectionNav = SectionNav & "			<option value='" & Url & "'"
	Url = lcase("/" & Application("Web") & "/" & sectionDir & "/content/" & Url)
	if left(CurrentNavPage, len(Url)) = Url then
		SectionNav = SectionNav & " selected"
		pageTitle = Item
	else
		'SectionNav = SectionNav & " url='" & Url & "' CurrentNavPage='" & CurrentNavPage & "' "
	end if
	SectionNav = SectionNav & ">" & Item & "</option>" & NL
end sub


'-------------------------------------------------------
'  This procedure sets the webmaster to a lower detail  
'  (ex. the section web master) and copies anything in  
'  sectionWebMaster to ccWebMaster.                     
'-------------------------------------------------------
sub SetWebMaster(webmaster)
	if ccWebMaster <> "" then ccWebMaster = ccWebMaster & ","
	ccWebMaster = ccWebMaster & sectionWebMaster
	sectionWebMaster = webmaster
end sub


'-------------------------------------------
'  This procedure sends an email based      
'  upon parameters passed to the function.  
'-------------------------------------------

sub SendMail (Subject, Recipients, CC, BodyTxt)
	dim val, i, email
	dim strTo, strCc
	

	Set CDONTSmail = CreateObject("CDONTS.NewMail")
		
	if session("Username") <> "" then CDONTSmail.From = Session("Username") & "@sark.com"
	'The following line was discovered not to work on 4-6-2000
	'if Session("Name") <> "" then CDONTSmail.From = Session("Name")
	CDONTSmail.Subject = Subject

	val = Recipients
	i = InStr(val, ",")
	while i > 0
		email = left(val, i-1)
		if InStr(email, "@") = 0 then email = email & "@sark.com"
		strTo = strTo & email & ";"
		val = trim(mid(val, i+1))
		i = InStr(val, ",")
		wend
	if val <> "" then
		if InStr(val, "@") = 0 then val = val & "@sark.com"
		strTo = strTo & val
		end if

	CDONTSmail.To = strTo	
		
	val = CC
	i = InStr(val, ",")
	while i > 0
		email = left(val, i-1)
		if InStr(email, "@") = 0 then email = email & "@sark.com"
		strCc = strCc & email & ";"
		val = trim(mid(val, i+1))
		i = InStr(val, ",")
		wend
	if val <> "" then
		if InStr(val, "@") = 0 then val = val & "@sark.com"
		strCc = strCc & val
		end if
	
	CDONTSmail.Cc = strCc	
	CDONTSmail.Body = BodyTxt
	CDONTSmail.Send
	
	set CDONTSmail = nothing
	
end sub


'------------------------------------------------------------------
' This function replaces all '+' characters in a URLEncoded string
' into the '%20' string. There have been problems when attempting to
' download a document, when its name has embedded spaces i.e. "This File.doc". 
' The URLEncode method properly converts the spaces to '+' characters
' i.e. "This+File%2Edoc". However, the web server can't find the encoded document. 
' The web server is able to find "This%20File%2Edoc".
'----------------------------------------------------------------
function PlusToHex(val)
	PlusToHex = Replace(val, "+", "%20")
end function


'----------------------------------------------------------------------------------
' Determines if the current user is the Branch manager, in sales or an account mgr.
' If so, they have permission to use the consultant skills search page.
'----------------------------------------------------------------------------------
'function CanSearchSkills
'	'If IsNull(Session("sessCanSearchSkills")) Then
'
'		sql = "SELECT e.employeeid FROM employee e, Employee_Title et " _
'			& "WHERE e.Username = '" & Lcase(session("Username")) & "' AND (e.Username = 'cdolan' OR e.Username = 'kdolan') " _
'			& " AND e.EmployeeTitle_ID = et.Employee_Title_ID " _
'			& " AND (title_function_id IN ('A','S') OR (title_function_id ='C' AND title_level >= 5)) "
'		set rs = DBQuery(sql)
'
'		If Not rs.eof Then
'			Session("sessCanSearchSkills")=True
'		Else
'			Session("sessCanSearchSkills")=False
'		End If
'		
'		rs.close
'		set rs=Nothing
'	'End If
'
'	If Lcase(session("Username"))="dpodnar" Then
'		Session("sessCanSearchSkills")=True	' Used for production testing
'	End If
'
'	CanSearchSkills = Session("sessCanSearchSkills")
'end function


'-------------------------------------
'  This procedure opens the generic   
'  database connection as necessary.  
'-------------------------------------
sub DBConnect
	if isempty(DataConn) then
		set DataConn = Server.CreateObject("ADODB.Connection")
		DataConn.ConnectionTimeout = Application("DataConn_ConnectionTimeout")
		DataConn.CommandTimeout = Application("DataConn_CommandTimeout")
		DataConn.Open Application("DataConn_ConnectionString"), Application("DataConn_RuntimeUserName"), Application("DataConn_RuntimePassword")
		if Application("debug") then response.write(NL & "<!-- DEBUG: Database connection opened -->" & NL)
		end if
end sub


'----------------------------------------------------------
'  This function strips out invalid characters that could  
'  disrupt the SQL statement when executing.               
'----------------------------------------------------------
function Clean(val)
	Clean = Replace(val, "'", "''")
end function


'-----------------------------------------------------------
'  This function executes a sql query against a database,   
'  creating a new database connection as necessary.         
'  It returns a recordset object as a result of the query.  
'-----------------------------------------------------------
function DBQuery(sql)
	DBConnect
	if Application("debug") then Response.Write(NL & "<!-- DEBUG: Executing sql = " & chr(34) & sql & chr(34) & " -->" & NL)
	set DBQuery = DataConn.execute(sql)
end function



'---------------------------------------------------------------
'  This function executes a sql query (once per day, per key)   
'  and caches it for later user (in the Application object).    
'  This is useful to reduce the number of database hits on      
'  common queries with few data updates.                        
'---------------------------------------------------------------
function DBCachedQuery(key, sql)
	
	dim UpdateRecs
	UpdateRecs = not (Application(key & "_date") = Date)
	
	if UpdateRecs then
		
		DBConnect
		Set rsCustomRec = Server.CreateObject("ADODB.Recordset")
		rsCustomRec.CursorLocation = adUseClient	' This is usable disconnected
		rsCustomRec.Open sql, DataConn, adOpenStatic, AdLockReadOnly
		rsCustomRec.ActiveConnection = Nothing		' Disconnect the Recordset
		
		Application.Lock
		Set Application(key) = rsCustomRec
		Application(key & "_date") = Date
		Application.Unlock
		end if
		
	Set DBCachedQuery = Application(key).Clone
	
end function



'--------------------------------------------
'  This procedure clears a cached recordset  
'  from the application object.              
'--------------------------------------------
sub DBClearCachedQuery(key)
	Application.Lock
	set Application(key) = nothing
	Application(key & "_date") = ""
	Application.UnLock
end sub

'---------------------------------------------
' Security functions
'---------------------------------------------
' hasRole returns true if the current user has been assigned the specified role
Function hasRole(role)
 'Guest user has no rights
 if isNull(Session("ID")) or Session("ID") = "" or Session("isGuest") then
  hasRole = false
  Exit Function
 end if
 'Check for webmaster (has all privileges)
 sql="SELECT COUNT(*) as roleCount FROM Security_User_Roles ur, Security_Roles r " & _
     "WHERE " & _
     "ur.role_id = r.role_id " & _
     "AND rtrim(upper(r.role_name)) = 'WEBMASTER' " & _
     "AND employee_id=" & Session("ID")
 set rsRole = DBQuery(sql)
 if not IsNull(rsRole) and not IsNull(rsRole("roleCount")) and rsRole("roleCount") <> "" and rsRole("roleCount") > 0 then
  hasRole = true
  Exit Function
 end if 
 
 ' Check for normal grants
 sql="SELECT COUNT(*) as roleCount FROM Security_User_Roles ur, Security_Roles r " & _
     "WHERE " & _
     "ur.role_id = r.role_id " & _
     "AND rtrim(upper(r.role_name)) = '" & trim(ucase(role)) & "' " & _
     "AND employee_id=" & Session("ID")
 set rsRole = DBQuery(sql)
 if not IsNull(rsRole) and not IsNull(rsRole("roleCount")) and rsRole("roleCount") <> "" and rsRole("roleCount") > 0 then
  hasRole = true
 else
  hasRole = false
 end if
End Function

' Returns the first user with WebMaster privileges
Function getWebMaster
 sql = "SELECT username FROM Employee e, Security_User_Roles ur, Security_Roles r " & _
       "WHERE e.employeeid = ur.employee_id " & _
       "AND ur.role_id = r.role_id " & _
       "AND upper(r.role_name)='WEBMASTER' " & _
       "ORDER BY e.employeeid"
 set rsWM = DBQuery(sql)
 if not rsWM.eof then
  getWebMaster = rsWM("username")
 else
  getWebMaster = ""
 end if
End Function
%>
