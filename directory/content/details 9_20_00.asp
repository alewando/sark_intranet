<!--
Developer:		GINGRICE
Date:			08/18/1998
Description:	Displays detailed listing of consultants
History:		12/11/1998 - KDILL - Added support for email display.-->


<% pageTitle = "Details" %>

<!-- #include file="../section.asp" -->

<%'For testing
'session("GotAcctMgr")= true 
%>
<SCRIPT LANGUAGE=javascript>
<!--

	EmpID = <%=Request.QueryString("EmpID")%>
	function fnTabClick( nTab )
	{
		nTab = parseInt(nTab);
		var oTab;
		var prevTab = nTab-1;
		var nextTab = nTab+1;
		event.cancelBubble = true;
		el = event.srcElement;

		for (var i = 0; i < newsContent.length; i++)
		{
			oTab = tabs[i];
			oTab.className = "clsTab";
			oTab.style.borderLeftStyle = "";
			oTab.style.borderRightStyle = "";
			newsContent[i].style.display = "none";
		}
		newsContent[nTab].style.display = "block";
		tabs[nTab].className = "clsTabSelected";
		oTab = tabs[nextTab];
		if (oTab) oTab.style.borderLeftStyle = "none";
		oTab = tabs[prevTab];
		if (oTab) oTab.style.borderRightStyle = "none";
		event.returnValue = false;
	}

	function fnSaveState() {
		var iTabSelected = 0;
		var iLength = tabs.length;
		for (var i = 0; i < iLength; i++) {
			if (tabs[i].className == "clsTabSelected") iTabSelected = i;
		}
		idTabs.setAttribute("tabstate", iTabSelected);
		return iTabSelected;
	}
	
	function fnGetState() {
		var iTabSelected = idTabs.getAttribute("tabstate");
		var iLength = tabs.length;
		for (var i = 0; i < iLength; i++) {
			if ( i != iTabSelected) {
				tabs[i].className = "clsTab";
				if (newsContent[i]) newsContent[i].style.display = "none";
			} else {
				tabs[i].className = "clsTabSelected";
				if (newsContent[i]) newsContent[i].style.display = "block";			
			}
		}	
	}
	
	function fnMaintainTab( nTab )
	{
		nTab = parseInt(nTab);
		var oTab;
		var prevTab = nTab-1;
		var nextTab = nTab+1;
		//event.cancelBubble = true;
		//el = event.srcElement;

		for (var i = 0; i < newsContent.length; i++)
		{
			oTab = tabs[i];
			oTab.className = "clsTab";
			oTab.style.borderLeftStyle = "";
			oTab.style.borderRightStyle = "";
			newsContent[i].style.display = "none";
		}
		newsContent[nTab].style.display = "block";
		tabs[nTab].className = "clsTabSelected";
		oTab = tabs[nextTab];
		if (oTab) oTab.style.borderLeftStyle = "none";
		oTab = tabs[prevTab];
		if (oTab) oTab.style.borderRightStyle = "none";
		//event.returnValue = false;
	}

//Functions added by SSeissiger and KWahoff -- added functionality - after a user updates info on tabs
//user will return to a refreshed version of the same tab (instead of going back to the
//profile tab as a default)

function childClosing() 
{
	alert('The Child Window is closing!');
}

function checkChildWindow() 
{
	var oTab;
	state = fnSaveState();
	
	if ( !window.winUpdate || !window.winUpdate.open || window.winUpdate.closed )
	{
		//window.winUpdate.opener = null;
		
		clearInterval(intervalID);
		window.document.location.href="http://www.sarkcincinnati.com/Intranet/directory/content/details.asp?empid=" + EmpID + "&tabState=" + state
				
		//fnMaintainTab(state);
	}
}
//--></script>
<LINK REL='stylesheet' TYPE='text/css' HREF='msstyle.css'>
<SCRIPT SRC="scriptlib.js" LANGUAGE="JavaScript"></SCRIPT>
<SCRIPT language=javascript>
<!--function OpenWin(page, EmpID, WinTitle){
	winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=450,width=600,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
	}function openWindow(URL, EmpID, WinTitle) 
{
		winUpdate=window.open(URL + "?empID=" + EmpID, WinTitle,"height=615,width=510,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
		intervalID=setInterval('checkChildWindow()', 100)
		state = fnSaveState();
}
function ShowNoIEError(){	window.alert("Sorry, this feature requires Microsoft Internet Explorer version 5.0 or higher.  There is a link on the main page that you can follow to download IE 5.0")	}
// --></SCRIPT>
<%
'-------------------------------------
'  Get Reference to FileSystemObject  
'-------------------------------------
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")

'---------------------------------------------
'	Execute Database Query for Employee Info  
'---------------------------------------------
ls_sql = "select e.*, t.title_desc  from employee e, employee_title t where e.employeetitle_id *= t.employee_title_id and e.employeeid = " & Request.QueryString("EmpID") 
set rs = DBQuery(ls_sql)
curTitle = curTitle & " (" & rs("firstname") & " " & rs("lastname") & ")"
if IsNull(rs("title_desc")) then
	ls_title = ""
else
	ls_title = rs("title_desc")
end if

'--------------------------
'	Display Employee Name  
'--------------------------
Response.Write("<center><font color=navy><b>" & rs("firstname") & " " & rs("lastname") & ",&nbsp;&nbsp;" & ls_title & "</b></font></center><br>")
Response.Write("<TABLE BORDER=0 WIDTH='100%'>")
ls_clientphone = rs("client_phone")ls_clientemail = rs("clientemail")
ldt_startdate = rs("startdate")
ls_username = lcase(rs("username"))
ls_empid = rs("employeeid")
ls_pager = rs("pagernumber")ls_pager_email = trim(rs("pageremail"))
ls_cellphone = rs("cellphone")
ls_voicemail = rs("voicemail")ls_fax = rs("fax")
ls_msngrid = rs("msngrid")
'-----------------------------------
'	Check if Employee Picture Exists
'-----------------------------------
imgEmployee = "images/" & rs("username") & ".jpg" 
FileName = Server.MapPath (imgEmployee)
if NOT FileObject.FileExists(FileName) then imgEmployee = "images/no_photo.jpg"
Response.Write("<TR><TD ALIGN=Center VALIGN=Top ROWSPAN=2><img src='" & imgEmployee & "' height=100 width=100></TD>")

'-------------------------
'	Display Personal Info
'-------------------------
Response.Write("<TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'><b>" & "Home Address:" & "</b><br>")
if NOT rs("unlistedpersonal") then
	Response.Write(rs("address") & "<br>" & rs("city") & "," & rs("state") & "  " & rs("zip") & "<br>" & rs("homephone") & "</TD>")
else
	Response.Write("Unlisted</TD>")
end if
Response.Write("<TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'><b>" & "Personal Information:" & "</b><br>")
if rs("spousename") <> "" then
	Response.Write("Spouse: ")
	Response.Write(rs("spousename")) & "<br>" 
end if
Response.Write("Birthday: ")
if not isnull(rs("birthday")) then
	if isdate(rs("birthday")) then
		Response.Write(MonthName(DatePart("m", rs("birthday"))) & " " & DatePart("d", rs("birthday"))) & "<br>"
	end if
end if
if rs("PersonalWebsite") <> "" then
	Response.Write("Website: ")
	'Response.Write("<a href='http://" & rs("PersonalWebsite") & "' & target="_blank"> " & rs("PersonalWebsite") & "</a>")
	Response.Write("<a target='blank' href='http://" & rs("PersonalWebsite") & "'> " & rs("PersonalWebsite") & "</a>")

end if
'---------------------------------------------
'	Execute Database Query for Client Info    
'---------------------------------------------
if not isnull(rs("client_id")) then
	ls_sql = "select * from client where client_id= " & rs("client_id")
	set rs = DBQuery(ls_sql)
	ls_clientname = rs("clientname")
	ln_sortoverride = rs("sortoverride")
	end if
Response.Write("</TD></TR><TR><TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'>")

'-------------------------
'	Display SARK Info     
'-------------------------
Response.Write("<b>" & "SARK Information: " & "</b><br>")
if ls_username <> "" then Response.Write("Email: <a href='javascript:SendEmail();'>" & ls_username & "@sark.com</a><br>")
Response.Write("Start Date: " & ldt_startdate & "</td>")

'---------------------------------------------------------------------------------
' If details page belongs to current user then display edit buttons.
'---------------------------------------------------------------------------------	if (UCase(trim(ls_username)) = UCase(session("Username"))) then
		userButtons = true
	end if	if ( UCase(Application("WebMaster")) = UCase(session("Username")) ) then
		webmasterButtons = true
	end if
if ln_sortoverride <> 2 then
	Response.Write("<TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'><b>" & "Client:" & "</b><br>")
	Response.Write(ls_clientname & "<br>" & ls_clientphone & "<br>" & ls_clientemail)	Response.Write("</TD></TR>")
end if

'if NOT IsNull(ls_voicemail) then
	' Display office phone numbers
	Response.Write("<TR><TD></TD><TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'>")
	Response.Write("<b>" & "Contact Information: " & "</b><br>")
	Response.Write("Pager: " & ls_pager & "<br>")	if ls_pager_email <> "" then
		Response.Write("Pager Email: <a href='javascript:SendPage();'>" & ls_pager_email & "</a><br>")	end if
	Response.Write("Mobile: " & ls_cellphone & "<br>")
	Response.Write("Voice Mail: " & ls_voicemail & "<br>")	Response.Write("Fax: " & ls_fax & "<br>")
	Response.Write("Instant Messenger ID: <b>" & ls_msngrid & "</b><br>")	Response.Write("</TD><TD><form>")		'---------------------------------------
		'Add Button to update employee info
		'---------------------------------------
		ls_empid = Request.QueryString("EmpID")		if ( userButtons or webmasterButtons ) then
			Response.Write("<INPUT TYPE=button class=button NAME='UpdateEmp' VALUE = 'Edit Employee' onClick=" & chr(34) & "OpenWin('editemp.asp', " & ls_empid  & ", 'EditSARKEmp')" & chr(34) & "><br><br>")
		end if		
		if oBD.getsSkills then			'----------------------------------------------------------------			' View Skills Inventory                                    
			'----------------------------------------------------------------
			Response.Write("<INPUT TYPE=button class=button NAME='SkillsInv' " & _		          "VALUE = '  Skills Inventory  ' onClick=" & chr(34) & _		           "OpenWin('SkillsInv.asp', " & ls_empid  & ", 'EditSARKInd')" & chr(34) & ">")
		else
			'----------------------------------------------------------------
			' Give alert dialog that demands IE5+
			'----------------------------------------------------------------
			Response.Write("<INPUT TYPE=button class=button NAME='SkillsInv' " & _		          "VALUE = '  Skills Inventory  ' onClick=" & chr(34) & _		          "ShowNoIEError()" & chr(34) & ">")
			
		end if	
	Response.Write("</form></TD></TR></TABLE>")
'end if
'**********
Response.Write(vbNewLine & vbNewLine & "<!--" & vbNewLine & _               "You're looking at   " & trim(ls_username) & vbNewLine & _
               "You're logged in as " & Session("Username") & vbNewline & _
               "-->" & vbNewLine & vbNewLine)'**********


%>

<DIV CLASS="clsDocBody">
<!--Begin Content-->
<!--Code added here by Ryan Dlugosz on 6/20/00 to implement tabbed information interface
	Code was adapted from an MSDN implementation -->
	<TABLE ID="idTabs" CELLPADDING="2" CELLSPACING="0" BORDER="0" BGCOLOR="#ffffcc" STYLE="color: #FFFFFF; display:none; behavior: url(#default#savehistory);" ONSAVE="fnSaveState()" ONLOAD="fnGetState()">
		<TR HEIGHT="25" VALIGN="middle">
			<TD ID="tabs" CLASS="clsTabSelected" ONCLICK="fnTabClick('0');">&nbsp;<A HREF="javscript:void();" ONCLICK="return fnTabClick('0');">Profile</A>&nbsp;</TD>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('1');">&nbsp;<A HREF="javscript:void();" ONCLICK="return fnTabClick('1');">Technologies</A>&nbsp;</TD>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('2');">&nbsp;<A HREF="javscript:void();" ONCLICK="return fnTabClick('2');">Certifications</A>&nbsp;</TD>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('3');">&nbsp;<A HREF="javscript:void();" ONCLICK="return fnTabClick('3');">Industry&nbsp;Experience</A>&nbsp;</TD>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('4');">&nbsp;<A HREF="javscript:void();" ONCLICK="return fnTabClick('4');">Projects</A>&nbsp;</TD>
		<!-- code added by sseissiger 7-25-00
	adds the security feature for the review and the vacation tabs
-->


			<% if (Session("GotAcctMgr")) then %>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('5');">&nbsp;<A HREF="javascript:void();" ONCLICK="return fnTabClick('5');">Reviews</A>&nbsp;
			</TD>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('6');">&nbsp;<A HREF="javascript:void();" ONCLICK="return fnTabClick('6');">Vacation</A>&nbsp;</TD>
			<TD ID="tabs" CLASS="clsTab" ONCLICK="fnTabClick('7');">&nbsp;<A HREF="javascript:void();" ONCLICK="return fnTabClick('7');">Exam</A>&nbsp;</TD>
			<% end if %>
		</TR>		
	</TABLE>
	
<SCRIPT> tabNum = 4;</SCRIPT>
 
 <% if (Session("GotAcctMgr")) then %>
	
	<SCRIPT>tabNum = 7; </SCRIPT>			
			
 <% end if %>
 
 <SCRIPT>tabs[tabNum].width="100%"; idTabs.style.display="block";</SCRIPT>


<!-- <SCRIPT>tabs[4].width="100%"; idTabs.style.display="block";</SCRIPT> -->


 <A NAME="ProfileTab"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadProfile">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><img HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR>
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><img HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Profile</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" HEIGHT="6" WIDTH="2"></TD>
	</TR> 
</TABLE>

<SCRIPT>tblHeadProfile.style.display="none";</SCRIPT>

<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td width=50% valign=top>
<%	
'--------------------------------------------
'	Execute Database Query for Profile Info  
'--------------------------------------------ls_sql = "select * from profile where employee_id = " & ls_empid
set rs = DBQuery(ls_sql)
Response.Write("<br>")if  rs.eof then
	'--------------------------------------------------
	'	Display Profile Info - Only display is not null
	'--------------------------------------------------
	 	Response.Write("<font size = 1 color = gray>None Specified</font><br><br>")
 	
else	
	Response.write("<table><tr><td><font size=1 face='ms sans serif, arial, geneva'>")
	if not (isnull(rs("college")) or trim(rs("college")) = "") then
		Response.Write("<b>" & "College:" & "<br></b>" & rs("college") & "<br><br>")
		end if
	if not (isnull(rs("hobbies_sports_interests")) or trim(rs("hobbies_sports_interests")) = "" ) then
		Response.Write("<b>" & "Hobbies, Sports & Interests:" & "<br></b>" & rs("hobbies_sports_interests") & "<br><br>")
		end if
	if not (isnull(rs("dream_vacation")) or trim(rs("dream_vacation")) = "" ) then
		Response.Write("<b>" & "Dream Vacation:" & "<br></b>" & rs("dream_vacation") & "<br><br>")
		end if
	if not (isnull(rs("favorite_show_ad")) or trim(rs("favorite_show_ad")) = "" ) then
		Response.Write("<b>" & "Favorite Show/Ad:" & "<br></b>" & rs("favorite_show_ad") & "<br><br>")
		end if
	if not (isnull(rs("person_inspired")) or trim(rs("person_inspired")) = "" ) then
		Response.Write("<b>" & "Person Who Inspired me the Most:" & "<br></b>" & rs("person_inspired") & "<br><br>")
		end if
	if not (isnull(rs("favorite_motto")) or trim(rs("favorite_motto")) = "" ) then
		Response.Write("<b>" & "Favorite Motto:" & "<br></b>" & rs("favorite_motto"))
		end if
	Response.write("</font></td>")	Response.write("<td width=50% valign=top><font size=1 face='ms sans serif, arial, geneva'>")
	if not (isnull(rs("funniest_embarrassing")) or trim(rs("funniest_embarrassing")) = "" ) then
		Response.Write("<b>" & "Funniest or Most Embarrassing Moment:" & "<br></b>" & rs("funniest_embarrassing") & "<br><br>")
		end if 
	if not (isnull(rs("lotto")) or trim(rs("lotto")) = "" ) then
		Response.Write("<b>" & "If I Won the Lotto...:" & "<br></b>" & rs("lotto") & "<br><br>")
		end if
	if not (isnull(rs("pet_peeve")) or trim(rs("pet_peeve")) = "" ) then
		Response.Write("<b>" & "Pet Peeve:" & "<br></b>" &  rs("pet_peeve") & "<br><br>")
		end if
	if not (isnull(rs("prized_possession")) or trim(rs("prized_possession")) = "" ) then
		Response.Write("<b>" & "Prized Possession:" & "<br></b>" & rs("prized_possession") & "<br><br>")
		end if
	if not (isnull(rs("biggest_challenge")) or trim(rs("biggest_challenge")) = "" ) then
		Response.Write("<b>" & "Biggest Challenge:" & "<br></b>" & rs("biggest_challenge") & "<br><br>")
		end if
	Response.write("</font></td></tr></table>")	end if	rs.close
		if (userButtons) then
		'----------------------------------------------------------------		' Edit SARK Profile                                    
		'----------------------------------------------------------------
		
		Response.Write("<form><INPUT TYPE=button class=button NAME='UpdateProfile' " & _					   "VALUE = '   Edit Profile  ' onClick=" & chr(34) & _					   "openWindow('editprofile.asp', " & ls_empid  & ", 'EditSARKProfile')" & chr(34) & "></form>")
	end if
	%>	
	</td>
	</TR>
</TABLE>

 <A NAME="Technologies"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadTech">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR>
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Technologies</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadTech.style.display="none";</SCRIPT>

<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td>
	<%

sql = "SELECT Tech_Area.Tech_Area, Tech.Tech_Name, Employee_Tech_Xref.Employee_ID, Tech.Tech_ID" _
		& " FROM Tech INNER JOIN Employee_Tech_Xref ON Tech.Tech_ID = Employee_Tech_Xref.Tech_ID " _
		& " INNER JOIN Tech_Area ON Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID " _
		& " WHERE (Employee_Tech_Xref.Employee_ID = " & ls_empid & ") AND" _
		& " (Tech.Approved = 1)" _
		& " ORDER BY Tech_Area.Tech_Area, Tech.Tech_Name"
set rs = DBQuery(sql)
Response.Write("<br>")if rs.eof then	
	Response.Write("<font size = 1 color = gray>None Specified</font><br>")else        Dim HoldTech     
    
   	HoldTech = rs("Tech_Area")       Response.Write("<b>" & HoldTech & "</b><br>")
        while not rs.eof
		if rs("Tech_Area") <> HoldTech then
			HoldTech = rs("Tech_Area")			Response.Write("<b>" & HoldTech & "</b><br>")
		end if													
		response.write("&nbsp;&nbsp;<A HREF='../../technology/content/tech.asp?techid=" & rs("Tech_ID") & "'>" & rs("Tech_Name") & "</a><br>")		rs.movenext

	wendend if	

Response.Write("</font>")

	if (userButtons) then
		'----------------------------------------------------------------		' Edit Technical Expertise                                    
		'----------------------------------------------------------------
				Response.Write("<br><form><INPUT TYPE=button class=button NAME='TechExpert' " & _		               "VALUE = ' Edit Expertise ' onClick=" & chr(34) & _		               "openWindow('techexperts.asp', " & ls_empid  & ", 'EditSARKTech')" & chr(34) & "></form>")	end if

%><br>
	</td>
	</TR>
	</TABLE>
<SCRIPT>newsContent[1].style.display="none";var sText=1</SCRIPT>
 <A NAME="Certs"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadCerts">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR>
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Certifications</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadCerts.style.display="none";</SCRIPT>

<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" width=100% ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td>
	<%

'SQL statement for gathering certification information.
sql = "SELECT Certifications.Cert_Name" _
	& " FROM Employee_Cert_Xref INNER JOIN" _
	& " Certifications ON" _
    & " Employee_Cert_Xref.Cert_ID = Certifications.Cert_ID" _
	& " WHERE (Employee_Cert_Xref.Employee_ID = " & ls_empid & ")" _
	& " ORDER BY Certifications.Cert_Name"			
set rs = DBQuery(sql)
Response.Write("<br>")
if rs.eof then
	Response.Write("<font size = 1 color = gray>None Specified</font><br>")else    while not rs.eof
		Response.Write("<b>" & rs("Cert_Name") & "</b><br>")
		rs.movenext

	wendend if	

rs.close
Response.Write("</font><br>")

if (webmasterButtons) then
	'----------------------------------------------------------------	' Edit Certifications                                    
	'----------------------------------------------------------------
		Response.Write("<form><INPUT TYPE=button class=button NAME='Certifictaions' " & _					   "VALUE = 'Edit Certifications' onClick=" & chr(34) & _		               "OpenWin('certifications.asp', " & ls_empid  & ", 'EditSARKCert')" & chr(34) & "></form>")end if             

%>
	</td>
	</TR>
	</TABLE>
<SCRIPT>newsContent[2].style.display="none";var sText=1</SCRIPT>
 <A NAME="Industry"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadIndust">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR>
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Industry&nbsp;Experience</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadIndust.style.display="none";</SCRIPT>

<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td>
	<%

'SQL statement for gathering industry experience information.
sql = "SELECT Industry.Industry_Name" _
		& " FROM Employee_Industry_Xref INNER JOIN" _
		& " Industry ON" _
		& " Employee_Industry_Xref.Industry_ID = Industry.Industry_ID" _
		& " WHERE (Employee_Industry_Xref.Employee_ID = " & ls_empid & ")" _
		& " ORDER BY Industry.Industry_Name"			
set rs = DBQuery(sql)
Response.Write("<br>")
if rs.eof then
	Response.Write("<font size = 1 color = gray>None Specified</font><br>")else    while not rs.eof
		Response.Write("<b>" & rs("Industry_Name") & "</b><br>")
		rs.movenext

	wendend if	

rs.close
Response.Write("</font>")
if (userButtons) OR (webmasterButtons) then		'----------------------------------------------------------------		' Edit Industry Experience                                    
		'----------------------------------------------------------------

		Response.Write("<form><INPUT TYPE=button class=button NAME='IndExperience' " & _		               "VALUE = '  Edit Industry  ' onClick=" & chr(34) & _		               "openWindow('indexperience.asp', " & ls_empid  & ", 'EditSARKInd')" & chr(34) & "></form>")end if

%><br>
	</td>
	</TR>
	</TABLE>
<SCRIPT>newsContent[3].style.display="none";var sText=1</SCRIPT>
 <A NAME="Projects"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadProjects">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR>
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Projects</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadProjects.style.display="none";</SCRIPT>

<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td>
	<%

'SQL statement for gathering the projects worked on.

ProjectSql="SELECT Employee_Project_Xref.Employee_ID, " _
    & "Employee_Project_Xref.Project_ID, Client.Client_ID, " _
    & "Client.ClientName, Project.Project_Name " _
	& "FROM Client INNER JOIN " _
    & "Project ON Client.Client_ID = Project.Client_ID INNER JOIN " _
    & "Employee_Project_Xref ON " _
    & "Project.Project_ID = Employee_Project_Xref.Project_ID " _
	& "WHERE (Employee_Project_Xref.Employee_ID = " & ls_empid & ")"

set ProjectRs=DBQuery(ProjectSql)
	Response.Write("<br>")

if ProjectRs.eof then
	Response.Write("<font size = 1 color = gray>None Specified</font><br><br>")
else
	while not ProjectRs.eof
		Response.Write("<a href='../../projects/content/details.asp?clientid=" & ProjectRs("Client_ID") & "'>" & ProjectRs("ClientName") & "</a> - <a href='../../projects/content/details.asp?clientid= " & ProjectRs("Client_ID") & "&projectid=" & ProjectRs("Project_ID") & "'>" & ProjectRs("Project_Name") & "</a><br>")
		ProjectRs.movenext
	wend
end if

ProjectRs.close
Response.Write("</font><br>")%>
</td>
	</TR>
</TABLE>
<!-- Added by Jason Moore and Mike Thierauf
     Provides functionality for the review tab -->

<SCRIPT>newsContent[4].style.display="none";var sText=1</SCRIPT>
<A NAME="Reviews"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadReviews">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR>
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Reviews</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadReviews.style.display="none";</SCRIPT>

<!-- Added by sseissiger to enable the security feature
-->
<% if (Session("GotAcctMgr")) then %>


<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td width=50% valign=top>
<%	
'--------------------------------------------
'	Execute Database Query for Review Info  
'--------------------------------------------
ls_sql = "select last_promotion, last_review, next_review, sponsor_ID, account_mgr_ID, " & _         "comments from review where employeeID = " & ls_empid
set rs = DBQuery(ls_sql)if not rs.EOF then	ActMgr = rs("account_mgr_ID")	Sponsor = rs("sponsor_ID")
	promotion = rs("last_promotion")
	if not (isnull(ActMgr))then		ls_sql = "select firstname, lastname from employee where employeeID = " & ActMgr		set rs_act = DBQuery(ls_sql)	end if
	if not (isnull(Sponsor))then
		ls_sql = "select firstname, lastname from employee where employeeID = " & Sponsor
		set rs_spons = DBQuery(ls_sql)
	end if
	Response.Write("<br>")

'----------------------------------------------------
'	Display Review Info - Only display if is not null''   code updated by SSeissiger 8-9-00 to allow for '   the account manager and sponsor to be set to none'   and have a null inserted into the table
'----------------------------------------------------	Response.write("<table><tr><td colspan=2 width=60% valign=top><font size=1 face='ms sans serif, arial, geneva'>")	if not (isnull(ActMgr)) then
		Response.Write("<b>" & "Account Manager:"  & "</b><br>" & _		      "<a href='details.asp?EmpId=" & ActMgr & "'>" & rs_act("firstname") & " " & _		      rs_act("lastname") & "</a><br><br>")		else			Response.Write("<b>" & "Account Manager:</b> <br>None Assigned" & "<br><br>")
	end if			if not (isnull(Sponsor)) then 		Response.Write("<b>" & "Sponsor:" & "<br></b><a href='details.asp?EmpId=" & Sponsor & "'>" & rs_spons("firstname") & " " & rs_spons("lastname") & "</a><br><br>")		else
			Response.Write("<b>" & "Sponsor:</b> <br>None Assigned" & "<br><br>")	end if 
	
	if not (isnull(rs("comments")) or trim(rs("comments")) = "") then
		Response.Write("<b>" & "Comments:" & "<br></b>" & rs("comments") & "<br><br>")
	end if		Response.Write("</font></td>")	Response.Write("<td colspan=2 valign=top><font size=1 face='ms sans serif, arial, geneva'>")	
	'If there is no last promotion date then the date is set to 9/9/9999 in editreview.asp	'the last promotion date is left blank on the details page	if (rs("last_promotion") = "9/9/9999")then
		Response.Write("<b>" & "Last promotion:" & "<br></b>" & "<br><br>") 
		elseif not (isnull(rs("last_promotion")) or trim(rs("last_promotion")) = "") then
			Response.Write("<b>" & "Last promotion:" & "<br></b>" & rs("last_promotion") & "<br><br>")	end if	
		if not (isnull(rs("last_review")) or trim(rs("last_review")) = "") then
		Response.Write("<b>" & "Last review:" & "<br></b>" & rs("last_review") & "<br><br>")
		end if	if not (isnull(rs("next_review")) or trim(rs("next_review")) = "") then
		Response.Write("<b>" & "Next review:" & "<br></b>" & rs("next_review") & "<br><br>")
		end if		
	Response.write("</font></td></tr></table>")	
	rs.close
	if not (isnull(ActMgr)) then
		rs_act.close
	end if
	if not (isnull(Sponsor)) then
		rs_spons.close
	end if

else
	Response.Write("<font size = 1 color = gray>None Specified</font><br><br>")end if
Response.Write("<center><form id=form3 name=form3><INPUT TYPE=button class=button NAME='UpdateReview' " & _    "VALUE = '  Edit Review Info  ' onClick=" & chr(34) & _    "openWindow('editreview.asp', " & ls_empid  & ", 'EditSARKRev')" & chr(34) & "></form></center>")
%>	
</td>
</TR>
</TABLE>

<!--
			Added by SSeissiger and KWahoff to display the vacation tab
-->

<SCRIPT>	newsContent[5].style.display="none";
			var sText=1
</SCRIPT>


<% end if %> <!-- end for the security feature -->


<A NAME="Vacation"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadVacation">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR> 
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Vacation</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadVacation.style.display="none";</SCRIPT>


<!-- Added by sseissiger to enable the security feature
-->
<% if (Session("GotAcctMgr")) then %>


<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td width=100% valign=top>
	
<%
'--------------------------------
'Execute query for vacation
'--------------------------------

today = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
					
ls_sql = "SELECT employee_id, vacation_id, vacation_start, vacation_end, comments " &_
			"FROM vacation  " &_
			"WHERE employee_ID = " & ls_empID &_
			"AND datediff(dd, GETDATE(), vacation_end) >= 0 " &_
			"ORDER BY vacation_start"

set rs = DBQuery(ls_sql)

Response.Write("<br>")

if  rs.eof then
	'------------------------------------------------------------
	'	Display Vacation info..only display if it is not null
	'------------------------------------------------------------ 	Response.Write("<font size = 1 color = gray>None Specified</font><br><br>")
end if

if not rs.eof then
Response.write("<table><tr><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<font size=1 face='ms sans serif, arial, geneva'>")
Response.write("<td colspan=1 width=90 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>Vacation ID:</b>")
Response.write("<td colspan=1 width=70 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>Start: </b>")
Response.write("<td colspan=1 width=70 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>End: </b>")
Response.write("<td colspan=1 width=200 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>Comments: </b>")
Response.write("</tr></table>")
end if

while not rs.eof 
	ls_empid = rs("employee_id")
	ls_vacid = rs("vacation_id")
	ls_vacstart = rs("vacation_start")
	ls_vacend = rs("vacation_end")
	ls_vaccomments = rs("comments")
		Response.write("<table><td><font size=1 face='ms sans serif, arial, geneva'>")	if not (isnull(ls_vacid) or trim(ls_vacid) = "") then		Response.Write("<td colspan=1 width=90 valign=middle align=left><b><a href=""javascript:addDocument(" & ls_vacid & ");"" onMouseOver=""top.status='Update Vacation'; return true;"" onMouseOut=""top.status=''; return true;"">" & ls_vacid & "</a>&nbsp;&nbsp;&nbsp;")
		end if
	if not (isnull(ls_vacstart) or trim(ls_vacstart) = "" ) then
		Response.Write("<td colspan=1 width=70 valign=middle align=left><b>" & "</b>" & ls_vacstart)
		end if
	if not (isnull(ls_vacend) or trim(ls_vacend) = "") then
		Response.Write("<td colspan=1 width=70 valign=middle align=left><b>" & "</b>" & ls_vacend)
		end if	if not (isnull(ls_vaccomments) or trim(ls_vaccomments) = "") then
		Response.Write("<td colspan=1 width=200 valign=middle align=left><b>" & "</b>" & ls_vaccomments & "<br><br>")
		end if	
	Response.write("</td></table>")
	rs.movenext
	wend

rs.closeset rs = nothing

Response.Write("<center><form id=form10 name=form10> " &_
"<INPUT TYPE=button class=button NAME='AddVacation' " &_	               "VALUE = '  Add Vacation  ' onClick=" & chr(34) & _	               "openWindow('addNewVacation.asp', " & ls_empid & ", 'EditSarkVac')" & chr(34) & "> " &_
"<INPUT TYPE=button class=button NAME='VacHistory' " &_	               "VALUE = '  Vacation History  ' onClick=" & chr(34) & _	               "openWindow('vacHistory.asp', " & ls_empid & ", 'EditSarkVac')" & chr(34) & "> " &_
"<INPUT TYPE=hidden NAME='tabID' VALUE = '6'> " &_
				   "</form></center>")
%>	

	</TD>
	</TR>
</TABLE>
	
<SCRIPT>newsContent[6].style.display="none";var sText=1</SCRIPT>	


<% end if %>  <!-- end for the security feature -->


<A NAME="Exams"></a>
	<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="tblHeadExams">
	<TR>
	<TD ALIGN="left" WIDTH="100%" HEIGHT="25" COLSPAN="2"><IMG HEIGHT="25" src="ts.gif" WIDTH="15"></TD>
	</TR>
	<TR> 
	<TD VALIGN="middle" HEIGHT="18" BGCOLOR="#ffffcc" NOWRAP><IMG HEIGHT="18" src="ts.gif" WIDTH="1">&nbsp;&nbsp;<font COLOR="#000000" SIZE="2"><B>Exams</B></FONT>&nbsp;&nbsp;</TD>
	<TD ALIGN="right" WIDTH="100%"></TD>
	</TR>
	<TR>
	<TD ALIGN="left" BGCOLOR="#ffffcc" VALIGN="top" WIDTH="100%" HEIGHT="6" COLSPAN="2"><IMG SRC="ts.gif" ALT HEIGHT="6" WIDTH="2"></TD>
	</TR>
</TABLE>

<SCRIPT>tblHeadExams.style.display="none";</SCRIPT>


<!-- Added by sseissiger to enable the security feature
-->
<% if (Session("GotAcctMgr")) then %>


<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%" ID="newsContent" bgcolor=#ffffcc class=clsTabContents>
	<TR>
	<td width=100% valign=top>
	



	
<%
	'--------------------------------
	'Execute query for Exams
	'--------------------------------
						
ls_sql = "SELECT Exam_List.Exam_ID, Exam_List.Exam_Name," _
	& " Exams_passed.Exam_ID, Exams_passed.Exams_passed_ID, Exams_passed.Exam_date," _
	& " Exams_passed.Comments" _
	& " FROM Exam_List   INNER JOIN" _
	& " Exams_passed ON" _
    & " Exams_passed.Exam_ID = Exam_List.Exam_ID" _
	& " WHERE (Exams_passed.EmployeeID = " & ls_empid & ")" _
	& " ORDER BY Exams_passed.Exam_date"



set rs = DBQuery(ls_sql)

Response.Write("<br>")

if  rs.eof then
	'------------------------------------------------------------
	'	Display Exam info..only display if it is not null
	'------------------------------------------------------------ 	Response.Write("<font size = 1 color = gray>None Specified</font><br><br>")

end if

if not rs.eof then
Response.write("<table><tr><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<font size=1 face='ms sans serif, arial, geneva'>")
Response.write("<td colspan=1 width=70 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>ID:</b>")
Response.write("<td colspan=1 width=240 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>Exam Name: </b>")
Response.write("<td colspan=1 width=80 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
Response.Write("<b>Date Passed: </b>")

Response.write("</tr></table>")
end if

while not rs.eof 
	
	ls_examid = rs("Exams_passed_ID")
	ls_examname = rs("Exam_Name")
	ls_examdate = rs("Exam_Date")
	ls_comments = rs("Comments")	

	
		Response.write("<table><td><font size=1 border=3 face='ms sans serif, arial, geneva'>")	if not (isnull(ls_examid) or trim(ls_examid) = "") then		Response.Write("<td colspan=1 width=70 valign=middle align=left><b><a href=""javascript:addDocument2(" & ls_examid & ");"" onMouseOver=""top.status='Update Exam'; return true;"" onMouseOut=""top.status=''; return true;"">" & ls_examid & "</a></b>&nbsp;&nbsp;&nbsp;")
		end if
	if not (isnull(ls_examname) or trim(ls_examname) = "" ) then
		Response.Write("<td colspan=1 width=250 valign=middle align=left><b>" & "</b>"& ls_examname)
		end if
	if not (isnull(ls_examdate) or trim(ls_examdate) = "") then
		Response.Write("<td colspan=1 width=50 valign=middle align=left><b>" & "</b>"& ls_examdate)
		end if
Response.write("</td></table>")

	rs.movenext
	wend
Response.write("<br>")
rs.closeset rs = nothing


Response.Write("<center><table><form id=form10 name=form10> " &_
"<INPUT TYPE=button class=button NAME='Add Exam' " &_	               "VALUE = '  Add Exam  ' onClick=" & chr(34) & _	               "openWindow('addexam.asp', " & ls_empid & ", 'AddExam')" & chr(34) & "> " &_
"<INPUT TYPE=hidden NAME='tabID' VALUE = '7'> " &_
				   "</form></table></center>")

	
%>
	</TD>
	</TR>
</TABLE>


<% end if %> 












<%
if request("tabState") <> "" then	intTabState = request("tabState")
else
	intTabState = 0end if%>

<!-- </DIV> -->

<script language=javascript>
<!--
fnMaintainTab(<%=intTabState%>);
function SendEmail(){winTech= window.open("../../email.asp?editto=yes&recipient=<%=ls_username%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendEmail","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}
function SendPage(){winTech= window.open("../../pager.asp?recipient=<% =ls_pager_email %>&from=<%=session("Username")%>&name=<%=Server.URLEncode(session("name"))%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendPage","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}

function addDocument(arg) 
{
	winTech=window.open("updateVacation.asp?vacid=" + arg, "AddDocument","height=615,width=500,toolbar = 0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbar=1")

}

function addDocument2(arg) 
{
	winTech=window.open("updateexam.asp?epid=" + arg, "AddDocument","height=615,width=500,toolbar = 0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbar=1")

}

// -->
</script>

<!-- #include file="../../footer.asp" -->
