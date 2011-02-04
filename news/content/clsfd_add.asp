<%Response.Buffer = True%>
<!--
Developer:    AHARSHBARGER
Date:         01/14/1999
Description:  Add a classified entry
-->
<%
dim user_nm
user_nm = Session("UserName")

SQL = ""
SQL = SQL + "SELECT FirstName, LastName, EmployeeID "
SQL = SQL + "FROM Employee "
SQL = SQL + "WHERE Username = '" & user_nm & "'"
set rs = DBQuery(SQL)
%>

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Add A New Classified Entry</TITLE>
<!-- #include file="../../style.htm" -->
</HEAD>

<% If Request("ACTION") = "" Then %>

<BODY BGCOLOR=SILVER onLoad=""><center>

<FORM NAME="clsInfo" ACTION="clsfd_add.asp" METHOD="POST">

<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH='90%'>
<TR>
	<TD COLSPAN=2 ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva"><b>Enter the classified information:</B></FONT></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="10"></TD></TR>
<TR>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva">Posted by:</FONT></TD>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva"><%=rs("FirstName") & " " & rs("LastName")%><INPUT TYPE=HIDDEN NAME="USER" VALUE="<%=user_nm%>"></FONT></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="2"></TD></TR>
<TR>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva">Posted date:</FONT></TD>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva"><%=Now()%><INPUT TYPE=HIDDEN NAME="ENTER_DATE" VALUE="<%=Now()%>"></FONT></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="2"></TD></TR>
<TR>
	<TD ALIGN=LEFT VALIGN=TOP><font size=1 face="ms sans serif, arial, geneva">Brief Description:</font></td>
	<TD ALIGN=LEFT VALIGN=TOP><TEXTAREA NAME="DESCRIPTION" ROWS="10" COLS="50" WRAP="VIRTUAL"></TEXTAREA></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="2"></TD></TR>
<TR>
	<TD VALIGN=TOP ALIGN=CENTER COLSPAN=2>
		<input type=submit class=button value="Save">
		<input type=button class=button value="Cancel" onClick="window.close()">
	</TD>
</TR>
</TABLE>

<INPUT TYPE="hidden" NAME="STATUS" VALUE="A">
<INPUT TYPE="HIDDEN" NAME="ACTION" VALUE="ADD">
</form>
</center>

<%Else%>
<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}
// -->
</script>


<html>
<head>
<title>Add Classified</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<center><font size=1 face="ms sans serif, arial, geneva">Saving information...</font></center>

<%	
	if Request("ACTION") = "ADD" then

		cls_user		=	Clean(Request("USER"))
		cls_status		=	Clean(Request("STATUS"))
		cls_enter_dt	=	Clean(Request("ENTER_DATE"))
		cls_dsc			=	Clean(Request("DESCRIPTION"))

		SQL = ""
		SQL = SQL + "SELECT employeeid "
		SQL = SQL + "FROM employee "
		SQL = SQL + "WHERE username = '" & cls_user & "'"
		set rs = DBQuery(SQL)

		SQL = ""
		SQL = SQL + "INSERT classifieds (employee_id, timestamp, status, description,lastmodified) "
		SQL = SQL + "VALUES (" & rs("employeeid") & ", '" & cls_enter_dt & "', '" & cls_status & "','" & cls_dsc & "',null)"
		DBQuery(SQL)

		SQL = ""
		SQL = SQL + "SELECT @@identity from classifieds"
		set rs = DBQuery(SQL)

		SQL = ""
		SQL = SQL + "UPDATE classifieds "
		SQL = SQL + "set Description = '" & Replace(cls_dsc,"'","''") & "' where classifieds_id = " & rs(0)
		DBQuery(SQL)
	end if

End If%>

</body>
</html>

<%rs.close()%>