<%Response.Buffer = True%>
<!--
Developer:    AHARSHBARGER
Date:         01/14/1999
Description:  Edit a classified entry
-->
<%
dim user_nm

user_nm = Session("UserName")
cls_id = Clean(Request("CLSID"))

SQL = ""
SQL = SQL + "SELECT cl.Classifieds_ID, em.FirstName, em.LastName, em.EmployeeID, "
SQL = SQL + "cl.TimeStamp, cl.Description, em.UserName "
SQL = SQL + "FROM Classifieds cl, Employee em "
SQL = SQL + "WHERE cl.Classifieds_ID = " & cls_id
SQL = SQL + " AND cl.Employee_ID = em.EmployeeID "
set rs = DBQuery (SQL)

desc_txt = rs("Description")
if isnull(desc_txt) or desc_txt="" then desc_txt = "[Description not available]"
%>

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Edit Classified Entry</TITLE>
<!-- #include file="../../style.htm" -->
</HEAD>

<% If Request("ACTION") = "" Then %>

<BODY BGCOLOR=SILVER>
<center>

<FORM NAME="clsInfo" ACTION="clsfd_edit.asp" METHOD="POST">
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH='90%'>
<TR>
	<TD COLSPAN=2 ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva"><b>Edit classified:</B></FONT></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="10"></TD></TR>
<TR>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva">Posted by:</FONT></TD>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva"><%=rs("FirstName") & " " & rs("LastName")%><INPUT TYPE=HIDDEN NAME="USER_ID" VALUE="<%=rs("EmployeeID")%>"></FONT></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="2"></TD></TR>
<TR>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva">Posted date:</FONT></TD>
	<TD ALIGN=LEFT VALIGN=TOP><FONT SIZE=1 FACE="ms sans serif, arial, geneva"><%=rs("TimeStamp")%></FONT></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="2"></TD></TR>
<TR>
	<TD ALIGN=LEFT VALIGN=TOP><font size=1 face="ms sans serif, arial, geneva">Brief Description:</font></td>
	<TD ALIGN=LEFT VALIGN=TOP><TEXTAREA NAME="DESCRIPTION" ROWS="10" COLS="50" WRAP="VIRTUAL"><%=desc_txt%></TEXTAREA></TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="2"></TD></TR>
<TR>
	<TD VALIGN=TOP ALIGN=CENTER COLSPAN=2>
		<input type=submit class=button value="Save">
		<input type=button class=button value="Cancel" onClick="window.close()">
	</TD>
</TR>
</TABLE>

<INPUT TYPE="HIDDEN" NAME="CLSID" VALUE="<%=cls_id%>">
<INPUT TYPE="HIDDEN" NAME="ACTION" VALUE="EDIT">
</FORM>
</center>

<%ElseIf Request("ACTION") = "EDIT" Then %>
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
<title>Edit Classified</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<center><font size=1 face="ms sans serif, arial, geneva">Saving information...</font></center>

<%	
	cls_id			=	Clean(Request("CLSID"))
	cls_user_id		=	Clean(Request("USER_ID"))
	cls_dsc			=	Clean(Request("DESCRIPTION"))

	SQL = ""
	SQL = SQL + "UPDATE CLASSIFIEDS "
	SQL = SQL + "SET LASTMODIFIED = '" & Now() & "' , "
	SQL = SQL + "DESCRIPTION = '" & Replace(cls_dsc,"'","''") & "' "
	SQL = SQL + "WHERE CLASSIFIEDS_ID = " & cls_id 
	SQL = SQL + " AND EMPLOYEE_ID = " & cls_user_id
	DBQuery(SQL)
End If%>

</body>
</html>

<%
rs.close()
set rs = nothing
%>