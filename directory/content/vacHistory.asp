<!--
Developer:    SSeissiger
Date:         08/21/2000
Description:  Allows A/M to view vacation history
-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Vacation History</TITLE>
<!-- #include file="../../style.htm" -->
<script language="JavaScript" type="text/javascript">

function addDocument(arg) 
{
	winVacHist=window.open("updateVacation.asp?vacid=" + arg, "AddDocument","height=615,width=500,toolbar = 0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbar=1")
    intervalID=setInterval('checkChildWindow()', 100)
}
			
function checkChildWindow() 
{
	var oTab;

	if ( !window.winVacHist || !window.winVacHist.open || window.winVacHist.closed )
	{
		clearInterval(intervalID);
		window.document.location.href="<%=Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING")%>"
	}
}
</SCRIPT>
</HEAD>

<BODY color=silver  bgcolor=silver onLoad="window.resizeTo(615, 410)"> 


<%
ls_empID = Request.QueryString("empid")

ls_sql = "SELECT employee_id, vacation_id, vacation_start, vacation_end, comments " &_
			"FROM vacation  " &_
			"WHERE employee_ID = " & ls_empID &_
			" ORDER BY vacation_start"

set rs = DBQuery(ls_sql)

if  rs.eof then
	'------------------------------------------------------------
	'	Display Vacation info..only display if it is not null
	'-----------------------------------------------------------
%> 	<font size = 1 color = red><b>Query Returned No Results</b></font><br><br>
<%end if %>

<%if not rs.eof then%>
<table border=0><tr><font size=1 face='ms sans serif, arial, geneva'>
<td colspan=1 width=100 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>
<b>Vacation ID:</b>
<td colspan=1 width=75 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>
<b>Start: </b>
<td colspan=1 width=75 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>
<b>End: </b>
<td colspan=1 width=300 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>
<b>Comments: </b>
</tr></font>
</table>
<%end if %>

<%while not rs.eof 
	ls_empid = rs("employee_id")
	ls_vacid = rs("vacation_id")
	ls_vacstart = rs("vacation_start")
	ls_vacend = rs("vacation_end")
	ls_vaccomments = rs("comments")
%>		<table border=0><tr><font size=1 face='ms sans serif, arial, geneva'>	<%if not (isnull(ls_vacid) or trim(ls_vacid) = "") then%>
	<% response.write("<td colspan=1 width=100 valign=top align=left><a href=""javascript:addDocument(" & ls_vacid & ");"" onMouseOver=""top.status='Update Vacation'; return true;"" onMouseOut=""top.status=''; return true;"">" & ls_vacid & "</a>")%>	
		<%end if%>
	<%if not (isnull(ls_vacstart) or trim(ls_vacstart) = "" ) then%>
		<td colspan=1 width=75 valign=top align=left><%=ls_vacstart%>
		<%end if%>
	<%if not (isnull(ls_vacend) or trim(ls_vacend) = "") then%>
		<td colspan=1 width=75 valign=top align=left><%=ls_vacend%>
		<%end if%>	<%if not (isnull(ls_vaccomments) or trim(ls_vaccomments) = "") then%>
		<td colspan=1 width=300 valign=top align=left><%=ls_vaccomments%>
		<%end if%>	</td></table><%
	rs.movenext
	wend

rs.closeset rs = nothing
%>
<form>

	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='window.close();' id=button2 name=button2>

</form>

</BODY>
</HTML>
