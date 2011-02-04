<%Response.Buffer = True%>

<% pageTitle = "Details" %>

<!--
Developer:    
Date:         
Description:  Allows the full view of a particular classified ad.
-->

<!-- #include file="../section.asp" -->

<%
dim cls_id, date_val, date_txt, SQL

cls_id = Clean(Request("CLSID"))

'SQL = ""
'SQL = SQL + "SELECT Classifieds_ID "
'SQL = SQL + "FROM Classifieds "
'SQL = SQL + "WHERE Classifieds_ID > " & cls_id
'set rs = DBQuery (SQL)
'next_id=rs("Classifieds_ID")

'SQL = ""
'SQL = SQL + "SELECT Classifieds_ID "
'SQL = SQL + "FROM Classifieds "
'SQL = SQL + "WHERE Classifieds_ID < " & cls_id
'set rs = DBQuery (SQL)

'while not rs.eof
'	prev_id=rs("CLASSIFIEDS_ID")
'	rs.movenext()
'wend


SQL = ""
SQL = SQL + "SELECT cl.Classifieds_ID, em.FirstName, em.LastName, em.EmployeeID, "
SQL = SQL + "cl.TimeStamp, cl.Description, cl.LastModified, em.UserName "
SQL = SQL + "FROM Classifieds cl, Employee em "
SQL = SQL + "WHERE cl.Classifieds_ID = " & cls_id
SQL = SQL + " AND cl.Employee_ID = em.EmployeeID "
set rs = DBQuery (SQL)

desc_txt = rs("Description")
if isnull(desc_txt) or desc_txt="" then desc_txt = "[Description not available]"

cls_lastModified = rs("LastModified")
%>

<FORM NAME="clsForm" ACTION="details.asp" METHOD="POST">

<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=2 WIDTH="100%">
<TR>
	<TD BGCOLOR=#FFFFFF ALIGN=LEFT>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva">
		[<a href="classifieds.asp" onMouseOver="top.status='View the full classified listing.'; return true;" onMouseOut="top.status=''; return true;">View Listing</a>]&nbsp;&nbsp;&nbsp;		
		<%if UCase(Session("username")) = rs("UserName") then%>
			[<a href="javascript:editEntry()" onMouseOver="top.status='Edit classified.'; return true;" onMouseOut="top.status=''; return true;">Edit</a>]&nbsp;&nbsp;&nbsp;
		<%end if%>
		</FONT>
	</TD>
	<TD BGCOLOR=#FFFFFF ALIGN=RIGHT>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva">
<!--		[<a href="clsfd_details.asp?CLSID="<%'=prev_id%>" onMouseOver="top.status='Previous classified.'; return true;" onMouseOut="top.status=''; return true;">Previous</a>]&nbsp;&nbsp;&nbsp;
		[<a href="clsfd_details.asp?CLSID="<%'=next_id%>" onMouseOver="top.status='View the next classified.'; return true;" onMouseOut="top.status=''; return true;">Next</a>]
-->		</FONT>
	</TD>
</TR>
<TR><TD><IMG SRC="/image/common/spacer.gif" HEIGHT=10></TD></TR>  
</TABLE>

<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=2 WIDTH="100%">
<TR>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT NOWRAP WIDTH=25%><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Posted by:</B></FONT></TD>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT WIDTH=75%><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><A HREF="../../directory/content/details.asp?EmpID=<%=rs("EmployeeID")%>"><%=rs("FirstName") & " " & rs("LastName")%></A></FONT></TD>
</TR>
<TR>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT NOWRAP><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Posted date:</B></FONT></TD>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><%=rs("TimeStamp")%></FONT></TD>
</TR>
<%if not(isnull(cls_lastModified)) or (cls_lastModified<>"") then %>
<TR> 
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT NOWRAP><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Last modified date:</B></FONT></TD>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><%=cls_lastModified%></FONT></TD>
</TR>
<%end if%>
<TR>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT VALIGN=TOP NOWRAP><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Description:</B></FONT></TD>
	<TD BGCOLOR=#FFFFCC ALIGN=LEFT><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><%=desc_txt%></FONT></TD>
</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
</FORM>
</CENTER>

<!-- #include file="../../footer.asp" -->

<SCRIPT LANGUAGE="javascript">
function editEntry() {
	winTech=window.open("clsfd_edit.asp?CLSID=<%=cls_id%>", "AddClassified","height=350,width=580,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")
}
</SCRIPT>

<%
rs.close()
set rs = nothing
%>
