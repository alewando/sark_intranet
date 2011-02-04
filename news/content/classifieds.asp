<%Response.Buffer = True%>
<!--
Developer:    
Date:         
Description:  Allows the view and deletion of classified ads.
-->

<!-- #include file="../section.asp" -->


<%
dim desc_length, desc_txt, SQL

desc_length=40
'SetWebMaster "aharshbarger"

If Request("ACTION") = "DELETE" Then
	if Request.Form("DELETE_CHKBX").Count then
	   x = 1
       while(x < Request.Form("DELETE_CHKBX").Count + 1)
			SQL = ""
			SQL = SQL + "DELETE FROM CLASSIFIEDS "
			SQL = SQL + "WHERE CLASSIFIEDS_ID = " & Request.Form("DELETE_CHKBX")(x)
			DBQuery(sql)
            x = x + 1
       wend
    end if
End If

SQL = ""
SQL = SQL + "SELECT cl.Classifieds_ID, em.FirstName, em.LastName, em.EmployeeID, "
SQL = SQL + "cl.TimeStamp, cl.Description, cl.Status, em.UserName "
SQL = SQL + "FROM Classifieds cl, Employee em "
SQL = SQL + "WHERE cl.Employee_ID = em.EmployeeID "
SQL = SQL + "ORDER BY cl.TimeStamp DESC"
set rs = DBQuery (SQL)
%>

<FORM NAME="clsForm" ACTION="classifieds.asp" METHOD="POST">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
<TR>
	<TD ALIGN=LEFT>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR=black>
		Are your closets filled with stuff that you never use?  Or, is that old 386 computer
		just sitting on the floor collecting dust bunnies?  Well this is the page for you!!
		The classifieds can be used to advertise items/services for sale or products that
		your are interested in buying. These ads can all be placed FREE of charge!!
		</FONT>
	</TD>
</TR>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=10></TD></TR>  
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=10></TD></TR>  
</TABLE>

<TABLE  BORDER=0 CELLSPACING=0 CELLPADDING=2 WIDTH="100%">
<TR>
	<TD BGCOLOR=#FFFFFF ALIGN=LEFT COLSPAN=6>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR=black>
		<% if not Session("isGuest") then %>
		[<a href="javascript:addEntry();" onMouseOver="top.status='Add classified.'; return true;" onMouseOut="top.status=''; return true;">Add</a>]&nbsp;&nbsp;&nbsp;
		[<a href="javascript:deleteEntry();" onMouseOver="top.status='Remove all the selected classified items.'; return true;" onMouseOut="top.status=''; return true;">Remove marked</a>]
		<% end if %>
		</FONT>
	</TD> 
</TR>
<TR><TD BGCOLOR=#FFFFFF ALIGN=LEFT COLSPAN=6><IMG SRC="../../common/images/spacer.gif" HEIGHT=10></TD></TR>  
<TR>
	<TD BGCOLOR=#FFFCCC ALIGN=Left><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Mark</B></FONT></TD>
	<TD BGCOLOR=#FFFCCC ALIGN=Left><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Description</B></FONT></TD>
	<TD BGCOLOR=#FFFCCC ALIGN=Left><FONT SIZE=1 FACE="ms sans serif, arial, geneva">&nbsp;&nbsp;</FONT></TD>
	<TD BGCOLOR=#FFFCCC ALIGN=Left><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Posted by</B></FONT></TD>
	<TD BGCOLOR=#FFFCCC ALIGN=Left><FONT SIZE=1 FACE="ms sans serif, arial, geneva">&nbsp;&nbsp;</FONT></TD>
	<TD BGCOLOR=#FFFCCC ALIGN=Left><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><B>Posted date</B></FONT></TD>
</TR>
<TR>
	<TD COLSPAN=6 BGCOLOR=gray HEIGHT=1></TD>
</TR>

<%while not rs.eof%>

<% 	
	date_val=rs("TimeStamp")
	desc_txt = Left(rs("Description"), desc_length)
	if isnull(desc_txt) or desc_txt="" then desc_txt = "[Description not available]"
%>

<TR VALIGN =TOP>
    <TD BGCOLOR=#FFFCCC ALIGN=CENTER VALIGN=CENTER>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366">
	<%if (Session("username")) = rs("UserName") or hasRole("WebMaster") then%>
		<INPUT TYPE=CHECKBOX VALUE=<%=rs("Classifieds_ID")%> NAME="DELETE_CHKBX">
	<%else%>
		---
	<%end if%>  	
		</FONT>
	</TD> 
	<TD BGCOLOR=#FFFCCC VALIGN=CENTER>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366">
		<A HREF="clsfd_details.asp?CLSID=<%=rs("Classifieds_ID")%>"><%=desc_txt%></A><%if (Len(desc_txt) = desc_length) then%>...<%end if%>
		</FONT>
	</TD>
  	<TD BGCOLOR=#FFFCCC VALIGN=CENTER><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366">&nbsp;</FONT></TD>
    <TD BGCOLOR=#FFFFCC VALIGN=CENTER>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366">
		<A HREF="../../directory/content/details.asp?EmpID=<%=rs("EmployeeID")%>"><%=rs("FirstName") & " " & rs("LastName")%></A>
		</FONT>
	</TD>
  	<TD BGCOLOR=#FFFCCC VALIGN=CENTER><FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366">&nbsp;</FONT></TD>
    <TD BGCOLOR=#FFFFCC ALIGN=CENTER VALIGN=CENTER>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR="#333366"><%=DatePart("m", date_val)%>/<%=DatePart("d", date_val)%>/<%=DatePart("yyyy", date_val)%></FONT>
	</TD>
</TR>
<%
rs.movenext()
Wend
%>
</TABLE>

<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
</FORM><P>
</CENTER>

<!-- #include file="../../footer.asp" -->

<SCRIPT LANGUAGE="javascript">

function addEntry() {
	winTech=window.open("clsfd_add.asp", "AddClassified","height=350,width=580,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbar=1")
}

function deleteEntry() {
	document.clsForm.ACTION.value = "DELETE";
	document.clsForm.submit();	
}
</SCRIPT>


<%
rs.close()
set rs = nothing
%>