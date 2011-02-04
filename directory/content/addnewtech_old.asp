<!--
Developer:    GTYLER
Date:         02/16/2000
Description:  Allows user to add a new technology
-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Add New Technology</TITLE>
<!-- #include file="../../style.htm" -->
<script language="JavaScript" type="text/javascript">
<!--

    function checkthensubmit(){
		if (document.frmNewTech.txtNewTechName.value == ""){
			alert('You must enter a new technology.');
			}
		else {
			if (document.frmNewTech.chkExpert.checked){
				document.frmNewTech.expert.value=1
				}
			else {
				document.frmNewTech.expert.value=0
				}
			document.frmNewTech.submit()
			}
		}
			
	function OpenWin(page, EmpID, WinTitle){
		winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=615,width=500,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
		}
			
// -->			
</script>
</HEAD>

<%
'-------------------------------------------------
'   Execute database query for technical areas
'-------------------------------------------------
techarea_sql = "SELECT Tech_Area_ID, Tech_Area FROM Tech_Area ORDER BY Tech_Area"
Set rs = DBQuery(techarea_sql)
%>

<BODY bgcolor=silver link=blue alink=purple vlink=blue><CENTER>

<table border=0 width='90%'><tr><td>
<font size=1 face="ms sans serif, arial, geneva" color=maroon>
<b>Note:</b> The technology you propose will <b>not</b>
appear immediately after you submit it.  It must first be approved
by management.
</font>
</td></tr></table>

<form NAME="frmNewTech" ACTION="addnewtech_process.asp">
	<INPUT type=hidden name="empID" value="<%=Request.QueryString("EmpID")%>">
	<INPUT type=hidden name="expert" value="">
	<INPUT TYPE=hidden NAME="hidUsername" VALUE="<%=Session("Username")%>">
	
	<table border=0 width='90%'>
	<tr><td><font size=1 face="ms sans serif, arial, geneva">
		Please specify the technical area to which the new
		technology should belong.
	</FONT></td></tr>
	<tr><td><font size=1 face="ms sans serif, arial, geneva">
    <%
    'List all available technical areas in a dropdown list
    %>
    <b>Area:</b><br>
    <center>
    <SELECT NAME=selArea>
        <%
        While Not rs.EOF
            Response.Write("<OPTION VALUE='" & rs.Fields("Tech_Area_ID") & _
                           "'>" & rs.Fields("Tech_Area") & "</OPTION>")
	        rs.MoveNext
        Wend
        %>
    </SELECT>
    </center>

	<P>&nbsp;</P>
	</font></td></tr>
	<tr><td><font size=1 face="ms sans serif, arial, geneva">
		Please enter the name of the proposed new technology.
	</font></td></tr>
	
	<tr><td><font size=1 face="ms sans serif, arial, geneva">
		<b>Name:</b><br>
		<center>
		<INPUT TYPE=TEXT NAME="txtNewTechName" ID="txtNewTechName" VALUE="" MAXLENGTH="50">
		</center>
	    <P>&nbsp;</P>
	</font></td></tr>
	
	<tr><td><font size=1 face="ms sans serif, arial, geneva">
		Please enter a brief description of the proposed new technology.
	</FONT></td></tr>
	<tr><td><font size=1 face="ms sans serif, arial, geneva">
		<b>Description:</b><br>
		<center>
		<INPUT TYPE=TEXT NAME="txtNewTechDesc" ID="txtNewTechDesc" VALUE="" SIZE="50" MAXLENGTH="200">
		</center>
	</font></td></tr>
	</table>
	
    <table border=0 width='90%'>
	<tr height='60' valign='center'><td><font size=1 face="ms sans serif, arial, geneva">
		<INPUT TYPE=checkbox NAME=chkExpert CHECKED>
		I consider myself an expert in this technology.
	</FONT></td></tr>
	</table>
	
	<INPUT TYPE=button CLASS=button VALUE="Submit" OnClick='checkthensubmit();'>
	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='window.close();OpenWin("techexperts.asp", "<%=Request.QueryString("EmpID")%>", "EditSARKTech");'>

</form>

</CENTER></BODY>
</HTML>
