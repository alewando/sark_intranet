<%
	Dim in_submitted
	in_submitted = Request("submitted")
	
	If in_submitted = "TRUE" Then
		'Redirect to the skills inventory page
	%>
		<script language=javascript>
			window.self.close();
			window.open('../../directory/content/SkillsInv.asp?EmpID=<%=Session("ID")%>','','height=600,widht=600,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,scrollbars=1');
		</script>
	<% 
	End If
%>
<HTML>
<HEAD>
	<TITLE>Skills Inventory Update</TITLE>
</HEAD>
<BODY bgcolor='silver'>
  <TABLE>
    <TR>
	  <TD><FONT color='red'>
		It has been at least three months since you last updated your skills inventory. Please take a
		moment to do it now. This message will continue to appear until your skills inventory 
		has been updated.
	  </TD></FONT>
	</TR>
	<TR><TD>&nbsp;</TD></TR>
	<TR><TD>&nbsp;</TD></TR>
	<TR>
	  <TD>
	    <CENTER>
	    <FORM id='skillUpdate' name='skillUpdate' action='skillsInvUpdate.asp' method='POST'>
		  <INPUT type='submit' value='Update Inventory' id='edit' name='edit'>
		  <INPUT type='button' value="I'll Do It Later" id='cancel' name='cancel' onClick='javascript:window.self.close();'>
		  <INPUT type='hidden' id='submitted' name='submitted' value='TRUE'>
	    </FORM>
		</CENTER>
	  </TD>
    </TR>
  </TABLE>
</BODY>
</HTML>