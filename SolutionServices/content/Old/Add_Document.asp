<%response.buffer = true%>
<!--
Developer:    Dave Podnar
Date:         7/7/2000
Description: Script provides the upload document form and moves uploaded files into the proper
             directory. It is called by:
             1) View_repository.asp when the user chooses "Add Document" into the repository
             2) After the upload is successful, this script is automatically called to move 
                the newly uploaded file and provide the upload form again. Look at the form's
                action (not the request variable) and you'll see 
                "scripts/cpshost.dll?PUBLISH?http://.../add_document.asp"
                This sends the uploading file and form variables to cpshost.dll and then cpshost.dll
                calls add_document.asp with the form variables.
-->
<%
'------------------------------------------------------------------------------
' Determines if the current user is in at least one solution services practice area
' If so, they have permission to view the repository and add documents to it
'------------------------------------------------------------------------------
If IsNull(Session("isSolutionServices")) Then

	sql = "SELECT Tech_Specialists.Tech_Specialist_Type_ID, " _
	& " Employee.Username, Employee.EmployeeID " _
	& " FROM Employee, Tech_Specialists " _
	& " WHERE Employee.Username = '" & Lcase(session("Username")) & "' " _
	& " AND Employee.EmployeeID = Tech_Specialists.Employee_ID " _
	& " AND (Tech_Specialists.Tech_Specialist_Type_ID = 1 OR Tech_Specialists.Tech_Specialist_Type_ID = 2) "

	set rs = DBQuery(sql)

	If Not rs.eof Or Lcase(session("Username"))="dschrader" Or Lcase(session("Username"))="ggelasi" Or Lcase(Session("username")) = Lcase(Application("WebMaster")) Then
		Session("isSolutionServices")=True
	Else
		Session("isSolutionServices")=False
	End If
	rs.close
End If

If Session("isSolutionServices")=True Then
	i=0
Else
	Response.Redirect NOT_PRIVILEGED_SCRIPT
End If
%>

<HTML>
<HEAD>
<TITLE>Upload a Document into the Repository</TITLE>
<!-- #include file="../../style.htm" -->
</HEAD>
<!-- <BODY BGCOLOR=SILVER onUnload="closeWindow();"><center> -->

<BODY BGCOLOR=SILVER ><center>

<!-- #include file="../../script.asp" -->

<!-- #include file="CommonVars.asp" -->

<%
dim relativePath 'Contains relative directory path

If  request.querystring("dirPath") & "x" <> "x" Then 
	relativePath = request.querystring("dirPath")
Elseif  Request("CURRENTFOLDER") & "x" <> "x" Then 
	relativePath = Request("CURRENTFOLDER")
Else
	'dirPath is null so we initialize
	relativePath = REPOSITORY_ROOT_DIR
End If

If Request("ACTION") = ACTION_UPLOAD Then
	
	Set fso = CreateObject("Scripting.FileSystemObject")

	'Response.Write "FilePath=" & Request("FilePath") & "<br>"
	'Response.Write "FileName=" & Request("FileName") & "<br>"
	'Response.Write "FileExtention=" & Request("FileExtention") & "<br>"

	uploadedFile= Request("FilePath") & Request("FileName") & Request("FileExtention")
	destFile=server.mappath(relativePath) & "\" & Request("FileName") & Request("FileExtention")
		
	If (fso.FileExists(destFile)) = True Then
		'----------------------------------------------------------------------
		' The file already exists in the directory, we need to delete it since
		' move gives an error if the file is already there.
		'----------------------------------------------------------------------
		fso.DeleteFile destFile, True
	End If
	
	fso.MoveFile uploadedFile, server.mappath(relativePath) & "\" 

	Set fso = Nothing
End If

%>

<TABLE BGCOLOR=SILVER BORDER=0 WIDTH=450>
<ul>
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
Take these actions to upload a document into the repository:</td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>1)</B>&nbsp;&nbsp; Click the "Browse..." button to open a "Choose File" dialog</td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>2)</B>&nbsp;&nbsp; Use the dialog to select the desired file to upload</td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>3)</B>&nbsp;&nbsp; Click the "Open" button</td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>4)</B>&nbsp;&nbsp; Click the "Upload Document" button to upload the document into the repository</td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>5)</B>&nbsp;&nbsp; Click the "Cancel" button when you are finished uploading documents</td></tr>
</ul>
</TABLE>

<form enctype="multipart/form-data" action="http://<%= Request.ServerVariables("SERVER_NAME") %>/scripts/cpshost.dll?PUBLISH?http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")%>" method=post id=uploadForm name=uploadForm onsubmit="return checkFilename()">
<TABLE BGCOLOR=SILVER BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH='90%'>
<tr><td align="left"><input name="THE_FILE" type="file" size="40" readonly></td></tr>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="3"></TD></TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="3"></TD></TR>
<TR>
	<TD VALIGN=TOP ALIGN=LEFT COLSPAN=2>
		<input type=submit class=button value="Upload Document" id=submit1 name=submit1>
		<input type=button class=button value="Cancel" onClick='javascript:closeWindow()';>
	</TD>
</TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="3"></TD></TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="3"></TD></TR>
<TR>
	<TD VALIGN=TOP ALIGN=LEFT COLSPAN=2>
	<%
		If Request("ACTION") = ACTION_UPLOAD Then
			Response.Write "<b>Successfully uploaded document '" & Request("FileName") & Request("FileExtention") & "'</b>"
		else 
			Response.Write "&nbsp;"
		end if
	%>
</TR>
	</TD>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="5"></TD></TR>
</TABLE>

<input type=hidden name="TargetURL" value="http://<%= Request.ServerVariables("SERVER_NAME") %>/<%=UPLOAD_DIR%>/">
<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
<INPUT TYPE=HIDDEN NAME=CURRENTFOLDER VALUE="<%Response.Write relativePath%>">
</form>

</body>
<SCRIPT LANGUAGE="javascript">

function closeWindow()
{
  window.close();
}

function checkFilename() 
{
	frm=document.uploadForm;
	len = frm.elements.length;
	var valid=false;
	
	if (frm.THE_FILE.value + 'x' != 'x')
	{
		valid = true;
		frm.ACTION.value = '<%=ACTION_UPLOAD%>';
	}	
	else 
	{
		alert('Please select a document!');			
	}	

	return valid
}
</SCRIPT>
