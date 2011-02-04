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
If hasRole("SolutionServices") Then
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
	
	If  Request("ISIE55") & "x" = "1x" Then
		'--------------------------------------------------------
		' A hidden form variable was used to identify the user as
		' having an IE 5.5 browser. These browsers can't receive
		' the "readonly" attribute for the file form field. This
		' information must be passed through a form field, since 
		' file uploads after the first file won't have valid 
		' values in the Request.ServerVariables("HTTP_USER_AGENT") 
		'--------------------------------------------------------
		isMsie=1
		is55=1

	End If
	
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
	
	If (fso.FolderExists(destFile)) = False Then
		' Prevents error of uploading a file with the same names as a folder
		fso.MoveFile uploadedFile, server.mappath(relativePath) & "\" 
	End If

	Set fso = Nothing

Else
	'--------------------------------------------------------------------------------
	' Prior to the first file upload we can determine the browser information.
	' Subsequent iterations will require a client script to set a hidden form field.
	'--------------------------------------------------------------------------------
	isMsie=instr(Lcase(Request.ServerVariables("HTTP_USER_AGENT")),"msie")
	is55=instr(Request.ServerVariables("HTTP_USER_AGENT"),"5.5")

End If

%>

<TABLE BGCOLOR=SILVER BORDER=0 WIDTH='90%'>
<tr><td COLSPAN=2 ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
Take the following actions to upload a document into the repository:</td></tr>
<TR><TD COLSPAN=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>1)</B>&nbsp;&nbsp; </td><td>Click the "Browse..." button to open a "Choose File" dialog</td></tr>
<TR><TD COLSPAN=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>2)</B>&nbsp;&nbsp; </td><td>Use the dialog to select the desired file to upload</td></tr>
<TR><TD COLSPAN=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>3)</B>&nbsp;&nbsp; </td><td>Click the "Open" button</td></tr>
<TR><TD COLSPAN=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>4)</B>&nbsp;&nbsp; </td><td>Click the "Upload Document" button to upload the document into the repository</td></tr>
<TR><TD COLSPAN=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top><font size=2 face="ms sans serif, arial, geneva"  color=black>
&nbsp;&nbsp;<B>5)</B>&nbsp;&nbsp; </td><td>Click the "Cancel" button when you are finished uploading documents</td></tr>
</TABLE>

<form enctype="multipart/form-data" action="http://<%= Request.ServerVariables("SERVER_NAME") %>/scripts/cpshost.dll?PUBLISH?http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")%>" method=post id=uploadForm name=uploadForm onsubmit="return checkFilename()">
<TABLE BGCOLOR=SILVER BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH='90%'>
<tr><td align="left"><input name="THE_FILE" type="file" size="40" <%If isMsie = 0 Or is55 = 0 Then%> readonly>
<%End If%>
</td></tr>
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
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="3"></TD></TR>
<TR><TD COLSPAN=2><IMG SRC="/images/common/spacer.gif" HEIGHT="3"></TD></TR>
<%If isMsie <> 0 And is55 <> 0 Then%>
<tr><td ALIGN=LEFT VALIGN=top><font size=4 face="ms sans serif, arial, geneva"  color=red>
To avoid upload errors, don't type directly into the above input field!</td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<%End If%>
</TABLE>

<input type=hidden name="TargetURL" value="http://<%= Request.ServerVariables("SERVER_NAME") %>/<%=UPLOAD_DIR%>/">
<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
<INPUT TYPE=HIDDEN NAME=CURRENTFOLDER VALUE="<%Response.Write relativePath%>">
<INPUT TYPE=HIDDEN NAME=ISIE55 VALUE="0">
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
	var agt=navigator.userAgent.toLowerCase();
	
	if (frm.THE_FILE.value + 'x' != 'x')
	{
		valid = true;
		frm.ACTION.value = '<%=ACTION_UPLOAD%>';

		isV55 = (agt.indexOf("5.5") != -1) && (agt.indexOf("msie") != -1);
		if ( isV55 )
			frm.ISIE55.value = '1';
	}	
	else 
	{
		alert('Please select a document!');			
	}	

	return valid
}
</SCRIPT>
