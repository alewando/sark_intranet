<%response.buffer = true%>
<!--
Developer:    Dave Podnar
Date:         7/7/2000
Description: Main script for the Document Repository. The repository is file system based
             and its root is a subdirectory of the scripts directory. The follow functionality 
             is provided: 
			- Listing of the current area's sub folders and documents
			- Solution services members of the current area can add/delete documents and
			  add sub folders
			- Only the web administrator can delete folders.

			Uses 3 hidden form variables:
			ACTION = has nothing or "DELETE". If "DELETE" then we need to delete the
			         marked files and folders.
			CURRENTFOLDER = The user's current folder in the repository. The user navigates
			                through the repository in a similar fashion to "cd dir1/dir2".
			NEWFOLDER = If a value is present, we need to create the new folder as a sub
			            folder of the current folder.
			
Scripts: CommonVars.asp = Several common constants used throughout the scripts
         View_Repository.asp = This file
         Add_Document.asp = Functionality to add documents into the repository and move
                            a document to the correct folder.
		 Not_Privileged.asp = If the current user has reached the repository by accident
		                      as is not a member of Solution Services, they a message.
-->

<!-- #include file="CommonVars.asp" -->

<%
If hasRole("SolutionServices") Then
	i=0
Else
	Response.Redirect NOT_PRIVILEGED_SCRIPT
End If
%>

<!-- #include file="../section.asp" -->

<%

dim relativePath 'Will contain the relative directory path

if  request.querystring("dirPath") & "x" <> "x" then 
	relativePath = request.querystring("dirPath")
elseif  Request("CURRENTFOLDER") & "x" <> "x" then 
	relativePath = Request("CURRENTFOLDER")
else
	'dirPath is null so we initialize
	relativePath = REPOSITORY_ROOT_DIR
end if

'------------------------------------------------------------------
' The lTechArea variable will contain a solution service's area or
' be set to the empty string.
'------------------------------------------------------------------
if relativePath <> REPOSITORY_ROOT_DIR then
	i=instr(relativePath,"/")+1
	j=instr(mid(relativePath,i),"/")
	if j=0 then
		'This is the top folder of an area
		lTechArea=mid(relativePath,i)
	else
		'This is a sub folder of an area
		lTechArea=mid(relativePath,i,j-1)
	end if
else
	lTechArea=""
end if

'Response.Write "lTechArea=" & lTechArea & "<br>"

'------------------------------------------------------------------------------
' Determines if the current user is in any solution services practice areas
'------------------------------------------------------------------------------
sql = "SELECT Tech_Area.Tech_Area, Tech_Specialists.Tech_Specialist_Type_ID, " _
	  & " Employee.Username, Employee.EmployeeID " _
      & " FROM Employee, Tech_Specialists, Tech_Area " _
      & " WHERE Employee.Username = '" & Lcase(session("Username")) & "' " _
      & " AND Tech_Area.Tech_Area = '" & Lcase(lTechArea) & "' " _
      & " AND (Tech_Specialists.Tech_Specialist_Type_ID = 1 " _
      & " OR Tech_Specialists.Tech_Specialist_Type_ID = 2) " _
      & " AND Employee.EmployeeID = Tech_Specialists.Employee_ID " _
      & " AND Tech_Specialists.Tech_Area_ID = Tech_Area.Tech_Area_ID "

set rs = DBQuery(sql)

If Not rs.eof Then
	canMaintainRepos=true		'User is in the current practice area, they can delete docs and add folders
Else 
	canMaintainRepos=false		'User NOT in the current practice area, they can't delete docs or add folders
End If

'Response.Write "canMaintainRepos=" & canMaintainRepos & "<br>"

rs.close

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = fso.GetFolder(server.mappath(relativePath))

permissionFlag=NO_PERMISSION_ERROR

If Request("ACTION") = ACTION_DELETE Then

	If Request.Form("DELETE_FILE").Count And (hasRole("WebMaster") Or canMaintainRepos) Then
		'----------------------------------------------------------------------------
		'User has requested to Delete several documents and has the necessary
		'permission to perform the delete. Iterate through the list and delete those 
		'which have been marked.
		'----------------------------------------------------------------------------
	   x = 1
       While(x < Request.Form("DELETE_FILE").Count + 1)
       
            'Response.Write "Deleting[" & x & "]=" & server.mappath(Request.Form("DELETE_FILE")(x)) & "<BR>"
            fso.DeleteFile(server.mappath(Request.Form("DELETE_FILE")(x)))
            x = x + 1
       Wend
    Elseif Request.Form("DELETE_FILE").Count Then
		'----------------------------------------------------------------------------
		'User has requested to Delete several documents and does NOT have the necessary
		'permission to perform the delete. Set flag for error message.
		'----------------------------------------------------------------------------
    	permissionFlag=NO_DELETE_DOCUMENT_PERMISSION
    End If

	If Request.Form("DELETE_FOLDER").Count And hasRole("WebMaster") Then
	   x = 1
       While(x < Request.Form("DELETE_FOLDER").Count + 1)
       
            'Response.Write "Deleting[" & x & "]=" & server.mappath(Request.Form("DELETE_FOLDER")(x)) & "<BR>"
            fso.DeleteFolder(server.mappath(Request.Form("DELETE_FOLDER")(x)))
            x = x + 1
       Wend
	Elseif Request.Form("DELETE_FOLDER").Count Then
		'----------------------------------------------------------------------------
		'User has requested to Delete several folders and does NOT have the necessary
		'permission to perform the delete. Set flag for error message.
		'----------------------------------------------------------------------------
    	permissionFlag=NO_DELETE_FOLDER_PERMISSION
    End If
End If

If Request("NEWFOLDER") & "x" <> "x" Then
	'Response.Write "Creating Folder:" & server.mappath(relativePath & "/" & Request("NEWFOLDER")) & "<BR>"

	newFolder = server.mappath(relativePath & "/" & Request("NEWFOLDER"))
	If (fso.FolderExists(newFolder)) = False And (fso.FileExists(newFolder)) = False Then
		fso.CreateFolder(newFolder)
	End If
End If

Set collFolders = objFolder.subfolders
Set collFiles = objFolder.files

%>
<FORM NAME="repositoryForm" ACTION="view_repository.asp" METHOD="POST">

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<font size=1 face="ms sans serif, arial, geneva"  color=black>
This section contains the Solution Service's document repository. The repository should only be used 
to upload and download documents, NOT to automatically open documents. Click a folder name to navigate through the repository. <br>
</font></td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td ALIGN=LEFT VALIGN=top>
<b><font size=1 face="ms sans serif, arial, geneva"  color=black>
To down load a document, right-click a document name and select "Save Target As" or "Save Link As". The document can then be opened via 
File Explorer or its respective application. <font color=red>Do NOT left-click a document name, since this often produces quirky application behavior.
</font></font></td></tr>
<TR></B><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<%If permissionFlag = NO_DELETE_FOLDER_PERMISSION Then%>
<tr><td ALIGN=LEFT VALIGN=top>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
Permission Error: You do NOT have permission to delete folders!
</font></td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<%ElseIf permissionFlag = NO_DELETE_DOCUMENT_PERMISSION Then%>
<tr><td ALIGN=LEFT VALIGN=top>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
Permission Error: You do NOT have permission to delete documents!
</font></td></tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<%End If%>
</TABLE>

<TABLE  BORDER=0 CELLSPACING=0 CELLPADDING=2 WIDTH="100%">
<TR>
	<TD BGCOLOR=#FFFFFF ALIGN=LEFT COLSPAN=3>
		<FONT SIZE=1 FACE="ms sans serif, arial, geneva" COLOR=black>
			[<a href="javascript:openWindow('<%=server.URLEncode(ADD_DOCUMENT_SCRIPT & "?dirPath=" & relativePath)%>');" onMouseOver="top.status='Add document.'; return true;" onMouseOut="top.status=''; return true;">Add Document</a>]&nbsp;&nbsp;&nbsp;
			<%If hasRole("WebMaster") Or canMaintainRepos = True Then
			'-------------------------------------------------------------------------------
			' User is WebMaster or a member of the current Solution Services practice area and
			' is allowed to perform the following:
			' - Add folders
			' - Delete documents
			'-------------------------------------------------------------------------------
			%>
			[<a href="javascript:addFolder();" onMouseOver="top.status='Add folder.'; return true;" onMouseOut="top.status=''; return true;">Add Folder</a>]&nbsp;&nbsp;&nbsp;
			[<a href="javascript:deleteEntry();" onMouseOver="top.status='Remove all marked items.'; return true;" onMouseOut="top.status=''; return true;">Remove marked</a>]&nbsp;&nbsp;&nbsp;
			[<a href="javascript:clearCheckBoxes();" onMouseOver="top.status='Clear all marks.'; return true;" onMouseOut="top.status=''; return true;">Clear Checkboxes</a>]
			<%end if %>
		</FONT>
	</TD> 
</tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
</table>

<table class=tableShadow width=100% cellspacing=0 cellpadding=2 border=0 bgcolor=#ffffcc>
<tr><td colspan=3 ><font size=2.5 face= "ms sans serif, arial, geneva"  color = black>
	Current Folder : <b>
	<%
	i=instr(relativePath,"/")+1
	
	if relativePath <> REPOSITORY_ROOT_DIR then
		Response.write mid(relativePath,i)
	else
		Response.write "Solution Services"
	end if
	%>
	</b></td></tr>
<tr><td colspan=3 bgcolor=gray height=1>
</td></tr>

<%if relativePath <> REPOSITORY_ROOT_DIR then%>
	<tr><td width='2%' ALIGN=left VALIGN=top><font size=1 face= 'ms sans serif, arial, geneva'  color = navy>&nbsp;&nbsp;&nbsp;</td>
	<td width='98%'><font size=1 face= 'ms sans serif, arial, geneva'  color = navy>
	<A HREF='<%=request.servervariables("script_name")%>?dirPath=<%=server.URLEncode(fso.GetParentFolderName(relativePath))%>' onMouseOver="top.status='Move up one level.'; return true;" onMouseOut="top.status=''; return true;">[&nbsp;Up One Level&nbsp;]</A></td></tr>
<%end if

if collFolders.count <> 0  then
	'-------------------
	'Display sub folders
	'If the current user is the webmaster
	'they get a checkbox to mark directories for deletion
	'-------------------
	For Each whatever in collFolders
		response.write "<tr><TD width='2%' ALIGN=left VALIGN=top><FONT SIZE=1 FACE='ms sans serif, arial, geneva' COLOR=black>"
		
		if hasRole("WebMaster") Then
			'Only the web master can delete folders
			response.write "<INPUT TYPE=CHECKBOX VALUE='" & relativePath & "/" & whatever.name & "' NAME='DELETE_FOLDER'>"
		else
			response.write "&nbsp;&nbsp;&nbsp;"
		end if
		
		Response.write "</FONT></TD>"

		Response.Write "<td width='98%' ><font size=1 face= 'ms sans serif, arial, geneva'  color = navy>"
		Response.write "<A HREF='" & request.servervariables("script_name")
		Response.write "?dirPath=" & server.URLEncode(relativePath & "/" & whatever.name) & "'>"
		Response.write whatever.name & "</A></td></tr>"
	Next
else
	Response.write "<tr><td width='2%' ALIGN=left VALIGN=top>&nbsp;&nbsp;&nbsp;</td>"
	Response.write "<td width='98%' ALIGN=left VALIGN=top><font size=1 face= 'ms sans serif, arial, geneva'  color = black>no sub folders </td></tr>"
end if
%>
<tr><td colspan=2 bgcolor=gray height=1>
</td></tr>
<tr><td colspan=3>
	<table border=0 width=100% ><%							if collFiles.count <> 0  then
		'------------------------------------------
		'Display files in the respective directory
		'If the current user has canMaintainRepos=true
		'they get a checkbox to mark files for deletion
		'------------------------------------------
		response.write "<tr>"		if hasRole("WebMaster") or canMaintainRepos=true Then
			response.write "<th Align=left vAlign=top width='10%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"			response.write "Mark</font></th>"
			response.write "<th Align=left vAlign=top width='52%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "Document Name</font></th>"
			response.write "<th Align=left vAlign=top width='26%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "Date Created</font></th>"
			response.write "<th Align=right vAlign=top width='12%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "Size</font></th>"
			response.write "</tr>"
		else
			response.write "<th Align=left vAlign=top width='1%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"			response.write "&nbsp;</font></th>"
			response.write "<th Align=left vAlign=top width='61%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "Document Name</font></th>"
			response.write "<th Align=left vAlign=top width='26%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "Date Created</font></th>"
			response.write "<th Align=right vAlign=top width='12%'><font size=2 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "Size</font></th>"
			response.write "</tr>"
		end if
		
		For Each whatever in collFiles
			response.write "<tr><TD ALIGN=left VALIGN=top><FONT SIZE=1 FACE='ms sans serif, arial, geneva' COLOR=black>"
			
			if hasRole("WebMaster") or canMaintainRepos=true Then 
				response.write "<INPUT TYPE=CHECKBOX VALUE='" & relativePath & "/" & whatever.name & "' NAME='DELETE_FILE'>"
			else
				response.write "&nbsp;"
			end if
		
			response.write "</FONT></TD>"
			response.write "<td vAlign=top><font size=1 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write "<A HREF='"
			
			'------------------------------------------------------------------
			' PlusToHex() replaces all '+' characters in a URLEncoded string
			' into the '%20' string. There have been problems when attempting to
			' download a document, when its name has embedded spaces i.e. "This File.doc". 
			' The URLEncode method properly converts the spaces to '+' characters
			' i.e. "This+File%2Edoc". However, the web server can't find the encoded document. 
			' The web server is able to find "This%20File%2Edoc".
			'----------------------------------------------------------------
			response.write PlusToHex(server.URLEncode(relativePath & "/" & whatever.name))
			response.write "'>" & whatever.name & "</A></font></td>"
			response.write "<td vAlign=top><font size=1 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write whatever.datecreated & "</font></td>"
			response.write "<td Align=right vAlign=top><font size=1 face= 'ms sans serif, arial, geneva'  color = black>"
			response.write round(whatever.size/1024,0) & "K</font></td>"
			response.write "</tr>"
		Next
	else
		response.write "<tr><td colspan=2><font size=1 face= 'ms sans serif, arial, geneva'  color = black>"
		response.write "&nbsp;&nbsp;&nbsp;no files </td></tr>"
	end if
%>					</table>
		
</td></tr>
	
</td>
</tr></table>

<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
<INPUT TYPE=HIDDEN NAME=CURRENTFOLDER VALUE="<%Response.Write relativePath%>">
<INPUT TYPE=HIDDEN NAME=NEWFOLDER VALUE="">
</FORM><P>


<%
'Need to De-reference all the objects
set collFiles = Nothing
set collFolders = Nothing
set objFolder = Nothing
set fso = Nothing
%>


<!-- #include file="../../footer.asp" -->


<SCRIPT LANGUAGE="javascript">

function isDigit(aChar)
{
	if ( aChar >= '0' && aChar <= '9')
		return true;
	else 
		return false;
}

function isAlpha(aChar)
{
	if ( aChar >= 'A' && aChar <= 'Z')
		return true;
	else 
		return false;
}

function addFolder() 
{
	var valid=false;
	var n;
	frm=document.repositoryForm;
	
	//-----------------------------------------------------------------
	// Need to validate the folder name and make sure the following:
	// Can only contain letters, numbers and spaces
	//-----------------
	while (valid != true)
	{
		folderName=prompt('Please type a name for the new folder:','');
	
		if (folderName == null)
		{
			//--------------------------------------------------		
			// The cancel button was clicked
			//--------------------------------------------------
			break;
		}
	
		var tmpString=folderName.toUpperCase();
	
		if ( tmpString.length == 0 )
		{
			valid=false;
			alert('A Folder name must be provided!');
		}
		else
		{
			valid=true;
	
			for (n = 0; valid== true && n < tmpString.length; n++) 
			{
				var sChar = tmpString.charAt(n);

				if ( sChar!=' ' && !isAlpha(sChar) && !isDigit(sChar) )
				{
					valid=false;
					alert('Folder name can only contain alphabetics, numerics and spaces!');
				}
			}
		}
	}
	
	if ( valid == true )
	{
		frm.NEWFOLDER.value = folderName;
		frm.submit();	
	}
	else
		frm.reset();
}

function checkChildWindow() 
{
	var scriptName='<%=VIEW_REPOSITORY_SCRIPT%>',
	    pathName='<%=server.URLEncode(relativePath)%>';
	    
	if ( !window.winUploadDoc || !window.winUploadDoc.open || window.winUploadDoc.closed )
	{
		clearInterval(intervalID);
	
		window.location = scriptName + '?dirPath=' + pathName;
	}
}

function openWindow(URL) 
{
	if ( window.winUploadDoc && window.winUploadDoc.open && !window.winUploadDoc.closed )
	{
		winUploadDoc.document.uploadForm.THE_FILE.focus();
	}
	else
	{
		winUploadDoc=window.open(URL, "OpenDocument","height=500,width=580,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,scrollbars=1");
		intervalID=setInterval('checkChildWindow()', 500);
	}
}

function clearCheckBoxes() 
{
	frm=document.repositoryForm;
	len = frm.elements.length;
	var i=0;

	for( i=0 ; i<len; i++) 
	{
		if (frm.elements[i].name=='DELETE_FILE' || frm.elements[i].name=='DELETE_FOLDER' )
			frm.elements[i].checked=false;
	}
}

function deleteEntry() {
    
	frm=document.repositoryForm;
	len = frm.elements.length;
	var i=0, found=false, deleteList='';

	for( i=0 ; i<len; i++) 
	{
		if (frm.elements[i].name=='DELETE_FILE' && frm.elements[i].checked || 
		    frm.elements[i].name=='DELETE_FOLDER' && frm.elements[i].checked)
		{
			found=true;
			deleteList = deleteList + '\n' + frm.elements[i].value;
		}
	}
    
    if (found == false )
	{
		alert('No items marked for delete!');
	}
	else if ( confirm("Click 'OK', to permanently delete the following folders/documents from the repository?\n" + deleteList) )
	{
		frm.ACTION.value = '<%=ACTION_DELETE%>';
		frm.submit();	
	}
	else
	{
		frm.reset();
	}
}
</SCRIPT>


	

