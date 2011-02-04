<!--
Developer:    GINRICE code modified by DMARTIN, MAPGAR
Date:         08/28/2000
Description:  Allows editing of class info
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Update Class Info "addclass.asp"</title>
<!-- #include file="../../style.htm" -->
</head>
<body bgcolor=silver onLoad="window.resizeBy(0, -80)"> 

<%
'-----------------------------------------------
'	Execute Database Query for Exam Info		
'-----------------------------------------------
empID = Request.QueryString("EmpID") 
set rs = DBQuery("SELECT Class_List.Class_Name, Class_List.Class_ID," _
    & " Classes.Class_Start_Date, Classes.Class_End_Date, Classes.Comments, Location.Location_ID, Location.Location_Name" _
	& " FROM Location, Class_list INNER JOIN" _
	& " Classes ON" _
    & " Classes.Class_ID = Class_list.Class_ID" _
	& " WHERE Classes.EmployeeID = " & empid & "")


if not rs.eof then
	
	'-----------------------------------------------
	'	Get Exam Info							
	'-----------------------------------------------
	ls_classid =	  trim(rs("Class_ID"))
	ls_classname = 	trim(rs("Class_Name"))
	ls_classStartdate = trim(rs("Class_Start_Date"))
	ls_classEnddate = trim(rs("Class_End_Date"))
	ls_locationid = trim(rs("Location_ID"))
	ls_locationname = trim(rs("location_name"))
	ls_comments =	trim(rs("Comments"))

	ls_insertupdate = "U"
else
	ls_insertupdate = "I"
end if
%>

<body bgcolor=silver><center>

<table border=0 width="90%" Height="" ><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
This form allows for updating class information. 
</b></font></td></tr></table>

</script>

<FORM action=updateclass2.asp name=frmInfo>
<INPUT TYPE="hidden" NAME="EmployeeID" VALUE="<%=empID%>">&nbsp; 

<table border=0 style="HEIGHT: 193px; WIDTH: 452px">
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Class Name:</font></td>
	<td>
		<Select NAME="Class_ID">
			<%set rs = DBQuery("select Class_ID, Class_Name from Class_List order by Class_ID")
				if not rs.eof and not rs.bof then
					do while not rs.eof%>
						<Option Value=<%=rs("Class_ID")%> <%if trim(rs("Class_ID")) = ls_classid then%> Selected<%end if%>><%=trim(rs("Class_Name"))%>
						<% rs.movenext
					loop
					rs.close
				end if%>			
		</Select>
	</td>
</tr>

<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Start Date:</font></td>
	<td><INPUT NAME="class_start_date" VALUE="<%=date()%>" size=11 maxlength=11></td>
</tr>

<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">End Date:</font></td>
	<td><INPUT NAME="class_end_date" VALUE="<%=date()+3%>" size=11 maxlength=11></td>
</tr>


<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Location:</font></td>
	<td>
		<Select NAME="Location_ID">
			<%set rs = DBQuery("select Location_ID, Location_Name from Location order by Location_ID")
				if not rs.eof and not rs.bof then
					do while not rs.eof%>
						<Option Value=<%=rs("Location_ID")%> <%if trim(rs("Location_ID")) = ls_locationid then%> Selected<%end if%>><%=trim(rs("Location_Name"))%>
						<% rs.movenext
					loop
					rs.close
				end if%>			
		</Select>
	</td>
</tr>

<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT NAME="comments" VALUE="" size=40 maxlength=255></td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva"><b>Additional SARKs?</b></font></td>
</tr>
<tr>
<td>&nbsp;</td>
<td>
<!--  information changed to allow multiple people to be sent to the database for classes -->
<table width=100% cellspacing=0 cellpadding=2 border=0>
                   <th width='45%'></th>
                   <th width='10%'></th>
                   <th width='45%'></th>

   <tr> 
      <td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Not Selected</b></td>
      <td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black>&nbsp;</td>
      <td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Selected<br></b></td>
   </tr>

   <TR>
      <TD colspan=3><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD>
   </TR>  

   <tr>
      <td><select MULTIPLE size=10 name="NotSelected" id=NotSelected style="{ width: 100%; }" onfocus="setSelectList(document.frmInfo.Selected,false)";> 
 
      <%
      'Here we are just querying the database for employee names and user names
      'These will be displayed in the table under "Not Selected" so they can be selected
        
      mysql = "select a.lastname, a.firstname, a.EmployeeID from employee a where a.EmployeeID <> '" & empID & "' order by a.lastname"
    
      'Notice all of the sizes here are hard coded. Since we don't have 300 people in the branch,
      'this should not cause a problem. However, if you need more than 300, you'll have to change
      'the sizes here.
      Dim TempArrayFName(300)  'store the person's first name
      Dim TempArrayLName(300)  'store the person's last name
      Dim TempArrayUName(300)  'store the person's user name
      Dim TempDisplayableName(300) 'We'll put first and last name here and use for output
      Dim TempSelected(300)    'so we can identify which people have been selected
      Dim SelectedUserNames(300)
      set rsName = DBQuery(mysql) 'result set
    
      'Manipulate the result set
      i=0
      While Not rsName.eof
        TempArrayFName(i) = rsName("FirstName")
        TempArrayLName(i) = rsName("LastName")
        TempArrayUName(i) = rsName("EmployeeID")
        TempSelected(i) = false
       
        'Now print the names out
        ' Response.Write "'" & TempArrayLName(i) & ", " & TempArrayFName(i) &"'<br>"
      
      %>
      
      <option value=<%=TempArrayUName(i)%>><%=(TempArrayLName(i) & ", " & TempArrayFName(i))%></option>
      
      <%
        rsName.movenext
        i=i+1
      Wend
  
      rsName.close
      set rsName = Nothing    'CLEAR!
      %>
   
      </select></td>
	  	  <td align=center>
	     <table border=0>
	     <tr>
	        <td><input type=button value=" --> " id=button2 name=button2 onClick="moveOption(document.frmInfo.NotSelected,document.frmInfo.Selected,false,false);"></td>
	     </tr>
     	     <tr>	        <td><input type=button value=" <-- " id=button3 name=button3 onClick="moveOption(document.frmInfo.Selected,document.frmInfo.NotSelected,true,false);"></td>	     </tr>     
         </table>	  </td>

	  <td ><select MULTIPLE size=10 name="Selected" id=Selected style="{ width: 100%; }" onfocus="setSelectList(document.frmInfo.NotSelected,false)";>
 	     <%
	        i=0
		    While i < 300 And TempArrayUName(i) & "x" <> "x"

			   If TempSelected(i)=true Then%>
					<option value=<%=TempArrayUName(i)%>><%=(TempArrayLName(i) & ", " & TempArrayFName(i))%></option>
			   <%
			   End If
			   i=i+1
 		    Wend
 	     %>
	  </select></td>
   </tr>

   <TR>
      <TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD>
   </TR>  
 
</table>
</tr>

  <TR>

		<td align=middle colspan=2>
	<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="<%=ls_insertupdate%>"><br>
            <input type=button class=button value="Submit" OnClick='javasctipt:doSubmit()' id=button1 name=button1>
			<input type=button class=button value="Cancel" OnClick='window.close();' id=button2 name=button2>
		</td></TR>
</table></FORM>
</center>

<SCRIPT LANGUAGE="javascript">

/* function that makes sure the information submitted is selected */
function doSubmit()
{
	frm=document.frmInfo;
	setSelectList(frm.Selected,true);	//Turn on all items for server side to process
	frm.submit();	
}

 
/*
 * Sets the "selected" property for all options of a select list.
 */
function setSelectList(oSelect,value)
{
	var i;
	
	for (i = 0; i < oSelect.length; i++) 
	{       
		oSelect[i].selected=value;
		//alert(oSelect[i].value);
	}
}  
   
/*
 * Function is used to move selected options between the "NotSelected" and "Selected" lists.
 *
 * The boolean argument "deSelectFlag" indicates whether the movement is from "Selected" to
 * "NotSelected". This is required since the option text and values in the "Selected" list are 
 * slightly different than the "NotSelected" list option values.
 * 
 * The boolean argument "resetFlag" indicates when a form clear is occuring
 * 
 */
function moveOption(from,to,deSelectFlag,clearFlag)
{
	var flag=false;
	var i,j,l;
	textList = new Array();
	valueList = new Array();
	
	l=0;
	for (i = 0; i < from.length; i++) 
	{       
		j = to.length;   
		if ( from[i].selected )
		{
    		oText =  from[i].text; 
			oValue = from[i].value;

			// Creates a new option in the destination list
			to[j] = new Option (oText, oValue, false, false);
			//to[j].selected = true;
			flag=true;					//Flag that a movement has occurred
		}
	} 

	if (flag)
	{
		//---------------------------------------------------------------------
		// At least one option was moved and the from list needs these
		// removed.
		//---------------------------------------------------------------------
		for (i = from.length-1; i >= 0; i--) 
		{   
			if (from[i].selected)
			{
				from[i].selected=false;
				from[i] = null;
				l=i;	//first option in list 
			}
		}
		if ( from.length > 0 )
		{
			if (l >= from.length)
				l=from.length-1;
				
			from[0].selected=true;
		}
	}
	else if (clearFlag)
	{
		// No message in this situation
	}
	else if (deSelectFlag)
	{
		alert("Please high-light a name first from the Selected list.");
	}
	else
	{
		alert("Please high-light a name first from the Not Selected list.");
	}
}

</SCRIPT>
</body>
</html>
