<%
response.buffer = True

'-------------------------------------------------------- 
' Test to see if user has necessary privileges for page.
'-------------------------------------------------------- 
If hasRole("Sales") Then
	i=0
Else
	Response.Redirect "Not_Privileged.asp"
End If
	
Const ACTION_SEARCH = "SEARCH"	'Set in hidden form field to indicate search is requested
Const DelimiterString = " "		'Delimiter for parsing the keyword string
%>
<!--
Developer:    Dave Podnar
Date:         9/5/2000
Description: Main script for searching the consultant skills inventory information.

Scripts: Search_Skills.asp = Search skills inventory for consultants
-->

<!-- #include file="../section.asp" -->

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=Center VALIGN=center>
<font size=3 face="ms sans serif, arial, geneva" color=black>
<b>Skills Inventory Search</b>
</font></td></tr>
<p>
<tr><td ALIGN=LEFT VALIGN=center>
<font size=2 face="ms sans serif, arial, geneva" color=black>
Use this page to search the SARK consultant skills inventory. Select the skills and minimum skill levels 
you are searching for.  You can select multiple skills by holding down the Shift or Ctrl keys while selecting
available skills.  You can change a skill level for a selected skill by highlighting selected skill(s) and then selecting a
new skill level.
</font></td></tr>
</TABLE>

<%
'For Each xyz in Session.Contents
'Response.Write Session.Contents(xyz) & "<br>"
'Next

'For Each abc in Session.StaticObjects
'Response.Write Session.StaticObjects(abc) & "<br>"
'Next

reqFrmAction=Request.Form("ACTION")

If reqFrmAction & "x" = ACTION_SEARCH & "x" Then 'Or Not IsNull(session("sessFrmSkillLevel")) Then
	' The user has submitted a search request or previously submitted
	  
	'Response.Write "Search was requested with the following paramters:<br>"
	'Response.write "count=" & Request("selected").count & "<BR>"
	
	If reqFrmAction & "x" = ACTION_SEARCH & "x" Then
		'--------------------------------------------
		' Assign Form variables to local variables.
		' Radio objects are treated like strings
		'--------------------------------------------
		set reqFrmSelected=Request.Form("selected") 'Select List
		reqFrmSkillLevel=Request.Form("SKILL_LEVEL") 'Radio
		reqFrmSearchString=Request.Form("SEARCH_STRING") ' Text
		reqFrmSearchMode=Request.Form("SEARCH_MODE") ' Radio
		reqFrmAndOr=Request.Form("AND_OR") ' Radio

		' Persist values for session
		'Set session("sessFrmSelected") = reqFrmSelected
		session("sessFrmSkillLevel") = reqFrmSkillLevel
		session("sessFrmSearchString") = reqFrmSearchString
		session("sessFrmSearchMode") = reqFrmSearchMode
		session("sessFrmAndOr") = reqFrmAndOr
	
	Else
		'Assign local variables from last session search
		If Not IsNull(session("sessFrmSelected")) Then
			Set reqFrmSelected = session("sessFrmSelected")
		End If
		If Not IsNull(session("sessFrmSkillLevel")) Then
			Set reqFrmSkillLevel = session("sessFrmSkillLevel")
		End If
		If Not IsNull(session("sessFrmSearchString")) Then
			reqFrmSearchString = session("sessFrmSearchString")
		End If
		If Not IsNull(session("sessFrmSearchMode")) Then
			Set reqFrmSearchMode = session("sessFrmSearchMode")
		End If
		If Not IsNull(session("sessFrmAndOr")) Then
			Set reqFrmAndOr=session("sessFrmAndOr")
		End If
	End If
	
	'----------------------------------------------------------------------------
	' These variables assist construction of the proper select statement
	'----------------------------------------------------------------------------
	If InStr(1,reqFrmAndOr,"OR") = 0 Then
		allSkills=True 
	Else 
		allSkills=False 
	End If
	selectedSkills=False
	keywordSkills=False
	
	'----------------------------------------------------------------------------
	' These variables are used to construct various parts of the select statement
	'----------------------------------------------------------------------------
	inClause=""
	whereClause=""
	selectText=""
	keyWordString=""
	
	For Each skill in reqFrmSelected
		'---------------------------------------------------------------------
		' The items in the SELECT list have values in the form of "x:y" where
		' x is the skill level and y is the skill id. These are parsed and
		' assigned to local variables skillLevel and skillID
		'---------------------------------------------------------------------
		If skill & "x" <> "x" Then
			selectedSkills=True
			i=InStr(1, skill,":")
			skillLevel=clean(trim(mid(skill,1,i-1)))
			skillID=clean(trim(mid(skill,i+1)))
			inClause = inClause & skillID & ","
			'-------------------------------------------------------------
			' The AND_OR Radio field specifically has a value of "AND" or 
			' "OR" and can be used as is 
			'-------------------------------------------------------------
			tmpString = whereClause & reqFrmAndOr & " B.R" & skillID & " >=" & skillLevel & " "
			whereClause = tmpString
			
			tmpString = selectText & "SUM(CASE skillid WHEN " & skillID & " THEN skillranking ELSE 0 END) AS R" & skillID & ", "
			selectText = tmpString
		End If
	Next
	
	If reqFrmSearchString & "x" <> "x" Then
		'---------------------------------------------------------------------------
		' A search string was provided. We will remove all single and double quotes
		' to simplify the creation of the WHERE clause.
		'---------------------------------------------------------------------------
		searchString=Replace(reqFrmSearchString,"'","")		'Remove single quote
		searchString=Replace(searchString,CHR(34),"")		'Remove double quote
		searchString= trim(searchString) & DelimiterString	'Remove leading spaces and append space to assist parsing

		keywordSkills=True
		
		If reqFrmSearchMode = "1" Then
			'---------------------------------------------------------------------
			' Treat all words separately, i.e. "word1 word2 word3" becomes
			' skillvalue LIKE '%word1%word2%word3%'
			'
			' Parse words using space as separator. 
			' i = the current pointer in the string
			' j = pointer to next space in the string
			'---------------------------------------------------------------------
			i=1
			j = InStr(i, searchString, DelimiterString)

			While j <> 0
				If j-i >= len(DelimiterString) Then
					keyWord = trim(mid(searchString,i,j-i))
					If i=1 Then
						tmpString = " SkillValue LIKE '%" & keyWord & "%"
					Else
						tmpString = keyWordString & keyWord & "%"
					End If
					keyWordString = tmpString
				End If
				
				i=j+1
				j = InStr(i, searchString, DelimiterString)
			Wend

			If Len(keyWordString) > 1 Then
				' End quote for literal
				tmpString = keyWordString & "'"
				keyWordString = tmpString
			Else
				' All spaces were found in the search string, so treat as nothing entered
				keywordSkills = False
			End If
        
		Else
			'----------------------------
			'Treat words as a phrase
			'----------------------------
			keyWordString = " SkillValue LIKE '%" & trim(searchString) & "%' "
		End If
	End If
	
	'------------------------------------------------------------------------------------------
	' Here is information regarding the SELECT statements. The Skills_inventory table contains
	' the necessary information regarding all consultants skills. To properly account for
	' all skills/levels selected by a user, i.e. skilla >= 3 and skillb >=4 and ... skillxyz >=3
	' each consultant's skills_inventory row's need transposed into columns. Assume the following 
	' 5 rows were the only rows in the skills_inventory table:
	'		employeeid skillid skillranking
	'		1				10	    3
	'		1				11		4
	'		1				12		5
	'		2				10		6
	'		2				13		5
	'
	' We would transponse those into two rows, where the skillid's are now columns and the rankings
	' are the column values like this:
	'		employeeid	R10	R11 R12 R13
	'		 1			3   4   5   0
	'        2          6   0   0   5
	'
	' There are 4 possible SELECT statements based on the user's input:
	' #1 User selected skills and NO key words
	' #2 User selected skills and key words
	' #3 User selected NO skills and key words
	' #4 User selected NO skills and NO key words
	'
	' The selected skills part of the SQL is considered table B.
	' The key words part of the SQL is considered table C.
	' select employeeid from employee a, 
	' select employeeid from skills inventory 
	
	'Response.Write "selectText=<b>" & selectText & "<b><br>"
	'Response.Write "whereClause=<b>" & whereClause & "<b><br>"
	'Response.Write "inClause=<b>" & inClause & "<b><br>"
	'Response.Write "keyWordString=<b>" & keyWordString & "<b><br>"

	sql = "SELECT a.EmployeeID, a.firstname, a.lastname FROM employee a, "
	
	If selectedSkills Then
		sql = sql & "(SELECT employeeid, " & mid(selectText,1,len(selectText)-2) & " FROM skills_inventory WHERE skillid IN (" & mid(inClause,1,len(inClause)-1) & ") GROUP BY employeeid) b"
	End If
	
	If allSkills Then
		' Path taken when the user has requested the consultant to have ALL skills
		If keywordSkills Then
			If selectedSkills Then
				sql = sql & ","
			End If

			sql = sql & " (SELECT DISTINCT employeeid FROM skills_inventory x, skills y WHERE x.skillid=y.skillid AND y.isOtherSkill=1 AND y.skillname != x.skillvalue AND (" & keyWordString & ")) c "
		End If
		
		sql = sql & " WHERE "
		If selectedSkills Then
			sql = sql & "(" & mid(whereClause,Len(reqFrmAndOr)+1) & ") AND a.employeeid=b.employeeid "

			If keywordSkills Then
				sql = sql & " AND "
			End If
		End If

		If keywordSkills Then
			sql = sql & "a.employeeid=c.employeeid "
		End If

	Else
		' Path taken when the user has requested the consultant to have ANY skills
		If selectedSkills Then
			sql = sql & " WHERE (" & mid(whereClause,Len(reqFrmAndOr)+1) & ") AND a.employeeid=b.employeeid "

			If keywordSkills Then
				sql = sql & " UNION SELECT a.EmployeeID, a.firstname, a.lastname FROM employee a, "
			End If
		End If

		If keywordSkills Then
			sql = sql & " skills_inventory x, skills y  WHERE x.skillid=y.skillid AND y.isOtherSkill=1 AND y.skillname != x.skillvalue AND (" & keyWordString & ") AND x.employeeid=a.employeeid " 
		End If
		
	End If

	sql = sql & " ORDER BY a.lastname, a.firstname, a.employeeID"
	
	'Response.Write "<b>" & sql & "<b><br>"
	
	If selectedSkills Or keywordSkills Then
		set rs = DBQuery(sql)
	Else
		'------------------------------------------------------------
		' User has not specified a keyword or selected any skills.
		' Need to exec a dummy query so the rs object is created and 
		' the following If doesn't give a runtime error
		'------------------------------------------------------------
		set rs = DBQuery("select 1 from employee where employeeid=0 and employeeid<>0")
	End If
	
	If (selectedSkills Or keywordSkills) And Not rs.eof Then%>
		<br><font size=3 face='ms sans serif, arial, geneva' color=black>
		The following consultants match the search criteria:<br></font>
		<TABLE BORDER=0 WIDTH=100% >
		<%While Not rs.eof%>
			<tr><td align=left><A HREF='details.asp?EmpId=<%=rs("EmployeeID")%>'> <%=rs("lastname") & ", " & rs("firstname")%></a></td></tr>
			<%rs.movenext
		wend
		rs.close
		set rs=Nothing
		Response.Write "</TABLE>"
	Else%>
		<font size=2 face="ms sans serif, arial, geneva"  color=blue>
		<b>No consultants match the search criteria!</b><br>
	<%End If

End If

%>

<FORM NAME="searchForm" onSubmit="doSearch()" ACTION="<%=request.servervariables("script_name")%>" METHOD="POST">
<table class=tableShadow width=100% cellspacing=0 cellpadding=2 border=0 bgcolor=#ffffcc>
<th width='45%'></th>
<th width='10%'></th>
<th width='45%'></th>
<tr>
<td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Available<br>Skill</b></td>
<td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black>&nbsp;</td>
<td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Selected<br>(Skill Level) and Skill</b></td>
</tr>
<TR><TD colspan=3><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr>
 <td><select MULTIPLE size=10 name=NotSelected id=NotSelected style="{ width: 100%; }" onfocus="setSelectList(document.searchForm.Selected,false)">

	<%	'-------------------------------------------------------------------------------------------------
	' There will be a drop down of skill names. Currently, there are skills where the skillname is not unique 
	' in the skills table. In these cases we display the "skillname [groupname]". It is important to 
	' display all these, not just the distinct names, it allows the back end searching to be simpler and 
	' gives the user finer control of the search critieria.	'-------------------------------------------------------------------------------------------------
	sql = "SELECT c.counter, RTRIM(CAST(b.groupname AS VARCHAR(60))) AS GroupName, RTRIM(CAST(a.skillname AS VARCHAR(60))) AS SkillName, a.skillid " _ 
		& "FROM skills a, skills_groups b, (SELECT skillname, count(*) AS COUNTER FROM skills WHERE isOtherSkill=0 GROUP BY skillname) c " _
		& "WHERE a.groupid=b.groupid AND a.skillname=c.skillname ORDER BY a.skillname, a.skillid"
		
	set rsSkill = DBQuery(sql)
	Dim TempArrayID(300)
	Dim TempArrayName(300)
	Dim TempSelected(300)
	
	i=0
	While Not rsSkill.eof
		TempArrayID(i)=rsSkill("SkillID")
		If rsSkill("counter") > 1 Then
			TempArrayName(i)=rsSkill("SkillName") & " [" & rsSkill("GroupName") & "]"
	    Else
			TempArrayName(i)=rsSkill("SkillName")
	    End If
		isDesiredSkill=InStr(1, "," & inClause, "," & rsSkill("SkillID")& ",")
	    If isDesiredSkill <= 0 Then
			TempSelected(i)=false
			' User selected this skill and will NOT show in this select list
		%>
		<option value=<%=TempArrayID(i)%>><%=TempArrayName(i)%></option>
		<%
		Else
			'---------------------------------------------------------------------
			' The items in the SELECT list have values in the form of "x:y" where
			' x is the public skill level and y is the skill id. These are parsed and
			' assigned to local variables skillLevel and skillID. Additionally,
			' they have text in the form of "(z) abcd" where z is the private skill level.
			' The public skill level is displayed to the user but the private skill level
			' is used internally for queries.
			'---------------------------------------------------------------------
			Response.Write "'" & TempArrayID(i) & "'<br>"
			For Each skill in reqFrmSelected
				Response.write "skill=" & skill & "<BR>"
				If skill & "x" <> "x" Then
					pos=InStr(1, skill,":")
					skillID=mid(skill,pos+1)
					Response.Write "'" & skillID & "'<br>"
					If skillID&"x" = TempArrayID(i)&"x" Then
						'------------------------------------------------------------
						' The user has previously requested this skill as a criteria
						' therefore we use the previously received x:y value
						'------------------------------------------------------------
						Response.Write "'" & skill & "'<br>"
						TempArrayID(i)=skill
					End If
				End If
			Next
			
			TempSelected(i)=true
		End If
		rsSkill.movenext
		i=i+1
	Wend
	rsSkill.close
	set rsSkill=Nothing
	%>
	    </select></td>
	    <td align=center>
	    <table border=0>
	    <tr><td><input type=button value=" --> " id=button2 name=button2 onClick="moveOption(document.searchForm.NotSelected,document.searchForm.Selected,false,false);"></td></tr>	    <tr><td><input type=button value=" <-- " id=button3 name=button3 onClick="moveOption(document.searchForm.Selected,document.searchForm.NotSelected,true,false);"></td></tr></table></td>

	    <td ><select MULTIPLE size=10 name=Selected id=Selected style="{ width: 100%; }" onfocus="setSelectList(document.searchForm.NotSelected,false)">
	    <%
	    i=0
		While i < 300 And TempArrayID(i) & "x" <> "x"

			If TempSelected(i)=true Then%>
					<option value=<%=TempArrayID(i)%>><%=TempArrayName(i)%></option>
			<%
			End If
			i=i+1
		Wend%>
	    </select></td>
	    </tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr><td colspan=3>
	<table width=100% >
	<tr><td width=43%>&nbsp;</td><td align=left width=57%><b>Skill Level</b></td></tr><%	'------------------------------------------------------------------------------
	' Identify the valid Skill Rankings. A skill ranking of 1 indicates a consultant
	' has no knowledge of a skill and is therefore absent from the displayed rankings.
	'------------------------------------------------------------------------------
	sql = "SELECT DISTINCT SkillRanking, RankingDescription FROM skills_rankings WHERE SkillRanking > 1" _
	     & " ORDER BY SkillRanking"
	  
	set rsLevel = DBQuery(sql)
	j=1

	While Not rsLevel.eof
		'--------------------------------------------------------------------------------
		' The values of the following Radio/Toggle are in the format of x:y where
		' x = The displayed (public) value, this always starts at 1 and increments
		' y = The database column (private) value from the selected row, used when server eventually 
		'     executes the query
		'--------------------------------------------------------------------------------
	%>	
			<tr><td>&nbsp;</td><td align=left valign=center>
			<INPUT TYPE="RADIO" NAME="SKILL_LEVEL" <%If reqFrmSkillLevel & "x" = "x" And j=1 Then%> CHECKED 
			                                       <%ElseIf reqFrmSkillLevel = j & ":" & rsLevel("SkillRanking") Then%> CHECKED 
			<%End If%>
			
			onclick="setSkillValue(document.searchForm.Selected);" VALUE=<%=j & ":" & rsLevel("SkillRanking")%>>(<%=j%>)&nbsp;=&nbsp;<%=rsLevel("RankingDescription")%>
			</td></tr>
			
		<%rsLevel.movenext
		j=j+1
	Wend
	rsLevel.close
	set rsLevel=Nothing
	%>
	</table></td></tr><tr>
<td colspan=3 bgcolor=gray height=1></td>
</tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr>
<td colspan=3 align=LEFT><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Other skills</b> keyword search&nbsp;<INPUT TYPE="TEXT" NAME="SEARCH_STRING" MAXLENGTH="100" 
VALUE="<%If reqFrmSearchString & "x" <> "x" Then
		Response.write Replace(reqFrmSearchString,CHR(34),"")	'Remove double quote
		End If%>" SIZE="30">
</td>
</tr><tr>
<td colspan=3 align=CENTER><INPUT TYPE="RADIO" NAME="SEARCH_MODE" VALUE="1" <%If reqFrmSearchMode <> "0" Then%>CHECKED
		                                         <%End If%>                                >Separate Words
<INPUT TYPE="RADIO" NAME="SEARCH_MODE" VALUE="0" <%If reqFrmSearchMode = "0" Then%>CHECKED
		                                         <%End If%>                                >Phrase
</td>
</tr><TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr>
<td colspan=3 bgcolor=gray height=1></td>
</tr>
<TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr>
<td colspan=3 align=center>		
		

<INPUT TYPE="RADIO" NAME="AND_OR" VALUE="AND" <%If reqFrmAndOr <> "OR" Then%>CHECKED
		                                      <%End If%>                            >Consultant MUST have all skills
</td>
</tr><tr>
<td colspan=3 align=center><INPUT TYPE="RADIO" NAME="AND_OR" VALUE="OR" <%If reqFrmAndOr = "OR" Then%>CHECKED
                                             <%End If%>                            >Consultant has at least one skill
</td>
</tr><TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
<tr>
<td colspan=3 align=center><input type=button class=button value="Search" onClick='javascript:doSearch()'; id=button1 name=button1>
&nbsp;&nbsp;&nbsp;
<input type=button class=button value="Clear" onClick='javascript:doClear()'; id=button2 name=button2>
&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr><TR><TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
</table>

<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
</FORM><P>

<SCRIPT LANGUAGE="javascript">
/*
 * The values are correct but the text is NOT. Need to retrieve the public values
 * for the skill level and prepend those to the text string.
 */
	var oSelect=document.searchForm.Selected;
	var oRadio=document.searchForm.SKILL_LEVEL;
	var flag;
	var i,j,x;
	//alert("listlength="+oSelect.length);
    for(i=0; i < oSelect.length; i++)
    {
		// Extract private skill level value
		oValue = oSelect[i].value.substring(0,oSelect[i].value.indexOf(":")); 
		//alert("private="+oValue);
		flag=false;
		
		for(j=0; j < oRadio.length && !flag; j++)
		{
			x=oRadio[j].value.indexOf(":");
			pubVal=oRadio[j].value.substring(0,x);
			privVal=oRadio[j].value.substring(x+1,oRadio[j].value.length);
			if (privVal == oValue)
			{
				flag=true;
				oText=oSelect[i].text;
				oSelect[i].text= "(" + pubVal + ") " + oText;
			}
		}
	}
</SCRIPT>

<!-- #include file="../../footer.asp" -->


<SCRIPT LANGUAGE="javascript">

function doClear()
{
	frm=document.searchForm;
	frm.reset();	
	frm.SEARCH_STRING.value = ""

	//------------------------------------------------------------
	// The select lists don't automatically get reset so 
	// these calls move all options back to the NotSelected list.
	//------------------------------------------------------------
	setSelectList(frm.Selected,true);
	moveOption(frm.Selected,frm.NotSelected,true,true);
	
	//------------------------------------------------------------
	// Set all radio buttons to default settion (first value is on)
	//------------------------------------------------------------
	frm.SKILL_LEVEL[0].checked=true;
	frm.SEARCH_MODE[0].checked=true;
	frm.AND_OR[0].checked=true;
}

function doSearch()
{
	frm=document.searchForm;
	frm.ACTION.value = "<%=ACTION_SEARCH%>";
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
	}
}

/* 
 * All options which are selected have their skill value changed
 * to the current radio setting.
 */
function setSkillValue(oSelect)
{
	var i = 1;
    for(i=0; i < oSelect.length; i++)
    {
		if (oSelect[i].selected)
		{
			// Extract persistent values
			oText = oSelect[i].text.substring(4,oSelect[i].text.length); 
			oValue = oSelect[i].value.substring(oSelect[i].value.indexOf(":")+1,oSelect[i].value.length); 
			
			oText = "(" + getPublicRadioValue(document.searchForm.SKILL_LEVEL) + ") " + oText;			
			oValue = getPrivateRadioValue(document.searchForm.SKILL_LEVEL) + ":" +oValue;

			oSelect[i].text=oText;
			oSelect[i].value=oValue;
		}
	}
}


function getRadioValue(oRadio)
{
	var i = 1, flag=true, val;

    for(i=0; i < oRadio.length && flag; i++)
    {
		if ( oRadio[i].checked )
		{
			val=oRadio[i].value;
			flag=false
		}
	}
	return val;
}
       
//--------------------------------------------------------------------------------
// See server code for creation of these objects. The value is in a format of x:y.
// We extract the y portion of the value.
//--------------------------------------------------------------------------------
function getPrivateRadioValue(oRadio)
{
	var i = 1, flag=true, val;

    for(i=0; i < oRadio.length && flag; i++)
    {
		if ( oRadio[i].checked )
		{
			j=oRadio[i].value.indexOf(":");
			val=oRadio[i].value.substring(j+1,oRadio[i].value.length);
			flag=false
		}
	}
	return val;
}
       
//--------------------------------------------------------------------------------
// See server code for creation of these objects. The value is in a format of x:y.
// We extract the x portion of the value.
//--------------------------------------------------------------------------------
function getPublicRadioValue(oRadio)
{
	var i = 1, flag=true, val;

    for(i=0; i < oRadio.length && flag; i++)
    {
		if ( oRadio[i].checked )
		{
			j=oRadio[i].value.indexOf(":");
			val=oRadio[i].value.substring(0,j);
			flag=false
		}
	}
	return val;
}
       
/*
 * Function is used to move selected options between the "NotSelected" and "Selected" lists.
 *
 * The boolean argument "deSelectFlag" indicates whether the movement is from "Selected" to
 * "NotSelected". This is required since the option text and values in the "Selected" list are 
 * slightly different than the "NotSelected" list option values.
 * 
 * The boolean argument "resetFlag" indicates when a form clear is occuring
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
			if (deSelectFlag)
			{
				//------------------------------------------------------
				// The text values are in the form of "(1) Textxyz..."
				// When the skills are moved from the select list to the 
				// avail list we don't move the "(1) ". Likewise, the
				// value field is in the form of "x:y" where x is the 
				// skill level and y is the skill ID. We remove the skill
				// level "x:".
				//------------------------------------------------------
				oText = from[i].text.substring(4,from[i].text.length); 
				oValue = from[i].value.substring(from[i].value.indexOf(":")+1,from[i].value.length); 
			}
			else
			{
				oText = "(" + getPublicRadioValue(document.searchForm.SKILL_LEVEL) + ") " + from[i].text; 
				oValue = getPrivateRadioValue(document.searchForm.SKILL_LEVEL) + ":" +from[i].value;
			}
			// Creates a new option in the destination list
			to[j] = new Option (oText, oValue, false, false);
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
		alert("Please high light a skill first from the Selected list.");
	}
	else
	{
		alert("Please high light a skill first from the Available list.");
	}
}

</SCRIPT>
