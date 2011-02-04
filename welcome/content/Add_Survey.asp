<%
'----------------------------------------------------------------
' There is a situation where this popup window is redisplayed and
' it occurs, when the user cancels or submits the survey and then 
' attempts a refresh. There is code labeled "=====REDISPLAY Prevention====="
' which automatically closes the window when it is opened. 
'
' The following header stuff is supposed to prevent the page from 
' being cached. I haven't found success with these, but I kept it in, in the
' event it starts working with newer browser versions.
'----------------------------------------------------------------
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>

<!--
Developer:    Dave Podnar
Date:         8/14/2000
Description: Script provides the online survey to a user. When it is filled in correctly
             the responses are inserted into the ONLINE_SURVEY table.
-->

<HTML>
<HEAD><TITLE>Intranet Survey
</TITLE>
<!-- #include file="../../style.htm" -->
</HEAD>
<BODY BGCOLOR=SILVER >

<!-- #include file="../../script.asp" -->

<%

'------------------------------------------------------------------------------
' Variables used to hold responses to questions 1 and 2. Question 1 can have
' have 5 values checked.
'------------------------------------------------------------------------------
q1_1=0
q1_2=0
q1_3=0
q1_4=0
q1_5=0
q2=0

If Request("ACTION") = "SUBMIT" Then
	Session("Surveyed")=True
	
	If Request.Form("USAGE_FREQ").Count Then
		'----------------------------------------------------------------------------
		' Put the check box value into separate integer variables
		'----------------------------------------------------------------------------
	   x = 1
       While(x < Request.Form("USAGE_FREQ").Count + 1)
       
            'Response.Write "Usage [" & x & "]=" & Request.Form("USAGE_FREQ")(x) & "<BR>"
            If Request.Form("USAGE_FREQ")(x) = "1" Then
				q1_1=1
			ElseIf Request.Form("USAGE_FREQ")(x) = "2" Then
				q1_2=1
			ElseIf Request.Form("USAGE_FREQ")(x) = "3" Then
				q1_3=1
			ElseIf Request.Form("USAGE_FREQ")(x) = "4" Then
				q1_4=1
			ElseIf Request.Form("USAGE_FREQ")(x) = "5" Then
				q1_5=1
            End If
            x = x + 1
       Wend
	End If
	
	If Request.Form("USEFUL").Count Then
		'---------------------------------------------
		' Put the radio value into an integer variable
		'---------------------------------------------
	   x = 1
       While(x < Request.Form("USEFUL").Count + 1)
       
            'Response.Write "Useful [" & x & "]=" & Request.Form("USEFUL")(x) & "<BR>"
            If Request.Form("USEFUL")(x) = "1" Then
				q2=1
			ElseIf Request.Form("USEFUL")(x) = "2" Then
				q2=2
			ElseIf Request.Form("USEFUL")(x) = "3" Then
				q2=3
			ElseIf Request.Form("USEFUL")(x) = "4" Then
				q2=4
			ElseIf Request.Form("USEFUL")(x) = "5" Then
				q2=5
            End If
            x = x + 1
       Wend
	End If
	
	sql = "INSERT INTO online_survey " _
	& " (employeeId, q1_1, q1_2, q1_3, q1_4, q1_5, q2, q3, q4) VALUES (" _
	& Session("ID") & "," & q1_1 & "," & q1_2 & "," & q1_3 & "," & q1_4 & "," & q1_5 & "," & q2 & ","
	
	If Request.Form("NEVER_USE") & "x" <> "x" then 
		sql = sql & "'" & clean(Request.Form("NEVER_USE")) & "',"
	Else
		sql = sql & "null,"
	End If
	
	If Request.Form("WANT_TO_USE") & "x" <> "x" then 
		sql = sql & "'" & clean(Request.Form("WANT_TO_USE")) & "')"
	Else
		sql = sql & "null)"
	End If
	
	'Response.Write "sql=" & sql
	set rs = DBQuery(sql)
	set rs = Nothing
	%>
	<br>
	<br>
	<br>
	<br>
	<font color=blue size=4 face="ms sans serif, arial, geneva">
	Your input has been saved, thanks for taking time to fill out the survey!<br>
	<br>
	<br>
	Click the "Close Window" button to end the survey.
	</font>
	<br>
	<br>
	<br>
	<form>
	<input type=button class=button value="Close Window" onClick='javascript:window.close();'; id=submit1 name=submit1>
	</form>
	</body></html>

<%ElseIf Session("Surveyed") Then
'------------------------------------------------------------------------
'  =====REDISPLAY Prevention=====
' I don't know why this happens, but refreshing the welcome page, causes
' this pop up window to be redisplayed. It is a hack, but closing it here
' makes sure the user doesn't see the survey again. At the top of this
' script, you'll notice the page expiration information, but that doesn't
' seem to help either. 
'------------------------------------------------------------------------
%>
<SCRIPT LANGUAGE="javascript">
window.close();
</script>
</body>


<%
Else
	Session("Surveyed")=True
%>
<H2>Intranet Survey</H2>
<font color=black size=2 face="ms sans serif, arial, geneva">
Please take a minute and fill out the Cincinnati Sark Intranet survey. <b>If you don't have time
right now, just click the "Cancel" button and we will try and get your input during your next login.</b>
<form action="http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")%>" method=post id=surveyForm name=surveyForm >
<B>1)</font>&nbsp;How often do you use the intranet?</B>&nbsp;&nbsp;[&nbsp;<font color=red>Required</font>&nbsp;]<BR>
&nbsp;&nbsp;&nbsp;&nbsp;(check all that apply)<br>
&nbsp;&nbsp;&nbsp;<INPUT NAME="USAGE_FREQ" TYPE="CHECKBOX" VALUE="1">Once a day<BR>
&nbsp;&nbsp;&nbsp;<INPUT NAME="USAGE_FREQ" TYPE="CHECKBOX" VALUE="2">More than once a day<BR>
&nbsp;&nbsp;&nbsp;<INPUT NAME="USAGE_FREQ" TYPE="CHECKBOX" VALUE="3">Once a week<BR>
&nbsp;&nbsp;&nbsp;<INPUT NAME="USAGE_FREQ" TYPE="CHECKBOX" VALUE="4">Almost every day<BR>
&nbsp;&nbsp;&nbsp;<INPUT NAME="USAGE_FREQ" TYPE="CHECKBOX" VALUE="5">On the weekend<BR>
<BR>
<B>2)</font>&nbsp;How useful do you find the intranet?</B>&nbsp;&nbsp;[&nbsp;<font color=red>Required</font>&nbsp;]<BR>
&nbsp;&nbsp;&nbsp;&nbsp;(1 is most useful and 5 is least useful)<br>
&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" NAME="USEFUL" VALUE="1">1
&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" NAME="USEFUL" VALUE="2">2
&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" NAME="USEFUL" VALUE="3">3
&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" NAME="USEFUL" VALUE="4">4
&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" NAME="USEFUL" VALUE="5">5
<BR><BR>
<B>3)</font>&nbsp;Enter one item that you never use:</B>&nbsp;&nbsp;[&nbsp;<font color=blue>Optional</font>&nbsp;]<BR>
&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="TEXT" NAME="NEVER_USE" MAXLENGTH="250" SIZE="60"><BR><BR>
<B>4)</font>&nbsp;Enter one item that you would like to see:</B>&nbsp;&nbsp;[&nbsp;<font color=blue>Optional</font>&nbsp;]<BR>
&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="TEXT" NAME="WANT_TO_USE" MAXLENGTH="250" SIZE="60"><BR><BR>
<IMG SRC="/images/common/spacer.gif" HEIGHT="3">
<input type=button class=button value="Submit Survey" onClick='javascript:validateSurvey()'; id=submit1 name=submit1>
<input type=button class=button value="Cancel" onClick='javascript:closeWindow()'; id=button1 name=button1>
<input type=button class=button value="Clear Survey" onClick='javascript:clearSurvey()'; id=button2 name=button2>
<INPUT TYPE=HIDDEN NAME=ACTION VALUE="">
</FORM>
</font>
</body>

<SCRIPT LANGUAGE="javascript">
function closeWindow()
{
  window.close();
}

function clearSurvey() 
{
	frm=document.surveyForm;
	frm.reset();
}

function isFieldValid(theField) 
{
	var valid=false;
	
	for( i=0 ; i<theField.length && valid==false; i++) 
	{
		if (theField[i].checked )
		{
			valid=true;
		}
	}
	return valid;
}

function validateSurvey() 
{
	frm=document.surveyForm;
	len = frm.elements.length;
	var radioValid=false, checkValid=false, usageList="";
	
	checkValid=isFieldValid(frm.USAGE_FREQ);
	radioValid=isFieldValid(frm.USEFUL);

    if (checkValid == false)
	{
		alert('Please answer question #1.');
	}
    else if (radioValid == false)
	{
		alert('Please answer question #2.');
	}
	else
	{
		frm.ACTION.value = 'SUBMIT';
		frm.submit();
	}
}
</SCRIPT>
<%End If%>
</html>
