<!--
Developers:   Jason Moore and Mike Thierauf
Date:         07/17/2000
Description:  Allows editing of review information
-->

<!-- #include file="../../script.asp" -->

<html>
<head>
<title>Update Employee Review Information</title>

<!-- #include file="../../style.htm" -->
<SCRIPT>
<!--
function checkDate(inputField)
{  //check for valid dates
	var errorFlag = 0;
	
	var inputDate = inputField.value;
	var in_month, in_day, in_year;
	
	// get index of first slash in date
	var firstSlashIndex = inputDate.indexOf("/");
	// pull out month from date
	in_month = inputDate.substring(0, firstSlashIndex);
	// get index of second slash in date
	var secondSlashIndex = inputDate.indexOf("/", firstSlashIndex+1);
	// pull out day from date
	in_day = inputDate.substring(firstSlashIndex+1, secondSlashIndex);
	// pull out year from date
	in_year = inputDate.substring(secondSlashIndex+1, inputDate.length);
	
	// error checking
	if (in_month > 12 || in_month < 1)
		errorFlag = 1;
	if (in_year < 0)
		errorFlag = 1;
	
	// check 30-day months
	if (in_month == 4 || in_month == 6 || in_month == 9 || in_month == 11)
	{
		if (in_day > 30)
			errorFlag = 1;
	}
	// check for February
	else if (in_month == 2)
	{
		if (parseInt(in_year, 10)/4 != 0)
		{
			if (in_day > 28)
				errorFlag = 1;
		}
	}
	// check for 31-day months
	else
	{
		if (in_day > 31)
			errorFlag = 1;
	}
	
	// check to make sure year is 2 or 4 digits
	if (in_year.length != 2 && in_year.length != 4)
	{	
		errorFlag = 1;
	}
	
	return errorFlag;
}

function addYear()
{
	var output, input;
	var in_day, in_month, in_year, out_day, out_month, out_year;
  
	if(checkDate(document.frmInfo.last_review) == 0)
	{
		input = document.frmInfo.last_review.value;
	  
		// get index of first slash in date
		var firstSlashIndex = input.indexOf("/");
		// pull out month from date
		in_month = input.substring(0, firstSlashIndex);
		// get index of second slash in date
		var secondSlashIndex = input.indexOf("/", firstSlashIndex+1);
		// pull out day from date
		in_day = input.substring(firstSlashIndex+1, secondSlashIndex);
		// pull out year from date
		in_year = input.substring(secondSlashIndex+1, input.length);
  
		out_year = (parseInt(in_year,10) + 1);
    
		if (out_year.toString().length == 3) 
		{
			out_year = out_year.toString().substring(1,3);
		}
		if (out_year.toString().length == 1) 
		{
			out_year = "0" + out_year;
		}
  
		//check for February 29
		if (parseInt(in_month) == 2 && parseInt(in_day) == 29) 
		{
			out_month = "03";
			out_day = "01";
		}
		else
		{
			out_month = in_month;
			out_day = in_day;
		}
   
		output = ""; 
		output += out_month + "/";
		output += out_day + "/";
		output += out_year;
		document.frmInfo.next_review.value = output;
	}
}

function submitData()
{
	var lpErrorFlag = 0;
	var lrErrorFlag = 0;
	
if(document.frmInfo.last_promotion.value != "9/9/9999")	
{
	if((document.frmInfo.last_promotion.value != "") && (checkDate(document.frmInfo.last_promotion) == 1))
	{
		document.frmInfo.last_promotion.replaceAdjacentText("afterEnd", " (Invalid Date)");
		lpErrorFlag = 1;
	}
	else
	{
		document.frmInfo.last_promotion.replaceAdjacentText("afterEnd", " ");
		lpErrorFlag = 0;
	}
}
	
	if (checkDate(document.frmInfo.last_review) == 1)
	{
		document.frmInfo.last_review.replaceAdjacentText("afterEnd", " (Invalid Date)");
		lrErrorFlag = 1;
	}
	else
	{
		document.frmInfo.last_review.replaceAdjacentText("afterEnd", " ");
		lrErrorFlag = 0;
	}
	
	if (lrErrorFlag == 0)
	{
		// if last review date is valid, update the next review date
		addYear( );
	}
	
	if((document.frmInfo.last_promotion.value == "") || (document.frmInfo.last_promotion.value == "1/1/1900") ||
		(document.frmInfo.last_promotion.value == "9/9/9999"))
	{
		document.frmInfo.last_promotion.value = "9/9/9999";
	}	
	
	if (lrErrorFlag == 0 && lpErrorFlag == 0)
	{
		// if all dates are valid, submit the review data
		document.frmInfo.submit( );
	}
}

function closeWindow()
{	
	window.close();
}

//-->
</SCRIPT>
</head>
<%
'-----------------------------------------------
'	Execute Database Query for Profile Info		
'-----------------------------------------------
empID = Request.QueryString("EmpID") 
set rs = DBQuery("select * from review where employeeID = " & empID)
if not rs.eof then
'-----------------------------------------------
'	Get Review Info							
'-----------------------------------------------
	ls_last_promotion =	trim(rs("last_promotion"))
	ls_last_review = 	trim(rs("last_review"))
	ls_next_review = 	trim(rs("next_review"))
	ls_sponsor_id =		trim(rs("sponsor_id"))
	ls_acct_mgr_id = 	trim(rs("account_mgr_id"))
	ls_comments =		trim(rs("comments"))
	 
	ls_insertupdate = "U"
	else
	ls_insertupdate = "I"
end if 

if ls_sponsor_id = 0 then
	ls_sponsor_id = null
end if

if ls_acct_mgr_id = 0 then
	ls_acct_mgr_id = null
end if
%>

<body bgcolor=silver onLoad="window.resizeBy(-10, -200)"> 
<center>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter employee review information.
</b></font></td></tr></table>


<form NAME="frmInfo" ACTION="updatereview.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>">

<table border=0>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Last Promotion:</font></td>
	<td><INPUT TYPE="Text" NAME="last_promotion" <% if ls_last_promotion = "9/9/9999" then %>
																Value=""
																<%else%>
																Value="<%=ls_last_promotion%>"
																<%end if%>
	size=10 maxlength=150 style="color: black"></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Last Review:</font></td>
	<td><INPUT TYPE="Text" NAME="last_review" VALUE="<%=ls_last_review%>" size=10 maxlength=150></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Next Review</font></td>
	<td><INPUT TYPE="Text" NAME="next_review" VALUE="<%=ls_next_review%>" size=10 maxlength=150 readonly></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Sponsor:</font></td>
	<td>
		<select name="sponsor_id">
			<option value="0"<%if (isnull(ls_sponsor_id)) then %>selected<%end if%>>None Selected
			<%
			
			  set rs_list = DBQuery("select employeeID, firstname, lastname from employee " & _
			                        "order by lastname")
			  
			  while not rs_list.EOF
			  
			  s_empID = trim(rs_list("employeeID"))
			  s_fn = trim(rs_list("firstname"))
			  s_ln = trim(rs_list("lastname"))
			%>
			  <option value='<%=s_empID%>'<%if s_empID = ls_sponsor_id then%>selected<%end if%>><%=s_ln & " " & s_fn%></option>
			
			<%
			    rs_list.movenext
			    wend
			  rs_list.close
			  set rs_list = nothing
			%>
		</select>
	</td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Account Manager:</font></td>
	<td>
		<select name="account_mgr_id">
			<option value="0"<%if (isnull(ls_acct__id)) then %>selected<%end if%>>None Selected			
			<%
			  'Note: EmployeeTitle_ID of 5, 6, or 16 (13 - Branch manager)corresponds to an account manager position
			  set rs_list = DBQuery("select employeeID, firstname, lastname from employee " & _
			                        "where EmployeeTitle_ID in (5, 6, 13) order by lastname")
			  
			  while not rs_list.EOF
			  
			    s_empID = trim(rs_list("employeeID"))
			    s_fn = trim(rs_list("firstname"))
			    s_ln = trim(rs_list("lastname"))
			%>
			 
			  <option value='<%=s_empID%>'<%if s_empID = ls_acct_mgr_id then%>selected<%end if%>><%=s_ln & ", " & s_fn%></option>
			
			<%
			    rs_list.movenext
			    wend
			  rs_list.close
			  set rs_list = nothing
			%>
		</select>
	</td>
</tr>
<tr>
	<td align=right valign=top><font size=1 face="ms sans serif, arial, geneva">Comments</font></td>
	<td><textarea name="comments" rows="8" cols="30"><%=ls_comments%></textarea></td>
</tr>
</table>


<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="<%=ls_insertupdate%>"><br>
<table border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update Review" OnClick='submitData()' id=button1 name=button1>
			<input type=button class=button value="Cancel" OnClick='window.close();' id=button2 name=button2>
		</td>
	</tr>
</table>

</form>
</center>

</body>
</html>