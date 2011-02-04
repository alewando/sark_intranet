<!--
Developer:    SSeissiger, KWahoff
Date:         07/17/2000
Description:  Allows admin to update a new vacation
-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Edit Vacation</TITLE>
<!-- #include file="../../style.htm" -->

<%
vacID = Request.QueryString("vacID")

Response.Write("<!-- vacid: " & vacid & "-->") 
%>

	<form name=frmInfo action='deleteVacation.asp'>
	<INPUT Type=Hidden Name='DelFlag' Value=''>	<INPUT Type=Hidden Name='vacid' Value=<% =vacid %>>		
	</form>
</HEAD>

<%
	'-------------------------------------------
	'	Execute Database Query		
	'-------------------------------------------

empID = Request.QueryString("empID")
vacID = Request.QueryString("vacID")


set rs = DBQuery("select * from vacation where vacation_id = " & vacID)

	'-----------------------------------------------
	'	Get vacation info							
	'-----------------------------------------------
	ls_vacid = trim(rs("vacation_id"))
	ls_vacstart = trim(rs("vacation_start"))
	ls_vacend = trim(rs("vacation_end"))
	ls_vaccomments = trim(rs("comments"))		
	
rs.close

%>

<body bgcolor=silver onLoad="window.resizeBy(0, -200)"> 
<!--<center> -->

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter employee vacation information.
</b></font></td></tr></table>


<form NAME="frmUpdate" ACTION="editPreviousVacation.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>"> 

<table border=0 height=200>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Vacation Id:</font></td>
	<td><INPUT TYPE="Text" NAME="txtVacation_id" VALUE="<%=ls_vacid%>" size=10 maxlength=50 readonly></td>
</tr>


<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Vacation Start Date:</font></td>
	<td><INPUT TYPE="Text" NAME="txtVacation_start" VALUE="<%=ls_vacstart%>" size=10 maxlength=50></td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Vacation End Date:</font></td>
	<td><INPUT TYPE="Text" NAME="txtVacation_end" VALUE="<%=ls_vacend%>" size=10 maxlength=50></td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT TYPE=text NAME="txtComments" VALUE="<%=ls_vacComments%>" SIZE=50 MAXLENGTH="50"></td>
</tr>
</table>
	<p>
	<INPUT TYPE=button CLASS=button VALUE="Submit" OnClick='subVacation()' id=button1 name=button1>
	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='window.close();' id=button2 name=button2>
	<INPUT type=button class=button value='Delete ' onclick='deleteVacation();' id=button3 name=button3>
</form>

</BODY>

<script language="JavaScript" type="text/javascript">
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

function checkDateStartEnd(inputDateStart, inputDateEnd)
{  //check to make sure start date is before end date
	var errorFlag = 0;
	
	var inputDateStart = inputDateStart.value;
	var s_month, s_day, s_year;
	
	var inputDateEnd = inputDateEnd.value;
	var e_month, e_day, e_year;
	
	// get index of first slash in start date
	var firstSlashIndexStart = inputDateStart.indexOf("/");
	// pull out month from start date
	s_month = inputDateStart.substring(0, firstSlashIndexStart);
	// get index of second slash in start date
	var secondSlashIndexStart = inputDateStart.indexOf("/", firstSlashIndexStart+1);
	// pull out day from start date
	s_day = inputDateStart.substring(firstSlashIndexStart+1, secondSlashIndexStart);
	// pull out year from start date
	s_year = inputDateStart.substring(secondSlashIndexStart+1, inputDateStart.length);
	
	// get index of first slash in end date
	var firstSlashIndexEnd = inputDateEnd.indexOf("/");
	// pull out month from end date
	e_month = inputDateEnd.substring(0, firstSlashIndexEnd);
	// get index of second slash in end date
	var secondSlashIndexEnd = inputDateEnd.indexOf("/", firstSlashIndexEnd+1);
	// pull out day from end date
	e_day = inputDateEnd.substring(firstSlashIndexEnd+1, secondSlashIndexEnd);
	// pull out year from end date
	e_year = inputDateEnd.substring(secondSlashIndexEnd+1, inputDateEnd.length);
	
	// error checking
	if (s_year > e_year)
		errorFlag = 1;
	
	if (s_year == e_year && s_month > e_month)
		errorFlag = 1;
		
	return errorFlag;
}

function subVacation()
{
	frm=document.frmUpdate;
	
	var i;
	var lpErrorFlag = 0;
	var lrErrorFlag = 0;
	var eErrorFlag = 0;
	
	// Check that all dates are valid	
	if (checkDate(document.frmUpdate.txtVacation_start) == 1)
	{
		document.frmUpdate.txtVacation_start.replaceAdjacentText("afterEnd", " (Invalid Date)");
		lpErrorFlag = 1;
	}
	else
	{
		document.frmUpdate.txtVacation_start.replaceAdjacentText("afterEnd", " ");
		lpErrorFlag = 0;
	}
	
	if (checkDate(document.frmUpdate.txtVacation_end) == 1)
	{
		document.frmUpdate.txtVacation_end.replaceAdjacentText("afterEnd", " (Invalid Date)");
		lrErrorFlag = 1;
	}
	else
	{
		document.frmUpdate.txtVacation_end.replaceAdjacentText("afterEnd", " ");
		lrErrorFlag = 0;
	}
	
	
	if (checkDateStartEnd(document.frmUpdate.txtVacation_start,document.frmUpdate.txtVacation_end ) == 1)
	{
		document.frmUpdate.txtVacation_end.replaceAdjacentText("afterEnd", " (Invalid Date)");
		eErrorFlag = 1;
	}
	else
	{
		document.frmUpdate.txtVacation_end.replaceAdjacentText("afterEnd", " ");
		eErrorFlag = 0;
	}
	
	
	if (lrErrorFlag == 0 && lpErrorFlag == 0 && eErrorFlag == 0)
	{
		// if all dates are valid, submit the review data
		document.frmUpdate.submit( );
	}
}

function closeWindow()
{	
	//parent.window.open( );
	window.close();
}

			
function OpenWin(page, EmpID, WinTitle){
	winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=100,width=300,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
}		
<!--

function deleteVacation()
{
	if (confirm("Are you sure you want to delete this vacation?"))
	{
		document.frmInfo.DelFlag.value = "True"
		document.frmInfo.submit()
	}
}

// -->
</script>
</HTML>