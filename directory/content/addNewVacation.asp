<!--
Developer:    SSeissiger, KWahoff
Date:         07/17/2000
Description:  Allows admin to add a new vacation
-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Add New Vacation</TITLE>
<!-- #include file="../../style.htm" -->
<script language="JavaScript" type="text/javascript">

			
function OpenWin(page, EmpID, WinTitle){
		winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=250,width=350,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
}function openWindow(URL, EmpID, WinTitle) 
{
	
	//alert(EmpID);
		//winUploadDoc=window.open(URL + "?empID=" + EmpID "& TabValue=" + TabValue, WinTitle,"height=615,width=510,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
		winUploadDoc=window.open(URL + "?empID=" + EmpID, WinTitle,"height=615,width=510,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
		intervalID=setInterval('checkChildWindow()', 100)
		
}

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

function checkDateSE(inputDateStart, inputDateEnd)
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

function submitVacdate()
{
	var lpErrorFlag = 0;
	var lrErrorFlag = 0;
	var eErrorFlag = 0;
	
	// Check that all dates are valid	
	if (checkDate(document.frmNewVacation.txtVacation_start) == 1)
	{
		document.frmNewVacation.txtVacation_start.replaceAdjacentText("afterEnd", " (Invalid Date)");
		lpErrorFlag = 1;
	}
	else
	{
		document.frmNewVacation.txtVacation_start.replaceAdjacentText("afterEnd", " ");
		lpErrorFlag = 0;
	}
	
	if (checkDate(document.frmNewVacation.txtVacation_end) == 1)
	{
		document.frmNewVacation.txtVacation_end.replaceAdjacentText("afterEnd", " (Invalid Date)");
		lrErrorFlag = 1;
	}
	else
	{
		document.frmNewVacation.txtVacation_end.replaceAdjacentText("afterEnd", " ");
		lrErrorFlag = 0;
	}
	
	if (checkDateSE(document.frmNewVacation.txtVacation_start, document.frmNewVacation.txtVacation_end ) == 1)
	{
		document.frmNewVacation.txtVacation_end.replaceAdjacentText("afterEnd", " (Invalid Date)");
		eErrorFlag = 1;
	}
	else
	{
		document.frmNewVacation.txtVacation_end.replaceAdjacentText("afterEnd", " ");
		eErrorFlag = 0;
	}
	
	if (lrErrorFlag == 0 && lpErrorFlag == 0 && eErrorFlag == 0)
	{
		// if all dates are valid, submit the vacation data
		document.frmNewVacation.submit( );
	}
}

function closeWindow()
{	
	//parent.window.open( );
	window.close();
}

//-->
</SCRIPT>
</HEAD>

<BODY color=silver  bgcolor=silver onLoad="window.resizeBy(0, -200)"> 

<TABLE border=0><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter employee vacation information.
</b></font></td></tr></TABLE>

<%
EmpID = Request.QueryString("empid")
tabID = Request("tabID")

Response.Write("<!-- EmpID: " & EmpID & "-->")
Response.Write("<!-- tabID: " & tabID & "-->")
%>

<form NAME="frmNewVacation" ACTION="editVacation.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<% =EmpID %>"> 

<table border=0 height=200>

<tr>
	<td width=110 align=left><font size=1 face="ms sans serif, arial, geneva">Vacation Start Date:</font></td>
	<td><INPUT TYPE="Text" NAME="txtVacation_start" VALUE="" size=10 maxlength=50></td>
</tr>

<tr>
	<td width=110 align=left><font size=1 face="ms sans serif, arial, geneva">Vacation End Date:</font></td>
	<td><INPUT TYPE="Text" NAME="txtVacation_end" VALUE="" size=10 maxlength=50></td>
</tr>

<tr>
	<td width=110 align=left><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT TYPE=text NAME="txtComments" VALUE="" SIZE=40 MAXLENGTH="50"></td>
</tr>

</table>

   
	<p>
	<INPUT TYPE=button CLASS=button VALUE="Submit" OnClick='submitVacdate()' id=button1 name=button1>
	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='window.close();' id=button2 name=button2>

</form>

</BODY>
</HTML>