<!--
Developer:    Kevin Dill
Date:         09/01/1998
Description:  Stock header.
-->


<html>

<head>
<title>Stock Lookup</title>
</head>


<!-- #include file="../../style.htm" -->



<body bgcolor=black link=red vlink=red alink=red><center>

<table border=0 cellpadding=0 cellspacing=0><tr>

<form name=frmInfo method=get action="stockResults.asp" target="StockResults"><td>
<font size=1 face='ms sans serif, geneva' color=white><b>Stock Lookup:</b></font>
<input type=text name=stock class=fldSearch size=25 maxlength=250 value="<%=request.querystring("stock")%>">
<input type=submit class=fldSearch value="Go"><input type=button class=fldSearch value="Back" onClick="parent.location.href='default.asp'">
</td></form>
</tr></table>

</center>
</body>
</html>