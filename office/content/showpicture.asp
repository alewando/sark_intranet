<HTML>
<HEAD>
  <TITLE>Software Architects Photo Album</TITLE>
  <STYLE>  
  
  .primaryColor
    {
    font-family: Arial;
    font-size: 12pt;
    background-color: #ffffff;
    color: #000000;
    }
  .secondaryColor
    {
    font-family: Arial;
    font-size: 12pt;
    background-color: #990000;
    color: white;
    text-align: center;
    }
  .highlightColor
    {
    font-family: Arial;
    font-size: 12pt;
    background-color: #cccccc;
    color: #ffffff;
    }
  </STYLE>
</HEAD>
<BODY bgcolor="#000000">
<CENTER>
  <TABLE class="secondaryColor" width="100%" border="0" cellspacing="4">
  <TR>
    <TD>
      <B>
        <% 
        dim strName, sfo, oFile
        set fso = CreateObject("Scripting.FileSystemObject")
        set oFile = fso.GetFile(server.mappath(".") & "/../album/" & request.querystring("id"))
        strName = left(oFile.name, len(oFile.name) - 4)
        response.write(strName) 
        set oFile = nothing
        set fso = nothing
        %>
      </B>
    </TD>
  </TR>
  <TR>
    <TD>
      <TABLE class="primaryColor" width="100%">
      <TR>
        <TD class="highlightColor" height="4"></TD>
      </TR>
      <TR>
        <TD align="center">
          <!-- PHOTO -->
          <IMG src="../album/<% response.write(request.querystring("id")) %>" border="0">
          <!-- PHOTO -->
        </TD>
      </TR>
      </TABLE>
    </TD>
  </TR>
  </TABLE>
</CENTER>
</BODY>
</HTML>
