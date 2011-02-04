<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

'
' SERVER SIDE BROWSER SNIFFING CODE
'
Dim SvrVarCollection
Set SvrVarCollection = Request.ServerVariables
Dim ua, os
Dim IE4RelOrNewer, IE5, IE4, IE3, IE302, AnyIE, Netscape3, Netscape4, Netscape5, AnyNetscape, BetaVer, IE4Release
Dim Win95, Win98, WinNT, AnyMac, PPC, Mac68K, Unix
Dim SC1, SC2, MB, MP
Dim IE4Up, NS4Up, NS5Up

''''''''''''''''''''
' Server Variables '
''''''''''''''''''''

ua  = SvrVarCollection("HTTP_USER_AGENT")
os  = SvrVarCollection("HTTP_UA_OS")


'''''''''''''''''''''''''
' Determine the Browser '
'''''''''''''''''''''''''

AnyIE = False
IE5 = False
IE4 = False
IE3 = False
IE302 = False
AnyNetscape = False
Netscape5 = False
Netscape4 = False
Netscape3 = False

If Instr(ua, "MSIE") Then
   AnyIE = True
   If Instr(ua, "MSIE 5.") Then
      IE5 = True
   ElseIf Instr(ua, "MSIE 4.") Then
      IE4 = True
   ElseIf Instr(ua, "MSIE 3.") Then
      IE3 = True
      If Instr(ua, "MSIE 3.02") Then
         IE302 = True
      End If
   End If
ElseIf Instr(ua, "Mozilla") And Instr(ua, "compatible") = 0 Then
   AnyNetscape = True
   If Instr(ua, "Mozilla/5") Then
      Netscape5 = True
   ElseIf Instr(ua, "Mozilla/4") Then
      Netscape4 = True
   ElseIf Instr(ua, "Mozilla/3") Then
      Netscape3 = True
   End If
End If

'''''''''''''''''''''
' Determine if Beta '
'''''''''''''''''''''
BetaVer = False

SC1 = Instr(ua, ";")
SC2 = Instr(SC1 + 1, ua, ";")
MB = Instr(SC1 + 1, ua, "b")
MP = Instr(SC1 + 1, ua, "p")

If (AnyIE And ((MB > SC1 And MB < SC2) Or (MP > SC1 And MP < SC2))) Or (AnyNetscape And Instr(ua, "b") < SC1) Then
   BetaVer = True
End If

''''''''''''''''''''
' Determine the OS '
''''''''''''''''''''

Win95 = False
Win98 = False
WinNT = False
AnyMac = False
PPC = False
Mac68K = False
Unix = False
Win31_WinNT351 = False

If InStr(ua, "Windows 95") Or InStr(ua, "Win95") Then
   Win95 = True
ElseIf InStr(ua, "Windows 98") Or InStr(ua, "Win98") Then
   Win98 = True
ElseIf InStr(ua, "Windows 3.1") Or InStr(os, "Win16") Then
   Win31_WinNT351 = True
ElseIf InStr(ua, "NT") Or InStr(os, "NT") Then
   WinNT = True
ElseIf InStr(ua, "Mac") Then
   AnyMac = True
   If InStr(ua, "PowerPC") Or InStr(ua, "PPC") Then
      PPC = True
   ElseIf InStr(ua, "68000") Or InStr(ua, "68K") Then
      Mac68K = True
   End If
ElseIf InStr(ua, "X11") Then
   Unix = true
End If


' Set Session Variables
Session("AnyIE") = AnyIE
Session("IE5") = IE5
Session("IE4") = IE4
Session("IE3") = IE3
Session("IE302") = IE302
Session("AnyNetscape") = AnyNetscape
Session("Netscape4") = Netscape4
Session("Netscape3") = Netscape3


IE4Up = false
NS4Up = false
NS5Up = false
if(AnyIE) then
 Session("BrowserType") = "IE"
 if IE4 then
  IE4Up = true
 end if
end if
if(AnyNetscape) then
 Session("BrowserType") = "Netscape"
 if(Netscape5) then
  NS5Up = true
  NS4Up = true
 elseif(Netscape4) then 
  NS4Up = true
 end if
end if
Session("IE4Up") = IE4Up
Session("NS4Up") = NS4Up
Session("NS5Up") = NS5Up

Session("BetaVer") = BetaVer

Session("Win95") = Win95
Session("Win98") = Win98
Session("WinNT") = WinNT
Session("AnyMac") = AnyMac
Session("PPC") = PPC
Session("Mac68K") = Mac68K
Session("Unix") = Unix
Session("Win31_WinNT351") = Win31_WinNT351
</SCRIPT>