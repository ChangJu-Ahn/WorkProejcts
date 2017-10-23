<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #include file="../../Inc/Common.asp"--> <%'Ãß°¡ÆÄÀÏ %>
<TITLE>
</TITLE>
<style>
.nomaltxt
{
    FONT-SIZE: 9pt;
    COLOR: #333333;
    LINE-HEIGHT: 17px;
    FONT-STYLE: normal;
    FONT-FAMILY: ±¼¸²Ã¼;
    TEXT-DECORATION: none
}
TD
{
    FONT-SIZE: 9pt;
    CURSOR: default;
    FONT-FAMILY: ±¼¸²Ã¼ 
}
INPUT
{
    BORDER-RIGHT: slategray 1px solid;
    BORDER-TOP: slategray 1px solid;
    FONT-SIZE: 9pt;
    BORDER-LEFT: slategray 1px solid;
    BORDER-BOTTOM: slategray 1px solid;
    FONT-FAMILY: ±¼¸²Ã¼;
    HEIGHT: 19px
}
SELECT
{
    BORDER-RIGHT: 1px solid;
    BORDER-TOP: 1px solid;
    BORDER-LEFT: 1px solid;
    BORDER-BOTTOM: 1px solid;
    FONT-FAMILY: ±¼¸²Ã¼ 
}
.TopMenuFont
{ 
    font-size:9pt; 
    font-family:±¼¸²Ã¼;
    color:#000000; 
    text-decoration:none;
}
.btntd02
{
    FONT-SIZE: 9pt;
    BACKGROUND-IMAGE: url(../image/login/buttonmiddle.gif);
    PADDING-TOP: 3px;
    FONT-FAMILY: ±¼¸²Ã¼;
    TEXT-ALIGN: center
}
.tdclass05
{
    FONT-SIZE: 9pt;
    FONT-FAMILY: ±¼¸²Ã¼;
    BACKGROUND-COLOR: #ffffff;
    TEXT-ALIGN: left
}
.tdclass06
{
    FONT-SIZE: 9pt;
    FONT-FAMILY: ±¼¸²Ã¼;
    BACKGROUND-COLOR: #ffffff;
    TEXT-ALIGN: center
}
.btntd02l
{
    FONT-SIZE: 9pt;
    FONT-FAMILY: ±¼¸²Ã¼;
    TEXT-ALIGN: right
}
.btntd02r
{
    FONT-SIZE: 9pt;
    FONT-FAMILY: ±¼¸²Ã¼;
    TEXT-ALIGN: left
}
.RADIO
{
    BORDER: 0PX;
}    
</style>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle01.css">		
</HEAD>
<Script Language="VBScript" SRC="../../inc/common01.vbs">       </Script>
<script language = vbs>

function fcSave()
	Call exeBiz()
End function

Sub exeBiz()
	'MousePT.style.visibility = "visible"
    form.submit
End Sub

</script>
<BODY topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" border=1 scroll=auto>
<center>
<form method=post action="biz.asp" name="form" >

<%

    Dim pRec
    Dim pTA
    Dim pWidth
    Dim pMaster
    Dim pSystem
    Dim pTran
    Dim i


'    pSystem = "#" & "FF6666"
'    pMaster = "#" & "E3E7E3"
'    pTran   = "#" & "E7F1D9"

'    pSystem = "#" & "E0F4FF"
 '   pMaster = "#" & "FFD7B2"
  '  pTran   = "#" & "FFFFDD"
  
    pSystem = "#" & "F4F1E7"
    pMaster = "#" & "DEEEE0"
    pTran   = "#" & "D8E7EF"  
    
    'FCFBF8
    'DEEEE0
    'CAE2D9
    pWidth = 690
    
    Call LoadBasisGlobalInf()
    MetaConnString   = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),gDataBase)


'    Response.Write "<table width = 98% border=0 cellspacing=1 cellpadding=1 class=TAB13>"
'    Response.Write "<tr> "
'    Response.Write "<td align=right> "
'    Response.Write "<INPUT TYPE=BUTTON VALUE='ÀúÀå' onclick='exeBiz()' id=BUTTON1 name=BUTTON1>"
'    Response.Write "</td> "
'    Response.Write "</tr> "
'    Response.Write "</table> "
    
    Set pRec = Server.CreateObject("ADODB.RecordSet")

    pTA = Trim(Request.QueryString("TA"))
    
    If pTA = "" Then
       pRec.Open "select * from z_table_list order by TABLE_ID ",MetaConnString 
    Else
       pRec.Open "select * from z_table_list where table_id like '" & pTA & "%' order by TABLE_ID ",MetaConnString 
    End If

    Response.Write "<table width = 98% border=0 cellspacing=1 cellpadding=1 class=TAB13>"
    Response.Write "<tr> "
    Response.Write "<td> "
    Response.Write "<table border=0 cellspacing=1 cellpadding=1 bgcolor=#cccccc width=100% >"
    Response.Write "<tr HEIGHT=30>"
    Response.Write "<td bgcolor=#e7e5ce         align=center>¼ø¹ø        </td>"
    Response.Write "<td bgcolor=#e7e5ce         align=center>Å×ÀÌºí     </td>"
    Response.Write "<td bgcolor=#e7e5ce         align=center>¼³¸í       </td>"
    Response.Write "<td bgcolor=" & pSystem & " align=center>&nbsp;System&nbsp;</td>"
    Response.Write "<td bgcolor=" & pMaster & " align=center>&nbsp;Master&nbsp;</td>"
    Response.Write "<td bgcolor=" & pTran   & " align=center>&nbsp;Transaction&nbsp;</td>"
    Response.Write "</tr>"


    Do While Not (pRec.EOF Or pRec.BOF)
       Response.Write "<tr> <td bgcolor=#e7e5ce align=right >&nbsp;&nbsp;" & i + 1 &  "&nbsp;</td>"
       Select Case pRec("Kind")
           Case "S"
                     Response.Write "<td bgcolor=" & pSystem & "  ><font color=''>&nbsp;" & pRec("TABLE_ID") &  "</font></td>"
                     Response.Write "<td bgcolor=" & pSystem & "  >&nbsp;" & pRec("TABLE_NM") &  "</td>"
                     Response.Write "<td bgcolor=" & pSystem & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_A name=" & pRec("TABLE_ID") &  " VALUE=S CLASS=RADIO CHECKED>"
                     Response.Write "<td bgcolor=" & pSystem & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_B name=" & pRec("TABLE_ID") &  " VALUE=M CLASS=RADIO        >"
                     Response.Write "<td bgcolor=" & pSystem & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_C name=" & pRec("TABLE_ID") &  " VALUE=T CLASS=RADIO        >"
           Case "M"
                     Response.Write "<td bgcolor=" & pMaster & "  ><font color=''>&nbsp;" & pRec("TABLE_ID") &  "</font></td>"
                     Response.Write "<td bgcolor=" & pMaster & "  >&nbsp;" & pRec("TABLE_NM") &  "</td>"
                     Response.Write "<td bgcolor=" & pMaster & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_A name=" & pRec("TABLE_ID") &  " VALUE=S CLASS=RADIO        >"
                     Response.Write "<td bgcolor=" & pMaster & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_B name=" & pRec("TABLE_ID") &  " VALUE=M CLASS=RADIO CHECKED>"
                     Response.Write "<td bgcolor=" & pMaster & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_C name=" & pRec("TABLE_ID") &  " VALUE=T CLASS=RADIO        >"
           Case Else
                     Response.Write "<td bgcolor=" & pTran   & "  >&nbsp;" & pRec("TABLE_ID") &  "</td>"
                     Response.Write "<td bgcolor=" & pTran   & "  >&nbsp;" & pRec("TABLE_NM") &  "</td>"
                     Response.Write "<td bgcolor=" & pTran   & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_A name=" & pRec("TABLE_ID") &  " VALUE=S CLASS=RADIO        >"
                     Response.Write "<td bgcolor=" & pTran   & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_B name=" & pRec("TABLE_ID") &  " VALUE=M CLASS=RADIO        >"
                     Response.Write "<td bgcolor=" & pTran   & "  align=center><INPUT TYPE=RADIO id=" & pRec("TABLE_ID") &  "_C name=" & pRec("TABLE_ID") &  " VALUE=T CLASS=RADIO CHECKED>"
           End Select          
       Response.Write "</td>" & vbCrLf
       Response.Write "</tr>" & vbCrLf
       pRec.MoveNext()
       i = i + 1
    Loop
    Response.Write "</table>"


       Response.Write "</td>" & vbCrLf
       Response.Write "</tr>" & vbCrLf

    Response.Write "</table>"

    pRec.Close
    
    Set pRec = Nothing 
    
'    Response.Write "<table width = 98% border=0 cellspacing=1 cellpadding=1 class=TAB13>"
'    Response.Write "<tr> "
'    Response.Write "<td align=right> "
'    Response.Write "<INPUT TYPE=BUTTON VALUE='ÀúÀå' onclick='exeBiz()' id=BUTTON1 name=BUTTON1>"
'    Response.Write "</td> "
'    Response.Write "</tr> "
'    Response.Write "</table> "
       
%>

<input type=hidden id=TA name=TA value=<%=Request.QueryString("TA")%> >

</form>

</BODY>
</HTML>