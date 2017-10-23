<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Yim Young Ju
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>
<%
    Const ggWidth = 700
    Const gpWidth = 500
    Session("SDB") = EnCode(Trim(UCase(Request.Form("SDB"))))
    Session("TDB") = EnCode(Trim(UCase(Request.Form("TDB"))))
%>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 

On Error Resume Next
	Err.Clear  
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function


Sub SelectDB
    top.location.href = "ZK000MA2.asp"
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
<%

Call LoadBasisGlobalInf()

%>
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO" >
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>데이터베이스생성</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>   
							<TR>
								<TD  CLASS=TD5 ALIGN=CENTER NOWRAP COLSPAN=2>서버정보</TD>
								</TD>															
							</TR>
							<TR>
						  		<TD CLASS=TD5 ALIGN=LEFT NOWRAP>데이타베이스서버</TD>
								<TD CLASS=TD6 NOWRAP><%=GetGlobalInf("gDBServerIP")%></TD>
						        </TD>
							</TR>
							<TR>
						  		<TD CLASS=TD5 ALIGN=LEFT NOWRAP>소스데이타베이스</TD>
								<TD CLASS=TD6 NOWRAP><%=DeCode(Session("SDB"))%></TD>
						        </TD>
							</TR>
							<TR>
						  		<TD CLASS=TD5 ALIGN=LEFT NOWRAP>생성데이타베이스</TD>
								<TD CLASS=TD6 NOWRAP><%=DeCode(Session("TDB"))%></TD>
						        </TD>
							</TR>
							
							<TR>
								<TD CLASS=TD5 ALIGN=CENTER  NOWRAP COLSPAN=2>현재진척도</TD>
								</TD>															
							</TR>
							
							<TR>
								<TD CLASS=TD5 WIDTH=20%>메시지</TD>
								<TD CLASS=TD6 ><SPAN NAME=txtMessage ID=txtMessage ><SPAN></TD>    
							</TR>
							
    					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
	
	<DIV ID="MousePT" NAME="MousePT" style='visibility:visible;    LEFT: expression((document.body.clientWidth-320)/2);    TOP: expression(document.body.clientHeight/2);' align=center>
		<table BORDER=1 width="320" border=1 cellpadding=1 cellspacing=1 bordercolor=#CCCCCC bordercolorlight=#CCCCCC bgcolor="buttonface" bordercolordark="#000000" vspace="0" hspace="0">
		<tr bgcolor="#CED3E7"> 
		<td bgcolor="#FFFFFF"><img src="../../image/net.gif" width="32" height="31" vspace="0" hspace="0" align="absmiddle">
		  <b>&nbsp;&nbsp;데이터 베이스를 생성하는 중입니다</b></td>
		</tr>
		</table>
	</DIV>


	<%
	    Dim SDBName,CDBName
	    Dim WorkingDir
	    Dim iTemp
	    
	    Dim gtotObjectCount
	    Dim gstaTimer

	    gtotObjectCount = 0 

	    If Response.Buffer Then Response.Flush

	    MasterConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),"master")   
	    MetaConnString   = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),gDataBase    )   
	    SourceConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),DeCode(Session("SDB")))   
	    TargetConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),DeCode(Session("TDB")))   
   
	    
	    SDBName	   = DeCode(Session("SDB")) 
	    CDBName	   = DeCode(Session("TDB")) 
	    
	    gstaTimer = Timer
	    
	    Call RefreshProgressBar("S")
	    
	    iTemp = CopyDBToDB(MetaConnString,SDBName,CDBName)    
	   
	    Call RefreshProgressBar("E")
	 
	    Call ShowMessage(iTemp)

	Sub RefreshProgressBar(ByVal pMark)
	%>
	    <SCRIPT LANGUAGE="VBS">
	       Dim intWidth
	       
	       Select Case "<%=pMark%>"
	           Case "S"  : document.all("txtMessage").innerText = "데이터베이스 생성중"                         
	           Case "E"  : document.all("MousePT").style.visibility ="hidden"
             
	       End Select
	    </SCRIPT>
	<%
	  If Response.Buffer Then Response.Flush
	  
	End Sub


	Sub ShowMessage(ByVal pData)
	%>
	    <SCRIPT LANGUAGE="VBS">
	       document.all("txtMessage").innerText  = "<%= pData %>"
	    </SCRIPT>
	<%
	End Sub

	%>


<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>


