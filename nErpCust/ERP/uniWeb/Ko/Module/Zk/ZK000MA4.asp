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
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>

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
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


Sub doWork()

	Dim X
	Dim i
	Dim IntRetCD
	
	X = 0
	For i=0 To frm1.DBList.length - 1
		If frm1.DBList(i).checked = True Then
			X = X + 1 
		End If
	Next
	
	If X > 0 Then
	
		IntRetCD = DisplayMsgBox("990053", parent.VB_YES_NO,"X","X")  'You want to delete this DataBase?
		If IntRetCD = vbNo Then
			Exit Sub
		Else
			MousePT.style.visibility = "visible"
			frm1.submit
		End If
    
    Else
    
		'Msgbox ("선택된 데이타베이스가 없습니다.")
		Call DisplayMsgBox("990054","X","X","X")
		Exit Sub
    
    End If
    
End Sub

Sub NonSelected()

    MousePT.style.visibility = "hidden"

End Sub

Sub OkProcess()
    MousePT.style.visibility = "hidden"
    document.location.href = "ZK000MA4.asp"
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<%
Call LoadBasisGlobalInf()

MasterConnString = MakeConnString(GetGlobalInf("gDBServerIP") ,GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),"master")      

%>
<BODY TABINDEX="-1" SCROLL="auto">
<FORM NAME=frm1  METHOD="POST" action="ZK000MB4.asp" target=MyBizASP >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>데이터베이스 삭제</font></td>
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
								<TD CLASS="TD6" NOWRAP>
								<%=DBRadioList(MasterConnString,"DBList")%>
								</TD>															
							</TR>
    					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:doWork()" Flag=1>삭제</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT" style='visibility:hidden;    LEFT: expression((document.body.clientWidth-320)/2);    TOP: expression(document.body.clientHeight/2);' align=center>
<table BORDER=1 width="320" border=1 cellpadding=1 cellspacing=1 bordercolor=#CCCCCC bordercolorlight=#CCCCCC bgcolor="buttonface" bordercolordark="#000000" vspace="0" hspace="0">
  <tr bgcolor="#CED3E7"> 
    <td bgcolor="#FFFFFF"><img src="../../image/net.gif" width="32" height="31" vspace="0" hspace="0" align="absmiddle">
      <b>&nbsp;&nbsp;데이터베이스를 삭제하는 중입니다.</b></td>
  </tr>
</table>
</DIV>


</BODY>
</HTML>


