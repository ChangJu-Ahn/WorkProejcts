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

<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->

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

Sub SetDB()
    If frm1.SDBC.checked = true Then
       frm1.sdb.value ="unierp272"
    Else
       frm1.sdb.value =""
    End If   
End Sub

Sub doWork()

	Dim X
	Dim i
	Dim TartgetDB
	Dim SourceDB
	
	X = 0
	For i=0 To frm1.SDBList.length - 1
		If frm1.SDBList(i).checked = True Then
			SourceDB = frm1.SDBList(i).value
			X = X + 1 
		End If
	Next
	
	If X = 0 Then
		Call DisplayMsgBox("990055","X","X","X")
		Exit Sub
	End If
	
	X = 0
	For i=0 To frm1.TDBList.length - 1
		If frm1.TDBList(i).checked = True Then
			TartgetDB = frm1.TDBList(i).value
			X = X + 1 
		End If
	Next	
	
	If X = 0 Then
		Call DisplayMsgBox("990056","X","X","X")
		Exit Sub
	End If	
	
	If UCASE(SourceDB) = UCASE(TartgetDB) Then
		Call DisplayMsgBox("990057","X","X","X")
		Exit Sub
	End If	
	
	
    frm1.submit
    
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

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1  action="ZK000MB6.asp" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%> ></TD>
	</TR>
	<TR HEIGHT=23>
		<TD >
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>테이블복사</font></td>
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
		<TD CLASS="Tab11" >
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD CLASS="TD5" width=50%><CENTER>Source</CENTER></TD><TD CLASS="TD5" width=50%><CENTER>Target</CENTER></TD>
				</TR>
				<TR>
					<TD bgcolor=white><%=DBRadioList(MasterConnString,"SDBList")%></TD><TD bgcolor=white><%=DBRadioList(MasterConnString,"TDBList")%></TD>
				</TR>
				
				<tr>
					<TD CLASS="TD5" >Range    </TD>
					<TD bgcolor=white>&nbsp;&nbsp;<INPUT TYPE=CHECKBOX name=RS_S ID=RS_S VALUE=S CLASS=RADIO DISABLED CHECKED>System<INPUT TYPE=CHECKBOX name=RS_M ID=RS_M VALUE=M  CLASS=RADIO>Master<INPUT TYPE=CHECKBOX name=RS_T ID=RS_T VALUE=T  CLASS=RADIO>Transaction</TD>
				</TR>
          
			</TABLE>
		</TD>
	</TR>
	
	
   <TR HEIGHT=20>
		<TD >
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:doWork()" Flag=1>복사시작</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>	
	
	<TR>
		<TD HEIGHT=20 ><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>


