
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(부서내부코드반영 Transaction)
'*  3. Program ID           : B2902ma2.asp
'*  4. Program Name         : B2902ma2.asp
'*  5. Program Desc         : 내부부서코드반영작업 
'*  6. Modified date(First) : 2000/10/02
'*  7. Modified date(Last)  : 2002/12/16
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "B2902mb2.asp"
Const BIZ_PGM_RESULT_ID = "B2902ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Sub SetDefaultVal()
    'Call CommonQueryRs(" ORGID,ORGDT ", " HORG_ABS ", " ORGID = (SELECT MAX(ORG_CHANGE_ID) FROM HORG_WORK_LIST WHERE WORK_FLAG > '')" , _
	Call CommonQueryRs(" ORGID,ORGDT ", " HORG_ABS ", " ORGID = (SELECT top 1 ORG_CHANGE_ID FROM HORG_WORK_LIST WHERE WORK_FLAG > '' order by insrt_dt desc )" , _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    	               
	frm1.txtOrgChangeID.value = Trim(Replace(lgF0,Chr(11),""))
    if Trim(Replace(lgF0,Chr(11),"")) <> "" then
	    frm1.txtChangeDt.Year  = Left(Trim(Replace(lgF1,Chr(11),"")), 4)
	    frm1.txtChangeDt.Month = Mid(Trim(Replace(lgF1,Chr(11),"")), 5, 2)
	    frm1.txtChangeDt.Day   = Right(Trim(Replace(lgF1,Chr(11),"")), 2)		
	End if		
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "B","NOCOOKIE","MA") %>

End Sub

Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  AND MINOR_CD <> " & FilterVar("*", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboModuleCd, lgF0, lgF1, Chr(11))	
End Sub

Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitComboBox
	Call SetDefaultVal
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
    
	frm1.cboModuleCd.focus
End Sub

Function FncQuery()

End Function

Function FncPrint()
    Call parent.FncPrint()
End Function

Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)
End Function

Function FncExit()
    FncExit = True
End Function

Function btnBatch_OnClick()
	Dim strVal
	dim intRetCD
	
	If frm1.cboModuleCd.value = "" Or frm1.cboModuleCd.value = "*" Then
		Call DisplayMsgBox("970029", "X", "업무", "X")
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=Gen"
	strVal = strVal & "&txtModuleCd=" & Trim(frm1.cboModuleCd.value)
	strVal = strVal & "&txtOrgChangeId=" & frm1.txtOrgChangeID.value  
	strVal = strVal & "&txtChangeDt=" & Trim(Replace(lgF1,Chr(11),""))
	strVAl = strVal & "&txtConfirm=" & "R"
	''strVal = strVal & "&txtChangeDt=" & frm1.txtChangeDt.text	
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Function

Function btnBatchCancel_OnClick()
	Dim strVal
	dim intRetCD
	
	If frm1.cboModuleCd.value = "" Or frm1.cboModuleCd.value = "*" Then
		Call DisplayMsgBox("970029", "X", "업무", "X")
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=Gen"
	strVal = strVal & "&txtModuleCd=" & Trim(frm1.cboModuleCd.value)
	strVal = strVal & "&txtOrgChangeId=" & frm1.txtOrgChangeID.value  
	strVal = strVal & "&txtChangeDt=" & Trim(Replace(lgF1,Chr(11),""))
	strVAl = strVal & "&txtConfirm=" & "C"
	''strVal = strVal & "&txtChangeDt=" & frm1.txtChangeDt.text	
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Function

Function Batch_OK()
	Dim i, j
	i = frm1.hTotal.value
	j = frm1.hSuccess.value
	
	Call DisplayMsgBox("183114", "X", i, j)
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>			
				<TR HEIGHT=20>
					<TD CLASS="TD5">&nbsp;</TD>
					<TD CLASS="TD6">&nbsp;</TD>
				</TR>				
				<TR>
					<TD CLASS="TD5">업무</TD>  
					<TD CLASS="TD6" COLSPAN=3><SELECT NAME="cboModuleCd" tag="12" STYLE="WIDTH: 160px;"></SELECT></TD>
				</TR>				
				<TR>
					<TD CLASS="TD5" NoWrap>반영시킬 개편ID</TD>
					<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtOrgChangeID" SIZE=21 tag="14" STYLE="TEXT-ALIGN:Center" readonly></TD>
				</TR>
				<TR>
					<TD CLASS="TD5">개편일자</TD>
					<TD CLASS="TD6">
					<OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtChangeDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="14" 	ALT="개편일자"></OBJECT></TD>
				</TR>
				<TR>
					<TD CLASS="TD5">&nbsp;</TD>
					<TD CLASS="TD6">&nbsp;</TD>
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
					<TD><BUTTON NAME="btnBatch" CLASS="CLSSBTN" Flag=1>반영</BUTTON>&nbsp;
					    <BUTTON NAME="btnBatchCancel" CLASS="CLSSBTN" Flag=2>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:PgmJump(BIZ_PGM_RESULT_ID)">내부부서코드 Table</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B2902mb2.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hTotal" tag="24">
<INPUT TYPE=HIDDEN NAME="hSuccess" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

