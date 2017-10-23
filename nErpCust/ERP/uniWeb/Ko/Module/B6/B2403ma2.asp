
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(부서확정)
'*  3. Program ID           : B2403ma2.asp
'*  4. Program Name         : B2403ma2.asp
'*  5. Program Desc         : 부서확정작업 
'*  6. Modified date(First) : 2000/10/30
'*  7. Modified date(Last)  : 2002/07/25
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Son Rak Hwan
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

Option Explicit																	'☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "B2403mb2.asp"
Const BIZ_PGM_RESULT_ID = "B2405ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop, Enddate

<% '------------------------------------------  OpenOrgId()  -------------------------------------------
'	Name : OpenOrgID()
'	Description : OrgId PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenOrgId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부서개편ID 팝업"		<%' 팝업 명칭 %>
	arrParam(1) = "horg_abs"				<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtOrgId.value		<%' Code Condition%>
	arrParam(3) = ""						<%' Name Cindition%>
	arrParam(4) = ""						<%' Where Condition%>
	arrParam(5) = "부서개편ID"			<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "orgid"					<%' Field명(0)%>
    arrField(1) = "orgnm"					<%' Field명(1)%>
    arrField(2) = "orgdt"					<%' Field명(2)%>
    
    arrHeader(0) = "부서개편ID"			<%' Header명(0)%>
    arrHeader(1) = "부서개편명"			<%' Header명(1)%>
    arrHeader(2) = "개편일자"			<%' Header명(2)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtOrgId.focus 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOrgId(arrRet)
	End If	
	
End Function

'------------------------------------------  SetOrgId()  --------------------------------------------
'	Name : SetOrgId()
'	Description : OrgId Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOrgId(Byval arrRet)
	With frm1
		.txtOrgId.value = arrRet(0)
		.txtOrgNm.value = arrRet(1)
		
		.txtChangeDt.Year  = Left(arrRet(2), 4)
		.txtChangeDt.Month = Mid(arrRet(2), 5, 2)
		.txtChangeDt.Day   = Right(arrRet(2), 2)		
	End With
End Function

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub Form_Load()

    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
        
    Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
    
	frm1.txtOrgId.focus
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
	
	If Trim(frm1.txtOrgId.value) = "" Then
		Call DisplayMsgBox("970029", "X", "부서개편ID", "X")
		frm1.txtOrgNm.Value = ""
		frm1.txtChangeDt.Value = ""
		frm1.txtOrgId.Focus
		Exit Function
	End If
		
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	frm1.txtOrgNm.value = ""
	frm1.txtChangeDt.value = ""
	
	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=Gen"
	strVal = strVal & "&txtOrgId=" & frm1.txtOrgId.value
	strVal = strVal & "&txtConfirm=" & "R"
	strVal = strVal & "&txtUsrId=" & parent.gUsrID

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Function

Function btnBatchCancel_OnClick()
	Dim strVal
	dim intRetCD
	
	If Trim(frm1.txtOrgId.value) = "" Then
		Call DisplayMsgBox("970029", "X", "부서개편ID", "X")
		frm1.txtOrgNm.Value = ""
		frm1.txtChangeDt.Value = ""
		frm1.txtOrgId.Focus
		Exit Function
	End If
		
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	frm1.txtOrgNm.value = ""
	frm1.txtChangeDt.value = ""
	
	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=Gen"
	strVal = strVal & "&txtOrgId=" & frm1.txtOrgId.value
	strVal = strVal & "&txtConfirm=" & "C"
	strVal = strVal & "&txtUsrId=" & parent.gUsrID
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Function

Function LookUp_OK()
	call CommonQueryRs(" ORGNM,ORGDT "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  	
	frm1.txtOrgnm.value = Trim(Replace(lgF0,Chr(11),""))
	if Trim(Replace(lgF0,Chr(11),"")) <> "" then
	    frm1.txtChangeDt.Year  = Left(Trim(Replace(lgF1,Chr(11),"")), 4)
	    frm1.txtChangeDt.Month = Mid(Trim(Replace(lgF1,Chr(11),"")), 5, 2)
	    frm1.txtChangeDt.Day   = Right(Trim(Replace(lgF1,Chr(11),"")), 2)		
	End if
End Function

Function Batch_OK()
	Dim i, j
	
	i = frm1.hTotal.value
	j = frm1.hSuccess.value
	
	Call LookUp_Ok()
		
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
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						  <TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
                          </TR>
                          <TR>
								<TD CLASS="TD5" NoWrap>부서개편ID</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtOrgId" SIZE=8 MAXLENGTH=5 tag="12XXXU"  ALT="부서개편ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgId" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId()">&nbsp;
												<INPUT TYPE=TEXT NAME="txtOrgNm" Size=30 tag="14">
								</TD>
						  </TR>
                          <TR>
								<TD CLASS="TD5">개편일자</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtChangeDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="14" ALT="개편일자"></OBJECT>');</SCRIPT></TD>
                          </TR>
                          <TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
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
					<TD><BUTTON NAME="btnBatch" CLASS="CLSSBTN" Flag=1>반영</BUTTON>&nbsp;
					    <BUTTON NAME="btnBatchCancel" CLASS="CLSSBTN" Flag=2>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:PgmJump(BIZ_PGM_RESULT_ID)">부서정보등록</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=40><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=40 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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