<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 세금계산서 일괄 발행 
'*  3. Program ID           : S5311BA2
'*  4. Program Name         : 
'*  5. Program Desc         : 매출관리 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/07/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'=======================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID = "S5311bb2.asp"

' Constant variables 
'========================================
Const C_PopBillToParty	= 1
Const C_PopTaxBizArea	= 2
Const C_PopSalesGrp		= 3
	
' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim IsOpenPop          
Dim lgBlnOpenedFlag

Dim EndDate, StartDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
StartDate = UNIGetFirstDay(EndDate, Parent.gDateFormat)

'========================================
Sub InitVariables()
End Sub

'========================================
Sub SetDefaultVal()
	With frm1
		.txtFromDt.Text = StartDate
		.txtToDt.Text = EndDate
		.txtWorkDt.Text = EndDate
	End With
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "BA") %>
End Sub

'========================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet
  
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 

	frm1.txtFromDt.focus
	lgBlnOpenedflag = True
End Sub
	
'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================
Function FncExit()
	FncExit = True
End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopBillToParty												
		iArrParam(1) = "dbo.b_biz_partner BP"			' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtBillToParty.value)	' Code Condition
		iArrParam(3) = ""								' Name Cindition
		iArrParam(4) = "EXISTS (SELECT * FROM dbo.b_biz_partner_ftn BF WHERE BP.bp_cd = BF.partner_bp_cd AND BF.partner_ftn = " & FilterVar("SBI", "''", "S") & ") " & _
					   "AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
		iArrParam(5) = "발행처"						' TextBox 명칭 
			
		iArrField(0) = "ED15" & Parent.gColSep & "BP.bp_cd"	' Field명(0)
		iArrField(1) = "ED30" & Parent.gColSep & "BP.bp_nm"	' Field명(1)
		    
		iArrHeader(0) = "발행처"					' Header명(0)
		iArrHeader(1) = "발행처명"					' Header명(1)

	Case C_PopTaxBizArea
		iArrParam(0) = "세금신고사업장"					
		iArrParam(1) = "dbo.b_tax_biz_area"
		iArrParam(2) = Trim(frm1.txtTaxBizArea.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = "세금신고사업장"							

		iArrField(0) = "ED15" & Parent.gColSep & "TAX_BIZ_AREA_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "TAX_BIZ_AREA_NM"

		iArrHeader(0) = "세금신고사업장"							
		iArrHeader(1) = "세금신고사업장명"							

	Case C_PopSalesGrp												
		iArrParam(1) = "dbo.B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"

	End Select
 
	iArrParam(0) = iArrParam(5)							' 팝업 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
End Function

'=======================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopBillToParty
			.txtBillToParty.value = pvArrRet(0) 
			.txtBillToPartyNm.value = pvArrRet(1)   
		Case C_PopTaxBizArea
			.txtTaxBizArea.value = pvArrRet(0) 
			.txtTaxBizAreaNm.value = pvArrRet(1)   
		Case C_PopSalesGrp
			.txtSalesGrp.value = pvArrRet(0) 
			.txtSalesGrpNm.value = pvArrRet(1)   
		End Select
	End With
	
	SetConPopup = True

End Function

'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If lgBlnOpenedFlag Then	GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'	Description : 세금신고 사업장 정보를 Fetch한다.
'====================================================
Function GetTaxBizArea(Byval pvStrFlag)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrTaxBizArea(1), iArrTemp
	
	GetTaxBizArea = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetTaxBizArea ('', '',  " & FilterVar(frm1.txtTaxBizArea.value, "''", "S") & ",  " & FilterVar(pvStrFlag, "''", "S") & ") "
	iStrWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrTaxBizArea(0) = iArrTemp(1)
		iArrTaxBizArea(1) = iArrTemp(2)
		GetTaxBizArea = SetConPopup(iArrTaxBizArea, C_PopTaxBizArea)
	Else
		If Err.number <> 0 Then	Err.Clear 

		' 세금 신고 사업장을 Editing한 경우 
		GetTaxBizArea = OpenConPopup(C_PopTaxBizArea)
	End if
End Function

'=======================================================
Function ExeReflect(ByVal pvStrWkFlag) 
	Call BtnDisabled(1)
	Dim strVal

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then
			Call BtnDisabled(0)
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtFromDt.text , Parent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtWorkDt.Text, Parent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, .txtWorkDt.alt)
			Call BtnDisabled(0)
			.txtFromDt.focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtToDt.text , Parent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtWorkDt.Text, Parent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, .txtWorkDt.alt)	
			Call BtnDisabled(0)
			.txtToDt.Focus()
			Exit Function
		End If

		' 작업을 수행 하시겠습니까?
		If DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X") = vbNo Then
			Call BtnDisabled(0)
			Exit Function
		End If

		strVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0006
		strVal = strVal     & "&txtFromDt="		& .txtFromDt.Text				<%'시작일 %>
		strVal = strVal     & "&txtToDt="		& .txtToDt.Text					<%'종료일 %>
		strVal = strVal     & "&txtWorkDt="		& .txtWorkDt.Text				<%'발행일 %>
		strVal = strVal     & "&txtBilltoparty=" & .txtBilltoparty.value		<%'발행처 %>
		strVal = strVal     & "&txtSalesGrp="	& .txtSalesGrp.value			<%'영업그룹 %>
		strVal = strVal     & "&txtTaxBizArea=" & .txtTaxBizArea.value			<%'세금신고사업장 %>
		strVal = strVal     & "&txtWkFlag=" & pvStrWkFlag						' E : Execute, C : Cancel

		strVal = strVal & "&txtUserId=" & Parent.gUsrID
	End With

	If LayerShowHide(1) = False then
		Call BtnDisabled(0)
		Exit Function 
	End if

	Call RunMyBizASP(MyBizASP, strVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'=======================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Call DisplayMsgBox("990000","X","X","X")
	Call SetFocusToDocument("M")
	frm1.txtFromDt.Focus
End Function

'=======================================================
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgBox("800161","X","X","X")
	Call SetFocusToDocument("M")
	frm1.txtFromDt.Focus
End Function

'   Event Desc : 발행처 
'==========================================
Function txtBillToParty_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtBillToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("SBI", "''", "S") & "", "default", "default", "default", "" & FilterVar("BF", "''", "S") & "", C_PopBillToParty) Then
				.txtBillToParty.value = ""
				.txtBillToPartyNm.value = ""
				.txtBilltoparty.focus
			ELSE
				.txtTaxBizArea.focus
			End If
			txtBillToParty_OnChange = False
		Else
			.txtBillToPartyNm.value = ""
		End If
	End With
End Function

'   Event Desc : 세금신고사업장 변경시 세금신고사업장 정보 Fetch
'==========================================
function txtTaxBizArea_OnChange()
	With frm1
		If Trim(.txtTaxBizArea.value) = "" Then
			.txtTaxBizAreaNm.value = ""
		Else
			IF Not GetTaxBizArea("NM") Then
				.txtTaxBizArea.value= ""
				.txtTaxBizAreaNm.value = ""
				.txtTaxBizArea.focus
			Else
				.txtSalesGrp.focus
			End if
			txtTaxBizArea_OnChange=false
		End if
	End With
End function

'   Event Desc : 영업그룹 
'==========================================
Function txtSalesGrp_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtSalesGrp.value = ""
				.txtSalesGrpNm.value = ""
				.txtSalesGrp.focus
			Else
				.txtFromDt.Focus
			End If
			txtSalesGrp_OnChange = False
		Else
			.txtSalesGrpNm.value = ""
		End If
	End With
End Function

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
		
	End If
End Sub

'========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If
End Sub

'========================================
Sub txtWorkDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtWorkDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtWorkDt.Focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서 일괄 발행</font></td>
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
								<TD CLASS="TD5" NOWRAP>발행일</TD>
								<TD CLASS="TD6" NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5311ba2_fpDateTime1_txtFromDt.js'></script>
											</TD>
											<TD>
												&nbsp;~&nbsp;
											</TD>
											<TD>
												<script language =javascript src='./js/s5311ba2_fpDateTime2_txtToDt.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>작업일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5311ba2_fpDateTime3_txtWorkDt.js'></script>
							</TR>
						    <TR>
								<TD CLASS=TD5>발행처</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBilltoparty" ALT="발행처" SIZE=10 MAXLENGTH=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopBillToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtBilltoPartyNm" SIZE=25 TAG="14"></TD>
								<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizArea" ALT="세금신고사업장" TYPE=TEXT MAXLENGTH=10 SIZE=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopup C_PopTaxBizArea">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=25 TAG="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick=ExeReflect("E") Flag=1>실행</BUTTON>&nbsp;
						<BUTTON NAME="btnCancel" CLASS="CLSSBTN" onclick=ExeReflect("C") Flag=1>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


