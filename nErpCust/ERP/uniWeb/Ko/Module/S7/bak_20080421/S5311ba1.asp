<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 세금계산서 자동 발행 
'*  3. Program ID           : S5311BA1
'*  4. Program Name         : 
'*  5. Program Desc         : 매출관리 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/07/19
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'**********************************************************************************************
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
Const BIZ_PGM_ID = "S5311bb1.asp"

' Constant variables 
'========================================
Const C_PopBillToParty	= 1
Const C_PopBillType		= 2
Const C_PopSalesGrp		= 3
Const C_PopTaxBizArea	= 4
	
' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim IsOpenPop          
Dim lgBlnOpenedFlag
Dim lgBlnRegChecked			' 등록/삭제 Check여부 

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
		lgBlnRegChecked = True

		.txtFromDt.Text = StartDate
		.txtToDt.Text = EndDate
		.txtIssuedDt.Text = EndDate
		.cboVatCalcType.value = "2" 
		Call chkByBillNo_OnClick
	End With
End Sub	

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "BA") %>
End Sub

'========================================
Sub InitComboBox()
    With frm1
		Call SetCombo(.cboVatCalcType, "1", "개별")
		Call SetCombo(.cboVatCalcType, "2", "통합")
		Call SetCombo(.cboVatCalcType, "3", "거래처정보참조")
	End With
End Sub

'========================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet

	Call InitComboBox()
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 

	frm1.txtFromDt.focus
	lgBlnOpenedflag = True
End Sub
	
'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
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
		
		frm1.txtBilltoparty.focus

	Case C_PopBillType												
		iArrParam(1) = "s_bill_type_config"
		iArrParam(2) = Trim(frm1.txtBillType.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "매출채권형태"

		iArrField(0) = "bill_type"
		iArrField(1) = "bill_type_nm"

		iArrHeader(0) = "매출채권형태"
		iArrHeader(1) = "매출채권형태명"

		frm1.txtBillType.focus
		
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
	    
	    frm1.txtSalesGrp.focus

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

		frm1.txtTaxBizArea.focus
	End Select
 
	iArrParam(0) = iArrParam(5)

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
		Case C_PopBillType
			.txtBillType.value = pvArrRet(0) 
			.txtBillTypeNm.value = pvArrRet(1)   
		Case C_PopSalesGrp
			.txtSalesGrp.value = pvArrRet(0) 
			.txtSalesGrpNm.value = pvArrRet(1)   
		Case C_PopTaxBizArea
			.txtTaxBizArea.value = pvArrRet(0) 
			.txtTaxBizAreaNm.value = pvArrRet(1)   
		End Select
	End With
	
	SetConPopup = True

End Function

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
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim strVal
	Dim strWkYear
	Dim strWkMonth
	Dim strWkYYYYMM
	Dim strYYYYMMDD
	Dim IntRetCD
	Dim strYear,strMonth,strDay

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

		If UniConvDateToYYYYMMDD(.txtFromDt.text , Parent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtIssuedDt.Text, Parent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, .txtIssuedDt.alt)
			Call BtnDisabled(0)
			.txtFromDt.focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtToDt.text , Parent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtIssuedDt.Text, Parent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, .txtIssuedDt.alt)	
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
		strVal = strVal     & "&txtFromDt="		& .txtFromDt.Text				'시작일 
		strVal = strVal     & "&txtToDt="		& .txtToDt.Text					'종료일 
		strVal = strVal     & "&txtIssuedDt="	& .txtIssuedDt.Text				'발행일 
		strVal = strVal     & "&txtBilltoparty=" & .txtBilltoparty.value		'발행처 
		strVal = strVal     & "&txtBillType="	& .txtBillType.value			'매출채권유형 
		strVal = strVal     & "&txtSalesGrp="	& .txtSalesGrp.value			'영업그룹 
		strVal = strVal     & "&txtTaxBizArea=" & .txtTaxBizArea.value			'세금계산서 

		'B/L포함여부 
		strVal = strVal     & "&txtBLFlag=" & "N"

		'발행여부 
		If .rdoPostFlag1.checked Then
			strVal = strVal     & "&txtPostFlag=" & "Y"
		Else
			strVal = strVal     & "&txtPostFlag=" & "N"
		End If

		'영업그룹별 
		If .chkBySalesGrp.checked Then
			strVal = strVal     & "&txtBySalesGrp=" & "Y"
		Else
			strVal = strVal     & "&txtBySalesGrp=" & "N"
		End If
		
		'매출채권유형별 
		If .chkByBillType.checked Then
			strVal = strVal     & "&txtByBillType=" & "Y"
		Else
			strVal = strVal     & "&txtByBillType=" & "N"
		End If
		
		'매출채권번호별 
		If .chkByBillNo.checked Then
			strVal = strVal     & "&txtByBillNo=" & "Y"
			strVal = strVal     & "&txtVatCalcType=" & "4"						'VAT적용기준 
			'매출채권 : 세금계산서 1 : 1
			If .rdoVatMixedFlag2.checked AND .rdoTaxbillDevidedFlag2.checked Then
				strVal = strVal     & "&txtByOnlyBillNo=" & "Y"
			Else
				strVal = strVal     & "&txtByOnlyBillNo=" & "N"
			End If
		Else
			strVal = strVal     & "&txtVatCalcType=" & .cboVatCalcType.value		'VAT적용기준 
			strVal = strVal & "&txtByBillNo=" &			  "N"
			strVal = strVal & "&txtByOnlyBillNo=" &		  "N"
		End If
		
		strVal = strVal & "&txtUserId=" & Parent.gUsrID
		
		' 등록 
		If .rdoWorkTypeReg.checked Then
			strVal = strVal & "&txtWorkType=C"
		' 삭제 
		Else
			strVal = strVal & "&txtWorkType=D"
		End If
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

'========================================
Sub rdoWorkTypeReg_OnClick()
	If Not lgBlnRegChecked Then
		lgBlnRegChecked = True
		idDateTitle.innerHTML = "매출채권일"
		Call ggoOper.SetReqAttr(frm1.txtIssuedDt,"N")
		Call ggoOper.SetReqAttr(frm1.cboVatCalcType,"N")
		Call ggoOper.SetReqAttr(frm1.chkByBillType,"D")
		Call ggoOper.SetReqAttr(frm1.chkBySalesGrp,"D")
		Call ggoOper.SetReqAttr(frm1.chkByBillNo,"D")
		Call ggoOper.SetReqAttr(frm1.txtBillType,"D")
		frm1.btnBillType.disabled = False
	End If
End Sub

'========================================
Sub rdoWorkTypeDel_OnClick()
	If lgBlnRegChecked Then
		lgBlnRegChecked = False
		idDateTitle.innerHTML = "발행일"
		Call ggoOper.SetReqAttr(frm1.txtIssuedDt,"Q")
		Call ggoOper.SetReqAttr(frm1.cboVatCalcType,"Q")
		Call ggoOper.SetReqAttr(frm1.chkByBillType,"Q")
		Call ggoOper.SetReqAttr(frm1.chkBySalesGrp,"Q")
		Call ggoOper.SetReqAttr(frm1.chkByBillNo,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBillType,"Q")
		frm1.btnBillType.disabled = True
	End If
End Sub

' 발행처 
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
				.txtBillType.focus
			End If
			txtBillToParty_OnChange = False
		Else
			.txtBillToPartyNm.value = ""
		End If
	End With
End Function

' 매출채권유형 
'==========================================
Function txtBillType_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtBillType.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("BT", "''", "S") & "", C_PopBillType) Then
				.txtBillType.value = ""
				.txtBillTypeNm.value = ""
				.txtBillType.focus
			Else
				.txtSalesGrp.focus
			End If
			txtBillType_OnChange = False
		Else
			.txtBillTypeNm.value = ""
		End If
	End With
End Function

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
				.txtTaxBizArea.focus
			End If
			txtSalesGrp_OnChange = False
		Else
			.txtSalesGrpNm.value = ""
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
				.cboVatCalcType.focus
			End if
			txtTaxBizArea_OnChange=false
		End if
	End With
End function

'   Event Desc : 매출채권번호별 집계 
'==========================================
Sub chkByBillNo_OnClick()
	' 세금계산서 관리방법 Check	
	With frm1
		if .chkByBillNo.checked Then
			Call ggoOper.SetReqAttr(.rdoVatMixedFlag1, "D")
			Call ggoOper.SetReqAttr(.rdoVatMixedFlag2, "D")
			Call ggoOper.SetReqAttr(.rdoTaxbillDevidedFlag1, "D")
			Call ggoOper.SetReqAttr(.rdoTaxbillDevidedFlag2, "D")
			Call ggoOper.SetReqAttr(.chkBySalesGrp, "Q")
			Call ggoOper.SetReqAttr(.chkByBillType, "Q")
			Call ggoOper.SetReqAttr(.cboVatCalcType, "Q")
		else
			Call ggoOper.SetReqAttr(.rdoVatMixedFlag1, "Q")
			Call ggoOper.SetReqAttr(.rdoVatMixedFlag2, "Q")
			Call ggoOper.SetReqAttr(.rdoTaxbillDevidedFlag1, "Q")
			Call ggoOper.SetReqAttr(.rdoTaxbillDevidedFlag2, "Q")
			Call ggoOper.SetReqAttr(.chkBySalesGrp, "D")
			Call ggoOper.SetReqAttr(.chkByBillType, "D")
			Call ggoOper.SetReqAttr(.cboVatCalcType, "N")
		end if
	End With
End Sub

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If
End Sub

Sub txtIssuedDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIssuedDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssuedDt.Focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서 일괄 등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>작업유형</TD>
							    <TD CLASS=TD6><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="Y" CHECKED ID="rdoWorkTypeReg"><LABEL FOR="rdoWorkTypeReg">등록</LABEL>&nbsp;
							                  <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="N" ID="rdoWorkTypeDel"><LABEL FOR="rdoWorkTypeDel">삭제</LABEL></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="idDateTitle" NOWRAP>매출채권일</TD>
								<TD CLASS="TD6" NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5311ba1_fpDateTime1_txtFromDt.js'></script>
											</TD>
											<TD>
												&nbsp;~&nbsp;
											</TD>
											<TD>
												<script language =javascript src='./js/s5311ba1_fpDateTime2_txtToDt.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>발행일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5311ba1_fpDateTime3_txtIssuedDt.js'></script>
							</TR>
						    <TR>
								<TD CLASS=TD5>발행처</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBilltoparty" ALT="발행처" SIZE=10 MAXLENGTH=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopBillToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtBilltoPartyNm" SIZE=25 TAG="14"></TD>
								<TD CLASS=TD5 NOWRAP>매출채권형태</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillType" ALT="매출채권형태" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopBillType">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizArea" ALT="세금신고사업장" TYPE=TEXT MAXLENGTH=10 SIZE=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopup C_PopTaxBizArea">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=25 TAG="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>VAT적용기준</TD>
	                        	<TD CLASS="TD6" NOWRAP>
                					<SELECT Name="cboVatCalcType" ALT="VAT적용기준" CLASS ="cbonormal" tag="12"><OPTION></OPTION></SELECT>
		                    	</TD>
								<TD CLASS=TD5 NOWRAP>발행여부</TD>
				                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlag" TAG="12X" VALUE="Y" CHECKED ID="rdoPostFlag1"><LABEL FOR="rdoPostFlag1">Y</LABEL>&nbsp;
				                                     <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlag" TAG="12X" VALUE="N" ID="rdoPostFlag2"><LABEL FOR="rdoPostFlag2">N</LABEL></TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100% CLASS=TD6 COLSPAN=4>
									<FIELDSET ID="fldAggration" CLASS="CLSFLD">
									<LEGEND ALIGN=LEFT><LABEL>집계단위</LABEL></LEGEND>
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS=TD5 NOWRAP>영업그룹별</TD>
											    <TD CLASS=TD6 NOWRAP title="영업그룹별"><INPUT TYPE=CHECKBOX NAME="chkBySalesGrp" tag="11" Class="Check"></TD>
												<TD CLASS=TD5 NOWRAP>매출채권유형별</TD>
											    <TD CLASS=TD6 NOWRAP title="매출채권유형별"><INPUT TYPE=CHECKBOX NAME="chkByBillType" tag="11" Class="Check"></TD>
											</TR>
											<TR>
												<TD HEIGHT=20 WIDTH=100% CLASS=TD6 COLSPAN=4>
													<FIELDSET ID="fldbyBillNo" CLASS="CLSFLD" TITLE="매출번호별">
													<LEGEND ALIGN=LEFT><LABEL FOR="chkByBillNo">매출번호별</LABEL>&nbsp;<INPUT TYPE=CHECKBOX NAME="chkByBillNo" tag="11" Class="Check"></LEGEND>
														<TABLE <%=LR_SPACE_TYPE_40%>>
															<TR>
																<TD CLASS=TD5 NOWRAP>부가세혼합여부</TD>
															    <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoVatMixedFlag" TAG="11X" VALUE="Y" ID="rdoVatMixedFlag1"><LABEL FOR="rdoVatMixedFlag1">Y</LABEL>&nbsp;
															                         <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoVatMixedFlag" TAG="11X" VALUE="N" CHECKED ID="rdoVatMixedFlag2"><LABEL FOR="rdoVatMixedFlag2">N</LABEL></TD>
																<TD CLASS=TD5 NOWRAP>계산서분할여부</TD>
															    <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxbillDevidedFlag" TAG="11X" VALUE="Y" ID="rdoTaxbillDevidedFlag1"><LABEL FOR="rdoTaxbillDevidedFlag1">Y</LABEL>&nbsp;
															                         <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxbillDevidedFlag" TAG="11X" VALUE="N" CHECKED ID="rdoTaxbillDevidedFlag2"><LABEL FOR="rdoTaxbillDevidedFlag2">N</LABEL></TD>
															</TR>
														</TABLE>
													</FIELDSET>
												</TD>
											<TR>
										</TABLE>
									</FIELDSET>
								</TD>
							<TR>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON></TD>
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


