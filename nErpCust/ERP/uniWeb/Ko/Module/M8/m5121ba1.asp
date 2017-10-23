<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : 구매 
'*  2. Function Name        : 매입일괄등록 
'*  3. Program ID           : M5121BA1
'*  4. Program Name         :
'*  5. Program Desc         : 매입일괄등록 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005/08/30
'*  8. Modified date(Last)  : 2005/09/08
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Shim Hae Young
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
' =======================================================================================================
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

<SCRIPT LANGUAGE=VBSCRIPT>		            '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "M5121bb1.asp"
Const BIZ_PGM_JUMP_ID = "M5111MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim ToDateOfDB

Dim IsOpenPop

ToDateOfDB = UNIConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gDateFormat)

'=========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
	lgBlnFlgChgValue = False

	With frm1
		If parent.gPurGrp <> "" And Trim(.txtPurGrp.value) = "" Then
			.txtPurGrp.value = parent.gSalesGrp
		End If

		.txtConFromDt.Text	= ToDateOfDB
		.txtConToDt.Text	= ToDateOfDB
		.txtIVDt.Text		= ToDateOfDB

		.txtConFromDt.Focus
	End With
End Sub

'===========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "BA") %>
End Sub

'===========================================
Function JumpIV()
	Call PgmJump(BIZ_PGM_JUMP_ID)
End Function

'=========================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'=========================================
Function FncExcel()
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'=========================================
Function FncFind()
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'=========================================
Function FncExit()
    FncExit = True
End Function

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConFromDt.focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConToDt.focus
	End If
End Sub

'========================================
Sub txtIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIvDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIvDt.focus
	End If
End Sub

Function OpenConMovType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtConMovType.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "입출고형태"
	arrParam(1) = "M_Mvmt_type"

	arrParam(2) = Trim(frm1.txtConMovType.Value)

	arrParam(4) = "USAGE_FLG='Y' "
	arrParam(5) = "입출고형태"

    arrField(0) = "IO_Type_Cd"
    arrField(1) = "IO_Type_NM"

    arrHeader(0) = "입출고형태"
    arrHeader(1) = "입출고형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtConMovTypeNm.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtConMovType.Value	= arrRet(0)
		frm1.txtConMovTypeNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtConMovType.focus
		Set gActiveElement = document.activeElement
	End If
End Function

Function OpenConSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtConSppl.className)=Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"
	arrParam(1) = "B_Biz_Partner"

	arrParam(2) = Trim(frm1.txtConSppl.Value)
	arrParam(3) = ""

	arrParam(4) = "Bp_Type in ('S','CS') AND usage_flag='Y' AND  in_out_flag = 'O' "		'사외거래처만"
	arrParam(5) = "공급처"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

	arrHeader(0) = "공급처"
	arrHeader(1) = "공급처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtConSppl.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtConSppl.Value = arrRet(0)
		frm1.txtConSpplNm.Value = arrRet(1)
		frm1.txtConSppl.focus
		Set gActiveElement = document.activeElement
	End If
End Function

Function OpenConPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtConPurGrp.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"
	arrParam(1) = "B_Pur_Grp"

	arrParam(2) = Trim(frm1.txtConPurGrp.Value)

	arrParam(4) = "B_Pur_Grp.USAGE_FLG='Y' "
	arrParam(5) = "구매그룹"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "구매그룹"
    arrHeader(1) = "구매그룹명"
    arrHeader(2) = "구매조직"
    arrHeader(3) = "구매조직명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtConPurGrp.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtConPurGrp.Value= arrRet(0)
		frm1.txtConPurGrpNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtConPurGrp.focus
		Set gActiveElement = document.activeElement
	End If
End Function

Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPurGrp.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"
	arrParam(1) = "B_Pur_Grp"

	arrParam(2) = Trim(frm1.txtPurGrp.Value)

	arrParam(4) = "B_Pur_Grp.USAGE_FLG='Y' "
	arrParam(5) = "구매그룹"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "구매그룹"
    arrHeader(1) = "구매그룹명"
    arrHeader(2) = "구매조직"
    arrHeader(3) = "구매조직명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtPurGrp.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrp.Value= arrRet(0)
		frm1.txtPurGrpNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtPurGrp.focus
		Set gActiveElement = document.activeElement
	End If
End Function



'========================================
Function ExeIVBatch()
	Call BtnDisabled(1)
	Dim iStrVal

	ExeIVBatch = False                                                          '⊙: Processing is NG

	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Or Not chkField(Document, "2") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	With frm1
		If Not ValidDateCheck(.txtConFromDt, .txtConToDt) Then
			Call BtnDisabled(0)
			Exit Function
		End If

		If Not ValidDateCheck(.txtConToDt, .txtIvDt) Then
			Call BtnDisabled(0)
			Exit Function
		End If

		If DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X") = vbNo Then
			Call BtnDisabled(0)
			Exit Function
		End If

		iStrVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0006
		iStrVal = iStrVal & "&txtConFromDt="	& .txtConFromDt.Text    '입출고일From
		iStrVal = iStrVal & "&txtConToDt="		& .txtConToDt.Text      '입출고일To
		iStrVal = iStrVal & "&txtConMovType="	& .txtConMovType.value  '입출고형태 
		iStrVal = iStrVal & "&txtConSppl="	    & .txtConSppl.value     '공급처 
		iStrVal = iStrVal & "&txtConPurGrp="	& .txtConPurGrp.value   '구매그룹 
		iStrVal = iStrVal & "&txtIvDt="		    & .txtIvDt.Text         '매입일 
		iStrVal = iStrVal & "&txtPurGrp="	    & .txtPurGrp.value   	'구매그룹 

		iStrVal = iStrVal & "&txtIvType="	    & .txtIvType.value       '매입형태 
		iStrVal = iStrVal & "&txtVAT="	        & .txtVAT.value          'VAT
		iStrVal = iStrVal & "&txtVatRt="	    & .txtVatRt.value         'VAT율 

		iStrVal = iStrVal & "&txtPayMeth="	    & .txtPayMeth.value          '결제방법 

		' VAT포함구분 
		If .rdoVatFlg1.checked Then
			iStrVal = iStrVal & "&txtVatFlg=1"      '별도 
        Else
			iStrVal = iStrVal & "&txtVatFlg=2"      '포함 
        End If
		iStrVal = iStrVal & "&txtUserId="		& Parent.gUsrID

		' 전자세금계산서여부 
		If .rdoIssueDTFg1.checked Then
			iStrVal = iStrVal & "&txtIssueDTFg=Y"      '별도 
        Else
			iStrVal = iStrVal & "&txtIssueDTFg=N"      '포함 
        End If

	End With

	If LayerShowHide(1) = False then
		Call BtnDisabled(0)
		Exit Function
	End if

	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeIVBatch = True                                                           '⊙: Processing is NG
End Function

'========================================
Function ExeIVBatchOk()				            '☆: 저장 성공후 실행 로직 
	Call DisplayMsgBox("990000","X","X","X")
End Function

'========================================
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgBox("800161","X","X","X")
End Function

Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5)
	If lblnWinEvent = True Or UCase(frm1.txtIvType.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True

	arrHeader(0) = "매입형태"						' Header명(0)
    arrHeader(1) = "매입형태명"						' Header명(1)

    arrField(0) = "IV_TYPE_CD"							' Field명(0)
    arrField(1) = "IV_TYPE_NM"							' Field명(1)

	arrParam(0) = "매입형태"						' 팝업 명칭 
	arrParam(1) = "M_IV_TYPE"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtIvType.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtIvType.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			' Name Cindition
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and import_flg=" & FilterVar("N", "''", "S") & " "						' Where Condition
	arrParam(5) = "매입형태"						' TextBox 명칭 

    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtIvType.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
    end if
    lblnWinEvent = False
    frm1.txtIvType.focus
	Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenCommPopup()  -------------------------------------------------
Function OpenCommPopup(arrHeader, arrField, arrParam, arrRet)


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		OpenCommPopup = False
	Else
		OpenCommPopup = True
		lgBlnFlgChgValue = True
	End If

End Function



Function OpenVat()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtVAT.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT형태"
	arrParam(1) = "B_MINOR,b_configuration"

	arrParam(2) = Trim(frm1.txtVAT.Value)

	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd "
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT형태"

    arrField(0) = "b_minor.MINOR_CD"
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"

    arrHeader(0) = "VAT형태"
    arrHeader(1) = "VAT형태명"
    arrHeader(2) = "VAT율"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtVAT.focus
		Exit Function
	Else
		frm1.txtVAT.Value    = arrRet(0)
		frm1.txtVATNm.Value    = arrRet(1)
		frm1.txtVatRt.text  = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		frm1.txtVAT.focus
		lgBlnFlgChgValue = True
	End If
	Set gActiveElement = document.activeElement
End Function


Function OpenPaymeth()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPayMeth.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "결제방법"
	arrParam(1) = "B_Minor,b_configuration"

	arrParam(2) = Trim(frm1.txtPayMeth.Value)
	'arrParam(3) = Trim(frm1.txtPayNm.Value)

	arrParam(4) = "b_minor.Major_Cd=" & FilterVar("B9004", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "결제방법"

    arrField(0) = "b_minor.Minor_Cd"
    arrField(1) = "b_minor.Minor_Nm"
    arrField(2) = "b_configuration.REFERENCE"

    arrHeader(0) = "결제방법"
    arrHeader(1) = "결제방법명"
    arrHeader(2) = "Reference"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPayMeth.focus
		Exit Function
	Else
		frm1.txtPayMeth.Value    = arrRet(0)
		frm1.txtPayMethNm.Value    = arrRet(1)
		frm1.txtPayMeth.focus
		lgBlnFlgChgValue = True
	End If
	Set gActiveElement = document.activeElement
End Function
'====================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)
	On Error Resume Next

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp

	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""

	Err.Clear

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, parent.gColSep)
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			If lgBlnOpenedFlag Then
				GetCodeName = OpenConPopup(pvIntWhere)
			End If
		Else
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear
		End If
	End if
End Function
'==========================================================================================


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입일괄등록</font></td>
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
		<TD CLASS="Tab11" VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" ID="idDateTitle" NOWRAP>입출고일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConFromDt" CLASS="FPDTYYYYMMDD" tag="12X1" Alt="시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtConToDt" CLASS="FPDTYYYYMMDD" tag="12X1" Alt="종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>입출고형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConMovType" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="입출고형태" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConMovType()">&nbsp;<INPUT NAME="txtConMovTypeNm" TYPE="Text" Alt="입출고형태명" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>공급처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConSppl" TYPE="Text" MAXLENGTH="10" SIZE=10 Alt="공급처" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSppl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSppl()">&nbsp;<INPUT NAME="txtConSpplNm" TYPE="Text" Alt="공급처명" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConPurGrp" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="구매그룹" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPurGrp">&nbsp;<INPUT NAME="txtConPurGrpNm" TYPE="Text" Alt="구매그룹명" SIZE=25 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
								<TD>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
										<TR>
										<TD>
											<TD CLASS="TD5" ID="idBpTitle" NOWRAP>매입일</TD>
											<TD CLASS="TD6" NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtIvDt" CLASS="FPDTYYYYMMDD" tag="22X1" Alt="매입일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>구매그룹</TD>
											<TD CLASS=TD6 nowrap><INPUT NAME="txtPurGrp" TYPE="Text" Alt="구매그룹" MAXLENGTH=4 SiZE=10 tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurGrp()">&nbsp;<INPUT NAME="txtPurGrpNm" TYPE="Text" Alt="구매그룹명" SIZE=25 tag="14"></TD>
										</TD>
										<TR>
										</TABLE>
								<TR>
								<TD>
										<FIELDSET CLASS="CLSFLD">
										<LEGEND>예외입고/반품 Default Value</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
												<TR>
													<!--TD CLASS=TD5 NOWRAP>매입형태</TD>
													<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIvType" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="매입유형" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvType()">&nbsp;<INPUT NAME="txtIVTypeNm" TYPE="Text" Alt="매입유형명" SIZE=25 tag="14"></TD-->
													<TD CLASS=TD5 NOWRAP>VAT</TD>
													<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVAT" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="VAT" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVAT" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenVAT">&nbsp;<INPUT NAME="txtVATNm" TYPE="Text" Alt="VAT명" SIZE=25 tag="14"></TD>
													<TD CLASS=TD5 NOWRAP>결제방법</TD>
													<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayMeth" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="결제방법" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayMeth" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayMeth()">&nbsp;<INPUT NAME="txtPayMethNm" TYPE="Text" Alt="결제방법명" SIZE=25 tag="14"></TD>
												</TR>
												<TR>
													
													<TD CLASS="TD5" nowrap>VAT율</TD>
													<TD CLASS="TD6" nowrap>
														<Table cellpadding=0 cellspacing=0>
															<TR>
																<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT NAME="txtVatRt" MAXLENGTH=10 CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 style="HEIGHT: 20px; WIDTH: 96px" tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
																</TD>
																<TD WIDTH="*" NOWRAP>%
																</TD>
															</TR>
														</Table>
													</TD>
													<TD CLASS="TD5" nowrap>VAT포함구분</TD>
													<TD CLASS="TD6" nowrap>
														<INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" checked id="rdoVatFlg1" tag="21X"><label for="rdoVatFlg">별도 </label>
														<INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" id="rdoVatFlg2"  tag="21X"><label for="rdoVatFlg">포함&nbsp;</label>
													</TD>
												</TR>
										</TABLE>
										</FIELDSET>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
										<TR>
										<TD>
											<TD CLASS="TD5" ID="idIssueDTFg" NOWRAP>전자세금계산서여부</TD>
											<TD CLASS="TD6" NOWRAP>
                                                <INPUT TYPE=radio NAME="rdoIssueDTFg" ALT="YES" CLASS="RADIO" id="rdoIssueDTFg1" tag="21X"><label for="rdoIssueDTFg">YES </label>
                                                <INPUT TYPE=radio NAME="rdoIssueDTFg" ALT="NO" CLASS="RADIO" checked id="rdoIssueDTFg2"  tag="21X"><label for="rdoIssueDTFg">NO&nbsp;</label>
											</TD>
											<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 nowrap>&nbsp;</TD>
										</TD>
										<TR>
										</TABLE>
								</TD>
								</TR>
							 </TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD VALIGN=TOP>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeIVBatch()" Flag=1>실행</BUTTON></TD>
					<TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpIV()">매입세금계산서</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
