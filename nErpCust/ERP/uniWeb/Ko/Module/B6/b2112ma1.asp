<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : B2112MA1
'*  4. Program Name         : TAX Biz Area(세금신고사업장정보등록)
'*  5. Program Desc         : 세금신고사업장정보등록 
'*  6. Component List       : ADO
'*  7. Modified date(First) : 2002/07/19
'*  8. Modified date(Last)  : 2002/09/25
'*  9. Modifier (First)     : LEENAMYO
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit


'==========================================================================================================
<%
StartDate = DateSerial(Year(Date),Month(Date),1)
StartDate = Year(StartDate) & "-" & Right("0" & Month(StartDate),2) & "-" & Right("0" & Day(StartDate),2)
EndDate = Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2)
%>

Const BIZ_PGM_ID = "b2112mb1.asp"											 '☆: 비지니스 로직 ASP명 

'============================================  1.2.2 Global 변수 선언  ===================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2. Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgIntGrpCount				'☜: Group View Size를 조사할 변수 
Dim lgIntFlgMode					'☜: Variable is for Operation Status

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal
Dim IsOpenPop
Dim lgCboKeyPress
Dim lgOldIndex
Dim lgOldIndex2


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenbizareaInfo()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
Function OpenbizareaInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "세금신고사업장 팝업"					' 팝업 명칭 
	arrParam(1) = "B_TAX_BIZ_AREA"							' TABLE 명칭 
	arrParam(2) = strCode									' Code Condition
	arrParam(3) = ""										' Name COndition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "세금신고사업장"

    arrField(0) = "TAX_BIZ_AREA_CD"							' Field명(0)
    arrField(1) = "TAX_BIZ_AREA_NM"							' Field명(1)

    arrHeader(0) = "세금신고사업장코드"					' Header명(0)
    arrHeader(1) = "세금신고사업장명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTaxBizAreaCd.focus
	    Exit Function
	Else
		Call SetbizareaInfo(arrRet,iWhere)
	End If
End Function

Function OpenZipCode(ByVal strCode, ByVal iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	iCalledAspName = AskPRAspName("ZipPopup")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZipPopup", "X")
		IsOpenPop = False
		Exit Function
	End If

	'//조회모드일경우 팝업되지 않게 ////
	If lgIntFlgMode = parent.OPMD_UMODE Then
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode
	arrParam(1) = ""
	arrParam(2) = parent.gCountry

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtZipCode.focus
	    Exit Function
	Else
		Call SetBizAreaInfo(arrRet,iWhere)
	End If

End Function


'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------------
Function SetBizAreaInfo(ByVal arrRet, ByVal iWhere)

	With frm1
		If iWhere = 0 Then
			.txtTaxBizAreaCd.focus
			.txtTaxBizAreaCd.value = arrRet(0)
			.txtTaxBizAreaNm.value = arrRet(1)
		ElseIf iWhere = 1 Then
			.txtZipCode.focus
			.txtZipCode.value = arrRet(0)
			.txtAddr1.value     = arrRet(1)

			lgBlnFlgChgValue = True
		End If
	End With

End Function

Function OpenTaxOffice(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	'//조회모드일경우 팝업되지 않게 ////
'	If lgIntFlgMode = parent.OPMD_UMODE Then
'		IsOpenPop = False
'		Exit Function
'	End If



	arrParam(0) = "세무서 팝업"						' 팝업 명칭 
	arrParam(1) = "B_TAX_OFFICE"						' TABLE 명칭 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "세무서"

    arrField(0) = "TAX_OFFICE_CD"						' Field명(0)
    arrField(1) = "TAX_OFFICE_NM"						' Field명(1)

    arrHeader(0) = "세무서코드"						' Header명(0)
    arrHeader(1) = "세무서명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTaxOfficeCd.focus
	    Exit Function
	Else
		Call SetTaxOffice(arrRet,iWhere)
	End If

End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetTaxOffice()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTaxOffice(Byval arrRet,byval iWhere)

	With frm1
		If iWhere = 1 Then
			.txtTaxOfficeCd.focus
			.txtTaxOfficeCd.value   = arrRet(0)
			.txtTaxOfficeNm.value = arrRet(1)
		End If

		lgBlnFlgChgValue = True
	End With

End Function


Function OpenCommonPopupInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	'//조회모드일경우 팝업되지 않게 ////
	If lgIntFlgMode = parent.OPMD_UMODE Then
		IsOpenPop = False
		Exit Function
	End If

	select case iwhere
		case 0
			arrParam(0) = "업태 팝업"					' 팝업 명칭 
			arrParam(1) = "B_MINOR"						' TABLE 명칭 
			arrParam(2) =  strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("B9003", "''", "S") & "  "			' Where Condition
			arrParam(5) = "업태"

			arrField(0) = "MINOR_CD"					' Field명(0)
			arrField(1) = "MINOR_NM"					' Field명(1)

			arrHeader(0) = "업태코드"					' Header명(0)
			arrHeader(1) = "업태명"					' Header명(1)

		case 1
			arrParam(0) = "업종 팝업"					' 팝업 명칭 
			arrParam(1) = "B_MINOR"						' TABLE 명칭 
			arrParam(2) =  strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("B9002", "''", "S") & "  "			' Where Condition
			arrParam(5) = "업종"

			arrField(0) = "MINOR_CD"					' Field명(0)
			arrField(1) = "MINOR_NM"					' Field명(1)

			arrHeader(0) = "업종코드"					' Header명(0)
			arrHeader(1) = "업종명"					' Header명(1)  

 	End Select

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		select case iwhere
			case 0
				frm1.txtInd_class.focus
			case 1
				frm1.txtInd_Type.focus
		End Select
	    Exit Function
	Else
		Call SetCommonPopupInfo(arrRet,iWhere)
	End If

End Function

Function SetCommonPopupInfo(Byval arrRet,byval iWhere)

	With frm1
		If iWhere = 0 Then
			.txtInd_class.focus
			.txtInd_class.value = arrRet(0)
			.txtInd_class_Nm.value = arrRet(1)
		Elseif iWhere = 1 Then
			.txtInd_Type.focus
			.txtInd_Type.value = arrRet(0)
			.txtInd_Type_Nm.value   = arrRet(1)
		End If

		lgBlnFlgChgValue = True
	End With

End Function


'==========================================================================================================
Sub Form_Load()

    Call InitVariables
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call SetToolBar("1110100000001111")
	frm1.txtTaxBizAreaCd.focus
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

	FncQuery = False
    Err.Clear
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    frm1.txtTaxBizAreaNm.value = ""

  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'☜: Query db data

    FncQuery = True

End Function


'========================================================================================
Function FncNew()
    Dim IntRetCD

	FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables

    Call SetToolBar("1110100000001111")
	frm1.txtTaxBizAreaCd.focus
    FncNew = True
End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete
    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 

    FncSave = False
    Err.Clear 

  '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '⊙: No data changed!!
        Exit Function
    End If

  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If

  '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave
    FncSave = True 
End Function


'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgIntFlgMode = parent.OPMD_CMODE

    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.LockField(Document, "N")

	lgBlnFlgChgValue = True
    frm1.txtTaxBizAreaCd_Body.value = ""
    frm1.txtTaxBizAreaCd_Body.focus
End Function


'========================================================================================
Function FncCancel()
	On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
	On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
	On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
	On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtTaxBizAreaCd = " & lgPrevNo

	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtTaxBizAreaCd=" & lgNextNo

	Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)
End Function


'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtTaxBizAreaCd_Body=" & Trim(frm1.txtTaxBizAreaCd_Body.value)
    strVal = strVal & "&txtOwnRgstNo=" & Trim(frm1.txtOwnRgstNo.value)

	Call RunMyBizASP(MyBizASP, strVal)
    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	lgBlnFlgChgValue = False
	Call FncNew()
End Function



'========================================================================================
Sub cboXCH_RATE_FG_OnChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtTaxBizAreaCd=" & Trim(frm1.txtTaxBizAreaCd.value)
    
    call RunMyBizASP(MyBizASP, strVal)
    DbQuery = True
End Function


'========================================================================================
Function DbQueryOk()
    Call SetToolBar("1111100000111111")
    Call ggoOper.LockField(Document, "Q")
    lgIntFlgMode = parent.OPMD_UMODE
End Function


'========================================================================================
Function DbSave() 
    Err.Clear
	DbSave = False

    Dim strVal
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function


'========================================================================================
Function DbSaveOk()
    frm1.txtTaxBizAreaCd.value = frm1.txtTaxBizAreaCd_Body.value
    lgBlnFlgChgValue = False
    FncQuery
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>세금신고사업장</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtTaxBizAreaCd" MAXLENGTH="10" SIZE=10 ALT ="세금신고사업장코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenbizareaInfo(frm1.txtTaxBizAreaCd.value,0)">
													<INPUT NAME="txtTaxBizAreaNm" MAXLENGTH="50" SIZE=30 ALT ="세금신고사업장명" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>세금신고사업장코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd_Body" ALT="세금신고사업장코드" MAXLENGTH="10" SIZE=10 tag = "23XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>세금신고사업장명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaNm_Body" ALT="세금신고사업장명" MAXLENGTH="50" tag="23"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세금신고사업장장문명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaFullNm" ALT="세금신고사업장장문명" MAXLENGTH="50" SIZE=30 tag ="23"></TD>
								<TD CLASS=TD5 NOWRAP>세금신고사업장영문명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaEngNm" ALT="세금신고사업장영문명" MAXLENGTH="50" SIZE=30 tag ="23"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사업자등록번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOwnRgstNo" ALT="사업자등록번호" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag ="23"></TD>
								<TD CLASS=TD5 NOWRAP>대표자명</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepreNm" ALT="대표자명" MAXLENGTH="50" STYLE="TEXT-ALIGN:left" tag  ="23"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세무서</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxOfficeCd"   ALT="세무서코드" Size = "12" MAXLENGTH="10" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenTaxOffice(frm1.txtTaxOfficeCd.value, 1)">
													 <INPUT NAME="txtTaxOfficeNm" MAXLENGTH="25" SIZE = "25" tag="24X"></TD>
								<TD CLASS=TD5 NOWRAP>업태</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_class" ALT="업태" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="23" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCommonPopupInfo(frm1.txtInd_class.value,0)">
													 <INPUT NAME="txtInd_class_Nm" ALT="업태" SIZE="20" tag = "24" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>업종</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="업종" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="23" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCommonPopupInfo(frm1.txtInd_Type.value,1)">
													<INPUT NAME="txtInd_Type_Nm" ALT="업종" SIZE="20" tag = "24" ></TD>
								<TD CLASS=TD5 NOWRAP>FAX번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFaxNo" ALT="FAX번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="23" ></TD>
							</TR>
							<TR>
 							    <TD CLASS=TD5 NOWRAP>전화번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo" ALT="전화번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="23"></TD>
 							    <TD CLASS="TD5" NOWRAP>회계담당자</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCharge" ALT="회계담당자" MAXLENGTH="10" SIZE="20" STYLE="TEXT-ALIGN:left" tag="2" ></TD>
 							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>우편번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZipCode" ALT="우편번호" MAXLENGTH="12" Size="11" STYLE="TEXT-ALIGN:left" tag  ="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZipCode(frm1.txtZipCode.value, 1)"></TD>
							    <TD CLASS=TD5 NOWRAP>전산담당자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIsCharge" ALT="전산담당자" MAXLENGTH="10" SIZE="20" STYLE="TEXT-ALIGN:left" tag="2" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주소</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr1"  ALT="주소"     MAXLENGTH="100" SIZE="80" STYLE="TEXT-ALIGN:left" tag="23" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr2"  ALT="주소"     MAXLENGTH="100" SIZE="80" STYLE="TEXT-ALIGN:left" tag="25" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>영문주소</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtEng1Addr" ALT="영문주소" MAXLENGTH="50" Size="80" STYLE="TEXT-ALIGN: left" tag ="23"></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtEng2Addr" ALT="영문주소" MAXLENGTH="50" Size="80" STYLE="TEXT-ALIGN: left" tag ="25"></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtEng3Addr" ALT="영문주소" MAXLENGTH="50" Size="80" STYLE="TEXT-ALIGN: left" tag ="25"></TD>	
							</TR>
							<% Call SubFillRemBodyTd5656(2) %>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

