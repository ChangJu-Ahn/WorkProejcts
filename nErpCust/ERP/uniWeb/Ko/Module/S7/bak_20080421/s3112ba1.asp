<%@ LANGUAGE="VBSCRIPT" %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : s3112ba1
'*  4. Program Name         : 수주내역 참조 
'*  5. Program Desc         : 매출채권내역등록에서 수주내역 참조 Popup
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/07
'*  8. Modified date(Last)  : 2002/05/07
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s3112bb1.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 19                                           '☆: key count of SpreadSheet
Const C_PopItemCd		= 1

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim lgArrReturn												'☜: Return Parameter Group
Dim lgBlnOpenedFlag
Dim	lgBlnItemCdChg

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

top.document.title = PopupParent.gActivePRAspName

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
        
    lgBlnItemCdChg = False
End Function

'=======================================================
Sub SetDefaultVal()
	Dim arrRowSep, arrColValue

	With frm1
		arrRowSep = Split(arrParent(1), PopupParent.gRowSep)
		arrColValue = Split(arrRowSep(0),PopupParent.gColSep)

		.txtFromDt.Text = UNIDateClientFormat(UniConvDateAToB(UniConvDateToYYYYMM(arrColValue(8), PopupParent.gDateFormat, "-") & "-01", PopupParent.gServerDateFormat ,PopupParent.gAPDateFormat))
		.txtToDt.Text = arrColValue(8)	

		<% '수주번호 %>
		.txtSoNo.value			= arrColValue(0)
		<% '주문처 %>
		.txtSoldToParty.value	= arrColValue(1)
		.txtSoldToPartyNm.value = arrColValue(2)
		<% '영업그룹 %>
		.txtSalesGrpCd.value	= arrColValue(3)
		.txtSalesGrpNm.value	= arrColValue(4)
		<% '결제방법 %>
		.txtPayTermsCd.value	= arrColValue(5)
		.txtPayTermsNm.value	= arrColValue(6)
		<% '화폐 %>
		.txtHCurrency.value		= arrColValue(7)
		<% '매출채권일 %>
		.txtHBillDt.value		= arrColValue(8)
		.txtHVatRate.value		= arrColValue(9)
		.txtHVatType.value		= arrColValue(10)
		.txtHVatCalcType.value	= arrColValue(11)
		.txtHVatIncflag.value	= arrColValue(12)
		.txtHXchgRate.value		= arrColValue(13)
		.txtHXchgOp.value		= arrColValue(14)		
		<% '매출채권형태 %>
		.txtHBillTypeCd.value	= arrColValue(15)
			
	End With

	Redim lgArrReturn(0,0)
	Self.Returnvalue = lgArrReturn
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>	
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub

'========================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S3112BA1","S","A","V20030318",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
							C_MaxKey, "X","X")		
	Call SetSpreadLock 	 	
		    
End Sub

'========================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False
'		ggoSpread.SpreadLock 1 , -1
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.OperationMode = 5
		.ReDraw = True
    End With
End Sub	

'========================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow

	With frm1
		If .vspdData.SelModeSelCount > 0 Then 
			intInsRow = 0

			Redim lgArrReturn(.vspdData.SelModeSelCount, .vspdData.MaxCols)

			For intRowCnt = 1 To .vspdData.MaxRows

				.vspdData.Row = intRowCnt

				If .vspdData.SelModeSelected Then
					For intColCnt = 1 To .vspdData.MaxCols - 1
						.vspdData.Col = GetKeyPos("A", intColCnt)
						lgArrReturn(intInsRow, intColCnt - 1) = .vspdData.Text
					Next

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
	End With
		
	Self.Returnvalue = lgArrReturn
	Self.Close()
End Function

'========================================
Function CancelClick()
	Self.Close()
End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Select Case pvIntWhere

	Case C_PopItemCd
		iArrParam(1) = "b_item"										' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtItemCd.value)					' Code Condition
		iArrParam(3) = ""											' Name Cindition
		iArrParam(4) = "valid_flg = " & FilterVar("Y", "''", "S") & "  AND valid_to_dt >=  " & FilterVar(UNIConvDate(frm1.txtHBillDt.value), "''", "S") & " " & _
						" AND valid_from_dt <=  " & FilterVar(UNIConvDate(frm1.txtHBillDt.value), "''", "S") & " "	' Where Condition
		iArrParam(5) = Trim(frm1.txtItemCd.alt)						' TextBox 명칭 

		iArrField(0) = "ED15" & PopupParent.gColSep & "item_cd"		' Field명(0)
		iArrField(1) = "ED30" & PopupParent.gColSep & "item_nm"
		iArrField(2) = "ED30" & PopupParent.gColSep & "spec"

		iArrHeader(0) = "품목"									' Header명(0)
		iArrHeader(1) = "품목명"								' Header명(1)
		iArrHeader(2) = "규격"									' Header명(2)
		
		frm1.txtItemcd.focus
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
	End If	
	
End Function

'========================================
Function OpenTrackingNo()
	Dim iCalledAspName
	Dim iStrRet
	Dim iArrTNParam(5)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	iCalledAspName = AskPRAspName("s3135pa3")	
	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa3", "x")
		gblnWinEvent = False
		Exit Function
	End if

	With frm1
		If Len(.txtSoldToParty.value) Then
			iArrTNParam(0) = .txtSoldToParty.value
		End If
	
		If Len(.txtSalesGrpCd.value) Then
			iArrTNParam(1) = .txtSalesGrpCd.value
		End If

		If Len(.txtItemcd.value) Then
			iArrTNParam(3) = .txtItemcd.value
		End If
			
		If Len(.txtSONo.value) Then
			iArrTNParam(4) = .txtSONo.value
		End If

		iArrTNParam(5) = "BS"
	
		iStrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, iArrTNParam), _
			"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		.txtTrackingNo.focus

		If iStrRet <> "" Then
			.txtTrackingNo.value = iStrRet 
		End If		
	End With
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=======================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopItemCd
		frm1.txtItemCd.value = pvArrRet(0) 
		frm1.txtItemNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
   
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	lgBlnOpenedFlag = True
	DbQuery()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================
Function txtItemCd_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnItemCdChg = True
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'====================================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnItemCdChg Then
		iStrCode = Trim(frm1.txtItemCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("IT", "''", "S") & "", C_PopItemCd) Then
				Call DisplayMsgBox("970000", "X", frm1.txtItemCd.alt, "X")
				frm1.txtItemCd.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtItemNm.value = ""
		End If
		lgBlnItemCdChg	= False
	End If

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
		' 화면 Open전에는 코드 Popup을 Display하지 않는다.
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'========================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then Call DbQuery
	End If
End Sub

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("P")
		frm1.txtFromDt.Focus
	End If
End Sub

'========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToDt.Focus
	End If
End Sub

'========================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtFromDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtHBillDt.value, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, .txtHBillDt.alt & "(" & .txtHBillDt.value & ")")
			.txtFromDt.focus	
			Exit Function
		End If
		If UniConvDateToYYYYMMDD(.txtToDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtHBillDt.value, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, .txtHBillDt.alt & "(" & .txtHBillDt.value & ")")
			.txtToDt.Focus
			Exit Function
		End If
	End With
   
    Call ggoOper.ClearField(Document, "2")

	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If

    Call InitVariables
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtHItemCd.Value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtHTrackingNo.value)
		Else
			' 처음 조회시 
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)				<%'☆: 조회 조건 데이타 %>
			If Len(Trim(.txtToDt.text)) Then
				strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			Else
				strVal = strVal & "&txtToDt=" & Trim(.txtHBillDt.value)
			End if
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		End If
		strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
		strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
		strVal = strVal & "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value)
		strVal = strVal & "&txtPayTermsCd=" & Trim(.txtPayTermsCd.value)
		strVal = strVal & "&txtCurrency=" & Trim(.txtHCurrency.value)
		strVal = strVal & "&txtBillTypeCd=" & Trim(.txtHBillTypeCd.value)
		
		strVal = strVal & "&txtVatRate=" & Trim(.txtHVatRate.value)
		strVal = strVal & "&txtVatType=" & Trim(.txtHVatType.value)
		strVal = strVal & "&txtVatCalcType=" & Trim(.txtHVatCalcType.value)
		strVal = strVal & "&txtXchgRate=" & Trim(.txtHXchgRate.value)
		strVal = strVal & "&txtXchgOp=" & Trim(.txtHXchgOp.value)
		strVal = strVal & "&txtVatIncflag=" & Trim(.txtHVatIncFlag.value)

        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With    
    
	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	With frm1
		If .vspdData.MaxRows > 0 Then
			If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
				lgIntFlgMode = PopupParent.OPMD_UMODE
				.vspdData.Row = 1	
				.vspdData.SelModeSelected = True
			End If
			.vspdData.Focus
		Else
			Call SetFocusToDocument("P")
			.txtFromDt.Focus
		End If
	End With

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>수주일</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s3112ba1_fpDateTime1_txtFromDt.js'></script>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s3112ba1_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6"><INPUT NAME="txtItemcd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopItemCd">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" SIZE=20 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6><INPUT NAME="txtTrackingNo" ALT="Tracking 번호" TYPE=TEXT MAXLENGTH=25 SIZE=30 TAG="11XXXU" TABINDEX=-1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTrackingNo()"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldToParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="14XXXU">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="14XXXU">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>결제방법</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="14XXXU">&nbsp;<INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" SIZE=30 MAXLENGTH=18 TAG="14XXXU"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3112ba1_vaSpread_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											  <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG>			</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTrackingNo" TAG="24">

<INPUT TYPE=HIDDEN NAME="txtHBillTypeCd" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHBillDt" tag="14" alt="매출채권일">
<INPUT TYPE=HIDDEN NAME="txtHVatRate" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHVatType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHVatCalcType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHXchgRate" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHXchgOp" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHVatIncflag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" tag="14">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
