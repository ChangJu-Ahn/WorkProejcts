<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales																		*
'*  2. Function Name        : 판매경비관리																*
'*  3. Program ID           : S6111MA2																	*
'*  4. Program Name         : 판매경비일괄처리															*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : PS9G115.dll, PS9G241.dll
'*  7. Modified date(First) : 2000/04/26																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Cho Sung Hyun																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/26 : 화면 design												*
'*							  2. 2000/09/22 : 4th Coding Start											*
'*							  3. 2001/12/19 : Date 표준적용												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBS">
Option Explicit					<% '☜: indicates that All variables must be declared in advance %>

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s6111mb2.asp"			<% '☆: 비지니스 로직 ASP명 %>
Const gstrProcessStepMajor = "S9014"

Dim C_Select	
Dim C_PostFlg	
Dim C_ChargeNo	
Dim C_ChargeCd		
Dim C_ProcessStep	
Dim C_ProcessStepNm
Dim C_BASNo		
Dim C_SalesGroups
Dim C_BpCd		
Dim C_ChargeDt	
Dim C_VATType	
Dim C_Cur		
Dim C_ChargeDocAmt
Dim C_XchRate	
Dim C_ChargeLocAmt
Dim C_VATRate	
Dim C_VATAmt	
Dim C_VATLocAmt	
Dim C_CostFlg	
Dim C_PayType	
Dim C_CheckNo	
Dim C_PayAccount
Dim C_PayBank
Dim C_Remark
Dim C_ChgFlg
	
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim gblnWinEvent					'~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
									'	PopUp Window가 사용중인지 여부를 나타내는 variable
'========================================================================================================
Sub initSpreadPosVariables()  

	C_Select			= 1
	C_PostFlg			= 2
	C_ChargeNo			= 3								<% '☆: Spread Sheet 의 Columns 인덱스 %>
	C_ChargeCd			= 4
	C_ProcessStep		= 5
	C_ProcessStepNm		= 6
	C_BASNo				= 7
	C_SalesGroups		= 8
	C_BpCd				= 9
	C_ChargeDt			= 10
	C_VATType			= 11
	C_Cur				= 12
	C_ChargeDocAmt		= 13	
	C_XchRate			= 14
	C_ChargeLocAmt		= 15
	C_VATRate			= 16
	C_VATAmt			= 17
	C_VATLocAmt			= 18
	C_CostFlg			= 19
	C_PayType			= 20
	C_CheckNo			= 21
	C_PayAccount		= 22
	C_PayBank			= 23
	C_Remark			= 24
	C_ChgFlg			= 25
	
End Sub
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE								<%'⊙: Indicates that current mode is Create mode%>
	lgBlnFlgChgValue = False								<%'⊙: Indicates that no value changed%>
	lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
	lgStrPrevKey = ""										<%'initializes Previous Key%>
	lgLngCurRows = 0 										<%'initializes Deleted Rows Count%>
		
	<% '------ Coding part ------ %>
	gblnWinEvent = False
	Call BtnDisabled(1)
End Function

'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate
	frm1.txtCharge.focus
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
    With frm1
    
		ggoSpread.Source = .vspdData
			
		.vspdData.MaxCols = C_ChgFlg
		.vspdData.MaxRows = 0
			
		.vspdData.ReDraw = False
			
		ggoSpread.Spreadinit "V20030301",,parent.gAllowDragDropSpread    

		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetCheck	C_Select,     "", 5,,,true
		ggoSpread.SSSetCheck	C_PostFlg,    "확정여부", 15,,,true
		ggoSpread.SSSetEdit		C_ChargeNo,   "경비관리번호", 15,,,,2
		ggoSpread.SSSetEdit		C_ChargeCd,   "경비항목", 12,,,,2
		ggoSpread.SSSetEdit		C_ProcessStep,"진행구분", 12,,,,2
		ggoSpread.SSSetEdit		C_ProcessStepNm,"진행구분명", 12
		ggoSpread.SSSetEdit		C_BASNo,      "발생근거번호", 18,,,,2
		ggoSpread.SSSetEdit		C_SalesGroups,"영업그룹", 12,,,,2
		ggoSpread.SSSetEdit		C_BpCd,       "거래처", 20
		ggoSpread.SSSetDate		C_ChargeDt,   "발생일",12,2,parent.gDateFormat
		ggoSpread.SSSetEdit		C_VATType,    "계산서종류", 12,,,,2
		ggoSpread.SSSetEdit		C_Cur,        "화폐", 10,,,,2
		ggoSpread.SSSetFloat    C_ChargeDocAmt,"발생금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat    C_XchRate,    "환율",15,parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat    C_ChargeLocAmt,"자국금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat    C_VATRate,     "VAT율" ,15,parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat    C_VATAmt,      "VAT금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat    C_VATLocAmt,   "VAT자국금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_CostFlg,     "물대포함", 12
		ggoSpread.SSSetEdit		C_PayType,     "지불유형", 12,,,,2
		ggoSpread.SSSetEdit		C_CheckNo,     "수표번호", 12
		ggoSpread.SSSetEdit		C_PayAccount,  "출금계좌", 12
		ggoSpread.SSSetEdit		C_PayBank,     "출금은행", 12
		ggoSpread.SSSetEdit		C_Remark,      "기타참조사항", 20
		ggoSpread.SSSetEdit		C_ChgFlg,      "Chgfg", 1, 2
			

		Call ggoSpread.SSSetColHidden(C_CostFlg,C_CostFlg,True)
		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
			
		SetSpreadLock "", 0, -1, ""

		.vspdData.ReDraw = True
	End With
End Sub

'========================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
		ggoSpread.Source = .vspdData
			
		ggoSpread.SpreadLock C_PostFlg, lRow, -1
		ggoSpread.SpreadLock C_ChargeNo, lRow, -1
		ggoSpread.SpreadLock C_ChargeCd, lRow, -1
		ggoSpread.SpreadLock C_ProcessStep, lRow, -1
		ggoSpread.SpreadLock C_ProcessStepNm, lRow, -1
		ggoSpread.SpreadLock C_BASNo, lRow, -1
		ggoSpread.SpreadLock C_SalesGroups, lRow, -1
		ggoSpread.SpreadLock C_BpCd, lRow, -1
		ggoSpread.SpreadLock C_ChargeDt, lRow, -1
		ggoSpread.SpreadLock C_VATType, lRow, -1
		ggoSpread.SpreadLock C_Cur, lRow, -1
		ggoSpread.SpreadLock C_ChargeDocAmt, lRow, -1
		ggoSpread.SpreadLock C_XchRate, lRow, -1
		ggoSpread.SpreadLock C_ChargeLocAmt, lRow, -1
		ggoSpread.SpreadLock C_VATRate, lRow, -1
		ggoSpread.SpreadLock C_VATAmt, lRow, -1
		ggoSpread.SpreadLock C_VATLocAmt, lRow, -1
		ggoSpread.SpreadLock C_CostFlg, lRow, -1
		ggoSpread.SpreadLock C_PayType, lRow, -1
		ggoSpread.SpreadLock C_CheckNo, lRow, -1
		ggoSpread.SpreadLock C_PayAccount, lRow, -1
		ggoSpread.SpreadLock C_PayBank, lRow, -1
		ggoSpread.SpreadLock C_Remark, lRow, -1
			
	End With
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
	    
		.Redraw = False

		ggoSpread.SSSetProtected C_PostFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChargeNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChargeCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProcessStep, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProcessStepNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BASNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SalesGroups, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BpCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChargeDt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Cur, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChargeDocAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_XchRate, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChargeLocAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATRate, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATLocAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CostFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PayType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CheckNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PayAccount, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PayBank, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Remark, pvStartRow, pvEndRow
		.Col = 1
		.Row = .ActiveRow
		.Action = 0
		.EditMode = True

		.ReDraw = True
	End With
End Sub

'========================================================================================================
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "거래처"							<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPayCharge.value)			<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "거래처"							<%' TextBox 명칭 %>

	arrField(0) = "BP_CD"								<%' Field명(0)%>
	arrField(1) = "BP_NM"								<%' Field명(1)%>

	arrHeader(0) = "거래처"							<%' Header명(0)%>
	arrHeader(1) = "거래처명"							<%' Header명(1)%>

	frm1.txtPayCharge.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet)
	End If
End Function

'========================================================================================================
Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos								<%' 팝업 명칭 %>
	arrParam(1) = "B_Minor"								<%' TABLE 명칭 %>
	arrParam(2) = Trim(strMinorCD)						<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		<%' Where Condition%>
	arrParam(5) = strPopPos								<%' TextBox 명칭 %>

	arrField(0) = "Minor_CD"							<%' Field명(0)%>
	arrField(1) = "Minor_NM"							<%' Field명(1)%>

	arrHeader(0) = strPopPos							<%' Header명(0)%>
	arrHeader(1) = strPopPos & "명"				<%' Header명(1)%>

	frm1.txtProcessStep.focus 
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCd(strMajorCd, arrRet)
	End If
End Function

'========================================================================================================
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "영업그룹"						<%' 팝업 명칭 %>
	arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtSalesGrp.value)			<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
	arrParam(5) = "영업그룹"						<%' TextBox 명칭 %>

	arrField(0) = "SALES_GRP"							<%' Field명(0)%>
	arrField(1) = "SALES_GRP_NM"						<%' Field명(1)%>

	arrHeader(0) = "영업그룹"						<%' Header명(0)%>
	arrHeader(1) = "영업그룹명"						<%' Header명(1)%>

	frm1.txtSalesGrp.focus 
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
End Function

'========================================================================================================
Function OpenSalesCharge()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol,TempCd

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "경비항목"
	arrParam(1) = "A_JNL_ITEM"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtCharge.value)			<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = "JNL_TYPE = " & FilterVar("EC", "''", "S") & ""						<%' Where Condition%>
	arrParam(5) = "경비항목"						<%' TextBox 명칭 %>

	arrField(0) = "JNL_CD"								<%' Field명(0)%>
	arrField(1) = "JNL_NM"								<%' Field명(1)%>

	arrHeader(0) = "경비항목"						<%' Header명(0)%>
	arrHeader(1) = "경비항목명"						<%' Header명(1)%>

	frm1.txtCharge.focus 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesCharge(arrRet)
	End If	
	
End Function

'========================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtPayCharge.Value = arrRet(0)
	frm1.txtPayChargeNm.Value = arrRet(1)
End Function

'========================================================================================================
Function SetMinorCd(strMajorCd, arrRet)
	Select Case strMajorCd
		Case gstrProcessStepMajor
			frm1.txtProcessStep.Value = arrRet(0)
			frm1.txtProcessStepNm.Value = arrRet(1)
		Case Else
	End Select
End Function

'========================================================================================================
Function SetSalesGroup(arrRet)
	frm1.txtSalesGrp.value = arrRet(0)
	frm1.txtSalesGrpNm.value = arrRet(1)
End Function
'========================================================================================================
Function SetSalesCharge(arrRet)
	frm1.txtCharge.value = arrRet(0)
	frm1.txtChargeNm.value = arrRet(1)
End Function
	
'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtBLNo.value

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
				
		If strTemp = "" then Exit Function
				
		frm1.txtBLNo.value =  strTemp
			
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
			
		Call FncQuery()
						
		WriteCookie CookieSplit , ""
			
	End If

End Function
	
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_Select			= iCurColumnPos(1)
			C_PostFlg			= iCurColumnPos(2)
			C_ChargeNo			= iCurColumnPos(3)
			C_ChargeCd			= iCurColumnPos(4)
			C_ProcessStep		= iCurColumnPos(5)
			C_ProcessStepNm		= iCurColumnPos(6)
			C_BASNo				= iCurColumnPos(7)
			C_SalesGroups		= iCurColumnPos(8)
			C_BpCd				= iCurColumnPos(9)
			C_ChargeDt			= iCurColumnPos(10)
			C_VATType			= iCurColumnPos(11)
			C_Cur				= iCurColumnPos(12)
			C_ChargeDocAmt		= iCurColumnPos(13)
			C_XchRate			= iCurColumnPos(14)
			C_ChargeLocAmt		= iCurColumnPos(15)
			C_VATRate			= iCurColumnPos(16)
			C_VATAmt			= iCurColumnPos(17)
			C_VATLocAmt			= iCurColumnPos(18)
			C_CostFlg			= iCurColumnPos(19)
			C_PayType			= iCurColumnPos(20)
			C_CheckNo			= iCurColumnPos(21)
			C_PayAccount		= iCurColumnPos(22)
			C_PayBank			= iCurColumnPos(23)
			C_Remark			= iCurColumnPos(24)
			C_ChgFlg			= iCurColumnPos(25)
    End Select    
End Sub

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029																<% '⊙: Load table , B_numeric_format %>
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											<% '⊙: Lock  Suitable  Field %>

	Call InitSpreadSheet															<% '⊙: Setup the Spread sheet %>
	Call SetDefaultVal
	Call CookiePage(0)	
	Call InitVariables
	<% '----------  Coding part  ------------------------------------------------------------- %>

	Call SetToolbar("11000000000011")												<% '⊙: 버튼 툴바 제어 %>
		
End Sub
	
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
Sub btnPayChargeOnClick()
	Call OpenBizPartner()
End Sub
'========================================================================================================
Sub btnChargeOnClick()
	Call OpenSalesCharge()
End Sub
'========================================================================================================
Sub btnProcessStepOnClick()
	Call OpenMinorCd(frm1.txtProcessStep.value, frm1.txtProcessStepNm.value, "진행구분", gstrProcessStepMajor)
End Sub
'========================================================================================================
Sub btnSalesGrpOnClick()
	Call OpenSalesGroup()
End Sub
'========================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7 
		Call SetFocusToDocument("M")   
		Frm1.txtFromDt.Focus
    End If
End Sub

Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtToDt.Focus
    End If
End Sub

'========================================================================================================
Sub txtFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	If lgIntFlgMode = parent.OPMD_CMODE Then Exit Sub

	If Col = C_Select And Row > 0 Then
	    Select Case ButtonDown
	    Case 0
			ggoSpread.Source = frm1.vspdData
			ggoSpread.EditUndo
			lgBlnFlgChgValue = False
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True
	    End Select
    End If

End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )

	Exit Sub 

	frm1.vspdData.Row = Row
	If Col = C_Select Then
		frm1.vspdData.Col = 0
		If  Trim(frm1.vspdData.Text) = "" Then
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
		Else
			ggoSpread.Source = frm1.vspdData
			ggoSpread.EditUndo Row
			lgBlnFlgChgValue = False					
		End If	
	End If
End Sub
	
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				DbQuery
			End If
		End If
	End With
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_Select Or NewCol <= C_Select Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If NewTop > oldTop Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
	End If
End Sub

'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													<% '⊙: Processing is NG %>

	Err.Clear															<% '☜: Protect system from crashing %>

	<% '------ Check previous data area ------ %>
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			<% '⊙: "Will you destory previous data" %>
'			IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	<% '------ Erase contents area ------ %>
	Call ggoOper.ClearField(Document, "2")								<% '⊙: Clear Contents  Field %>
	Call InitVariables													<% '⊙: Initializes local global variables %>


	Call ggoOper.LockField(Document, "N")									<%'⊙: Lock  Suitable  Field%>
	Call SetToolbar("11000000000011")										<% '⊙: 버튼 툴바 제어 %>

	<% '------ Query function call area ------ %>
	Call DbQuery()														<% '☜: Query db data %>

	FncQuery = True														<% '⊙: Processing is OK %>
End Function
	
'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False                                                          <%'⊙: Processing is NG%>

	<% '------ Check previous data area ------ %>
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
'			IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	<% '------ Erase condition area ------ %>
	<% '------ Erase contents area ------ %>
	Call ggoOper.ClearField(Document, "A")									<%'⊙: Clear Condition,contents Field%>
	Call ggoOper.LockField(Document, "N")									<%'⊙: Lock  Suitable  Field%>
	Call InitVariables														<%'⊙: Initializes local global variables%>
	Call SetToolbar("11000000000011")										<% '⊙: 버튼 툴바 제어 %>
	Call SetDefaultVal

	Set gActiveElement = document.ActiveElement   
		
	FncNew = True															<%'⊙: Processing is OK%>

End Function
	
'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False												<% '⊙: Processing is NG %>
		
	<% '------ Precheck area ------ %>
	If lgIntFlgMode <> parent.OPMD_UMODE Then								<% 'Check if there is retrived data %>
		Call DisplayMsgBox("900002", "X", "X", "X")
'			Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	<% '------ Delete function call area ------ %>
	Call DbDelete													<% '☜: Delete db data %>

	FncDelete = True												<% '⊙: Processing is OK %>
End Function
	
'========================================================================================================
Function FncSave()
	Dim IntRetCD
	Dim lRow
		
	FncSave = False																		<% '⊙: Processing is NG %>
		
	Err.Clear																			<% '☜: Protect system from crashing %>
		
	lgBlnFlgChgValue = False

	For lRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 1

		If frm1.vspdData.Text = 1 Then
			lgBlnFlgChgValue = True
		End If
	Next
		
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                   <%'No data changed!!%>
	    Exit Function
	End If
			
	<% '------ Check contents area ------ %>
	ggoSpread.Source = frm1.vspdData

	If Not chkField(Document, "2") Then								<% '⊙: Check contents area %>
		Exit Function
	End If
		
	If ggoSpread.SSDefaultCheck = False Then
		Exit Function
	End If
		
	<% '------ Save function call area ------ %>
	Call DbSave																			<% '☜: Save db data %>
		
	FncSave = True																		<% '⊙: Processing is OK %>
End Function

'========================================================================================================
Function FncCopy()

	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False     
    
	If frm1.vspdData.Maxrows < 1 Then Exit Function

	frm1.vspdData.ReDraw = False

	ggoSpread.Source = frm1.vspdData	
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True

'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If Err.number = 0 Then	
	   FncCopy = True                                                            '☜: Processing is OK
	End If
		
	Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
Function FncCancel() 

	If frm1.vspdData.Maxrows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo														<%'☜: Protect system from crashing%>

End Function

'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'========================================================================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    
	Call ggoSpread.ReOrderingSpreadData()	

    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Cur,C_ChargeDocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData, -1, -1 ,Parent.gCurrency,C_ChargeLocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Cur,C_VATAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData, -1, -1 ,Parent.gCurrency,C_VatLocAmt,"A","I","X","X")

End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			<%'⊙: "Will you destory previous data"%>

'			IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>
		
	If frm1.rdoPostingflg1.checked Then
		frm1.txtRadio.value = ""
	ElseIf frm1.rdoPostingflg2.checked Then
		frm1.txtRadio.value = "Y"
	ElseIf frm1.rdoPostingflg3.checked Then
		frm1.txtRadio.value = "N"
	End If		 
		
	Dim strVal

	Call LayerShowHide(1)

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtCharge=" & Trim(frm1.txtHCharge.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtHSalesGrp.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtPayCharge=" & Trim(frm1.txtHPayCharge.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtFromDt=" & Trim(frm1.txtHFromDt.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtToDt=" & Trim(frm1.txtHToDt.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtProcessStep=" & Trim(frm1.txtHProcessStep.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtHRadio.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtCharge=" & Trim(frm1.txtCharge.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtSalesGrp.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtPayCharge=" & Trim(frm1.txtPayCharge.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromDt.text)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtToDt=" & Trim(frm1.txtToDt.text)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtProcessStep=" & Trim(frm1.txtProcessStep.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtRadio.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If

	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
	
	DbQuery = True														<%'⊙: Processing is NG%>
End Function
	
'========================================================================================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDel
	Dim intInsrtCnt

	DbSave = False														<% '⊙: Processing is OK %>
    
	Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID

		lGrpCnt = 1

		strVal = ""
		intInsrtCnt = 1

		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 1

			If .vspdData.Text = 1 Then
					strVal = strVal & lRow & parent.gColSep			<% '☜: C=Create, Row위치 정보 %>

					.vspdData.Col = C_ChargeNo								<% '2 %>
					strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
						
					lGrpCnt = lGrpCnt + 1
					intInsrtCnt = intInsrtCnt + 1
			End If
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		If Len(strVal) Then
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)						<% '☜: 비지니스 ASP 를 가동 %>
		Else Exit Function
		End If
	End With

	DbSave = True														<% '⊙: Processing is NG %>
End Function
	
'========================================================================================================
Function DbDelete()
End Function
	
'========================================================================================================
Function DbQueryOk()													<% '☆: 조회 성공후 실행로직 %>
	<% '------ Reset variables area ------ %>
	lgIntFlgMode = parent.OPMD_UMODE
	lgBlnFlgChgValue = False											<% '⊙: Indicates that current mode is Update mode %>

	Call ggoOper.LockField(Document, "Q")								<% '⊙: This function lock the suitable field %>
	Call SetToolbar("11101000000111")												<% '⊙: 버튼 툴바 제어 %>

	If frm1.vspdData.MaxRows > 0 Then
        frm1.vspdData.Focus		
	Else
		frm1.txtCharge.focus
    End If     

End Function
	
'========================================================================================================
Function DbSaveOk()			<%'☆: 저장 성공후 실행 로직 %>
	Call ggoOper.ClearField(Document, "2")
	Call InitVariables
	Call FncQuery()
End Function
	
'========================================================================================================
Function DbDeleteOk()													<%'☆: 삭제 성공후 실행 로직 %>
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>경비일괄처리</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>경비항목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCharge" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="경비항목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharge" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnChargeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtChargeNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnSalesGrpOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayCharge" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayCharge" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnPayChargeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPayChargeNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>발생일</TD>						
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s6111ma2_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s6111ma2_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>진행구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProcessStep" SIZE=10 MAXLENGTH=5 TAG="11XXXU" ALT="진행구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessStep" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnProcessStepOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtProcessStepNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="11X" VALUE="A" CHECKED ID="rdoPostingflg1">
										<LABEL FOR="rdoPostingflg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="11X" VALUE="Y" ID="rdoPostingflg2">
										<LABEL FOR="rdoPostingflg2">확정</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="11X" VALUE="N" ID="rdoPostingflg3">
										<LABEL FOR="rdoPostingflg3">미확정</LABEL>
									</TD>
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
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/s6111ma2_vaSpread_vspdData.js'></script>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="Hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHCharge" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayCharge" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24"> 
<INPUT TYPE=HIDDEN NAME="txtHProcessStep" TAG="24"> 
<INPUT TYPE=HIDDEN NAME="txtHRadio" TAG="24"> 
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="24"> 
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
