<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4111MA1
'*  4. Program Name         : 검사결과등록 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID ="q4111mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "q4111mb2.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "q4111mb3.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_CONFIRM_RELEASE_ID = "q4111mb4.asp"
Const BIZ_PGM_CANCEL_RELEASE_ID = "q4111mb5.asp"
'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - START */
Const BIZ_PGM_LOOKUP_ID ="q4111mb6.asp"							
'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - END */
Const BIZ_PGM_JUMP1_ID = "Q4112MA1"
Const BIZ_PGM_JUMP2_ID = "Q4113MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop

Dim lgstatusflag	'검사진행상태 코드값 저장 변수 
Dim lgInspClassCd	'검사분류코드 저장 변수 
Dim lgReceivingInspType		'수입검사유형(입고전검사:B, 입고후검사:A)

Dim lgAutoPR		'자동입고여부(예:Y, 아니오:Y가 아니면 모두)
Dim lgAutoST		'자동재고이동여부(예:Y, 아니오:Y가 아니면 모두)
Dim lgIFYesNo		'자체검사의뢰 여부(아니오:N, 예:N이 아니면 모두)

Dim lgReleaseBtnFlag
Dim strtxtLotSize

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMD	'#####
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)                                           '☆: 초기화면에 뿌려지는 시작 날짜 -----
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE                       				'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                	              		'⊙: Indicates that no value changed
    lgIntGrpCount = 0                                               '⊙: Initializes Group View Size
    
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'☆: 사용자 변수 초기화 
    lgstatusflag = ""
    lgInspClassCd = ""
	lgReceivingInspType = ""
	lgAutoPR = ""
	lgAutoST = ""
	lgIFYesNo = ""
	
	lgReleaseBtnFlag = "R"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
		
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If
		
	If ReadCookie("txtInspReqNo") <> "" Then
		frm1.txtInspReqNo1.Value = ReadCookie("txtInspReqNo")
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""	
	
'	frm1.txtInspDt.Text = CompanyYMD
	frm1.cboDecision.value = "A"
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0010", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboDecision , lgF0, lgF1, Chr(11))
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Sub OpenPlant() 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Sub

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "공장코드"		
	arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)		
		frm1.txtPlantNm.Value = arrRet(1)
	End If	

	frm1.txtPlantCd.Focus	
	Set gActiveElement = document.activeElement
End Sub

'------------------------------------------  OpenInspReqNo1() -------------------------------------------------
'	Name : OpenInspReqNo1()
'	Description : InspReqNo1 PopUp
'--------------------------------------------------------------------------------------------------------- 
Sub OpenInspReqNo1()       
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Sub
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Sub	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo1.Value)
	Param4 = ""			'검사분류 
	Param5 = ""			'판정 
		
	iCalledAspName = AskPRAspName("Q4111PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "Q4111PA1", "X")
		IsOpenPop = False
		Exit Sub
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtInspReqNo1.value = arrRet(0)
		Call ChangingFieldByInspClass(arrRet(30))
	End If	

	frm1.txtInspReqNo1.Focus	
	Set gActiveElement = document.activeElement	
End Sub

'------------------------------------------  OpenInspReqNo2()  -------------------------------------------------
'	Name : OpenInspReqNo2()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Sub OpenInspReqNo2()        
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Sub
	
	If UCase(frm1.txtInspReqNo1.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Sub
	End If
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Sub	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo2.Value)
	Param4 = ""				'검사분류 
	Param5 = "N"			'검사진행현황(미검사)
	
	iCalledAspName = AskPRAspName("q2512pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q2512pa1", "X")
		IsOpenPop = False
		Exit Sub
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	Set gActiveElement = document.activeElement
	
	If arrRet(0) <> "" Then
		Call SetInspReqNo2(arrRet)
	End If

	frm1.txtInspReqNo2.Focus
	Set gActiveElement = document.activeElement	
End Sub

'------------------------------------------  OpenInspector() -------------------------------------------------
'	Name : OpenInspector()
'	Description :Inspector PopUp
'--------------------------------------------------------------------------------------------------------- 
Sub OpenInspector()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Sub
	
	IsOpenPop = True

	arrParam(0) = "검사원팝업"	
	arrParam(1) = "B_Minor"				
	arrParam(2) = Trim(frm1.txtInspectorCd.Value)
	arrParam(3) = ""
	arrParam(4) = "Major_Cd = " & FilterVar("Q0002", "''", "S") & " "      ' Where Condition
	arrParam(5) = "검사원"			

    arrField(0) = "Minor_CD"	
    arrField(1) = "Minor_NM"	

    arrHeader(0) = "검사원코드"		
    arrHeader(1) = "검사원명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtInspectorCd.Value = arrRet(0)		
		frm1.txtInspectorNm.Value = arrRet(1)		
		lgBlnFlgChgValue = True
	End If	
	
	frm1.txtInspectorCd.Focus
	Set gActiveElement = document.activeElement
End Sub

'------------------------------------------  SetInspReqNo2()  --------------------------------------------------
'	Name : SetInspReqNo2()
'	Description : OpenInspReqNo2 Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Sub SetInspReqNo2(Byval arrRet)
	
	Dim blnRet
	
	With frm1
		.txtInspReqNo2.value = arrRet(0)
		.hInspReqNo2.value = arrRet(0)
		
		.txtInspClass.value = arrRet(1)
		.txtItemCd.value = arrRet(2)
		.txtItemNm.value = arrRet(3)
		.txtSpec.value = arrRet(4)
		
		Call ChangingFieldByInspClass(arrRet(36))
		
		Select Case UCase(Trim(arrRet(36)))
			Case "R"
				.txtSupplierCd.value = arrRet(5)
				.txtSupplierNm.value = arrRet(6)
				.txtSLCd1.value = arrRet(32)
				.txtSLNm1.value = arrRet(33)
				.txtRoutNo.value = ""
				.txtRoutNoDesc.value = ""
				.txtOprNo.value = ""
				.txtOprNoDesc.value = ""
				.txtWcCd.value = ""
				.txtWcNm.value = ""
				.txtSLCd2.value = ""
				.txtSLNm2.value = ""
				.txtBpCd.value = ""
				.txtBpNm.value = ""
			Case "P"
				.txtSupplierCd.value = ""
				.txtSupplierNm.value = ""
				.txtSLCd1.value = ""
				.txtSLNm1.value = ""
				.txtRoutNo.value = arrRet(7)
				.txtRoutNoDesc.value = arrRet(8)
				.txtOprNo.value = arrRet(9)
				.txtOprNoDesc.value = arrRet(10)
				.txtWcCd.value = arrRet(11)
				.txtWcNm.value = arrRet(12)
				.txtSLCd2.value = ""
				.txtSLNm2.value = ""
				.txtBpCd.value = ""
				.txtBpNm.value = ""
			Case "F"
				.txtSupplierCd.value = ""
				.txtSupplierNm.value = ""
				.txtSLCd1.value = ""
				.txtSLNm1.value = ""
				.txtRoutNo.value = ""
				.txtRoutNoDesc.value = ""
				.txtOprNo.value = ""
				.txtOprNoDesc.value = ""
				.txtWcCd.value = ""
				.txtWcNm.value = ""
				.txtSLCd2.value = arrRet(32)
				.txtSLNm2.value = arrRet(33)
				.txtBpCd.value = ""
				.txtBpNm.value = ""
			Case "S"
				.txtSupplierCd.value = ""
				.txtSupplierNm.value = ""
				.txtSLCd1.value = ""
				.txtSLNm1.value = ""
				.txtRoutNo.value = ""
				.txtRoutNoDesc.value = ""
				.txtOprNo.value = ""
				.txtOprNoDesc.value = ""
				.txtWcCd.value = ""
				.txtWcNm.value = ""
				.txtSLCd2.value = ""
				.txtSLNm2.value = ""
				.txtBpCd.value = arrRet(13)
				.txtBpNm.value = arrRet(14)
		End Select
		.txtLotNo.value = arrRet(16)
		.txtLotSubNo.value = arrRet(17)
		.txtLotSize.Text = arrRet(18)
		.txtUnit.value = arrRet(19)
		.txtInspReqDt.Text = arrRet(20)
		'환경설정의 검사일에 대한 기본표시값 확인 후 처리: 현재일, 검사의뢰일 
		blnRet =  CommonQueryRs2by2(" BASIC_MARK_FOR_INSP_DT ", " Q_CONFIGURATION ", " PLANT_CD =  " & FilterVar(.txtPlantCd.value, "''", "S") & " ", lgF0)
		If blnRet = False Then
			.txtInspDt.Text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
		Else
			lgF0 = Split(lgF0, Chr(11))
		
			If Trim(lgF0(1)) = "1" Then '현재일 
				.txtInspDt.Text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
			Else
				.txtInspDt.Text = arrRet(20)
			End If
		End If
	End With
    
    lgBlnFlgChgValue = True
End Sub

'------------------------------------------  SetInspector()  --------------------------------------------------
'	Name : SetInspector()
'	Description : OpenInspector Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Sub SetInspector(byval arrRet)
	frm1.txtInspectorCd.Value = arrRet(0)		
	frm1.txtInspectorNm.Value = arrRet(1)		
	lgBlnFlgChgValue = True
End Sub

'=============================================  2.5.1 LoadIDispositionOfNCM()======================================
'=	Event Name : LoadIDispositionOfNCM
'=	Event Desc :
'========================================================================================================
Function LoadIDispositionOfNCM()
	Dim intRetCD
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo1.value)
	End With	
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.1 LoadNoticeOfRejection()======================================
'=	Event Name : LoadNoticeOfRejection
'=	Event Desc :
'========================================================================================================
Function LoadNoticeOfRejection()
	Dim intRetCD
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo1.value)
	End With	
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=============================================  2.6.1 ChangingFieldByInspClass()======================================
'=	Sub Name : ChangingFieldByInspClass
'=	Sub Desc : 검사분류별 Field 변경(공급처, 작업장, 거래처)
'========================================================================================================
Sub ChangingFieldByInspClass(Byval sInspClass)

	Select Case sInspClass
		Case "R"
			Receiving.style.display = ""
			Process1.style.display = "none"
			Process2.style.display = "none"
			Final.style.display = "none"
			Shipping.style.display = "none"
			
		Case "P"
			Receiving.style.display = "none"
			Process1.style.display = ""
			Process2.style.display = ""
			Final.style.display = "none"
			Shipping.style.display = "none"
			
		Case "F"
			Receiving.style.display = "none"
			Process1.style.display = "none"
			Process2.style.display = "none"
			Final.style.display = ""
			Shipping.style.display = "none"
			
		Case "S"
			Receiving.style.display = "none"
			Process1.style.display = "none"
			Process2.style.display = "none"
			Final.style.display = "none"
			Shipping.style.display = ""
			
		Case Else
			Receiving.style.display = "none"
			Process1.style.display = "none"
			Process2.style.display = "none"
			Final.style.display = "none"
			Shipping.style.display = "none"
			
	End Select 
End Sub

'=============================================  2.6.2 ChangingReleaseBtn()======================================
'=	Sub Name : ChangingReleaseBtn
'=	Sub Desc : Release 버튼 캡션변경 및 Protect 처리 
'========================================================================================================
Sub ChangingReleaseBtn(Byval strInspStatus)
	With frm1
		Select Case strInspStatus
			Case "D"
				.btnRelease.value = "Release"
				.btnRelease.disabled = False
				lgReleaseBtnFlag = "R"
			
			Case "R"
				.btnRelease.value = "Release취소"
				.btnRelease.disabled = False
				lgReleaseBtnFlag = "C"
				
			Case Else
				.btnRelease.value = "Release"
				.btnRelease.disabled = True
				lgReleaseBtnFlag = "R"
				
		End Select
	End With
End Sub

'=============================================  2.6.3 ProtectResultFields()======================================
'=	Sub Name : ProtectResultFields
'=	Sub Desc : Release일, 양품창고, 불량품창고 Protect 처리 
'========================================================================================================
Sub ProtectResultFields(Byval strInspStatus)
	With frm1
		If strInspStatus = "D" Then				'판정완료 일경우 
			Call ggoOper.SetReqAttr(.txtInspectorCd, "N")
			Call ggoOper.SetReqAttr(.txtInspDt, "N")
			Call ggoOper.SetReqAttr(.txtInspQty, "N")
			Call ggoOper.SetReqAttr(.txtDefectQty, "N")
			Call ggoOper.SetReqAttr(.cboDecision, "N")
			Call ggoOper.SetReqAttr(.txtRemark, "D")
		ElseIf lgstatusflag = "R" Then			'Release완료인 경우 
			Call ggoOper.SetReqAttr(.txtInspectorCd, "Q")
			Call ggoOper.SetReqAttr(.txtInspDt, "Q")
			Call ggoOper.SetReqAttr(.txtInspQty, "Q")
			Call ggoOper.SetReqAttr(.txtDefectQty, "Q")
			Call ggoOper.SetReqAttr(.cboDecision, "Q")
			Call ggoOper.SetReqAttr(.txtRemark, "Q")
		Else
			Call ggoOper.SetReqAttr(.txtInspectorCd, "N")
			Call ggoOper.SetReqAttr(.txtInspDt, "N")
			Call ggoOper.SetReqAttr(.txtInspQty, "N")
			Call ggoOper.SetReqAttr(.txtDefectQty, "N")
			Call ggoOper.SetReqAttr(.cboDecision, "N")
			Call ggoOper.SetReqAttr(.txtRemark, "D")
			
		End If
	End With
End Sub

'=============================================  2.6.6 Release()======================================
'=	Sub Name : Release
'=	Sub Desc : Release 처리 
'========================================================================================================
Sub Release()
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Sub
    End If
    
	If lgReleaseBtnFlag = "R" Then

		Dim arrParam1, arrParam2
		Dim arrRet
		Dim iCalledAspName, IntRetCD
		
		If IsOpenPop = True Then Exit Sub
	
		IsOpenPop = True
		
		Redim arrParam1(6)
		Redim arrParam2(5)

		'환경설정의 Release일에 대한 기본표시값 확인 후 처리: 현재일, 검사일 
		Call CommonQueryRs2by2(" BASIC_MARK_FOR_RELEASE_DT ", " Q_CONFIGURATION ", " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " ", lgF0)
	
		lgF0 = Split(lgF0, Chr(11))
	
		arrParam1(0) = frm1.hReleaseDt.value
		arrParam1(1) = frm1.hGoodsQty.value
		arrParam1(2) = frm1.hDefectivesQty.value
		arrParam1(3) = frm1.hGoodsSLCd.value
		arrParam1(4) = frm1.hGoodsSLNm.value
		arrParam1(5) = frm1.hDefectivesSLCd.value
		arrParam1(6) = frm1.hDefectivesSLNm.value
		
		arrParam2(0) = lgIFYesNo
		arrParam2(1) = lgInspClassCd
		arrParam2(2) = lgReceivingInspType
		arrParam2(3) = lgAutoPR
		arrParam2(4) = lgAutoST
		arrParam2(5) = Trim(frm1.txtPlantCd.value)
		
		iCalledAspName = AskPRAspName("q4111pa2")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q4111pa2", "X")
			IsOpenPop = False
			Exit Sub
		End If
		
		arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2), _
			"dialogWidth=820px; dialogHeight=180px; center: Yes; help: No; resizable: No; status: No;")
		
		IsOpenPop = False
	
		If arrRet(0) = "" Then
			Exit Sub
		Else
			frm1.txtReleaseDt.Text = arrRet(0)
			frm1.txtGoodsSLCd.value = arrRet(1)
			frm1.txtGoodsSLNm.value = arrRet(2)
			frm1.txtDefectivesSLCd.value = arrRet(3)
			frm1.txtDefectivesSLNm.value = arrRet(4)
		End If
		
		frm1.btnRelease.disabled = True
		
		If Not DbConfirmRelease Then
			frm1.btnRelease.disabled = False
		End If
		
	ElseIf lgReleaseBtnFlag = "C" Then
		frm1.btnRelease.disabled = True
		
		If Not DbCancelRelease Then
			frm1.btnRelease.disabled = False
		End If
	End If
End Sub

'=============================================  2.6.7 DbConfirmRelease()======================================
'=	Function Name : DbConfirmRelease
'=	Function Desc : 
'========================================================================================================
Function DbConfirmRelease()
	DbConfirmRelease = False
	
	Dim strVal
	
	LayerShowHide(1)
       
    strVal = BIZ_PGM_CONFIRM_RELEASE_ID & "?txtInspReqNo=" & Trim(frm1.txtInspReqNo2.value) _
										& "&txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
										& "&txtReleaseDt=" & frm1.txtReleaseDt.Text _
										& "&txtGoodsSLCd=" & Trim(frm1.txtGoodsSLCd.value) _
										& "&txtDefectivesSLCd=" & Trim(frm1.txtDefectivesSLCd.value)
    
	Call RunMyBizASP(MyBizASP, strVal)	
	
	DbConfirmRelease = True
End Function

'=============================================  2.6.8 DbConfirmReleaseOK()======================================
'=	Sub Name : DbConfirmReleaseOK
'=	Sub Desc : 
'========================================================================================================
Sub DbConfirmReleaseOK()
	lgBlnFlgChgValue = False
    Call MainQuery()
End Sub

'========================================================================================
' Sub Name : DbConfirmReleaseNotOK
' Sub Desc : DbConfirmRelease가 성공하지 못할 경우 
'========================================================================================
Sub DbConfirmReleaseNotOK()		
	frm1.btnRelease.disabled = False
End Sub

'=============================================  2.6.9 DbCancelRelease()======================================
'=	Function Name : DbCancelRelease
'=	Function Desc : 
'========================================================================================================
Function DbCancelRelease()
	DbCancelRelease = False
	
	Dim strVal
    
    LayerShowHide(1)
       
    strVal = BIZ_PGM_CANCEL_RELEASE_ID & "?txtInspReqNo=" & Trim(frm1.txtInspReqNo2.value) _
									   & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
									   & "&txtReleaseDt=" & frm1.txtReleaseDt.Text _
									   & "&txtGoodsSLCd=" & Trim(frm1.txtGoodsSLCd.value) _
									   & "&txtDefectivesSLCd=" & Trim(frm1.txtDefectivesSLCd.value)
    
	Call RunMyBizASP(MyBizASP, strVal)	
	
	DbCancelRelease = True
End Function

'=============================================  2.6.10 DbCancelReleaseOK()======================================
'=	Sub Name : DbCancelReleaseOK
'=	Sub Desc : 
'========================================================================================================
Sub DbCancelReleaseOK()
	lgBlnFlgChgValue = False
    Call MainQuery()
End Sub

'========================================================================================
' Sub Name : DbCancelReleaseNotOK
' Sub Desc : DbCancelRelease가 성공하지 못할 경우 
'========================================================================================
Sub DbCancelReleaseNotOK()		
	frm1.btnRelease.disabled = False
End Sub

 '==========================================  3.1.1 Form_load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																	'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "3", "2")
	Call AppendNumberPlace("8", "15", "2")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolbar("11101000000011")
	Call InitVariables																		'⊙: Initializes local global variables
	
	Call ChangingFieldByInspClass("")
	Call ChangingReleaseBtn("")
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus
	Else
		frm1.txtInspReqNo1.focus
    End If
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - START */
'=======================================================================================================
'   Event Name : txtInspReqNo2_OnChange
'   Event Desc : 
'=======================================================================================================
Sub txtInspReqNo2_OnChange()
	Dim strInspReqNo
	
	If gLookUpEnable = False Then Exit Sub
	
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
	
    strInspReqNo = Trim(frm1.txtInspReqNo2.value)
    If strInspReqNo = "" Then Exit Sub
    If frm1.hInspReqNo2.value = strInspReqNo Then Exit Sub
    
    'Clear
    Call ChangingFieldByInspClass("")
    
    With frm1
		.hInspReqNo2.value = ""

		.txtInspClass.value = ""
		.txtItemCd.value = ""
		.txtItemNm.value = ""
		.txtSpec.value = ""
		.txtLotNo.value = ""
		.txtLotSubNo.Text = ""
		.txtLotSize.Text = ""
		.txtUnit.value = ""

		.txtSupplierCd.value = ""
		.txtSupplierNm.value = ""
		.txtSLCd1.value = ""
		.txtSLNm1.value = ""
		
		.txtRoutNo.value = ""
		.txtRoutNoDesc.value = ""
		.txtOprNo.value = ""
		.txtOprNoDesc.value = ""
		.txtWcCd.value = ""
		.txtWcNm.value = ""
				
		.txtSLCd2.value = ""
		.txtSLNm2.value = ""
			
		.txtBPCd.value = ""
		.txtBPNm.value = ""
		.txtInspDt.Text = ""
    End With

    Call LookUpInspReqNo(strInspReqNo)

End Sub

'=======================================================================================================
'	Sub Name : LookUpInspReqNo																			   
'	Sub Desc :																						
'========================================================================================================
Sub LookUpInspReqNo(Byval pvInspReqNo)
	Dim strVal
    
    Call LayerShowHide(1)
       
    strVal = BIZ_PGM_LOOKUP_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							   & "&txtInspReqNo=" & pvInspReqNo		
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

End Sub

'=======================================================================================================
'	Sub Name : LookUpInspReqNoOk																			   
'	Sub Desc :																						
'========================================================================================================
Sub LookUpInspReqNoOk(Byval pvInspClass)
	Dim blnRet
	'환경설정의 검사일에 대한 기본표시값 확인 후 처리: 현재일, 검사의뢰일 
	blnRet = CommonQueryRs2by2(" BASIC_MARK_FOR_INSP_DT ", " Q_CONFIGURATION ", " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " ", lgF0)
	If blnRet = False Then		'Default: 현재일 
		frm1.txtInspDt.Text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	Else
		lgF0 = Split(lgF0, Chr(11))
		
		If Trim(lgF0(1)) = "1" Then '현재일 
			frm1.txtInspDt.Text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
		Else
			frm1.txtInspDt.Text = frm1.txtInspReqDt.Text
		End If
	End If

	Call ChangingFieldByInspClass(pvInspClass)
	
End Sub
'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - START */

'=======================================================================================================
'   Event Name : cboDecision_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboDecision_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInspDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReleaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReleaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReleaseDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_Change
'   Event Desc : 
'=======================================================================================================
Sub txtInspDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtReleaseDt_Change
'   Event Desc : 
'=======================================================================================================
Sub txtReleaseDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspQty_Change
'   Event Desc : 
'=======================================================================================================
Sub txtInspQty_Change()
    lgBlnFlgChgValue = True
    With frm1
		If UNICDBL(.txtDefectQty.text) <= UNICDBL(.txtInspQty.text) Then
			Call Update_Rate
		End If
    End With
End Sub

'=======================================================================================================
'   Event Name : txtInspQty_OnBlur
'   Event Desc : 
'=======================================================================================================
Sub txtInspQty_OnBlur()
    With frm1
		If UNICDBL(.txtDefectQty.text) > UNICDBL(.txtInspQty.text) Then
			Call DisplayMsgBox("221310","X","X","X") 		'불량수가 검사수보다 클 수 없습니다.
			.txtDefectQty.value = .htxtDefectQty.value
			.txtInspQty.value = .htxtInspQty.value
			.txtInspQty.Focus
			Set gActiveElement = document.activeElement
		End If
    End With
End Sub

'=======================================================================================================
'   Event Name : txtDefectQty_Change
'   Event Desc : 
'=======================================================================================================
Sub txtDefectQty_Change()
    lgBlnFlgChgValue = True
    With frm1
		If UNICDBL(.txtDefectQty.text) <= UNICDBL(.txtInspQty.text) Then
			Call Update_Rate
		End If    
    End With
End Sub

'=======================================================================================================
'   Event Name : txtDefectQty_OnBlur
'   Event Desc : 
'=======================================================================================================
Sub txtDefectQty_OnBlur()
    With frm1
		If UNICDBL(.txtDefectQty.text) > UNICDBL(.txtInspQty.text) Then
			Call DisplayMsgBox("221310","X","X","X") 		'불량수가 검사수보다 클 수 없습니다.
			.txtDefectQty.value = .htxtDefectQty.value
			.txtInspQty.value = .htxtInspQty.value
			.txtDefectQty.Focus
			Set gActiveElement = document.activeElement
		End If
    End With
End Sub

'=======================================================================================================
'   Event Name : Update_Rate
'   Event Desc : 
'=======================================================================================================
Sub Update_Rate()
    With frm1
		If UNICDBL(.txtInspQty.text) <> UNICDBL(0) Then
			.txtDefectiveRate.Text = UNIFormatNumber(CStr(UNICDbl(.txtDefectQty.Text) / UNICDbl(.txtInspQty.Text) * UNICDbl(100)), 2, -2, 0, 3, 0)
			.htxtDefectQty.value = .txtDefectQty.text
			.htxtInspQty.value = .txtInspQty.text
		Else
			.txtDefectiveRate.Text = UNIFormatNumber(0, 2, -2, 0, 3, 0)
		End If
    End With
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	FncQuery = False
	
	Dim IntRetCD 
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
    End If

	'-----------------------
	'Check condition area
	'-----------------------
	If Not ChkField(Document, "1") Then	Exit Function
	
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
   	Call InitVariables
   	
   	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function									'☜: Query db data
	
	FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	FncNew = False
	
	Dim IntRetCD 
    	
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call ChangingFieldByInspClass("")
	Call ChangingReleaseBtn("")
	
	Call SetDefaultVal
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetToolbar("11101000000011")
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus
	Else
		frm1.txtInspReqNo2.focus
    End If
	Set gActiveElement = document.activeElement 
    
	FncNew = True 									'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	FncDelete = False
	
	Dim IntRetCD 
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True                                                        '⊙: Processing is OK                   							'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	FncSave = False
		
	Dim IntRetCD 
	strtxtLotSize = frm1.txtLotSize.text   
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
    
    '-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then Exit Function
    
    If frm1.cboDecision.Value = "N" Then
		Call DisplayMsgBox("221324", "X", "X", "X")  	'판정을 내리셔야 합니다 
		Exit Function
	End If
		
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then Exit Function                              '☜: Save db data
    
	FncSave = True                                      	                    '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim IntRetCD 
    
    FncPrev = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then Exit Function
    End If
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then	Exit Function
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbPrev = False Then Exit Function							'☜: Query db data
    
	FncPrev = True
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim IntRetCD 
    
    FncNext = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then	Exit Function
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbNext = False Then Exit Function							'☜: Query db data
    
	FncNext = False
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = False
	
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If

    Call ggoOper.ClearField(Document, "1")                        
    Call ggoOper.LockField(Document, "N")						
	
	Call InitVariables
	Call ChangingFieldByInspClass("")
	Call ChangingReleaseBtn("")
	
	frm1.txtInspReqNo2.value = ""
	frm1.txtInspClass.value = ""
	frm1.txtItemCd.value = ""
	frm1.txtItemNm.value = ""
	frm1.txtSpec.value = ""
	frm1.txtLotNo.value = ""
	frm1.txtLotSubNo.Text = ""
	frm1.txtLotSize.Text = ""
	frm1.txtUnit.value = ""
	frm1.txtStatusFlag.value = ""
	frm1.txtGoodsQty.Text = ""
	frm1.txtDefectivesQty.Text = ""
	frm1.txtReleaseDt.Text = ""
	frm1.txtGoodsSLCd.value = ""
	frm1.txtDefectivesSLCd.value = ""
	
	lgBlnFlgChgValue = True
	
	Call SetToolbar("11101000000011")
	Call DisableToolBar(TBC_COPY)	
	frm1.txtInspReqNo2.focus
	Set gActiveElement = document.activeElement  
	
	FncCopy = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
    Call parent.FncExport(Parent.C_SINGLE)											'☜: 화면 유형 
    FncExcel = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE , False)                                   '☜:화면 유형, Tab 유무 
    FncFind = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = False
	
	Dim IntRetCD
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If

    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '⊙: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1)
       
    strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value) _
							& "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbPrev()
    DbPrev = False                                                         '⊙: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value)	_
							& "&PrevNextFlg=" & "P"									'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	DbPrev = True
End Function

'========================================================================================
' Function Name : DbNext
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbNext()
    DbNext = False                                                         '⊙: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value) _
							& "&PrevNextFlg=" & "N"									'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	DbNext = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	DbQueryOk = False
    
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    
    Call ProtectResultFields(lgstatusflag)
    
    '공급처/작업장/창고/거래처 Display 처리 
    Call ChangingFieldByInspClass(lgInspClassCd)
    
    'Release 버튼 캡션변경 및 Protect 처리 
    Call ChangingReleaseBtn(lgstatusflag)
	
	'Toolbar Setting
	Select Case lgstatusflag
		Case "D"
			Call SetToolbar("11111000111111")
		Case "R"
			Call SetToolbar("11101000111111")
		Case Else
			Call SetToolbar("11111000111111")
	End Select
	
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement 
    
    DbQueryOk = True
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	DbSave = False															'⊙: Processing is NG

	LayerShowHide(1)
		
	With frm1
		.txtFlgMode.value = lgIntFlgMode
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()		
	DbSaveOk = False
	frm1.txtInspReqNo1.value = frm1.txtInspReqNo2.value 
	lgBlnFlgChgValue = False
    Call MainQuery()
    DbSaveOk = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = False
	
	Call LayerShowHide(1)
	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value)
	
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True			                                                   			'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()
	DbDeleteOk = false
	lgBlnFlgChgValue = False												'☆: 삭제 성공후 실행 로직 
	Call MainNew()
	DbDeleteOk = true
End Function

</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- TAB, REFERENCE AREA START -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>검사결과 등록</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- TAB, REFERENCE AREA END -->
	<!-- CONDITION/CONTENT AREA START -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<!-- CONDITION AREA START-->
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantPopup ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X">
									</TD>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtInspReqNo1" SIZE="20" MAXLENGTH="18" ALT="검사의뢰번호" TAG="12XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnInspReqNoPopup1 ONCLICK=vbscript:OpenInspReqNo1() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<!-- CONDITION AREA END-->
				<!-- CONTENT AREA START-->
				<TR>
					<TD HEIGHT=* WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_50%>>
							<!-- 검사의뢰내용 START -->
							<TR>
								<TD>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>검사의뢰</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtInspReqNo2" SIZE="20" MAXLENGTH="18" ALT="검사의뢰번호" TAG="23XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnInspReqNoPopup2 ONCLICK=vbscript:OpenInspReqNo2() OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()" SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">
												</TD>
												<TD CLASS="TD5" NOWPAP>검사분류</TD>
												<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtInspClass" SIZE="20" MAXLENGTH="40" ALT="검사분류" TAG="24" ></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>품목</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE="20" MAXLENGTH="18" ALT="품목" TAG="24">&nbsp;<INPUT NAME="txtItemNm" TAG="24"></TD>
												<TD CLASS="TD5" NOWRAP>규격</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE="40" MAXLENGTH="50" ALT="규격" TAG="24"></TD>
											</TR>
											<TR>
							                	
							                	<TD CLASS="TD5" NOWRAP>로트번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE="20" MAXLENGTH="25" ALT="LOT NO" TAG="24">&nbsp;
							                		<script language =javascript src='./js/q4111ma1_txtLotSubNo_txtLotSubNo.js'></script>
													</TD>
							                	<TD CLASS="TD5" NOWRAP>로트크기</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtLotSize_txtLotSize.js'></script>&nbsp;<INPUT TYPE=TEXT NAME="txtUnit" SIZE="5" MAXLENGTH="3" TAG="24" ALT="단위">
												</TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWRAP>검사의뢰일</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtInspReqDt_txtInspReqDt.js'></script>
												</TD>
												<TD CLASS=TD5 NOWRAP></TD>
												<TD CLASS=TD6 NOWRAP></TD>
											</TR>
							                <TR ID="Receiving">
												<TD CLASS=TD5 NOWRAP>공급처</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE="10" MAXLENGTH="10" ALT="공급처" TAG="24">&nbsp;<INPUT NAME="txtSupplierNm" TAG="24"></TD>
												<TD CLASS=TD5 NOWRAP>창고</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd1" SIZE="10" MAXLENGTH="7" ALT="창고" TAG="24">&nbsp;<INPUT NAME="txtSLNm1" TAG="24"></TD>
											</TR>
											<TR ID="Process1">
												<TD CLASS="TD5" NOWRAP>라우팅</TD>
				                				<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=20 MAXLENGTH=20 ALT="라우팅" tag="24">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutNoDesc" SIZE=20 MAXLENGTH=20 tag="24"></TD>
												<TD CLASS="TD5" NOWRAP>공정</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=5 MAXLENGTH=3 ALT="공정" tag="24">&nbsp;<INPUT TYPE=TEXT NAME="txtOprNoDesc" SIZE=20 MAXLENGTH=20 tag="24"></TD>
											</TR>
											<TR ID="Process2">
												<TD CLASS=TD5 NOWRAP>작업장</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE="10" MAXLENGTH="7" ALT="작업장" TAG="24">&nbsp;<INPUT NAME="txtWcNm" TAG="24"></TD>
												<TD CLASS=TD5 NOWRAP></TD>
												<TD CLASS=TD6 NOWRAP></TD>
											</TR>
											<TR ID="Final">
												<TD CLASS=TD5 NOWRAP>창고</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd2" SIZE="10" MAXLENGTH="7" ALT="창고" TAG="24">&nbsp;<INPUT NAME="txtSLNm2" TAG="24"></TD>
											</TR>
											<TR ID="Shipping">
												<TD CLASS=TD5 NOWRAP>거래처</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE="10" MAXLENGTH="10" ALT="거래처" TAG="24">&nbsp;<INPUT NAME="txtBpNm" TAG="24"></TD>
												<TD CLASS=TD5 NOWRAP></TD>
												<TD CLASS=TD6 NOWRAP></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<!-- 검사의뢰내용 END -->
							<!-- 검사결과 START -->
							<TR>
								<TD>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>검사결과</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사진행상태</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtStatusFlag" SIZE="20" MAXLENGTH="40" ALT="검사진행상태" TAG="24"></TD>
												<TD CLASS="TD5" NOWPAP></TD>
												<TD CLASS="TD6" NOWPAP></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사원</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtInspectorCd" SIZE="15" MAXLENGTH="10" ALT="검사원" TAG="22XXU"><IMG ALIGN=top HEIGHT=20 NAME=btnInspectorPopup ONCLICK=vbscript:OpenInspector() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtInspectorNm" TAG="24">
												</TD>
												<TD CLASS="TD5" NOWRAP>검사일</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtInspDt_txtInspDt.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사수</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtInspQty_txtInspQty.js'></script>
												</TD>
							                	<TD CLASS="TD5" NOWRAP>불량수</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtDefectQty_txtDefectQty.js'></script>
												</TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWRAP>판정</TD>        
												<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDecision" ALT="판정" STYLE="WIDTH: 150px" TAG="22"></SELECT></TD>
							                	<TD CLASS="TD5" NOWRAP>불량률</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtDefectiveRate_txtDefectiveRate.js'></script>&nbsp;%
												</TD>
							                </TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>비고사항</TD>
												<TD CLASS="TD6" NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtRemark" SIZE="100" MAXLENGTH="200" TAG="21" ALT="비고사항"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWRAP HEIGHT=5 colspan=3></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<!-- 검사결과 END -->
							<!-- Release START -->
							<TR>
								<TD>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>Release</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>양품수</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtGoodsQty_txtGoodsQty.js'></script>
												</TD>
							                	<TD CLASS="TD5" NOWRAP>불량품수</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtDefectivesQty_txtDefectivesQty.js'></script>
												</TD>
							                </TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>Release일</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q4111ma1_txtReleaseDt_txtReleaseDt.js'></script>
												</TD>
												<TD CLASS="TD5" NOWRAP></TD>
												<TD CLASS="TD6" NOWRAP></TD>
												
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>양품창고</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtGoodsSLCd" SIZE="10" MAXLENGTH="7" ALT="양품창고" TAG="24XXXU">&nbsp;<INPUT NAME="txtGoodsSLNm" TAG="24">
												</TD>
							                	<TD CLASS="TD5" NOWRAP>불량품창고</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtDefectivesSLCd" SIZE="10" MAXLENGTH="7" ALT="불량품창고" TAG="24XXXU">&nbsp;<INPUT NAME="txtDefectivesSLNm" TAG="24">
												</TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<!-- Release END -->
						</TABLE>
					</TD>
				</TR>	
				<!-- CONTENT AREA END-->
			</TABLE>
		</TD>
	</TR>
	<!-- CONDITION/CONTENT AREA END -->
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
	      	<TD WIDTH="100%" >
	      		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	        		<TR>
	        			<TD WIDTH=10>&nbsp;</TD>
	        			<TD><BUTTON NAME="btnRelease" CLASS="CLSMBTN" ONCLICK="vbscript:Release()">Release</BUTTON></TD>
	        			<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadIDispositionOfNCM">부적합 처리</A></A>&nbsp;|&nbsp;<A href="vbscript:LoadNoticeOfRejection">불합격 통지 등록</A></TD>
	        		</TR>
	      		</TABLE>
	      	</TD>
         </TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtDefectQty" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="htxtInspQty" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspReqNo2" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hReleaseDt" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hGoodsQty" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hDefectivesQty" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hGoodsSLCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hGoodsSLNm" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hDefectivesSLCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hDefectivesSLNm" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
