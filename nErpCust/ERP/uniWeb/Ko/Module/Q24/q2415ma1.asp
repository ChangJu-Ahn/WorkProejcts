<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2415MA1
'*  4. Program Name         : 판정 
'*  5. Program Desc         : Quality Configuration
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

Const BIZ_PGM_QRY_ID = "Q2415MB1.asp"	 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_QRY2_ID = "Q2415MB4.asp"
Const BIZ_PGM_SAVE_ID = "Q2415MB2.asp"	 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "Q2415MB3.asp"

Const BIZ_PGM_JUMP1_ID = "Q2411MA1"
Const BIZ_PGM_JUMP2_ID = "Q2416MA1"
Const BIZ_PGM_JUMP3_ID = "Q2417MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_DecisionNm '= 1
Dim C_InspOrder '= 2	
Dim C_InspItemCd '= 3	
Dim C_InspItemNm '= 4
Dim C_InspSeries '= 5
Dim C_SampleQty '= 6
Dim C_DefectQty '= 7
Dim C_AcceptanceNumber '= 8
Dim C_RejectionNumber '= 9
Dim C_AcceptanceCoefficient '= 10
Dim C_MaxDefectRatio '= 11
'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - START */
Dim C_InspMethodCd '= 12
Dim C_InspMethodNm '= 13
Dim C_DecisionCd '= 14
'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - END */

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgNextNo					'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo					' ""

Dim strInspClass

Dim IsOpenPop      

'==========================================  2.1.1 InitVariables()======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE   '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '⊙: Indicates that no value changed
    lgIntGrpCount = 0           '⊙: Initializes Group View Size
    	
    '----------  Coding part -------------------------------------------------------------
    	
    IsOpenPop = False					'☆: 사용자 변수 초기화 
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    
    '###검사분류별 변경부분 Start###
    strInspClass = "S"
	'###검사분류별 변경부분 End###
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	
	'/* 10월 정기패치 : 판정의 Default 값을 '미판정'에서 '합격'으로 변경 - START */
	frm1.cboDecision.value = "A"
	'/* 10월 정기패치 : 판정의 Default 값을 '미판정'에서 '합격'으로 변경 - END */
	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If
	
	If ReadCookie("txtInspReqNo") <> "" Then
		frm1.txtInspReqNo.Value = ReadCookie("txtInspReqNo")
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""
	'/* 9월 정기패치: 화면 초기화 시 검사일 Default값 입력안함. */
End Sub

'=============================================== 2.2.3 InitSpreadSheet()========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021225", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		
		.MaxCols = C_DecisionCd + 1
    	.MaxRows = 0
    		
    	Call GetSpreadColumnPos("A")
    		
		Call AppendNumberPlace("6", "3","0")
		Call AppendNumberPlace("7", "15","4")
		
		ggoSpread.SSSetCombo C_DecisionCd, "판정코드", 10, 0, False
		ggoSpread.SSSetCombo C_DecisionNm, "판정", 10, 0, False
		ggoSpread.SSSetFloat C_InspOrder, "검사순서", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit C_InspItemCd, "검사항목코드", 12, 0, -1, 5, 2
		ggoSpread.SSSetEdit C_InspItemNm, "검사항목명", 20, 0, -1, 40
		ggoSpread.SSSetFloat C_InspSeries, "차수", 5, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
   		ggoSpread.SSSetFloat C_SampleQty, "시료수", 12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
	    ggoSpread.SSSetFloat C_AcceptanceNumber, "합격판정개수", 12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
	    ggoSpread.SSSetFloat C_RejectionNumber, "불합격판정개수", 12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_AcceptanceCoefficient, "합격판정계수", 12, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_MaxDefectRatio, "최대허용불량률", 12, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - START */
		ggoSpread.SSSetEdit C_InspMethodCd, "검사방식코드", 12, 0, -1, 4, 2
		ggoSpread.SSSetEdit C_InspMethodNm, "검사방식명", 30, 0, -1, 40
		'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - END */
		ggoSpread.SSSetFloat C_DefectQty, "불량수", 12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

		ggoSpread.SSSetRequired C_DecisionCd, 1, .MaxRows	'@@@변경 

 		Call ggoSpread.SSSetColHidden(C_DecisionCd, C_DecisionCd, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    .ReDraw = true
	    
	    Call SetSpreadLock
	End With
End Sub

'================================== 2.2.5 SetSpreadLock()==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	frm1.vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
	'ggoSpread.SpreadLock 1, -1, frm1.vspdData.MaxCols
	frm1.vspdData.ReDraw = True
End Sub

'================================== 2.2.7 SetSpreadColor()==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_DecisionNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspOrder, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspSeries, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SampleQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcceptanceNumber, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RejectionNumber, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcceptanceCoefficient, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MaxDefectRatio, pvStartRow, pvEndRow
		'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - START */
		ggoSpread.SSSetProtected C_InspMethodCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspMethodNm, pvStartRow, pvEndRow
		'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - END */
		'ggoSpread.SSSetProtected C_DefectQty, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_DefectQty, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_DecisionNm = 1
	C_InspOrder= 2	
	C_InspItemCd= 3	
	C_InspItemNm = 4
	C_InspSeries= 5
	C_SampleQty = 6
	C_DefectQty = 7
	C_AcceptanceNumber = 8
	C_RejectionNumber = 9
	C_AcceptanceCoefficient = 10
	C_MaxDefectRatio = 11
	'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - START */
	C_InspMethodCd = 12
	C_InspMethodNm = 13		
	C_DecisionCd = 14	
	'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - END */
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_DecisionNm = iCurColumnPos(1)
			C_InspOrder= iCurColumnPos(2)	
			C_InspItemCd= iCurColumnPos(3)	
			C_InspItemNm = iCurColumnPos(4)
			C_InspSeries= iCurColumnPos(5)
			C_SampleQty = iCurColumnPos(6)
			C_DefectQty = iCurColumnPos(7)
			C_AcceptanceNumber = iCurColumnPos(8)
			C_RejectionNumber = iCurColumnPos(9)
			C_AcceptanceCoefficient = iCurColumnPos(10)
			C_MaxDefectRatio = iCurColumnPos(11)
			'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - START */
			C_InspMethodCd = iCurColumnPos(12)
			C_InspMethodNm = iCurColumnPos(13)
			C_DecisionCd = iCurColumnPos(14)
			'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - END */
 	End Select
End Sub

'==========================================  2.2.6 InitComboBox()=======================================
'	Name : InitComboBox
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","major_cd=" & FilterVar("Q0010", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDecision,lgF0  ,lgF1  ,Chr(11))
End Sub

'==========================================  2.2.6 InitSpreadComboBox()=======================================
'	Name : InitComboBox
'	Description : Combo Display
'========================================================================================================= 
Sub InitSpreadComboBox()
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","major_cd=" & FilterVar("Q0009", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	lgF0 = Replace(lgF0,chr(11),vbTab)
	lgF1 = Replace(lgF1,chr(11),vbTab)
	
	ggoSpread.SetCombo lgF0,C_DecisionCd
	ggoSpread.SetCombo lgF1,C_DecisionNm
End Sub

'------------------------------------------  OpenPlant() -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

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

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam,arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenPlant = true		
End Function

'------------------------------------------  OpenInspReqNo() -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo()        
	OpenInspReqNo = false
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo.Value)	
	'###검사분류별 변경부분 Start###	
	Param4 = strInspClass 		'검사분류 
	'###검사분류별 변경부분 End###
	Param5 = ""			'판정 
	Param6 = ""			'검사진행상태 
	
	iCalledAspName = AskPRAspName("Q4111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "Q4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	frm1.txtInspReqNo.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspReqNo.Value    = arrRet(0)		
		frm1.txtInspReqNo.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspReqNo = true
End Function

'------------------------------------------  OpenInspector() -------------------------------------------------
'	Name : OpenInspector()
'	Description :Inspector PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspector()
	OpenInspector = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	If UCase(frm1.txtInspectorCd.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If

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
	
	frm1.txtInspectorCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspectorCd.Value    = arrRet(0)
		frm1.txtInspectorNm.Value    = arrRet(1)				
		frm1.txtInspectorCd.Focus	
		lgBlnFlgChgValue = True	
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspector = true
End Function

'=============================================  2.5.1 LoadInspection()======================================
'=	Event Name : LoadInspection
'=	Event Desc :
'========================================================================================================
Function LoadInspection()
	Dim intRetCD
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
		
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 LoadDisposition()  ======================================
'=	Event Name : LoadDisposition
'=	Event Desc :
'========================================================================================================
Function LoadDisposition()
	Dim intRetCD
	
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
		
	End With
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=============================================  2.5.3 LoadRejectReport()  ======================================
'=	Event Name : LoadRejectReport
'=	Event Desc :
'========================================================================================================
Function LoadRejectReport()	'수입검사에만 있슴.
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
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
		
	End With
	
	PgmJump(BIZ_PGM_JUMP3_ID)
End Function

'=============================================  2.5.4 LoadRelease()  ======================================
'=	Event Name : LoadRelease
'=	Event Desc :
'========================================================================================================
Function LoadRelease()
	Dim intRetCD
	
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	If strInspClass = "R" Then
		PgmJump(BIZ_PGM_JUMP4_ID)
	ElseIf strInspClass = "P" Then
		PgmJump(BIZ_PGM_JUMP3_ID)
	ElseIf strInspClass = "F" Then
		PgmJump(BIZ_PGM_JUMP3_ID)
	ElseIf strInspClass = "S" Then
		PgmJump(BIZ_PGM_JUMP3_ID)
	End If
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey			'Sort in Descending
 			lgSortKey = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub
 
'==========================================  3.1.1 Form_Load()======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     '⊙:Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "Q")                                   '⊙:Lock  Suitable  Field
	Call InitSpreadSheet                                                    '⊙:Setup the Spread sheet
	Call InitComboBox
    Call InitSpreadComboBox
	'----------  Coding part -------------------------------------------------------------
	Call SetDefaultVal
	Call SetToolBar("11100000000011")		'⊙: 버튼 툴바 제어 
	Call InitVariables                                                      '⊙:Initializes local global variables

	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If    	
	lgBlnFlgChgValue = False
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call InitSpreadComboBox
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
	Call DbQueryOk
 	'------ Developer Coding part (End) 	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )    	
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'=======================================================================================================
'   Event Name : txtInspQty_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtInspQty_Change()
	'/* 2003-04 정기 패치: 검사분류별 불량률 단위 사용하도록 되있는 내용 중 잘못된 것 수정 - START */
	lgBlnFlgChgValue = True
	With frm1
	    If .txtInspQty.Text = "" Then 
			.txtDefectRatio.Value = UNIFormatNumber(0, 2, -2, 0, 3, 0)
			Exit Sub
		End If
	    
	    If UNICDbl(.txtInspQty.Text) = 0 Then 
			.txtDefectRatio.Value = UNIFormatNumber(0, 2, -2, 0, 3, 0)
		Else
			.txtDefectRatio.Value = UNIFormatNumber(CStr(UNICDbl(.txtDefectQty.Text) / UNICDbl(.txtInspQty.Text) * UNICDbl(100)), 2, -2, 0, 3, 0)
	    End If
	End With
    '/* 2003-04 정기 패치: 검사분류별 불량률 단위 사용하도록 되있는 내용 중 잘못된 것 수정 - END */
End Sub

'=======================================================================================================
'   Event Name : txtDefectQty_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtDefectQty_Change()
    '/* 2003-04 정기 패치: 검사분류별 불량률 단위 사용하도록 되있는 내용 중 잘못된 것 수정 - START */
	lgBlnFlgChgValue = True
    With frm1
	    If .txtDefectQty.Text = "" Then 
			.txtDefectRatio.Value = UNIFormatNumber(0, 2, -2, 0, 3, 0)
			Exit Sub
		End If
	    
	    If UNICDbl(.txtInspQty.Text) = 0 Then 
			.txtDefectRatio.Value = UNIFormatNumber(0, 2, -2, 0, 3, 0)
		Else
			.txtDefectRatio.Value = UNIFormatNumber(CStr(UNICDbl(.txtDefectQty.Text) / UNICDbl(.txtInspQty.Text) * UNICDbl(100)), 2, -2, 0, 3, 0)
	    End If
   End With
'/* 2003-04 정기 패치: 검사분류별 불량률 단위 사용하도록 되있는 내용 중 잘못된 것 수정 - END */
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
'   Event Name : txtInspDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInspDt_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		.Row = Row
		Select Case Col
			Case  C_DecisionNm
				.Col = Col
				intIndex = .Value
				.Col = C_DecisionCd
				.Value = intIndex
		End Select
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft, ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	 '----------  Coding part -------------------------------------------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  Then	'☜: 재쿼리 체크 
		If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" Then		'⊙:다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
		
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub

'=======================================================================================================
'   Event Name : cboDecision_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboDecision_onchange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
	Dim IntRetCD
	
	FncQuery = False     '⊙: Processing is NG
	
	Err.Clear            '☜: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")		'⊙: Clear Contents Field
	Call InitVariables
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then
		Exit Function
	End If																'☜: Query db data
	
	FncQuery = True		
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
	Dim IntRetCD
	
	FncNew = False                			'⊙: Processing is NG
	Err.Clear                            		'☜: Protect system from crashing
	
	'-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "Q")              '⊙: Lock  Suitable Field
	Call InitVariables				   '⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolBar("11100000000011")		'⊙: 버튼 툴바 제어 
	
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If    	
	
	lgBlnFlgChgValue = False
	FncNew = True
End Function

'========================================================================================
' Function Name : Fnc
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()
	Dim IntRetCD
	
	FncDelete = False									'⊙: Processing is NG
	
	  '-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then
		Exit Function
	End If

	  '-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then
		Exit Function
	End If
	
	FncDelete = True   
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave()
	Dim IntRetCD
	Dim lRow
	
	FncSave = False             '⊙: Processing is NG
	
	Err.Clear		     '☜: Protect system from crashing
	
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = False  and ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
	
	  '-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "1") Then
       		Exit Function
    End If

    If Not chkField(Document, "2") Then
    	Exit Function
    End If
    	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSDefaultCheck = False Then    				'⊙: Check contents area
    	Exit Function
    End If
    	
    With frm1
		If UNICDbl(.txtInspQty.Text) <= 0 Then
			Call DisplayMsgBox("221325", "X", "X", "X")  	'검사수는 0보다 커야 합니다.
			Exit Function
		End If
	    	
		If .cboDecision.Value = "N" Then
			Call DisplayMsgBox("221324", "X", "X", "X")  	'판정을 내리셔야 합니다 
			Exit Function
		End If
	    	
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C_DecisionCd				
			If .vspdData.Text = "N" Then
				Call DisplayMsgBox("221324", "X", "X", "X")  	'판정을 내리셔야 합니다 
				Exit Function
			End If
		Next
		
	    If Len(Trim(.txtRemark.Value)) > 200 Then
			Call MsgBox("비고는 200자를 초과할 수 없습니다", vbInformation)
			.txtRemark.Focus
			Exit Function
		End If
	End With
	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then	
		Exit Function
	End If		               '☜: Save db data
	
	FncSave = True             '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
	FncCopy = False
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	FncCancel = False
	
	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	
	FncCancel = True
End Function

'=============================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'===============================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = False
	
	If frm1.vspdData.MaxRows < 1 then
	   Call DisplayMsgBox("900002","X", "X", "X")	'조회를 먼저하십시요 
	   Exit function
	End If

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)

	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If
	
    With frm1	
    	.vspdData.ReDraw = False
     	.vspdData.focus
    	Parent.ggoSpread.Source = .vspdData
    	Parent.ggoSpread.InsertRow .vspdData.ActiveRow, imRow
    	'SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
    	.vspdData.ReDraw = True
    End With
    '----------  Coding part  -------------------------------------------------------------  
    
	Call SetActiveCell(frm1.vspdData,C_DecisionNm,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement	  
    If Err.number = 0 Then FncInsertRow = True
End Function

'===============================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'===============================================================================
Function FncDeleteRow()
	FncDeleteRow = false
End Function

'===============================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'===============================================================================
Function FncPrint()
	Call Parent.FncPrint()
End Function

'===============================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'===============================================================================
Function FncPrev()
	FncPrev = false
	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgPrevNo = "" Then
	 	Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
	 	Exit Function
	End If
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
	strVal = strVal & "&txtInspReqNo=" & lgPrevNo						'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncPrev = true
End Function

'===============================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'===============================================================================
Function FncNext()	
	FncNext = false
	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태값 
	strVal = strVal & "&txtInspReqNo=" & lgNextNo						'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncNext = true
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI , False)                          
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'===============================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel
'===============================================================================
Function FncExit()	
	Dim IntRetCD	
	FncExit = False
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True  Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If	
	FncExit = True
End Function

'===============================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'===============================================================================
Function DbDelete()
	Err.Clear                                                           
	Call LayerShowHide(1)
	DbDelete = False			
	Dim strVal
	strVal = BIZ_PGM_DEL_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtInspReqNo=" & Trim(frm1.txtInspReqNo.value)
	
	Call RunMyBizASP(MyBizASP, strVal)
	DbDelete = True
End Function

'===============================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'===============================================================================
Function DbDeleteOk()	
	DbDeleteOk = false
	lgBlnFlgChgValue = False
	'/* 9월 정기패치: 판정 삭제 후 조회 실행 - START */
	Call MainQuery()
	'/* 9월 정기패치: 판정 삭제 후 조회 실행 - END */
	DbDeleteOk = true
End Function

'===============================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'===============================================================================
Function DbQuery()
	'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - START */
	Dim strVal
	
	DbQuery = False
	Err.Clear
	
	Call LayerShowHide(1)
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & .hPlantCd.value		
			strVal = strVal & "&txtInspReqNo=" & .hInspReqNo.value
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)
			strVal = strVal & "&txtInspReqNo=" & Trim(.txtInspReqNo.value)
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)
		DbQuery = True                            
	End With
	'/* 2003-04 정기패치: 계수형 검사방식인 경우 검사항목별 자동 판정 기능 추가 - END */
	
End Function

'===============================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'===============================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	'검사진행현황이 '판정완료' 혹은 Release	완료가 아닌 경우에만 수정가능토록 한다.
	ggoSpread.Source = frm1.vspdData
   	If frm1.hStatusFlag.value <> "D" And frm1.hStatusFlag.value <> "R" Then
   		Dim i
   		With frm1.vspdData
   			for i = 1 to .MaxRows
   				.Row = i
   				.Col = 0
   				.Text = ggoSpread.UpdateFlag
   			Next 
   		End With

   		Call ggoOper.LockField(Document, "N")              '⊙: Lock  Suitable Field
   		ggoSpread.SpreadUnLock C_DecisionNm, 1, C_DecisionNm
   		ggoSpread.SpreadUnLock C_DefectQty, 1, C_DefectQty
   		Call SetSpreadColor(1, frm1.vspdData.MaxRows)
   		Call SetToolBar("11111001000111")		'⊙: 버튼 툴바 제어	
   	Else
   		Call ggoOper.LockField(Document, "Q")			'⊙: This function lock the suitable field	
   		Call SetSpreadColor(1, frm1.vspdData.MaxRows)
   		ggoSpread.SSSetProtected C_DecisionNm, -1, -1
   		ggoSpread.SSSetProtected C_DefectQty, -1, -1
   		Call SetToolBar("11111000000111")
   	End If
   	lgBlnFlgChgValue = False
   	lgIntFlgMode = Parent.OPMD_UMODE	'⊙: Indicates that current mode is Update mode
   	DbQueryOk = true
End Function

'===============================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'===============================================================================
Function DbSave()
	Dim lRow    
	Dim lGrpCnt 
	Dim strVal
	Dim strDel
	
	Call LayerShowHide(1)
	
	DbSave = False                          '⊙:Processing is NG

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1

		strVal = ""
    	strDel = ""
    		
    	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			If .vspdData.Text = ggoSpread.UpdateFlag Then  		'☜: 수정 
				strVal = strVal & "U" & Parent.gColSep			'☜: U=Update
				.vspdData.Col = C_InspItemCd		'1
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_InspSeries		'2
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_DefectQty		'3
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_DecisionCd		'4
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				strVal = strVal & CStr(lRow) & Parent.gRowSep	'5
				lGrpCnt = lGrpCnt + 1
			End If
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With
	
	DbSave = True
End Function

'===============================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function
'===============================================================================
Function DbSaveOk()
	DbSaveOk = false
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
	DbSaveOk = true
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
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>출하검사 판정</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
     									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>								
     									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20  MAXLENGTH=18 ALT="검사의뢰번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo()"></TD>							
															</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD>
									<FIELDSET CLASS="CLSFLD">
										<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>품목</TD>
					                						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="품목" tag="24">
													<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="24" ></TD>
												<TD CLASS="TD5" NOWRAP>거래처</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=4 ALT="거래처" tag="24">
													<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 tag="24" ></TD>
											</TR>
							                <TR>
					                			<TD CLASS="TD5" NOWRAP>로트번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=15 MAXLENGTH=12 ALT="LOT NO" tag="24">
					                				<INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=10 MAXLENGTH=5 tag="24" STYLE="Text-Align: Right"></TD>
					                			<TD CLASS="TD5" NOWRAP>로트크기</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2415ma1_fpDoubleSingle1_txtLotSize.js'></script>
												</TD>
						                	</TR>
					                		<TR>
												<TD CLASS="TD5" NOWRAP>검사의뢰일</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2415ma1_txtInspReqDt_txtInspReqDt.js'></script>
												</TD>		
												<TD CLASS="TD5" NOWRAP></TD>
				     							<TD CLASS="TD6" NOWRAP></TD>								
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>			
							<TR>
								<TD>
									<FIELDSET CLASS="CLSFLD">
										<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사원</TD>
				     								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspectorCd" SIZE=10 MAXLENGTH=10 ALT="검사원" tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspector" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspector()">
													<INPUT TYPE=TEXT NAME="txtInspectorNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>								
												<TD CLASS="TD5" NOWRAP>검사일</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2415ma1_fpDateTime1_txtInspDt.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사수</TD>            
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2415ma1_fpDoubleSingle2_txtInspQty.js'></script>
												</TD>
												<TD CLASS="TD5" NOWRAP>불량수</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2415ma1_fpDoubleSingle3_txtDefectQty.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>판정</TD>
												<TD CLASS="TD6" NOWRAP>
													<SELECT Name="cboDecision" ALT="판정" STYLE="WIDTH: 100px" tag="23"></SELECT></TD>
												<TD CLASS="TD5" NOWRAP>불량률</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDefectRatio" SIZE=15 MAXLENGTH=15 ALT="불량률" tag="24X3" STYLE="Text-Align: Right" >
														<INPUT TYPE=TEXT NAME="txtDefectRatioUnit" SIZE=3 MAXLENGTH=3 ALT="불량률단위" tag="24" STYLE="Text-Align: Center">	
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>비고사항</TD>
												<TD CLASS="TD6" NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtRemark" style="width:650px;" MAXLENGTH=200 TAG="25" ALT="비고사항"></TD>
											</TR>
										</TABLE>
									</FIELDSET>		
								</TD>
							</TR>							
							<TR>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<script language =javascript src='./js/q2415ma1_I220323789_vspdData.js'></script>
								</TD>	
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
      		<TD WIDTH=100%>
      			<TABLE <%=LR_SPACE_TYPE_30%>>
        				<TR>        				
        					<TD WIDTH=10>&nbsp;</TD>
	        				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspection">출하검사</A>&nbsp;|&nbsp;<A href="vbscript:LoadDisposition">부적합처리</A>&nbsp;|&nbsp;<A href="vbscript:LoadRelease">Release</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
       				</TR>
      			</TABLE>
      		</TD>
    	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread Width=100% tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspReqNo" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" tabindex=-1>
<!-- '/* 10월 정기패치 : 판정의 Default값을 '합격'으로 보여주는 것 관련 추가 - START */ -->
<INPUT TYPE=HIDDEN NAME="hStatusFlag" tag="24">
<!-- '/* 10월 정기패치 : 판정의 Default값을 '합격'으로 보여주는 것 관련 추가 - END */ -->
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


