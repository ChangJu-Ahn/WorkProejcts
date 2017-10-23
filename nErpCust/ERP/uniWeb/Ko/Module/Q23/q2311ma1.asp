<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2311MA1
'*  4. Program Name         : 검사등록 
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

Const BIZ_PGM_QRY_ID = "Q2311MB1.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "Q2311MB2.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "Q2311MB3.asp"
Const BIZ_PGM_QRY2_ID = "Q2311MB4.asp"
'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - START */
Const BIZ_PGM_LOOKUP_ID ="q2311mb5.asp"							
'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - END */
Const BIZ_PGM_JUMP1_ID = "Q2312ma1"
Const BIZ_PGM_JUMP2_ID = "Q2313ma1"
Const BIZ_PGM_JUMP3_ID = "Q2314ma1"
Const BIZ_PGM_JUMP4_ID = "Q2315ma1"
Const BIZ_PGM_JUMP5_ID = "Q2316ma1"
Const BIZ_PGM_JUMP6_ID = "Q2317ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspItemCd '= 1
Dim C_InspItemPopup '= 2
Dim C_InspItemNm '= 3
Dim C_InspOrder '= 4
Dim C_InspSeries '= 5
Dim C_SampleQty '= 6
Dim C_AcctncNumber '= 7
Dim C_RejtnNumber '= 8
Dim C_AccptncCoefficient '= 9
Dim C_MaxDefectRatio '= 10
Dim C_InspMthdNm '= 11
Dim C_InspUnitIndctnNm '= 12 
Dim C_InspSpec '= 13
Dim C_LSL '= 14
Dim C_USL '= 15
Dim C_MsmtEqpmtNm '= 16
Dim C_MsmtUnit '= 17
'------------- Hidden Column -----------
Dim C_InspMthdCd '= 18
Dim C_InspUnitIndctnCd '= 19 
Dim C_MsmtEqpmtCd '= 20
'---------------------------------------

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgNextNo					'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo					' ""
Dim strInspClass

Dim IsOpenPop          
Dim IsDbQueryOk
Dim blnIsInspectionRequest
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE                                               	'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False                                                	'⊙: Indicates that no value changed
	lgIntGrpCount = 0                                                     	  	'⊙: Initializes Group View Size
	'----------  Coding part  -------------------------------------------------------------
	IsOpenPop = False						'☆: 사용자 변수 초기화 
	lgStrPrevKey1 = ""
	lgStrPrevKey2 = ""
			
	IsDbQueryOk = False

	'###검사분류별 변경부분 Start###
	strInspClass = "F"
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
	
	blnIsInspectionRequest = ReadCookie("IsInspectionRequest")
	If ReadCookie("txtInspReqNo") <> "" Then
		If blnIsInspectionRequest = "True" Then
			frm1.txtInspReqNo2.Value = ReadCookie("txtInspReqNo")
			frm1.txtInspReqNo2.focus
			Call OnCookiesLoadingLookUpInspReqNo		
		Else
			frm1.txtInspReqNo.Value = ReadCookie("txtInspReqNo")
			frm1.txtInspReqNo.focus
		End If
	End If

	Set gActiveElement = document.activeElement
	
	WriteCookie "IsInspectionRequest", ""
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""	
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20040518", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
   		
   		.MaxCols = C_MsmtEqpmtCd + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
    		
    	Call AppendNumberPlace("6", "3","0")
    	Call AppendNumberPlace("7", "15","4")
    	
		ggoSpread.SSSetEdit		C_InspItemCd, "검사항목코드", 12, 0, -1, 5, 2
		ggoSpread.SSSetButton	C_InspItemPopup
		ggoSpread.SSSetEdit		C_InspItemNm, "검사항목명", 15, 0, -1, 40		
    	ggoSpread.SSSetFloat	C_InspOrder, "검사순서", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "P"
   		ggoSpread.SSSetFloat	C_InspSeries, "차수", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "P"
		ggoSpread.SSSetFloat	C_SampleQty, "시료수", 14, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat	C_AcctncNumber, "합격판정개수", 14, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat	C_RejtnNumber, "불합격판정개수", 14, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat	C_AccptncCoefficient, "합격판정계수", 14, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_MaxDefectRatio, "최대허용불량률", 14, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit		C_InspMthdCd, "검사방식코드", 10, 0, -1, 40, 2
    	ggoSpread.SSSetEdit		C_InspMthdNm, "검사방식", 20, 0, -1, 40
   		ggoSpread.SSSetEdit		C_InspUnitIndctnCd, "검사단위 품질표시코드", 10, 0, -1, 1, 2
   		ggoSpread.SSSetEdit		C_InspUnitIndctnNm, "검사단위 품질표시", 20, 0, -1, 40
   		ggoSpread.SSSetEdit		C_InspMthdNm, "검사방식", 20, 0, -1, 40
   		ggoSpread.SSSetEdit		C_InspSpec, "검사규격", 20, 2, -1, 20
   		ggoSpread.SSSetFloat	C_LSL, "하한규격", 14, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
   		ggoSpread.SSSetFloat	C_USL, "상한규격", 14, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
 		ggoSpread.SSSetEdit		C_MsmtEqpmtCd, "측정기코드", 10, 0, -1, 40, 2
 		ggoSpread.SSSetEdit		C_MsmtEqpmtNm, "측정기", 20, 0, -1, 40
 		ggoSpread.SSSetEdit		C_MsmtUnit, "측정단위", 11, 0, -1, 3
				
 		Call ggoSpread.MakePairsColumn(C_InspItemCd, C_InspItemPopup)
 		Call ggoSpread.SSSetColHidden(C_InspMthdCd, C_InspMthdCd, True)
 		Call ggoSpread.SSSetColHidden(C_InspUnitIndctnCd, C_InspUnitIndctnCd, True)
 		Call ggoSpread.SSSetColHidden(C_MsmtEqpmtCd, C_MsmtEqpmtCd, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)	    
	    
	    .ReDraw = true	
	    
	    Call SetSpreadLock
	End With
End Sub

'================================== 2.2.5 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	With frm1.vspdData
		.ReDraw = False
    	ggoSpread.SpreadLock C_InspItemNm, -1, C_InspSeries
   		ggoSpread.SpreadLock C_InspMthdCd, -1, C_MsmtUnit
   		Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
		.ReDraw = True
	End With    
End Sub

'================================== 2.2.7 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_InspItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspOrder, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspSeries, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_SampleQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspMthdCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspMthdNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspUnitIndctnCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspUnitIndctnNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspSpec, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LSL, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_USL, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MsmtEqpmtCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MsmtEqpmtNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MsmtUnit, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
	End With    
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_InspItemCd		= 1
	C_InspItemPopup		= 2
	C_InspItemNm		= 3
	C_InspOrder			= 4
	C_InspSeries		= 5
	C_SampleQty			= 6
	C_AcctncNumber		= 7
	C_RejtnNumber		= 8
	C_AccptncCoefficient= 9
	C_MaxDefectRatio	= 10
	C_InspMthdNm		= 11
	C_InspUnitIndctnNm	= 12 
	C_InspSpec			= 13
	C_LSL				= 14
	C_USL				= 15
	C_MsmtEqpmtNm		= 16
	C_MsmtUnit			= 17
	'------------- Hidden Column -----------
	C_InspMthdCd		= 18
	C_InspUnitIndctnCd	= 19 
	C_MsmtEqpmtCd		= 20
	'---------------------------------------	
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
					
			C_InspItemCd		= iCurColumnPos(1)
			C_InspItemPopup		= iCurColumnPos(2)
			C_InspItemNm		= iCurColumnPos(3)
			C_InspOrder			= iCurColumnPos(4)
			C_InspSeries		= iCurColumnPos(5)
			C_SampleQty			= iCurColumnPos(6)
			C_AcctncNumber		= iCurColumnPos(7)
			C_RejtnNumber		= iCurColumnPos(8)
			C_AccptncCoefficient= iCurColumnPos(9)
			C_MaxDefectRatio	= iCurColumnPos(10)
			C_InspMthdNm		= iCurColumnPos(11)
			C_InspUnitIndctnNm	= iCurColumnPos(12)
			C_InspSpec			= iCurColumnPos(13)
			C_LSL				= iCurColumnPos(14)
			C_USL				= iCurColumnPos(15)
			C_MsmtEqpmtNm		= iCurColumnPos(16)
			C_MsmtUnit			= iCurColumnPos(17)
			'------------- Hidden Column -----------
			C_InspMthdCd		= iCurColumnPos(18)
			C_InspUnitIndctnCd	= iCurColumnPos(19)
			C_MsmtEqpmtCd		= iCurColumnPos(20)
			'---------------------------------------
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
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
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

'------------------------------------------  OpenInspReqNo()  -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenInspReqNo()
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6
	Dim iCalledAspName
	
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
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "Q4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtInspReqNo.Value    = arrRet(0)		
	End If	
	
	frm1.txtInspReqNo.Focus
	Set gActiveElement = document.activeElement
	OpenInspReqNo = true	
End Function

'------------------------------------------  OpenInspReqNo2()  -------------------------------------------------
'	Name : OpenInspReqNo2()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo2()        
	OpenInspReqNo2 = false

	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtInspReqNo2.ClassName) = UCase(Parent.UCN_PROTECTED)  Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo2.Value)
	'###검사분류별 변경부분 Start###	
	Param4 = strInspClass 		'검사분류 
	'###검사분류별 변경부분 End###
	Param5 = "N"			'검사진행현황 
	
	iCalledAspName = AskPRAspName("q2512pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q2512pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtInspReqNo2.Focus
	If arrRet(0) <> "" Then

		SELECT CASE Trim(arrRet(34))
			CASE "I"
				Call DisplayMsgBox("223716", "X","X","X") 
				Exit Function
			CASE "D"
				Call DisplayMsgBox("223717", "X","X","X")
				Exit Function
			CASE "R"
				Call DisplayMsgBox("223718", "X","X","X") 
				Exit Function
		END SELECT

		frm1.txtInspReqNo2.value = arrRet(0)
		frm1.txtItemCd.value = arrRet(2)
		frm1.txtItemNm.value = arrRet(3)
		frm1.txtLotNo.value = arrRet(16)
		frm1.txtLotSubNo.value = arrRet(17)
		frm1.txtLotSize.Text = arrRet(18)
		'/* 2003-05 정기패치: 검사의뢰번호 LOOK UP 관련 수정 - START */
    	If frm1.hInspReqNo2.value <> Trim(frm1.txtInspReqNo2.value) Then
			frm1.vspdData.MaxRows = 0
			frm1.btnAllInspItem.disabled = False
			frm1.hInspReqNo2.value = Trim(frm1.txtInspReqNo2.value)
		End If
		'/* 2003-05 정기패치: 검사의뢰번호 LOOK UP 관련 수정 - END */
		lgBlnFlgChgValue = True
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspReqNo2 = true
End Function

'------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : Inspection Item By Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenInspItem(Byval strCode)
	Dim arrRet
	Dim Param1, Param2
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtInspReqNo2.Value) = "" then
		Call DisplayMsgBox("221323", "X", "X", "X")		'검사의뢰번호가 필요합니다 
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtInspReqNo2.value)		
	Param2 = frm1.txtLotSize.Text
	
	iCalledAspName = AskPRAspName("q2112pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "q2112pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	Call SetActiveCell(frm1.vspdData,C_InspItemCd,frm1.vspdData.ActiveRow,"M","X","X")
	If arrRet(0) <> "" Then
		If SetInspItem(arrRet) = False Then	Exit Function
	End If	
	OpenInspItem = true
End Function

'------------------------------------------  SetInspItem()  --------------------------------------------------
'	Name : SetInspItem()
'	Description : OpenInspItem Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetInspItem(Byval arrRet)
	SetInspItem = false
	
	With frm1.vspdData
		Call .SetText(C_InspItemCd,	.ActiveRow , arrRet(0))
		Call .SetText(C_InspItemNm,	.ActiveRow , arrRet(1))
		Call .SetText(C_InspOrder,	.ActiveRow , arrRet(2))
		Call .SetText(C_InspMthdCd,	.ActiveRow , arrRet(3))
		Call .SetText(C_InspMthdNm,	.ActiveRow , arrRet(4))
		Call .SetText(C_InspUnitIndctnCd,	.ActiveRow , arrRet(5))
		Call .SetText(C_InspUnitIndctnNm,	.ActiveRow , arrRet(6))
		Call .SetText(C_InspSeries,	.ActiveRow , arrRet(7))
		Call .SetText(C_SampleQty,	.ActiveRow , arrRet(8))
		Call .SetText(C_AcctncNumber,	.ActiveRow , arrRet(9))
		Call .SetText(C_RejtnNumber,	.ActiveRow , arrRet(10))
		Call .SetText(C_AccptncCoefficient,	.ActiveRow , arrRet(11))
		Call .SetText(C_MaxDefectRatio,	.ActiveRow , arrRet(12))
		Call .SetText(C_InspSpec,	.ActiveRow , arrRet(13))
		Call .SetText(C_LSL,		.ActiveRow , arrRet(14))
		Call .SetText(C_USL,		.ActiveRow , arrRet(15))
		Call .SetText(C_MsmtEqpmtCd,.ActiveRow , arrRet(16))
		Call .SetText(C_MsmtEqpmtNm,.ActiveRow , arrRet(17))
		Call .SetText(C_MsmtUnit,	.ActiveRow , arrRet(18))
		
		Call vspdData_Change(.Col, .Row)		 ' 변경이 읽어났다고 알려줌 
		Call SetActiveCell(frm1.vspdData,C_InspItemCd,frm1.vspdData.ActiveRow,"M","X","X")
	End With
	Set gActiveElement = document.activeElement
	SetInspItem = true
End Function

'=============================================  2.5.1 LoadInspDetails()  ======================================
'=	Event Name : LoadInspDetails
'=	Event Desc :
'========================================================================================================
Function LoadInspDetails()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	With frm1
		Parent.WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		Parent.WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		Parent.WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With	
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 LoadDefectType()  ======================================
'=	Event Name : LoadDefectType
'=	Event Desc :
'========================================================================================================
Function LoadDefectType()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	With frm1
		Parent.WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		Parent.WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		Parent.WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With	
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=============================================  2.5.5 LoadRemote()  ======================================
'=	Event Name : LoadRemote
'=	Event Desc :
'========================================================================================================
Function LoadRemote()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	With frm1
		Parent.WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		Parent.WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		Parent.WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	PgmJump(BIZ_PGM_JUMP3_ID)
End Function

'=============================================  2.5.3 LoadDecision()  ======================================
'=	Event Name : LoadDecision
'=	Event Desc :
'========================================================================================================
Function LoadDecision()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	With frm1
		Parent.WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		Parent.WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		Parent.WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	PgmJump(BIZ_PGM_JUMP4_ID)
End Function

'=============================================  2.5.4 LoadDisposition()  ======================================
'=	Event Name : LoadDisposition
'=	Event Desc :
'========================================================================================================
Function LoadDisposition()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	With frm1
		Parent.WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		Parent.WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		Parent.WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	PgmJump(BIZ_PGM_JUMP5_ID)
End Function

'=============================================  2.5.6 LoadRelease()  ======================================
'=	Event Name : LoadRelease
'=	Event Desc :
'========================================================================================================
Function LoadRelease()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	With frm1
		Parent.WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		Parent.WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		Parent.WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	PgmJump(BIZ_PGM_JUMP6_ID)
End Function

'=============================================  SetAllInspStandard()  ======================================
'=	Event Name : SetAllInspStandard
'=	Event Desc :
'========================================================================================================
Sub SetAllInspStandard()
	
	Dim strVal
	
	Err.Clear                                                               					'☜: Protect system from crashing
	'/* 2003-05 정기패치: 검사의뢰번호 LOOK UP 관련 수정 - START */
	If CheckRunningBizProcess = True Then Exit Sub
			
	Call LayerShowHide(1)
	
	With frm1	
		.vspdData.MaxRows = 0
			
		strVal = BIZ_PGM_QRY2_ID & "?txtInspReqNo=" & Trim(.txtInspReqNo2.Value) _
								 & "&txtLotSize=" & .txtLotSize.Text _
								 & "&txtMaxRows=" & .vspdData.MaxRows		
	End With		
	'/* 2003-05 정기패치: 검사의뢰번호 LOOK UP 관련 수정 - END */
	
	Call RunMyBizASP(MyBizASP, strVal)
End Sub

'=============================================  SetAllInspStandardOk()  ======================================
'=	Event Name : SetAllInspStandardOk
'=	Event Desc :
'========================================================================================================
Sub SetAllInspStandardOk()
	Dim lRow 
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = 0
			.Text = ggoSpread.InsertFlag
		Next 
	End With
	ggoSpread.SpreadUnLock C_InspItemCd, 1, C_InspItemCd
	ggoSpread.SpreadUnLock C_InspItemPopup, 1, C_InspItemPopup
	Call SetSpreadColor(1, frm1.vspdData.MaxRows)
	ggoSpread.SpreadUnLock C_SampleQty, 1, -1
	ggoSpread.SSSetRequired C_SampleQty, 1, -1
	ggoSpread.SpreadUnLock C_AcctncNumber, 1, -1
	ggoSpread.SpreadUnLock C_RejtnNumber, 1, -1
	ggoSpread.SpreadUnLock C_AccptncCoefficient, 1, -1
	ggoSpread.SpreadUnLock C_MaxDefectRatio, 1, -1
	
	frm1.btnAllInspItem.disabled = True	
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData

	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
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
    If Row <= 0 Then Exit Sub
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
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

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetDefaultVal
	Call SetToolBar("11101101000011")
	Set gActiveElement = document.activeElement
	
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		If blnIsInspectionRequest = "True" Then
			frm1.txtInspReqNo2.focus 
		Else
			frm1.txtInspReqNo.focus 
		End If
	End If
	Set gActiveElement = document.activeElement
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
    '------ Developer Coding part (Start)
    If IsDbQueryOk = True AND frm1.btnAllInspItem.disabled = True Then
		Call ggoSpread.ReOrderingSpreadData
		Call DbQueryOk
	Else
		frm1.btnAllInspItem.disabled = False
 	End If
 	'------ Developer Coding part (End)	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode)
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_InspItemPopup Then
			.Col = C_InspItemCd
			.Row = Row
			Call OpenInspItem(.Text)
		End If    
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	With frm1
		If Col = C_InspSeries Or Col = C_SampleQty Or Col = C_AcctncNumber Or Col = C_RejtnNumber Then
			Call CheckMinNumSpread(.vspdData, Col, Row)
		End If
		ggoSpread.Source = .vspdData
    	ggoSpread.UpdateRow Row
    		
    	.vspdData.Col = Col
    End With
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub vspddata_KeyPress(KeyAscii )
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	
	'----------  Coding part  -------------------------------------------------------------
	if frm1.vspdData.MaxRows < NewTop  + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크 
		If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then Exit Sub
		
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End if    
End Sub

'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - START */
'=======================================================================================================
'   Event Name : txtInspReqNo2_OnChange
'   Event Desc : 
'=======================================================================================================
Sub txtInspReqNo2_OnChange()
	Call OnCookiesLoadingLookUpInspReqNo
End Sub

'=======================================================================================================
'   Event Name : OnCookiesLoadingLookUpInspReqNo
'   Event Desc : 
'=======================================================================================================
Sub OnCookiesLoadingLookUpInspReqNo()
	Dim strInspReqNo
	
	If CheckRunningBizProcess = True Then Exit Sub
	
    strInspReqNo = Trim(frm1.txtInspReqNo2.value)
    If strInspReqNo = "" Then 
		frm1.hInspReqNo2.value = ""
		Exit Sub
	End If
    If frm1.hInspReqNo2.value = strInspReqNo Then Exit Sub
    
    With frm1
		.vspdData.MaxRows = 0
		.btnAllInspItem.disabled = False
		
		.hInspReqNo2.value = ""

		.txtItemCd.value = ""
		.txtItemNm.value = ""
		.txtLotNo.value = ""
		.txtLotSubNo.value = ""
		.txtLotSize.Text = ""
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
'/* 2003-05 정기패치 : 검사의뢰번호 LOOK UP 기능 추가 - END */

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	
	FncQuery = False                                                        							'⊙: Processing is NG
	
	Err.Clear                                                            		   					'☜: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
    End If

	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then	Exit Function

	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")								'⊙: Clear Contents  Field
	Call InitVariables
	
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
    frm1.btnAllInspItem.disabled = False
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function			'☜: Query db data
	
	FncQuery = True		
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                            					'⊙: Processing is NG
	Err.Clear                            							'☜: Protect system from crashing
	  '-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	
	  '-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                       		'⊙: Lock  Suitable  Field
	Call InitVariables																'⊙: Initializes local global variables
	Call SetDefaultVal
	
	Call SetToolBar("11101101000011")		'⊙: 버튼 툴바 제어 
	frm1.btnAllInspItem.Disabled = False
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If 
	
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete()
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
	If IntRetCD = vbNo Then	Exit Function

	  '-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True        
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim SampleQty
    Dim AccptDecisionQty
    Dim RejtDecisionQty
    Dim i
    	
	FncSave = False                                                         					'⊙: Processing is NG
	
	Err.Clear						                                                        '☜: Protect system from crashing
	
	  '-----------------------
	'Precheck area
	'-----------------------
	If frm1.vspdData.MaxRows = 0 Then
		Call DisplayMsgBox("221337", "X", "X", "X")		'검사자료를 입력하십시오.
		Exit Function
	End If	

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
	
	  '-----------------------
	'Check content area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If Not chkField(Document, "2") Then	Exit Function
    	
    If ggoSpread.SSDefaultCheck = False Then Exit Function
    	
    With frm1
    	For i = 1 To .vspdData.MaxRows
    		.vspdData.Row = i
			.vspdData.Col = C_SampleQty
			If Trim(.vspdData.Text) <> "" Then
				SampleQty = UNICDbl(.vspdData.Text)
				If SampleQty > UNICDbl(frm1.txtLotSize.Text) Then
					Call DisplayMsgBox("220825", "X", "X", "X")		'시료수가 Lot 크기보다 클 수 없습니다.
					.vspdData.Action = 0
					Exit Function
				End If
			End If
				
			.vspdData.Col = C_InspMthdCd
			If Left(Trim(.vspdData.Text), 1) <> "2" AND SampleQty <> UNICDbl(frm1.txtLotSize.Text) Then
				'계수형인 경우 
				.vspdData.Col = C_AcctncNumber
				If Trim(.vspdData.Text) <> "" Then
					AccptDecisionQty = UNICDbl(.vspdData.Text)
					If AccptDecisionQty >= SampleQty Then
						Call DisplayMsgBox("220820", "X", "X", "X")		'합격판정개수가 시료수와 같거나 클 수 없습니다.
						.vspdData.Action = 0
						Exit Function
					End If
				End If

				.vspdData.Col = C_RejtnNumber
				If Trim(.vspdData.Text) <> "" Then
	    			RejtDecisionQty = UNICDbl(.vspdData.Text)
	    			If RejtDecisionQty > SampleQty Then
					Call DisplayMsgBox("220821", "X", "X", "X")		'불합격판정개수가 시료수보다 클 수 없습니다.
					.vspdData.Action = 0
					Exit Function
					End If
				End If

				If AccptDecisionQty >= RejtDecisionQty Then
					Call DisplayMsgBox("220822", "X", "X", "X")		'불합격판정개수는 합격판정개수보다 커야 합니다.
					.vspdData.Action = 0
					Exit Function
				End If
			End If
			
			.vspdData.Col = C_AccptncCoefficient
   			If Trim(.vspdData.Text) <> "" Then
				If Not IsNumeric(Trim(.vspdData.Text)) Then
					Call DisplayMsgBox("220823", "X", "X", "X")		'합격판정계수에는 숫자를 입력하셔야 합니다.
					.vspdData.Action = 0
					Exit Function
				End If
			End If
			.vspdData.Col = C_MaxDefectRatio
			If Trim(.vspdData.Text) <> "" Then
				If Not IsNumeric(Trim(.vspdData.Text)) Then
					Call DisplayMsgBox("220824", "X", "X", "X")		'최대허용불량률에는 숫자를 입력하셔야 합니다.
					.vspdData.Action = 0
					Exit Function
				End If
			End If
    	Next
    End With
	  '-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then Exit Function            		                '☜: Save db data
	
	FncSave = True                                                        					  '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel= false
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo    
	                             					            '☜: Protect system from crashing
	If frm1.vspdData.MaxRows = 0 Then
		frm1.btnAllInspItem.disabled = False
	End If	
	FncCancel = true
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = false
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
	End If	
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
    	ggoSpread.InsertRow .vspdData.ActiveRow, imRow
    	Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1)
		.vspdData.ReDraw = True
    End With
    
    FncInsertRow = true
    
	Call SetActiveCell(frm1.vspdData,C_InspItemCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement	
    If Err.number = 0 Then FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows
    
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 

	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgPrevNo = "" Then
	 	Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
	 	Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 _
						& "&txtInspReqNo=" & lgPrevNo						'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 

	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 _
						& "&txtInspReqNo=" & lgNextNo						'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	
	FncExit = True

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Dim strVal
	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	Call LayerShowHide(1)	
	DbDelete = False									'⊙: Processing is NG
	
	strVal = BIZ_PGM_DEL_ID & "?txtInspReqNo=" & Trim(frm1.txtInspReqNo2.value)	_
							& "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)	 	
	
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()									'☆: 삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                       		'⊙: Lock  Suitable  Field
	Call InitVariables																'⊙: Initializes local global variables
	Call SetDefaultVal
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If 
	Call SetToolBar("11101101000111")		'⊙: 버튼 툴바 제어 
	frm1.btnAllInspItem.Disabled = False
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	Call LayerShowHide(1)
	
	DbQuery = False
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtInspReqNo=" & .hInspReqNo.Value _
									& "&txtPlantCd=" & .hPlantCd.Value _
									& "&lgStrPrevKey1=" & lgStrPrevKey1 _
									& "&lgStrPrevKey2=" & lgStrPrevKey2 _
									& "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtInspReqNo=" & Trim(.txtInspReqNo.Value) _
									& "&txtPlantCd=" & Trim(.txtPlantCd.Value) _
									& "&lgStrPrevKey1=" & lgStrPrevKey1 _
									& "&lgStrPrevKey2=" & lgStrPrevKey2 _
									& "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)					'☜: 비지니스 ASP 를 가동 
		
		DbQuery = True                                                          			'⊙: Processing is NG
	End With		
	
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
		
	DbQuery = True                                                          				'⊙: Processing is NG
	
End Function

'========================================================================================
' Function Name : DbQueryOk

' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE									'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    	
    Call SetSpreadColor(1, frm1.vspdData.MaxRows)
    ggoSpread.SSSetProtected C_InspItemCd, 1, -1
    ggoSpread.SSSetProtected C_InspItemPopup, 1, -1
	'검사Master작성 및 검사내역 등록 중 일 때만 수정 가능 
	If frm1.hStatusFlag.Value = "M" Or  frm1.hStatusFlag.Value = "V" Then
		If frm1.hStatusFlag.Value = "M" Then
			ggoSpread.SpreadUnLock C_SampleQty, 1, -1
    		ggoSpread.SSSetRequired C_SampleQty, 1, -1
    	Else
    		ggoSpread.SSSetProtected C_SampleQty, 1, -1
    	End If
    	
    	ggoSpread.SpreadUnLock C_AcctncNumber, 1, -1
    	ggoSpread.SpreadUnLock C_RejtnNumber, 1, -1
    	ggoSpread.SpreadUnLock C_AccptncCoefficient, 1, -1
    	ggoSpread.SpreadUnLock C_MaxDefectRatio, 1, -1
    		
    	Call SetToolBar("11111111000111")
	Else
		ggoSpread.SSSetProtected C_SampleQty, 1, -1
    	ggoSpread.SSSetProtected C_AcctncNumber, 1, -1
    	ggoSpread.SSSetProtected C_RejtnNumber, 1, -1
    	ggoSpread.SSSetProtected C_AccptncCoefficient, 1, -1
    	ggoSpread.SSSetProtected C_MaxDefectRatio, 1, -1
		Call SetToolBar("11110000000111")
	End If
	
    Call ggoOper.LockField(Document, "Q")								'⊙: This function lock the suitable field
    	
    If frm1.vspdData.MaxRows = 0 Then 
		frm1.btnAllInspItem.disabled = False
	Else
		frm1.btnAllInspItem.disabled = True
	End If
	
	IsDbQueryOk = True
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lInsCnt     
	Dim lDelCnt     
	Dim arrVal 
	Dim arrDel
	Dim ColSep
	Dim RowSep
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		    
		'-----------------------
		'Data manipulate area
		'-----------------------
		ColSep = Parent.gColSep
		RowSep = Parent.gRowSep
		
		lInsCnt = 0
 		lDelCnt = 0
   
		ReDim arrVal(0) 
		ReDim arrDel(0)

		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
			
				Case ggoSpread.InsertFlag					'☜: 신규 

					ReDim Preserve arrVal(lInsCnt) 
					
					arrVal(lInsCnt) = "C" & ColSep _
										  & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_InspSeries,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_SampleQty,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_AcctncNumber,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_RejtnNumber,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_AccptncCoefficient,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_MaxDefectRatio,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_MsmtEqpmtCd,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_MsmtUnit,lRow,"X","X") & ColSep _
										  & CStr(lRow) & RowSep		'10
					lInsCnt = lInsCnt + 1

				Case ggoSpread.UpdateFlag				'☜: 수정 

					ReDim Preserve arrVal(lInsCnt) 
					
					arrVal(lInsCnt) = "U" & ColSep _
										  & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_InspSeries,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_SampleQty,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_AcctncNumber,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_RejtnNumber,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_AccptncCoefficient,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_MaxDefectRatio,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_MsmtEqpmtCd,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_MsmtUnit,lRow,"X","X") & ColSep _
										  & CStr(lRow) & RowSep		'10

					lInsCnt = lInsCnt + 1

				Case ggoSpread.DeleteFlag				'☜: 삭제 
					ReDim Preserve arrDel(lDelCnt) 
					
					arrDel(lDelCnt) = "D" & ColSep _
										  & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & ColSep _
										  & GetSpreadText(.vspdData,C_InspSeries,lRow,"X","X") & ColSep _
										  & CStr(lRow) & RowSep
										  
					lDelCnt = lDelCnt + 1

			End Select
		Next

		.txtMaxRows.value = lDelCnt + lInsCnt
		.txtSpread.value = Join(arrDel, "") & Join(arrVal, "")

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'☜: 비지니스 ASP 를 가동 
		
	End With
	DbSave = True 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================
Function DbSaveOk()
	With frm1															'☆: 저장 성공후 실행 로직 
		.txtInspReqNo.value = .txtInspReqNo2.value 
	End With
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>최종검사 등록</FONT></TD>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20  MAXLENGTH=18 ALT="검사의뢰번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo()"></TD>							
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
									<FIELDSET CLASS="CLSFLD"><LEGEND>검사의뢰내용</LEGEND>
										<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo2" SIZE=20 MAXLENGTH=18 ALT="검사의뢰번호" tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo2()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()"></TD>
												<TD CLASS="TD5" NOWRAP></TD>
				                							<TD CLASS="TD6" NOWRAP></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>품목</TD>
				       							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="품목" tag="24">
													<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="24" ></TD>
												<TD CLASS="TD5" NOWRAP></TD>
												<TD CLASS="TD6" NOWRAP></TD>
											</TR>
				                						<TR>
				                							<TD CLASS="TD5" NOWRAP>로트번호</TD>
											   	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=15 MAXLENGTH=12 ALT="LOT NO" tag="24">
				                								<INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=10 MAXLENGTH=5 tag="24" STYLE="Text-Align: Right"></TD>
				                							<TD CLASS="TD5" NOWRAP>로트크기</TD>            
												<TD CLASS="TD6" NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtLotSize CLASS=FPDS140 title=FPDOUBLESINGLE ALT="LOT SIZE" tag="24X3"> <PARAM Name="AllowNull" Value="-1"> <PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
												</TD>
				                						</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<TR>
								<TD>
									<FIELDSET CLASS="CLSFLD"><LEGEND>검사결과</LEGEND>
										<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>검사수</TD>            
												<TD CLASS="TD6" NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtInspQty CLASS=FPDS140 title=FPDOUBLESINGLE ALT="검사수" tag="24X3"> <PARAM Name="AllowNull" Value="-1"> <PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
												</TD>
												<TD CLASS="TD5" NOWRAP>불량수</TD>
												<TD CLASS="TD6" NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtDefectQty CLASS=FPDS140 title=FPDOUBLESINGLE ALT="불량수" tag="24X3"> <PARAM Name="AllowNull" Value="-1"> <PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>판정</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtDecision" SIZE=20 MAXLENGTH=40  ALT="판정" tag="24"></TD>
												<TD CLASS="TD5" NOWRAP>검사진행상태</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtStatusFlag" SIZE=20 MAXLENGTH=40 ALT="검사진행상태" tag="24"></TD>
											</TR>
										</TABLE>
									</FIELDSET>		
								</TD>	
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD"> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	        				<TD><BUTTON NAME="btnAllInspItem" CLASS="CLSMBTN" ONCLICK="vbscript:SetAllInspStandard()">검사기준 불러오기</BUTTON></TD>
        					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspDetails">검사내역</A>&nbsp;|&nbsp;<A href="vbscript:LoadDefectType">불량유형</A>&nbsp;|&nbsp;<A href="vbscript:LoadRemote">불량원인</A>&nbsp;|&nbsp;<A href="vbscript:LoadDecision">판정</A>&nbsp;|&nbsp;<A href="vbscript:LoadDisposition">부적합처리</A>&nbsp;|&nbsp;<A href="vbscript:LoadRelease">Release</A></TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspReqNo" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspReqNo2" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hPlantCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspItemCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspSeries" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hStatusFlag" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

