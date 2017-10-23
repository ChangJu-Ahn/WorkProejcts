<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC400MA1
'*  4. Program Name         : 납입지시취소 
'*  5. Program Desc         : 납입지시취소 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003-02-25
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "mc400mb1.asp"							'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "mc400mb2.asp"							'☆: 비지니스 로직 ASP명 

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

' Grid 1(vspdData) - Operation
Dim C_Select		
Dim C_ProdOrderNo	'제조오더번호 
Dim C_ItemCd		'품목 
Dim C_ItemNm		'품목명 
Dim C_Spec			'규격 
Dim C_ReqDt			'필요일 
Dim C_ReqQty		'필요량 
Dim C_BaseUnit		'필요단위 
Dim C_DoQty			'필요납입지시량 
Dim C_BpCd			'공급처 
Dim C_BpNm			'공급처명 
Dim C_PoNo			'발주번호 
Dim C_PoSeqNo		'발주순번 
Dim C_DoQtyPoUnit	'발주납입지시량 
Dim C_PoUnit		'발주단위 
Dim C_OprNo			'공정 
Dim C_Seq			'부품예약순서 
Dim C_SubSeq		'납입지시순번 
Dim C_WcCd			'작업장 
Dim C_WcNm			'작업장명 
Dim C_PlanStartDt	'착수예정일 
Dim C_PlanComptDt	'완료예정일 
Dim C_ReleaseDt		'작업지시일 
Dim C_TrackingNo	'TrackingNo

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim lgBlnFlgChgValue							<%'Variable is for Dirty flag%>
Dim lgIntGrpCount								<%'Group View Size를 조사할 변수 %>
Dim lgIntFlgMode								<%'Variable is for Operation Status%>
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgSortKey

Dim strDate
Dim iDBSYSDate

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgSortKey = 1
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtDoDate.text		= UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
	
	Call SetToolbar("1110000000001111")
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","M","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122", ,Parent.gAllowDragDropSpread
				
		.ReDraw = false
				
		.MaxCols = C_TrackingNo + 1    
		.MaxRows = 0    
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck	C_Select ,		"선택"			,8,,,1
		ggoSpread.SSSetEdit 	C_ProdOrderNo,  "제조오더번호"	,20
		ggoSpread.SSSetEdit 	C_ItemCd,       "품목"			,20
		ggoSpread.SSSetEdit 	C_ItemNm,       "품목명"		,20
		ggoSpread.SSSetEdit 	C_Spec,			"규격"			,20
		ggoSpread.SSSetEdit 	C_ReqDt,		"필요일"		,12, 2
		ggoSpread.SSSetFloat	C_ReqQty,		"필요량"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BaseUnit,		"필요단위"		,8
		ggoSpread.SSSetFloat	C_DoQty,		"납입지시량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BpCd,			"공급처"		,15
		ggoSpread.SSSetEdit 	C_BpNm,			"공급처명"		,20
		ggoSpread.SSSetEdit 	C_PoNo,			"발주번호"		,20
		ggoSpread.SSSetEdit 	C_PoSeqNo,		"발주순번"		,10, 1
		ggoSpread.SSSetFloat	C_DoQtyPoUnit,	"발주단위납입지시량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_PoUnit,		"발주단위"		,8
		ggoSpread.SSSetEdit 	C_OprNo,		"공정"			,8
		ggoSpread.SSSetEdit 	C_Seq,			"부품예약순서"	,16
		ggoSpread.SSSetEdit 	C_SubSeq,		"납입지시순번"	,16
		ggoSpread.SSSetEdit 	C_WcCd,			"작업장"		,10
		ggoSpread.SSSetEdit 	C_WcNm,			"작업장명"		,20
		ggoSpread.SSSetEdit 	C_PlanStartDt,  "착수예정일"	,12, 2
		ggoSpread.SSSetEdit 	C_PlanComptDt,	"완료예정일"	,12, 2
		ggoSpread.SSSetEdit 	C_ReleaseDt,	"작업지시일"	,12, 2
		ggoSpread.SSSetEdit 	C_TrackingNo,	"Tracking No."	,20
		
		Call ggoSpread.SSSetColHidden(C_Seq,		C_Seq,		True)
		Call ggoSpread.SSSetColHidden(C_SubSeq,		C_SubSeq,	True)
		Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols,	True)
		
		ggoSpread.SSSetSplit2(3)
		
		Call SetSpreadLock 
		
		.ReDraw = true    
    End With
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	Dim i
	ggoSpread.Source = frm1.vspdData
	
	For i=2 To frm1.vspdData.MaxCols
		ggoSpread.SSSetProtected i, -1, -1
	Next
End Sub

'========================== 2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2110", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDoTime, lgF0, lgF1, Chr(11))
End Sub

'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()
	C_Select		= 1
	C_ProdOrderNo	= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_ReqDt			= 6
	C_ReqQty		= 7
	C_BaseUnit		= 8
	C_DoQty			= 9
	C_BpCd			= 10
	C_BpNm			= 11
	C_PoNo			= 12
	C_PoSeqNo		= 13
	C_DoQtyPoUnit	= 14
	C_PoUnit		= 15
	C_OprNo			= 16
	C_Seq			= 17
	C_SubSeq		= 18
	C_WcCd			= 19
	C_WcNm			= 20
	C_PlanStartDt	= 21
	C_PlanComptDt	= 22
	C_ReleaseDt		= 23
	C_TrackingNo	= 24
End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_Select		= iCurColumnPos(1)
			C_ProdOrderNo	= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_ReqDt			= iCurColumnPos(6)
			C_ReqQty		= iCurColumnPos(7)
			C_BaseUnit		= iCurColumnPos(8)
			C_DoQty			= iCurColumnPos(9)
			C_BpCd			= iCurColumnPos(10)
			C_BpNm			= iCurColumnPos(11)
			C_PoNo			= iCurColumnPos(12)
			C_PoSeqNo		= iCurColumnPos(13)
			C_DoQtyPoUnit	= iCurColumnPos(14)
			C_PoUnit		= iCurColumnPos(15)
			C_OprNo			= iCurColumnPos(16)
			C_Seq			= iCurColumnPos(17)
			C_SubSeq		= iCurColumnPos(18)
			C_WcCd			= iCurColumnPos(19)
			C_WcNm			= iCurColumnPos(20)
			C_PlanStartDt	= iCurColumnPos(21)
			C_PlanComptDt	= iCurColumnPos(22)
			C_ReleaseDt		= iCurColumnPos(23)
			C_TrackingNo	= iCurColumnPos(24)
    End Select
End Sub    

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'----------------------------------------------------------------------------------------------------------
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
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""									'ProdFromDt
	arrParam(2) = ""									'ProdToDt
	arrParam(3) = ""									'From Status
	arrParam(4) = ""									'To Status
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = ""	
	arrParam(8) = ""									'cboOrderType
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtProdOrderNo.focus
		Exit Function
	Else
	    frm1.txtProdOrderNo.Value	= arrRet(0)
		frm1.txtProdOrderNo.focus
	End If
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	 If Trim(frm1.txtPlantCd.value) = "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	'iCalledAspName = AskPRAspName("B1B11PA1")
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'--------------------------------------  OpenTrackingInfo()  ---------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingInfo()
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = ""	
	arrParam(4) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
	End If
End Function

'--------------------------------------  OpenPoNo()  -----------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
		Dim strRet
		Dim arrParam(2)
		Dim iCalledAspName
		Dim IntRetCD
		
		If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		IsOpenPop = True
		
		arrParam(0) = "N"	'Return Flag
		arrParam(1) = "Y"	'Release Flag
		arrParam(2) = ""	'STO Flag
		
		iCalledAspName = AskPRAspName("M3111PA1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
			IsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False
		
		If strRet(0) = "" Then
			frm1.txtPoNo.focus
			Exit Function
		Else
			frm1.txtPoNo.value = strRet(0)
			frm1.txtPoNo.focus
		End If	
		
End Function

'--------------------------------------  OpenBpInfo()  -----------------------------------------------------
'	Name : OpenBpInfo()
'	Description : Open Bp Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBpInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""	'Trim(frm1.txtBpNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value    = arrRet(0)		
		frm1.txtBpNm.Value    = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

'------------------------------------------  OpenWcCd()  -------------------------------------------------
'	Name : OpenWcCd()
'	Description : Work Center Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWcCd()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	 If Trim(frm1.txtPlantCd.value) = "" Then
		Call parent.DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "작업장팝업"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = frm1.txtWcCd.value  
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & parent.FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
				 
	arrParam(5) = "작업장"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"	
    
    arrHeader(0) = "작업장"		
    arrHeader(1) = "작업장명"		
    arrHeader(2) = "작업장타입"		
    arrHeader(3) = "작업장담당자"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtWcCd.focus
		Exit Function
	Else
		frm1.txtWcCd.value = arrRet(0)
		frm1.txtWcNm.value = arrRet(1)
		frm1.txtWcCd.focus
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call InitComboBox

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.cboDoTime.focus 
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFromReqDt.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtFromReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToReqDt.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtToReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtDoDate_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDoDate_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDoDate.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtDoDate.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtDoDate_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then							'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
    Else
        
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	lgBlnFlgChgValue = True    
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData
		.Row = Row
		.Col = C_Select
		
		ggoSpread.Source = frm1.vspdData
		
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
	End With
End Sub

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    'On Error Resume Next                                   
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.cboDoTime.focus 
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
    FncNew = True                                                           

End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False														'⊙: Processing is NG
    Err.Clear																'☜: Protect system from crashing

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	If Trim(frm1.txtPlantCd.value) = "" Then frm1.txtPlantNm.value = "" 
	If Trim(frm1.txtItemCd.value) = "" Then frm1.txtItemNm.value = "" 	
	If Trim(frm1.txtBpCd.value) = "" Then frm1.txtBpNm.value = "" 	
	If Trim(frm1.txtWcCd.value) = "" Then frm1.txtWcNm.value = "" 
	
    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables														'⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'☜: Query db data
	End If
	
	Set gActiveElement = document.activeElement
    FncQuery = True															'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    On Error Resume Next                                                    '☜: Protect system from crashing
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Trim(frm1.txtPlantCd.value) = "" Then frm1.txtPlantNm.value = "" 
	If Trim(frm1.txtItemCd.value) = "" Then frm1.txtItemNm.value = "" 	
	If Trim(frm1.txtBpCd.value) = "" Then frm1.txtBpNm.value = "" 	
	If Trim(frm1.txtWcCd.value) = "" Then frm1.txtWcNm.value = "" 
	    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                      '☜: Save db data
    
	Set gActiveElement = document.activeElement
    FncSave = True 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'☜: Protect system from crashing
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
	Set gActiveElement = document.activeElement
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
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
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
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    Dim strYear1
    Dim strMonth1
    Dim strDay1
    Dim strDate1
       
    DbQuery = False
    
	Call LayerShowHide(1)
    
	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtDoDate="			& Trim(.hDoDate.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&cboDoTime="			& Trim(.hDoTime.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.hItemCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtBpCd="			& UCase(Trim(.hBpCd.value))				'☆: 조회 조건 데이타		
			strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.hProdOrderNo.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtPoNo="			& UCase(Trim(.hPoNo.value))				'☆: 조회 조건 데이타		
			strVal = strVal & "&txtWcCd="			& UCase(Trim(.hWcCd.value))				'☆: 조회 조건 데이타		
			strVal = strVal & "&txtTrackingNo="		& UCase(Trim(.hTrackingNo.value))		'☆: 조회 조건 데이타		
			strVal = strVal & "&lgStrPrevKey1="		& lgStrPrevKey1							'
			strVal = strVal & "&lgStrPrevKey2="		& lgStrPrevKey2							'
			strVal = strVal & "&lgStrPrevKey3="		& lgStrPrevKey3
			strVal = strVal & "&lgStrPrevKey4="		& lgStrPrevKey4							'
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtDoDate="			& Trim(.txtDoDate.text)					'☆: 조회 조건 데이타 
			strVal = strVal & "&cboDoTime="			& Trim(.cboDoTime.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.txtItemCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtBpCd="			& UCase(Trim(.txtBpCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.txtProdOrderNo.value))	'☆: 조회 조건 데이타 
			strVal = strVal & "&txtPoNo="			& UCase(Trim(.txtPoNo.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtWcCd="			& UCase(Trim(.txtWcCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtTrackingNo="		& UCase(Trim(.txtTrackingNo.value))		'☆: 조회 조건 데이타		
			strVal = strVal & "&lgStrPrevKey1="		& ""
			strVal = strVal & "&lgStrPrevKey2="		& ""
			strVal = strVal & "&lgStrPrevKey3="		& ""
			strVal = strVal & "&lgStrPrevKey4="		& ""
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		End If
	End With

    Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11101000000111")														'⊙: 버튼 툴바 제어 
	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True	
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	On Error Resume Next
	Err.Clear

    Dim lRow 
    Dim strVal
	Dim lGrpCnt
	Dim igColSep,igRowSep
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
    Dim ii
	
    DbSave = False                                                          '⊙: Processing is NG
	
    LayerShowHide(1)
	
	igColSep = parent.gColSep
	igRowSep = parent.gRowSep
	strVal = ""
	lGrpCnt = 1
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

    With frm1
		.txtMode.Value = parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.Value = lgIntFlgMode									'☜: 신규입력/수정 상태 
		
    For lRow = 1 To .vspdData.MaxRows
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
            Case ggoSpread.UpdateFlag
				
				strVal = "U" & igColSep & lRow & igColSep			'☜: U=Update
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProdOrderNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X"))    & igRowSep
				
                lGrpCnt = lGrpCnt + 1
        End Select
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.UpdateFlag
		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
			 
		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
			       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		End Select   
    Next

	.txtMaxRows.value = lGrpCnt-1
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
    DbSave = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	frm1.txtPlantCd.value = frm1.hPlantCd.value
	
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
	IsOpenPop = False
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
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
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    ggoSpread.Source = gActiveSpdSheet
		
	Call ggoSpread.ReOrderingSpreadData()
End Sub 

'------------------------------------------  ChkBtnAll()  --------------------------------------------------
'	Name : ChkBtnAll()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Sub ChkBtnAll()
	Dim LngRow
	If frm1.vspdData.MaxRows <= 0 Then Exit Sub
	
	With frm1.vspdData
		For LngRow = 1 To .MaxRows
			.Row = LngRow
			.Col = C_Select
			.Value = 1

			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow LngRow
		Next
	End With
	lgBlnFlgChgValue = True
End Sub


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시확정취소</font></td>
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
			 						<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14" ALT="공장명"></TD>
									<TD CLASS=TD5 NOWRAP>납입지시일</TD> 
									<TD CLASS=TD6 NOWRAP>
										<TABLE><TR>
										<TD>
										<script language =javascript src='./js/mc400ma1_OBJECT3_txtDoDate.js'></script>
										</TD>
										<TD>
										&nbsp;
										<SELECT NAME="cboDoTime" ALT="납입지시시간" STYLE="Width: 98px;" tag="12"></SELECT>
										</TD>
										</TR></TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MaxLength=40 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>공급처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=15 MAXLENGTH=10 tag="11xxxU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MaxLength=40 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제조오더 번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="제조오더 번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS=TD5 NOWRAP>발주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="발주번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=20 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MaxLength=40 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=20 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<script language =javascript src='./js/mc400ma1_OBJECT4_vspdData.js'></script>
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
					<TD>
						<BUTTON NAME="btnCopy" CLASS="CLSMBTN" Flag=1 ONCLICK=ChkBtnAll>전체선택</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hDoDate" tag="24">
<INPUT TYPE=HIDDEN NAME="hDoTime" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hPoNo" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

