<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc800qb1
'*  4. Program Name         : 납입지시현황조회 
'*  5. Program Desc         : 납입지시현황조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/24
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Woo Guen
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
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   =====================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
Const BIZ_PGM_QRY_ID = "mc800qb1.asp"								'☆: Head Query 비지니스 로직 ASP명 

Dim C_ProdtOrderNo		
Dim C_ItemCd					
Dim C_ItemNm					
Dim C_Specification	
Dim C_ReqDt	
Dim C_ReqQty		
Dim C_BaseUnit	
Dim C_DoQty		 
Dim C_RcptQty		 
Dim C_BpCd		 
Dim C_BpNm
Dim C_DoDt
Dim C_DoTime	
Dim C_DoTimeDesc	
Dim C_DoStatus	
Dim C_DoStatusDesc
Dim C_TrackingNo				
Dim C_PoNo				
Dim C_PoSeqNo				
Dim C_DoQtyPoUnit			
Dim C_RcptQtyPoUnit		
Dim C_PoUnit		
Dim C_OprNo		
Dim C_Seq					
Dim C_SubSeq						
Dim C_WcCd						
Dim C_WcNm	
Dim C_PlanStartDt			
Dim C_PlanComptDt				
Dim C_ReleaseDt	

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop										'Popup

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False					'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2115", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboDlvyOrderStatus, lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim LocSvrDate
	
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtReqFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -5, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtReqToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 10, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	Call SetToolbar("1100000000001111")
End Sub
   
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    
    Call InitSpreadPosVariables()
    
    With frm1
    
	    ggoSpread.Source = .vspdData
	    ggoSpread.Spreadinit "V20030107", , Parent.gAllowDragDropSpread

	 	.vspdData.ReDraw = false
	    .vspdData.MaxCols = C_ReleaseDt + 1
	    .vspdData.MaxRows = 0

	    Call GetSpreadColumnPos("A")

	    ggoSpread.SSSetEdit		C_ProdtOrderNo, "제조오더번호", 16,,,,2
	    ggoSpread.SSSetEdit		C_ItemCd,		"품목", 20,,,,2
	    ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 25
	    ggoSpread.SSSetEdit		C_Specification,"규격", 25
	    ggoSpread.SSSetDate 	C_ReqDt,		"필요일", 12, 2, parent.gDateFormat    
		ggoSpread.SSSetFloat	C_ReqQty,		"필요수량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_BaseUnit,		"필요단위", 10
	    ggoSpread.SSSetFloat	C_DoQty,		"필요납입지시수량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_RcptQty,		"필요단위입고수량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_BpCd,			"공급처", 10
	    ggoSpread.SSSetEdit		C_BpNm,			"공급처명", 20
	    ggoSpread.SSSetDate 	C_DoDt,			"납입지시일", 12, 2, parent.gDateFormat    
		ggoSpread.SSSetCombo	C_DoTime,		"납입지시시간", 04
		ggoSpread.SSSetEdit		C_DoTimeDesc,	"납입지시시간", 12
	    ggoSpread.SSSetCombo	C_DoStatus,		"납입지시상태", 04
		ggoSpread.SSSetEdit		C_DoStatusDesc, "납입지시상태", 12
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
		ggoSpread.SSSetEdit 	C_PoNo,			"발주번호", 20
		ggoSpread.SSSetEdit 	C_PoSeqNo,		"발주순번", 10,1
		ggoSpread.SSSetFloat	C_DoQtyPoUnit,	"발주납입지시수량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_RcptQtyPoUnit,"발주단위입고수량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_PoUnit,		"발주단위", 10
	    ggoSpread.SSSetEdit		C_OprNo,		"공정", 10
	    ggoSpread.SSSetEdit		C_Seq,			"부품예약순서", 10
	    ggoSpread.SSSetEdit		C_SubSeq,		"납입지시순번", 10
	    ggoSpread.SSSetEdit		C_WcCd,			"작업장", 10
	    ggoSpread.SSSetEdit		C_WcNm,			"작업장명", 16
	    ggoSpread.SSSetDate 	C_PlanStartDt,	"착수계획일정", 10, 2, parent.gDateFormat
	    ggoSpread.SSSetDate 	C_PlanComptDt,	"완료계획일정", 10, 2, parent.gDateFormat
	    ggoSpread.SSSetDate 	C_ReleaseDt,	"작업지시일", 10, 2, parent.gDateFormat
    
	    Call ggoSpread.SSSetColHidden(C_DoTime, C_DoTime, True)
	    Call ggoSpread.SSSetColHidden(C_DoStatus, C_DoStatus, True)
	    Call ggoSpread.SSSetColHidden(C_Seq, C_Seq, True)
	    Call ggoSpread.SSSetColHidden(C_SubSeq, C_SubSeq, True)
	    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		
		.vspdData.ReDraw = true
		
		  ggoSpread.Source = frm1.vspdData
		  ggoSpread.SpreadLockWithOddEvenRowColor()

    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_ProdtOrderNo					= 1
	C_ItemCd						= 2
	C_ItemNm						= 3
	C_Specification					= 4
	C_ReqDt							= 5	
	C_ReqQty						= 6
	C_BaseUnit						= 7
	C_DoQty							= 8
	C_RcptQty						= 9 
	C_BpCd							= 10
	C_BpNm							= 11
	C_DoDt							= 12
	C_DoTime						= 13
	C_DoTimeDesc					= 14
	C_DoStatus						= 15
	C_DoStatusDesc					= 16
	C_TrackingNo					= 17	
	C_PoNo							= 18
	C_PoSeqNo						= 19
	C_DoQtyPoUnit					= 20
	C_RcptQtyPoUnit					= 21
	C_PoUnit						= 22
	C_OprNo							= 23
	C_Seq							= 24
	C_SubSeq						= 25		
	C_WcCd							= 26	
	C_WcNm							= 27
	C_PlanStartDt					= 28
	C_PlanComptDt					= 29	
	C_ReleaseDt						= 30
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
			
			C_ProdtOrderNo				= iCurColumnPos(1)
			C_ItemCd					= iCurColumnPos(2)
			C_ItemNm					= iCurColumnPos(3)
			C_Specification				= iCurColumnPos(4)  
			C_ReqDt						= iCurColumnPos(5)  
			C_ReqQty					= iCurColumnPos(6)  
			C_BaseUnit					= iCurColumnPos(7)  
			C_DoQty						= iCurColumnPos(8)  
			C_RcptQty					= iCurColumnPos(9)  
			C_BpCd						= iCurColumnPos(10) 
			C_BpNm						= iCurColumnPos(11) 
			C_DoDt						= iCurColumnPos(12) 
			C_DoTime					= iCurColumnPos(13) 
			C_DoTimeDesc				= iCurColumnPos(14) 
			C_DoStatus					= iCurColumnPos(15) 
			C_DoStatusDesc				= iCurColumnPos(16) 
			C_TrackingNo				= iCurColumnPos(17) 
			C_PoNo						= iCurColumnPos(18) 
			C_PoSeqNo					= iCurColumnPos(19) 
			C_DoQtyPoUnit				= iCurColumnPos(20) 
			C_RcptQtyPoUnit				= iCurColumnPos(21) 
			C_PoUnit					= iCurColumnPos(22) 
			C_OprNo						= iCurColumnPos(23) 
			C_Seq						= iCurColumnPos(24) 
			C_SubSeq					= iCurColumnPos(25) 
			C_WcCd						= iCurColumnPos(26) 
			C_WcNm						= iCurColumnPos(27) 
			C_PlanStartDt				= iCurColumnPos(28) 
			C_PlanComptDt				= iCurColumnPos(29) 
			C_ReleaseDt					= iCurColumnPos(30)
 	End Select
End Sub
 
'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
	arrField(0) = "PLANT_CD"					' Field명(0)
	arrField(1) = "PLANT_NM"					' Field명(1)
	
	arrHeader(0) = "공장"					' Header명(0)
	arrHeader(1) = "공장명"					' Header명(1)
    
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
		Set gActiveElement = document.activeElement	
	End If
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim iCalledAspName
	Dim arrParam(5), arrField(2)
	
	Dim IntRetCD
	Dim arrRet

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Exit Function
	End If

    '공장 체크 함수 호출 
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	'------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","X")
		IsOpenPop = False
		Exit Function
	End If
	
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
		Set gActiveElement = document.activeElement
	End If
End Function

'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Biz Partner PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtBpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
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
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNo()
	Dim iCalledAspName
	Dim arrParam(8)
	
	Dim arrRet
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtReqFromDt.Text
	arrParam(2) = frm1.txtReqToDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "RL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value) 
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = ""
	'arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = ""
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtProdOrderNo.focus
		Exit Function
	Else
		frm1.txtProdOrderNo.Value    = arrRet(0)
		frm1.txtProdOrderNo.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim iCalledAspName
	Dim arrParam(2)
	
	Dim strRet
	Dim IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	arrParam(0) = "N"	'Return Flag
	arrParam(1) = "N"	'Release Flag
	arrParam(2) = ""	'STO Flag

	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		lblnWinEvent = False
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
		Set gActiveElement = document.activeElement
	End If	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrParam(4)
	
	Dim arrRet
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If IsOpenPop = True  Then Exit Function
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtReqFromDt.Text
	arrParam(4) = frm1.txtReqToDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
		Set gActiveElement = document.activeElement	
	End If
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     				'⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          			'⊙: Lock  Suitable  Field
    Call InitSpreadSheet 

	Call SetDefaultVal
	Call InitVariables		'⊙: Initializes local global variables
 	Call InitComboBox

	 'Plant Code, Plant Name Setting 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
         
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Sub
		If Row < 1 Then
			ggoSpread.Source = frm1.vspdData
			 
 			If lgSortKey = 1 Then
 				ggoSpread.SSSort Col					'Sort in Ascending
 				lgSortKey = 2
 			Else
 				ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 				lgSortKey = 1
			End If 
		End If
	
    End With
	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then Exit Sub			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
    If OldLeft <> NewLeft Then Exit Sub
     '----------  Coding part  -------------------------------------------------------------
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
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
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtReqFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReqFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqFromDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtReqFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReqToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReqToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtReqToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtReqFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtReqToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = "" 
	End If
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then Exit Function										'⊙: This function check indispensable field

    If ValidDateCheck(frm1.txtReqFromDt, frm1.txtReqToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function										'☜: Query db data
       
    Set gActiveElement = document.ActiveElement   
    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                           '☜: Protect system from crashing
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
   
    Err.Clear							'☜: Protect system from crashing

    DbQuery = False                                                         			'⊙: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☜: 삭제 조건 데이타 
		strVal = strVal & "&txtReqFromDt=" & Trim(frm1.hReqFromDt.value)		'☜: 삭제 조건 데이타 
		strVal = strVal & "&txtReqToDt=" & Trim(frm1.hReqToDt.value)			'☜: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				'☜: 삭제 조건 데이타 
		strVal = strVal & "&txtBpCd=" & Trim(frm1.hBpCd.value)					'☜: 삭제 조건 데이타		
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	'☜: 조회 조건 데이타 
		strVal = strVal & "&txtPoNo=" & Trim(frm1.hPoNo.value)					'☜: 조회 조건 데이타 
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'☜: 삭제 조건 데이타  
		strVal = strVal & "&cboDlvyOrderStatus=" & Trim(frm1.hDlvyOrderStatus.value)	'☜: 삭제 조건 데이타		  
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☜: 삭제 조건 데이타 
		strVal = strVal & "&txtReqFromDt=" & Trim(frm1.txtReqFromDt.text)		'☜: 삭제 조건 데이타 
		strVal = strVal & "&txtReqToDt=" & Trim(frm1.txtReqToDt.text)			'☜: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'☜: 삭제 조건 데이타 
		strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.value)					'☜: 삭제 조건 데이타		
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'☜: 조회 조건 데이타 
		strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)					'☜: 조회 조건 데이타 
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)		'☜: 삭제 조건 데이타  
		strVal = strVal & "&cboDlvyOrderStatus=" & Trim(frm1.cboDlvyOrderStatus.value)	'☜: 삭제 조건 데이타	  
	End If    

    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'☆: 조회 성공후 실행로직 
	
	Call SetToolBar("11000000000111")											'⊙: 버튼 툴바 제어 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field

	If frm1.vspdData.MaxRows <= 0 Then Exit Function

    lgIntFlgMode = Parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
	
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시현황조회</font></td>
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlantCd()">
														 <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>필요일</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/mc800qa1_OBJECT1_txtReqFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/mc800qa1_OBJECT2_txtReqToDt.js'></script>
									</TD>
								</TR>
								<TR><TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd">
														 <INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>제조오더 번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="제조오더 번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo"  SIZE=20 MAXLENGTH=18 ALT="발주번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
								</TR>
								<TR><TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value,0"></TD>
									<TD CLASS="TD5" NOWRAP>납입지시상태</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDlvyOrderStatus" ALT="납입지시상태" STYLE="Width: 165px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
								<TD HEIGHT=* WIDTH=100%>
									<script language =javascript src='./js/mc800qa1_vspdData_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hReqToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hDlvyOrderStatus" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
