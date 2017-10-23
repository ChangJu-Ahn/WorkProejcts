<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory List onhand stock
'*  2. Function Name          : 
'*  3. Program ID             : I2321qa1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : Tracking별 재고현황조회 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2005/03/07
'*  8. Modified date(Last)    : 2005/10/28
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--########################################################################################################
						1. 선 언 부																		#
########################################################################################################-->
<!--********************************************  1.1 Inc 선언  ***************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ==================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=  "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit

'==========================================  1.2.2 Global 변수 선언  ==================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKeyIndex2
Dim IsOpenPop 

Dim gblnWinEvent
Dim strReturn
Dim lgOldRow

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "I2321qb1.asp"								
Const BIZ_PGM_QRY2_ID	= "I2321qb2.asp"								

'==========================================  1.2.1 Global 상수 선언  ======================================
 ' Grid 1(vspdData1) - Operation 
Dim C_ItemCd
Dim C_ItemNm
Dim C_TrackingNo
Dim C_Spec
Dim C_Unit
Dim C_Location
Dim C_TotQty
Dim C_TotAmt
Dim C_Price 
Dim C_PriceFlag
Dim C_PrevTotQty
Dim C_PrevTotAmt
Dim C_PrevPrice
Dim C_PrevPriceFlag

 ' Grid 2(vspdData2) - Operation 
Dim C_SlCd
Dim C_SlNm
Dim C_ItemCd2
Dim C_ItemNm2
Dim C_TrackingNo2
Dim C_GoodQty
Dim C_BadQty
Dim C_InspQty
Dim C_TransQty
Dim C_SchdRcptQty
Dim C_SchdIssueQty
Dim C_PrevGoodQty
Dim C_PrevBadQty 
Dim C_PrevInspQty
Dim C_PrevTrnsQty
Dim C_AllocationQty
Dim C_PickingQty



'==========================================  2.1.1 InitVariables()  =====================================
'	Name : InitVariables()																				=
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgIntGrpCount     = 0						
	lgBlnFlgChgValue  = False
	lgStrPrevKeyIndex = ""                      
	lgLngCurRows      = 0
    lgOldRow		  = 0
	
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End if
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtTrackingNo.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdNo)
	If pvSpdNo = "" Or pvSpdNo = "A" Then  
	
		Call InitSpreadPosVariables(pvSpdNo)

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20050307", , Parent.gAllowDragDropSpread
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1

			.ReDraw = false
				 
			.MaxCols = C_PrevPriceFlag+1											
			.MaxRows = 0
				
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit     C_ItemCd,       "품목",				18
			ggoSpread.SSSetEdit     C_ItemNm,       "품목명",			25
			ggoSpread.SSSetEdit     C_TrackingNo,   "Tracking No.",		20
			ggoSpread.SSSetEdit     C_Spec,			"규격",				20
			ggoSpread.SSSetEdit     C_Unit,			"단위",				10,2
			ggoSpread.SSSetEdit     C_Location,     "Location",			20
			ggoSpread.SSSetFloat    C_TotQty,       "현재고수량",		15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_TotAmt,       "현재고금액",		15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_Price,        "단가",				15, parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec 
			ggoSpread.SSSetEdit     C_PriceFlag,    "단가구분",			10,2
			ggoSpread.SSSetFloat    C_PrevTotQty,   "전월재고수량",		15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_PrevTotAmt,   "전월재고금액",		15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_PrevPrice,    "전월단가",			15, parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec 
			ggoSpread.SSSetEdit     C_PrevPriceFlag,"전월단가구분",		13,2	

 			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetSplit2(2)
				
			.ReDraw = true
    
		End With
	End If
		
	If pvSpdNo = "" Or pvSpdNo = "B" Then    

		Call InitSpreadPosVariables(pvSpdNo)

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20050307", , Parent.gAllowDragDropSpread

		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2

			.ReDraw = false

			.MaxCols = C_PickingQty +1										
			.MaxRows = 0

			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit     C_SlCd,         "창고",				10,,,,2
			ggoSpread.SSSetEdit		C_SlNm,         "창고명",			20
			ggoSpread.SSSetEdit		C_ItemCd2,		"품목",				18
			ggoSpread.SSSetEdit     C_ItemNm2,      "품목명",			25
			ggoSpread.SSSetEdit     C_TrackingNo2,  "Tracking No.",		25
			ggoSpread.SSSetFloat	C_GoodQty,      "양품재고량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_BadQty,       "불량재고량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_InspQty,      "검사중수량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_TransQty,     "이동중수량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_SchdRcptQty,  "입고예정량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_SchdIssueQty, "출고예정량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PrevGoodQty,  "전월양품재고량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PrevBadQty,   "전월불량재고량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PrevInspQty,  "전월검사중수량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PrevTrnsQty,	"전월이동중수량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_AllocationQty,"재고할당량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PickingQty,	"PICKING수량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			
 			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetSplit2(2)
			.ReDraw = true
		End With
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "" Or pvSpdNo = "A" Then						
	' Grid 1(vspdData1) - Operation 
		C_ItemCd             = 1
		C_ItemNm             = 2
		C_TrackingNo		 = 3
		C_Spec				 = 4
		C_Unit				 = 5
		C_Location           = 6
		C_TotQty             = 7
		C_TotAmt             = 8
		C_Price              = 9
		C_PriceFlag          = 10
		C_PrevTotQty         = 11
		C_PrevTotAmt         = 12
		C_PrevPrice          = 13
		C_PrevPriceFlag      = 14
	End If
	If pvSpdNo = "" Or pvSpdNo = "B"  Then							
		' Grid 2(vspdData2) - Operation 
		C_SlCd				= 1
		C_SlNm              = 2
		C_ItemCd2			= 3
		C_ItemNm2			= 4
		C_TrackingNo2		= 5
		C_GoodQty		    = 6
		C_BadQty		    = 7
		C_InspQty           = 8
		C_TransQty          = 9
		C_SchdRcptQty       = 10
		C_SchdIssueQty      = 11
		C_PrevGoodQty       = 12
		C_PrevBadQty        = 13
		C_PrevInspQty       = 14
		C_PrevTrnsQty		= 15	
		C_AllocationQty     = 16
		C_PickingQty		= 16
	End If
End Sub


'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData1 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		' Grid 1(vspdData1) - Operation
		C_ItemCd        = iCurColumnPos(1)
		C_ItemNm		= iCurColumnPos(2)
		C_TrackingNo	= iCurColumnPos(3)
		C_Spec			= iCurColumnPos(4)
		C_Unit			= iCurColumnPos(5)
		C_Location		= iCurColumnPos(6)
		C_TotQty		= iCurColumnPos(7)
		C_TotAmt		= iCurColumnPos(8)
		C_Price			= iCurColumnPos(9)
		C_PriceFlag		= iCurColumnPos(10)
		C_PrevTotQty	= iCurColumnPos(11)
		C_PrevTotAmt	= iCurColumnPos(12)
		C_PrevPrice		= iCurColumnPos(13)
		C_PrevPriceFlag	= iCurColumnPos(14)
		
	Case "B"
 		ggoSpread.Source = frm1.vspdData2 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

 		' Grid 2(vspdData2) - Operation
 		C_SlCd				= iCurColumnPos(1)
		C_SlNm              = iCurColumnPos(2)
		C_ItemCd2			= iCurColumnPos(3)
		C_ItemNm2			= iCurColumnPos(4)
		C_TrackingNo2		= iCurColumnPos(5)
		C_GoodQty		    = iCurColumnPos(6)
		C_BadQty		    = iCurColumnPos(7)
		C_InspQty           = iCurColumnPos(8)
		C_TransQty          = iCurColumnPos(9)
		C_SchdRcptQty       = iCurColumnPos(10)
		C_SchdIssueQty      = iCurColumnPos(11)
		C_PrevGoodQty       = iCurColumnPos(12)
		C_PrevBadQty        = iCurColumnPos(13)
		C_PrevInspQty       = iCurColumnPos(14)
		C_PrevTrnsQty		= iCurColumnPos(15)
		C_AllocationQty     = iCurColumnPos(16)
		C_PickingQty		= iCurColumnPos(17)

 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
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

'------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenItem()
' Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	 
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value," ","S")), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X", "X", "X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus 
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtItemCd.Value) 
	arrParam(2) = ""
	arrParam(3) = ""
	 
	arrField(0) = 1  
	arrField(1) = 2  
	arrField(2) = 9  
	arrField(3) = 6  

	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
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

'------------------------------------------  OpenItemAcct()  --------------------------------------------------
' Name : OpenItemAcct()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function
 
 IsOpenPop = True

 arrParam(0) = "품목계정 팝업"    
 arrParam(1) = "B_MINOR"      
 arrParam(2) = Trim(frm1.txtItemAcct.Value) 
 arrParam(3) = ""       
 arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""  
 arrParam(5) = "품목계정"   
 
 arrField(0) = "MINOR_CD"      
 arrField(1) = "MINOR_NM"      
 
 arrHeader(0) = "품목계정"     
 arrHeader(1) = "계정명"      
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtItemAcct.focus
	Exit Function
 Else
	frm1.txtItemAcct.Value		= arrRet(0)
	frm1.txtItemAcctNm.Value	= arrRet(1)
	frm1.txtItemAcct.focus
 End If 
End Function


'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Tracking No."	
	arrParam(1) = "S_SO_TRACKING"				
	arrParam(2) = Trim(frm1.txtTrackingNo.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "Tracking No."			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking No."		
    arrHeader(1) = "품목"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
	End If	
End Function


'------------------------------------------  OpenOnhandDtlRef()  -------------------------------------------------
'	Name : OpenOnhandDtlRefCode()
'	Description : OnahndStock detail Reference
'--------------------------------------------------------------------------------------------------------- 

Function OpenOnhandDtlRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1 
	Dim Param2
	Dim Param3
	Dim Param4	
	Dim Param5	
	Dim Param6
	Dim Param7 
	Dim Param8 
	Dim Param9
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.vspdData1.MaxRows = 0 and frm1.vspdData2.MaxRows = 0 Then
		Call DisplayMsgBOX("900002","X","X","X")
		frm1.txtItemCd.focus
		Exit Function
	End if

	Param4 = Trim(frm1.txtPlantCd.value)
	Param5 = "I"				
	
	ggoSpread.Source = frm1.vspdData2    

	With frm1.vspdData2	    
		If .MaxRows = 0 Then
		    Call DisplayMsgBox("169902","X", "X", "X")  
			Exit Function
		else
			.Col = C_SlCd
			.Row = .ActiveRow
			 Param1 = Trim(.Text )
			.Col = C_SlNm
			.Row = .ActiveRow
			 Param7 = Trim(.Text )
			
		End If	
    End With
    
	ggoSpread.Source = frm1.vspdData1
	
	With frm1.vspdData1	
		If .MaxRows = 0 Then
			Exit Function
		else
			.Col = C_ItemCd
			.Row = .ActiveRow
			Param2 = Trim(.Text )
			 .Col = C_TrackingNo
			.Row = .ActiveRow
			Param3 = Trim(.Text )
			.Col = C_ItemNm
			.Row = .ActiveRow
			Param8 = Trim(.Text)
			.Col = C_Unit
			.Row = .ActiveRow
			Param9 = Trim(.Text)
			
		End If	
	
		if Param4 = "" then
			Call DisplayMsgBox("169901","X", "X", "X")   
			Exit Function
		End IF
		
	End With
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2212RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2212RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData1.focus
		Exit Function
	End If	
	
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")        

 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
 		Exit Sub
 	End If
 	
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

 	gMouseClickStatus = "SP2C"   
    
 	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
 		Exit Sub
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
End Sub  

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
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
    
    Select case gActiveSpdSheet.id
    case "vaSpread1"
		Call InitSpreadSheet("A")	
	case "vaSpread2"
		Call InitSpreadSheet("B")
	End Select

    
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc :
'==========================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, Byval NewCol, Byval NewRow, Byval Cancel)
	
	If NewRow <= 0 Or Row = NewRow Then
		Exit Sub
	End If
	
	'frm1.vspdData2.MaxRows = 0
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	lgStrPrevKeyIndex2 = ""
	If DbDtlQuery(NewRow) = False Then	
		Exit Sub
	End If	
	
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End if
	End if
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop)
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then
		If lgStrPrevKeyIndex2 <> "" Then
			If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then
				Exit Sub
			End if
		End if
	End if
End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery
	Dim strFrom
	
	
	FncQuery = False
	
	on Error resume next
	Err.Clear
	
	Call InitVariables
	
	If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkFieldByCell(frm1.txtItemCd, "A",1) Then Exit Function
    
    Call SetToolbar("11000000000111")
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
	
	'-----------------------
	'공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value,"''","S")), _
	 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  
	 Call DisplayMsgBox("125000","X", "X", "X")    
	 frm1.txtPlantNm.value = ""
	 frm1.txtPlantCd.focus 
	 Exit function
	End If
	
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	   
	frm1.txtItemNm.value = ""
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
	 
			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItemNm.value = lgF0(0)
		End If
	End If
	
	If DbQuery = False Then	
		Exit Function
	End If
	
	FncQuery = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)								
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , True)      
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()	
    FncExit = True
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

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
    Err.Clear												                   
	    
    DbQuery = False											                   
	    
    Call LayerShowHide(1)
	    
    Dim strVal
    Dim strFlag
    
    If frm1.RadioOutputType.rdoCase1.Checked Then
		strFlag = "Y"
	Else
		strFlag = "N"
	End if
	
	With frm1
		
		if lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID	&"?txtPlantCd="			& .hPlantCd.value				& _
										"&txtItemCd="			& .hItemCd.value				& _
										"&txtTrackingNo="		& .hTrackingNo.value			& _
										"&txtItemAcct="			& .hItemAcct.value				& _
										"&txtFlag="				& strFlag						& _
										"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
										"&txtMaxRows="			& .vspdData1.MaxRows

		Else
			strVal = BIZ_PGM_QRY1_ID	&"?txtPlantCd="			& Trim(.txtPlantCd.value)		& _
										"&txtItemCd="			& Trim(.txtItemCd.value)		& _
										"&txtTrackingNo="		& Trim(.txtTrackingNo.value)	& _
										"&txtItemAcct="			& Trim(.txtItemAcct.value)		& _
										"&txtFlag="				& strFlag						& _
										"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
										"&txtMaxRows="			& .vspdData1.MaxRows
		End if
	End With								

    Call RunMyBizASP(MyBizASP, strVal)						                    

    DbQuery = True                                                              

End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()															
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									
    lgStrPrevKeyIndex2 = ""
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	lgOldRow = 1
	frm1.vspdData1.focus
	
	Call DbDtlQuery(1)

End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByRef Row) 

	Dim strVal
	Dim strItemCd
	Dim strTrackingNo			
   
	DbDtlQuery = False
			
	frm1.vspdData1.Row = Row
	frm1.vspdData1.Col = C_ItemCd
	strItemCd = frm1.vspdData1.Text
	
	frm1.vspdData1.Row = Row
	frm1.vspdData1.Col = C_TrackingNo
	strTrackingNo = frm1.vspdData1.Text
	
	'strTrackingNo = frm1.txtTrackingNo.Value
			
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("125000","X", "X", "X")
		frm1.txtPlantCd.focus 
		Exit Function
	End if 

	Call LayerShowHide(1)
	'"&txtTrackingNo="      & strTrackingNo					& _
	strVal = BIZ_PGM_QRY2_ID	& "?txtMode="			 & parent.UID_M0001					& _
								  "&txtPlantCd="         & Trim(frm1.txtPlantCd.value)		& _
								  "&txtItemCd="          & strItemCd						& _
								  "&txtTrackingNo="      & strTrackingNo					& _
								  "&lgStrPrevKeyIndex="	 & lgStrPrevKeyIndex2				& _
								  "&txtMaxRows="         & frm1.vspdData2.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)										

	DbDtlQuery = True

End Function

Function DbDtlQueryOk()											
	frm1.vspdData1.focus					
End Function




'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "MA") %>
End Sub
 
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
	Call InitSpreadSheet("")    
    Call InitVariables
    Call SetDefaultVal
    
    Call SetToolbar("11000000000011")

End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%> >
		</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE  <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Tracking No.별 재고현황</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenOnhandDtlRef()">재고상세조회</A></TD>					
					<TD WIDTH=10></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> >
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
					<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%> >
						<TR>
							<TD CLASS="TD5">공장</TD>
							<TD CLASS="TD6">
								<INPUT NAME="txtPlantCd" CLASS=required STYLE="Text-Transform: uppercase" TYPE="TEXT" MAXLENGTH=4 tag="12XXXU" ALT="공장" SIZE=10 ><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" CLASS=protected readonly=true TABINDEX="-1" TYPE="TEXT" MAXLENGTH="40" SIZE=35 tag="14N"></TD>
							<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
							<TD CLASS="TD6" NOWRAP>
								<INPUT NAME="txtTrackingNo" CLASS=required TYPE="TEXT" MAXLENGTH=25 tag="12XXXU" ALT="Tracking No." SIZE=25 ><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenTrackingNo()"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5">품목계정</TD>
							<TD CLASS="TD6">
								<INPUT NAME="txtItemAcct" CLASS=required TYPE="TEXT" MAXLENGTH=2 tag="12XXXU" ALT="품목계정" SIZE=10 ><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenItemAcct()">&nbsp;<INPUT NAME="txtItemAcctNm" CLASS=protected readonly=true TABINDEX="-1" TYPE="TEXT" MAXLENGTH="40" SIZE=35 tag="14N"></TD>
							<TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6" NOWRAP><INPUT NAME="txtItemCd" STYLE="Text-Transform: uppercase" TYPE="TEXT" MAXLENGTH=18 tag="11XXXU" ALT="품목" SIZE=15 ><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT NAME="txtItemNm" CLASS=protected readonly=true TABINDEX="-1" TYPE="TEXT" MAXLENGTH="40" SIZE=35 tag="14N"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5">수량유무</TD>
							<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">수량있음</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">전품목</LABEL>
							</TD>
							<TD CLASS="TD5" NOWRAP></TD>
							<TD CLASS="TD6" NOWRAP></TD>
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> >
							</TD>
						</TR>
					</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="70%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 width="100%" tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>	
								</TD>
							</TR>
							<TR HEIGHT="30%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" NAME=vspdData2 WIDTH="100%" tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>	
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> >
		</TD>	
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0>
		</IFRAME>
		</TD>	
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>	
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>	
	
	
	
	
	
