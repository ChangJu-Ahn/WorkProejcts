<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : MRP 전개 근거 조회 
'*  3. Program ID           : p2342ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/12/06
'*  9. Modifier (First)     : Lee Hyun Jae
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_QRY1_ID	= "p2342mb1.asp"
Const BIZ_PGM_QRY2_ID	= "p2342mb2.asp"

'==========================================================================================================

 ' Grid 1(vspdData1) - Operation 
Dim C_ItemCd		
Dim C_ItemNm		
Dim C_Spec				
Dim C_TrackingNo	
Dim C_TrackingFlg
Dim C_ProcType
Dim C_MRPController
Dim C_ProdMgr		
Dim C_PurOrg		
Dim C_PurOrg_Nm
Dim C_PurGrp
Dim C_PurGrp_Nm
Dim C_Suppl		
Dim C_Suppl_Nm   

 ' Grid 2(vspdData2) - Operation 
Dim C_Date			
Dim C_IssuePlanQty	
Dim C_DependencyRequiredQty
Dim C_RcptPlanQty			
Dim C_InvQty				
Dim C_PlanQty				

'==========================================================================================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'==========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim  lgStrPrevKey11
Dim  lgStrPrevKey12
Dim  lgStrPrevKey2

'========================================================================================================= 
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
         

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  
	If pvSpdNo = "A" Then							'☜: 대상이 vspdData1일때 
			 ' Grid 1(vspdData1) - Operation 
		C_ItemCd				= 1
		C_ItemNm				= 2
		C_Spec					= 3
		C_TrackingNo			= 4
		C_TrackingFlg			= 5
		C_ProcType				= 6
		C_MRPController			= 7
		C_ProdMgr				= 8
		C_PurOrg				= 9
		C_PurOrg_Nm				= 10
		C_PurGrp				= 11
		C_PurGrp_Nm				= 12
		C_Suppl					= 13
		C_Suppl_Nm				= 14
	
	ElseIf pvSpdNo = "B" Then						'☜: 대상이 vspdData2일때 
		 ' Grid 2(vspdData2) - Operation 
		C_Date					= 1
		C_IssuePlanQty			= 2
		C_DependencyRequiredQty	= 3
		C_RcptPlanQty			= 4
		C_InvQty				= 5
		C_PlanQty				= 6
	Else											'☜: 대상이 모든 Spread일때 
			 ' Grid 1(vspdData1) - Operation 
		C_ItemCd				= 1
		C_ItemNm				= 2
		C_Spec					= 3
		C_TrackingNo			= 4
		C_TrackingFlg			= 5
		C_ProcType				= 6
		C_MRPController			= 7
		C_ProdMgr				= 8
		C_PurOrg				= 9
		C_PurOrg_Nm				= 10
		C_PurGrp				= 11
		C_PurGrp_Nm				= 12
		C_Suppl					= 13
		C_Suppl_Nm				= 14
		 ' Grid 2(vspdData2) - Operation 
		C_Date					= 1
		C_IssuePlanQty			= 2
		C_DependencyRequiredQty	= 3
		C_RcptPlanQty			= 4
		C_InvQty				= 5
		C_PlanQty				= 6
	End If

End Sub

'==========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
 Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0 
        
    lgStrPrevKey11 = ""
    lgStrPrevKey12 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey    = 1
End Sub

'==========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
 Sub InitSpreadSheet(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Then
		Call initSpreadPosVariables(pvSpdNo)  
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20031128",,parent.gAllowDragDropSpread    
    
		.ReDraw = false
		    
		.MaxCols = C_Suppl_Nm + 1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm,       "품목명"		, 25
		ggoSpread.SSSetEdit		C_Spec,			"규격"			, 25
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No."	, 25
		ggoSpread.SSSetEdit		C_TrackingFlg,	"Tracking No."	, 15
		ggoSpread.SSSetEdit 	C_ProcType,		"조달구분"		, 10
		ggoSpread.SSSetEdit		C_MRPController,"MRP 담당자", 12
		ggoSpread.SSSetEdit		C_ProdMgr,		"생산담당자", 12
		ggoSpread.SSSetEdit		C_PurOrg,		"구매조직", 12
		ggoSpread.SSSetEdit		C_PurOrg_Nm,	"구매조직명", 12
		ggoSpread.SSSetEdit		C_PurGrp,		"구매그룹", 12
		ggoSpread.SSSetEdit		C_PurGrp_Nm,	"구매그룹명", 12
		ggoSpread.SSSetEdit		C_Suppl,		"공급처", 12
		ggoSpread.SSSetEdit		C_Suppl_Nm,		"공급처명", 12
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		Call ggoSpread.SSSetColHidden(C_TrackingFlg ,C_TrackingFlg ,True)
    
		.ReDraw = true
    
		End With
	
	ElseIf pvSpdNo = "B" Then
		Call initSpreadPosVariables(pvSpdNo)  
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
	
		.ReDraw = false
	
		.MaxCols = C_PlanQty + 1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetDate 		C_Date,		 			"일자"	,		11, 2, parent.gDateFormat    
		ggoSpread.SSSetFloat		C_IssuePlanQty,			"출고예정"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_DependencyRequiredQty,"소요량"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_RcptPlanQty, 			"입고예정"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_InvQty, 				"전일가용재고"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_PlanQty, 				"계획수량"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
    
		ggoSpread.SSSetSplit2(1)
	
		.ReDraw = true
	
		End With
		
	Else											'☜: 대상이 모든 Spread일때 
		Call initSpreadPosVariables(pvSpdNo)  
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20031128",,parent.gAllowDragDropSpread    
    
		.ReDraw = false
		    
		.MaxCols = C_Suppl_Nm + 1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm,       "품목명"		, 25
		ggoSpread.SSSetEdit		C_Spec,			"규격"			, 25
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No."	, 25
		ggoSpread.SSSetEdit		C_TrackingFlg,	"Tracking No."	, 15
		ggoSpread.SSSetEdit 	C_ProcType,		"조달구분"		, 10
		ggoSpread.SSSetEdit		C_MRPController,"MRP 담당자", 12
		ggoSpread.SSSetEdit		C_ProdMgr,		"생산담당자", 12
		ggoSpread.SSSetEdit		C_PurOrg,		"구매조직", 12
		ggoSpread.SSSetEdit		C_PurOrg_Nm,	"구매조직명", 12
		ggoSpread.SSSetEdit		C_PurGrp,		"구매그룹", 12
		ggoSpread.SSSetEdit		C_PurGrp_Nm,	"구매그룹명", 12
		ggoSpread.SSSetEdit		C_Suppl,		"공급처", 12
		ggoSpread.SSSetEdit		C_Suppl_Nm,		"공급처명", 12
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		Call ggoSpread.SSSetColHidden(C_TrackingFlg ,C_TrackingFlg ,True)
    
		.ReDraw = true
    
		End With
		
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
	
		.ReDraw = false
	
		.MaxCols = C_PlanQty + 1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetDate 		C_Date,		 			"일자"	,		11, 2, parent.gDateFormat    
		ggoSpread.SSSetFloat		C_IssuePlanQty,			"출고예정"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_DependencyRequiredQty,"소요량"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_RcptPlanQty, 			"입고예정"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_InvQty, 				"전일가용재고"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C_PlanQty, 				"계획수량"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
    
		ggoSpread.SSSetSplit2(1)
	
		.ReDraw = true
	
		End With
		
	End IF 
	
	Call SetSpreadLock 
    
End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
 Sub SetSpreadLock()

    '--------------------------------
    'Grid 1
    '--------------------------------
    ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
    
    '--------------------------------
    'Grid 2
    '--------------------------------
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
			
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6, i
	Dim iProdMgrArr, iProdMgrNmArr, iMRPMgrArr, iMRPMgrNmArr
	
    On Error Resume Next
    Err.Clear
	
	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Item Account
	'-----------------------------------------------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1015", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iProdMgrArr = Split(lgF0, Chr(11))
    iProdMgrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iProdMgrArr) - 1
		Call SetCombo(frm1.cboProdMgr, UCase(iProdMgrArr(i)), iProdMgrNmArr(i))
	Next
	
	'-----------------------------------------------------------------------------------------------------
	' List Minor code for MRP담당자 
	'-----------------------------------------------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1011", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iMRPMgrArr = Split(lgF0, Chr(11))
    iMRPMgrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iMRPMgrArr) - 1
		Call SetCombo(frm1.cboMrpMgr, UCase(iMRPMgrArr(i)), iMRPMgrNmArr(i))
	Next
	
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec		    = iCurColumnPos(3)    
			C_TrackingNo    = iCurColumnPos(4)    
			C_TrackingFlg	= iCurColumnPos(5)
			C_ProcType		= iCurColumnPos(6)
			C_MRPController	= iCurColumnPos(7)
			C_ProdMgr		= iCurColumnPos(8)
			C_PurOrg		= iCurColumnPos(9)
			C_PurOrg_Nm		= iCurColumnPos(10)
			C_PurGrp		= iCurColumnPos(11)
			C_PurGrp_Nm		= iCurColumnPos(12)
			C_Suppl			= iCurColumnPos(13)
			C_Suppl_Nm		= iCurColumnPos(14)
			
		Case "B"
            ggoSpread.Source = frm1.vspdData2
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Date					= iCurColumnPos(1)
			C_IssuePlanQty			= iCurColumnPos(2)
			C_DependencyRequiredQty	= iCurColumnPos(3)    
			C_RcptPlanQty			= iCurColumnPos(4)
			C_InvQty				= iCurColumnPos(5)
			C_PlanQty				= iCurColumnPos(6)
			
    End Select    

End Sub


'-----------------------------------  OpenConItemInfo()  -------------------------------------------------
'	Name : OpenConItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemInfo(arrRet)
	End If	

End Function

'---------------------------------------------  OpenConPlant()  -----------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

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
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	Dim iCalledAspName, IntRetCD
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
	
End Function


'------------------------------------------  OpenPurOrg()  -------------------------------------------------
'	Name : OpenPurOrg()	구매조직 
'	Description : PurOrg PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPurOrg()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPurOrg.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매조직팝업"	
	arrParam(1) = "B_PUR_ORG"				
	arrParam(2) = Trim(frm1.txtPurOrg.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "구매조직"
	
    arrField(0) = "PUR_ORG"	
    arrField(1) = "PUR_ORG_NM"	
    
    arrHeader(0) = "구매조직"		
    arrHeader(1) = "구매조직명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPurOrg(arrRet)
	End If	
	
End Function


'------------------------------------------  OpenPurGrp()  -------------------------------------------------
'	Name : OpenPurGrp()	구매그룹 
'	Description : OpenPurGrp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPurGrp.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_PUR_GRP"				
	arrParam(2) = Trim(frm1.txtPurGrp.Value)
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLG = " & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPurGrp(arrRet)
	End If	
	
End Function


'------------------------------------------  OpenSuppl()  -------------------------------------------------
'	Name : OpenSuppl()	주공급처 
'	Description : OpenSuppl PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSuppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSuppl.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.txtSuppl.Value)
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN(" & FilterVar("S", "''", "S") & " , " & FilterVar("CS", "''", "S") & ")"			
	arrParam(5) = "공급처"
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"
    
    arrHeader(0) = "공급처"		
    arrHeader(1) = "공급처명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSuppl(arrRet)
	End If	
	
End Function


'------------------------------------------  OpenPeggInfo()  -------------------------------------------------
'	Name : OpenPeggInfo()
'	Description : Pegging Info. Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPeggInfo()
	Dim strItemCd
	Dim arrRet
	Dim arrParam(1), arrField(1)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

    With frm1
 
		.vspdData1.Focus
		Set gActiveElement = document.activeElement
    
		ggoSpread.Source = .vspdData1

		If .vspdData1.ActiveRow < 1 Then
			Call DisplayMsgBox("202250", "X", "X", "X")
			Exit Function
		End If
		
		Call .vspdData1.GetText(C_ItemCd, .vspdData1.ActiveRow, strItemCd)
			
	End With
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.hPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(strItemCd)   ' Item Code

	iCalledAspName = AskPRAspName("P2341RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P2341RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=1024px; dialogHeight=768px; center: Yes; help: No; resizable: No; status: No;")
'		"dialogWidth=800px; dialogHeight=500px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'--------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(ByRef arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
		.txtItemCd.focus
		Set gActiveElement = document.activeElement	
	End With
End Function

'-----------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement		
End Function

Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetPurOrg()  --------------------------------------------------
'	Name : SetPurOrg()
'	Description : PurOrg Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPurOrg(byval arrRet)
	frm1.txtPurOrg.Value    = arrRet(0)
	frm1.txtPurOrgNm.Value    = arrRet(1)				
End Function

'------------------------------------------  SetPurGrp()  --------------------------------------------------
'	Name : SetPurGrp()
'	Description : PurGrp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPurGrp(byval arrRet)
	frm1.txtPurGrp.Value    = arrRet(0)
	frm1.txtPurGrpNm.Value    = arrRet(1)				
End Function

'------------------------------------------  SetPurOrg()  --------------------------------------------------
'	Name : SetSuppl()
'	Description : Suppl Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSuppl(byval arrRet)
	frm1.txtSuppl.Value    = arrRet(0)
	frm1.txtSupplNm.Value    = arrRet(1)				
End Function

'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    
    Call InitSpreadSheet("*")   
    Call InitVariables
	Call InitComboBox
	
    Call SetToolbar("11000000000011")
    
    If parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If		
	
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

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
    
    If lgOldRow <> Row Then
		
		frm1.vspdData1.Col = 1
		frm1.vspdData1.Row = row
	
		lgOldRow = Row
		
		frm1.vspdData2.MaxRows = 0
		
		Call LayerShowHide(1)
		  		
		Call DisableToolBar(Parent.TBC_QUERY)
		
        If DbDtlQuery = False Then 
           Call RestoreToolBar()
           Exit Sub
        End If 
		
	End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )

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

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey11 <> "" Then
			Call DisableToolBar(parent.TBC_QUERY) 
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
		End If     
    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey2 <> "" Then	
			Call LayerShowHide(1)
			Call DbDtlQuery
		End If     
    End if
    
End Sub

Sub txtPlantCd_OnChange()
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If	
End Sub

Sub txtItemCd_OnChange()
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If	
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim pvSpdNo
	
    ggoSpread.Source = gActiveSpdSheet
    pvSpdNo = gActiveSpdSheet.id
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(pvSpdNo)
    
    If pvSpdNo = "A" Then
		ggoSpread.Source = frm1.vspdData1
	Else
		ggoSpread.Source = frm1.vspdData2
	End If
	
	Call ggoSpread.ReOrderingSpreadData()

End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False

    Err.Clear

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")  
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
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
    If DbQuery = False Then
		Exit Function
	End If	

    FncQuery = True	
    
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
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False) 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear
        
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&lgStrPrevKey11=" & lgStrPrevKey11
		strVal = strVal & "&lgStrPrevKey12=" & lgStrPrevKey12
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&rdoProcType=" & .hProcType.value 	
		strVal = strVal & "&cboMrpMgr=" & .hMrpMgr.value
		strVal = strVal & "&cboProdMgr=" & Trim(.hProdMgr.value)
		strVal = strVal & "&txtPurOrg=" & Trim(.hPurOrg.value)
		strVal = strVal & "&txtPurGrp=" & Trim(.hPurGrp.value)
		strVal = strVal & "&txtSuppl=" & Trim(.hSuppl.value)
    Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&lgStrPrevKey11=" & lgStrPrevKey11
		strVal = strVal & "&lgStrPrevKey12=" & lgStrPrevKey12
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		If frm1.rdoProcType(0).checked Then
			strVal = strVal & "&rdoProcType=A"
		ElseIf frm1.rdoProcType(1).checked Then
			strVal = strVal & "&rdoProcType=M"
		ElseIf frm1.rdoProcType(2).checked Then
			strVal = strVal & "&rdoProcType=P"
		ElseIf frm1.rdoProcType(3).checked Then
			strVal = strVal & "&rdoProcType=O"	
		End If
		strVal = strVal & "&cboMrpMgr=" & Trim(.cboMrpMgr.value)
		strVal = strVal & "&cboProdMgr=" & Trim(.cboProdMgr.value)
		strVal = strVal & "&txtPurOrg=" & Trim(.txtPurOrg.value)
		strVal = strVal & "&txtPurGrp=" & Trim(.txtPurGrp.value)
		strVal = strVal & "&txtSuppl=" & Trim(.txtSuppl.value)
    End If
	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()	
    Call SetToolbar("11000000000111")

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DbDtlQuery
	End If
	
	frm1.vspdData1.Focus
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 
    Dim strVal
    Dim strItemCd
    Dim strTrackingNo
    Dim strTrackingFlg    
    
	With frm1.vspdData1
		strItemCd		= Trim(GetSpreadText(frm1.vspdData1,C_ItemCd,.ActiveRow,"X","X"))	
		strTrackingNo	= UCase(Trim(GetSpreadText(frm1.vspdData1,C_TrackingNo,.ActiveRow,"X","X")))
	    strTrackingFlg	= UCase(Trim(GetSpreadText(frm1.vspdData1,C_TrackingFlg,.ActiveRow,"X","X")))
    End With
    
    DbDtlQuery = False    
    
    Err.Clear
        
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtItemCd=" & Trim(strItemCd)
		strVal = strVal & "&txtTrackingNo=" & Trim(strTrackingNo)
		strVal = strVal & "&txtTrackingFlg=" & Trim(strTrackingFlg)		
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		
	Else
				
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtItemCd=" & Trim(strItemCd)
		strVal = strVal & "&txtTrackingNo=" & Trim(strTrackingNo)
		strVal = strVal & "&txtTrackingFlg=" & Trim(strTrackingFlg)		
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		
    End If
	Call RunMyBizASP(MyBizASP, strVal)    
    
    End With
    
    DbDtlQuery = True

End Function


Function DbDtlQueryOk()	

    lgIntFlgMode = parent.OPMD_UMODE
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP전개근거조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="VBScript:OpenPeggInfo()">Pegging정보</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
			 						<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo()"></TD>
									<TD CLASS=TD5 NOWRAP>MRP담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboMrpMgr" ALT="MRP담당자" STYLE="Width: 98px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=13 MAXLENGTH=4 tag="11XXXU" ALT="구매조직"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurOrg()">&nbsp;<INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>생산담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdMgr" ALT="생산담당" STYLE="Width: 98px;" tag="11XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=13 MAXLENGTH=4 tag="11XXXU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurGrp()">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>공급처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSuppl" SIZE=13 MAXLENGTH=10 tag="11XXXU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSuppl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSuppl()">&nbsp;<INPUT TYPE=TEXT NAME="txtSupplNm" SIZE=30 tag="14"></TD>
								</TR>
									<TD CLASS=TD5 NOWRAP>조달구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType4" CLASS="RADIO" tag="1X" Value="A" CHECKED><LABEL FOR="rdoProcType4">전체</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType1" CLASS="RADIO" tag="1X" Value="M"><LABEL FOR="rdoProcType1">제조</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType2" CLASS="RADIO" tag="1X" Value="P"><LABEL FOR="rdoProcType2">구매</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType3" CLASS="RADIO" tag="1X" Value="O"><LABEL FOR="rdoProcType3">외주</LABEL></TD>			
									</TD>																			
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								<TR>
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
							<TR HEIGHT="100%">
								<TD WIDTH="40%">
									<script language =javascript src='./js/p2342ma1_A_vspdData1.js'></script>
								</TD>							
								<TD WIDTH="60%">
									<script language =javascript src='./js/p2342ma1_B_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hMrpMgr" tag="24"><INPUT TYPE=HIDDEN NAME="hProdMgr" tag="24">
<INPUT TYPE=HIDDEN NAME="hPurOrg" tag="24"><INPUT TYPE=HIDDEN NAME="hPurGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hSuppl" tag="24"><INPUT TYPE=HIDDEN NAME="hProcType" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
