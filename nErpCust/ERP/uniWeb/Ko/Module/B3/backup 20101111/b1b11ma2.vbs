Const BIZ_PGM_QRY_ID = "b1b11mb5.asp"			

Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3

Dim C_ItemCd
Dim C_ItemNm
Dim C_Unit
Dim C_Account
Dim C_ProcType
Dim C_ProdEnv
Dim C_ItemClass
Dim C_PhantomType
Dim C_MPSItemFlg
Dim C_TrackingFlg
Dim C_CollectFlg
Dim C_WcCd
Dim C_AvailableFlg
Dim C_MPSMgr
Dim C_StdTime
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_MrpFlg
Dim C_LotSizing
Dim C_VarLT
Dim C_RoundPeriod
Dim C_FixOrderQty
Dim C_MinOrderQty
Dim C_MaxOrderQty
Dim C_RoundQty
Dim C_RoundFlg
Dim C_MRPMgr
Dim C_DamperFlg
Dim C_DamperMaxQty
Dim C_DamperMinQty
Dim C_MfgOrderUnit
Dim C_MfgOrderLT
Dim C_ProdMgr
Dim C_MfgScrapRate
Dim C_PurOrderUnit
Dim C_PurOrderLT
Dim C_PurOrg
Dim C_PurScrapRate
Dim C_SLCd
Dim C_IssueType
Dim C_IssueSLCd
Dim C_IssueUnit
Dim C_LotFlg
Dim C_SFStockQty
Dim C_ReOrderPnt
Dim C_InvFlg
Dim C_OverRcptFlg
Dim C_OverRcptRate
Dim C_CycleCntPerd
Dim C_ABCFlg
Dim C_InvMgr
Dim C_PurInspType
Dim C_MfgInspType
Dim C_FinalInspType
Dim C_IssueInspType
Dim C_MfgIssueLT
Dim C_PurIssueLT
Dim C_InspecMgr
Dim C_PrcCtrlInd
Dim C_StdPrice
Dim C_MoveAvgPrice
Dim C_LineNo
Dim C_OrderFrom
Dim C_AtpLt
Dim C_CalType	
Dim C_ItemSpec
Dim C_TrackingNo

Dim lgNextNo
Dim lgPrevNo
Dim lgStrPrevKey1
Dim lgOldRow
Dim IsOpenPop
Dim gSelframeFlg							'���� TAB�� ��ġ�� ��Ÿ���� Flag
Dim gblnWinEvent							'~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
	
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCd				= 1
	C_ItemNm				= 2
	C_Unit					= 3
	C_Account				= 4
	C_ProcType				= 5
	C_ProdEnv				= 6
	C_ItemClass				= 7
	C_PhantomType			= 8
	C_MPSItemFlg			= 9   
	C_TrackingFlg			= 10
	C_CollectFlg			= 11
	C_WcCd					= 12
	C_AvailableFlg			= 13 
	C_MPSMgr				= 14
	C_StdTime				= 15
	C_ValidFromDt			= 16
	C_ValidToDt				= 17
	C_MrpFlg				= 18
	C_LotSizing				= 19
	C_VarLT					= 20
	C_RoundPeriod			= 21
	C_FixOrderQty			= 22
	C_MinOrderQty			= 23    
	C_MaxOrderQty			= 24
	C_RoundQty				= 25
	C_RoundFlg				= 26
	C_MRPMgr				= 27
	C_DamperFlg				= 28
	C_DamperMaxQty			= 29
	C_DamperMinQty			= 30
	C_MfgOrderUnit			= 31
	C_MfgOrderLT			= 32
	C_ProdMgr				= 33
	C_MfgScrapRate			= 34
	C_PurOrderUnit			= 35
	C_PurOrderLT			= 36
	C_PurOrg				= 37    
	C_PurScrapRate			= 38
	C_SLCd					= 39
	C_IssueType				= 40
	C_IssueSLCd				= 41
	C_IssueUnit				= 42
	C_LotFlg				= 43
	C_SFStockQty			= 44
	C_ReOrderPnt			= 45
	C_InvFlg				= 46
	C_OverRcptFlg			= 47
	C_OverRcptRate			= 48
	C_CycleCntPerd			= 49
	C_ABCFlg				= 50
	C_InvMgr				= 51  
	C_PurInspType			= 52
	C_MfgInspType			= 53
	C_FinalInspType			= 54
	C_IssueInspType			= 55
	C_MfgIssueLT			= 56
	C_PurIssueLT			= 57
	C_InspecMgr				= 58
	C_PrcCtrlInd			= 59
	C_StdPrice				= 60
	C_MoveAvgPrice			= 61
	C_LineNo				= 62
	C_OrderFrom				= 63
	C_AtpLt					= 64
	C_CalType				= 65
	C_ItemSpec				= 66
	C_TrackingNo			= 67
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	DeScription : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE							
	lgBlnFlgChgValue = False							
	lgIntGrpCount = 0
	lgOldRow = 0										
    lgSortKey = 1                                       '��: initializes sort direction
	lgStrPrevKey1 = ""
	
	gblnWinEvent = False
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	DeScription : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtStartDt.Text	= StartDate
	frm1.txtEndDt.Text		= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Dim i

	Call initSpreadPosVariables()
	        
    With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20051125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_TrackingNo	+ 1
		.MaxRows = 0

		Call AppendNumberPlace("7", "2", "0")
		Call AppendNumberPlace("8", "2", "2")
		Call AppendNumberPlace("9", "3", "0")
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd		,"ǰ��"				, 15
		ggoSpread.SSSetEdit		C_ItemNm		,"ǰ���"			, 30
		ggoSpread.SSSetEdit		C_Unit			,"����"				, 10
		ggoSpread.SSSetEdit		C_Account		,"ǰ�����"			, 15
		ggoSpread.SSSetEdit		C_ProcType		,"���ޱ���"			, 15
		ggoSpread.SSSetEdit		C_ProdEnv		,"��������"			, 15
		ggoSpread.SSSetEdit		C_ItemClass		,"�����ǰ��Ŭ����"	, 15
		ggoSpread.SSSetEdit		C_PhantomType	,"Phantom����"		, 15, 2
		ggoSpread.SSSetEdit		C_MPSItemFlg	,"MPSǰ��"			, 10, 2
		ggoSpread.SSSetEdit		C_TrackingFlg	,"Tracking����"		, 10, 2
		ggoSpread.SSSetEdit		C_CollectFlg	,"�ܰ�������"		, 10, 2
		ggoSpread.SSSetEdit		C_WcCd		    ,"�۾���"			, 15
		ggoSpread.SSSetEdit		C_AvailableFlg  ,"��ȿǰ��"			, 10, 2
		ggoSpread.SSSetEdit		C_MPSMgr		,"������ �˻�����", 15
		ggoSpread.SSSetFloat	C_StdTime		,"ǥ�� ST"			, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetDate		C_ValidFromDt	,"������"			, 12, 2, parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt		,"������"			, 12, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_MrpFlg		,"������������"		, 10, 2
		ggoSpread.SSSetEdit		C_LotSizing		,"Lot Sizing"		, 15
		ggoSpread.SSSetEdit		C_VarLT			,"����L/T"			, 15, 1
		ggoSpread.SSSetEdit		C_RoundPeriod	,"�ø��Ⱓ"			, 15
		ggoSpread.SSSetFloat	C_FixOrderQty	,"������������"		, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_MinOrderQty	,"�ּҿ�������"		, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_MaxOrderQty	,"�ִ��������"		, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_RoundQty		,"�ø���"			, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_RoundFlg		,"�ҿ䷮�ø�����"	, 10, 2
		ggoSpread.SSSetEdit		C_MRPMgr		,"MRP�����"		, 20
		ggoSpread.SSSetEdit		C_DamperFlg		,"Damper����"		, 10, 2
		ggoSpread.SSSetFloat	C_DamperMaxQty	,"���� L/T"			, 15, "7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_DamperMinQty	,"Damper�ּ���"		, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_MfgOrderUnit	,"������������"		, 10
		ggoSpread.SSSetEdit		C_MfgOrderLT	,"��������L/T"		, 15, 1
		ggoSpread.SSSetEdit		C_ProdMgr		,"��������"		, 15
		ggoSpread.SSSetFloat	C_MfgScrapRate	,"����ǰ��ҷ���"	, 15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_PurOrderUnit	,"���ſ�������"		, 10
		ggoSpread.SSSetEdit		C_PurOrderLT	,"���ſ���L/T"		, 15, 1
		ggoSpread.SSSetEdit		C_PurOrg		,"��������"			, 20
		ggoSpread.SSSetFloat	C_PurScrapRate	,"����ǰ��ҷ���"	, 15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_SLCd			,"�԰�â��"			, 15
		ggoSpread.SSSetEdit		C_IssueType		,"�����"			, 15
		ggoSpread.SSSetEdit		C_IssueSLCd		,"���â��"			, 15
		ggoSpread.SSSetEdit		C_IssueUnit		,"������"			, 10
		ggoSpread.SSSetEdit		C_LotFlg		,"LOT����"			, 10, 2
		ggoSpread.SSSetFloat	C_SFStockQty	,"�������"		, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_ReOrderPnt	,"������"			, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_InvFlg		,"�������üũ"		, 10, 2
		ggoSpread.SSSetEdit		C_OverRcptFlg	,"���԰���뿩��"	, 10, 2	
		ggoSpread.SSSetFloat	C_OverRcptRate	,"���԰������"		, 15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_CycleCntPerd  ,"���ǻ��ֱ�"		, 15,1
		ggoSpread.SSSetEdit		C_ABCFlg		,"ǰ��ABC����"		, 10, 2
		ggoSpread.SSSetEdit		C_InvMgr		,"�������"		, 15
		ggoSpread.SSSetEdit		C_PurInspType	,"���԰˻翩��"		, 10, 2
		ggoSpread.SSSetEdit		C_MfgInspType	,"�����˻翩��"		, 10, 2 
		ggoSpread.SSSetEdit		C_FinalInspType ,"�����˻翩��"		, 10, 2
		ggoSpread.SSSetEdit		C_IssueInspType ,"���ϰ˻翩��"		, 10, 2
		ggoSpread.SSSetEdit		C_MfgIssueLT	,"�����˻�L/T"		, 15,1
		ggoSpread.SSSetEdit		C_PurIssueLT    ,"���Ű˻�L/T"		, 15,1
		ggoSpread.SSSetEdit		C_InspecMgr     ,"���Ž� �˻�����", 15
		ggoSpread.SSSetEdit		C_PrcCtrlInd    ,"�ܰ�����"			, 10
		ggoSpread.SSSetFloat	C_StdPrice		,"ǥ�شܰ�"			, 15, parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_MoveAvgPrice	,"�̵���մܰ�"		, 15, parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_LineNo		,"���μ�"			, 6, "7", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_OrderFrom		,"������������"		, 6
		ggoSpread.SSSetFloat	C_AtpLt			,"ATP L/T"			, 8, "9", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_CalType		,"Į����Ÿ��"		, 6
		ggoSpread.SSSetEdit		C_ItemSpec		,"ǰ��԰�"			, 40
		ggoSpread.SSSetEdit		C_TrackingNo		,"Tracking No"			, 20
	
		Call ggoSpread.SSSetColHidden(C_CollectFlg, C_WcCd, True)
		Call ggoSpread.SSSetColHidden(C_MPSMgr, C_StdTime, True)
		Call ggoSpread.SSSetColHidden(C_MrpFlg, C_ItemSpec, True)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
  
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_Unit				= iCurColumnPos(3)
			C_Account			= iCurColumnPos(4)
			C_ProcType			= iCurColumnPos(5)
			C_ProdEnv			= iCurColumnPos(6)
			C_ItemClass			= iCurColumnPos(7)
			C_PhantomType		= iCurColumnPos(8)
			C_MPSItemFlg		= iCurColumnPos(9)
			C_TrackingFlg		= iCurColumnPos(10)
			C_CollectFlg		= iCurColumnPos(11)
			C_WcCd				= iCurColumnPos(12)
			C_AvailableFlg		= iCurColumnPos(13)
			C_MPSMgr			= iCurColumnPos(14)
			C_StdTime			= iCurColumnPos(15)
			C_ValidFromDt		= iCurColumnPos(16)
			C_ValidToDt			= iCurColumnPos(17)
			C_MrpFlg			= iCurColumnPos(18)
			C_LotSizing			= iCurColumnPos(19)
			C_VarLT				= iCurColumnPos(20)
			C_RoundPeriod		= iCurColumnPos(21)
			C_FixOrderQty		= iCurColumnPos(22)
			C_MinOrderQty		= iCurColumnPos(23)
			C_MaxOrderQty		= iCurColumnPos(24)
			C_RoundQty			= iCurColumnPos(25)
			C_RoundFlg			= iCurColumnPos(26)
			C_MRPMgr			= iCurColumnPos(27)
			C_DamperFlg			= iCurColumnPos(28)
			C_DamperMaxQty		= iCurColumnPos(29)
			C_DamperMinQty		= iCurColumnPos(30)
			C_MfgOrderUnit		= iCurColumnPos(31)
			C_MfgOrderLT		= iCurColumnPos(32)
			C_ProdMgr			= iCurColumnPos(33)
			C_MfgScrapRate		= iCurColumnPos(34)
			C_PurOrderUnit		= iCurColumnPos(35)
			C_PurOrderLT		= iCurColumnPos(36)
			C_PurOrg			= iCurColumnPos(37)
			C_PurScrapRate		= iCurColumnPos(38)
			C_SLCd				= iCurColumnPos(39)
			C_IssueType			= iCurColumnPos(40)
			C_IssueSLCd			= iCurColumnPos(41)
			C_IssueUnit			= iCurColumnPos(42)
			C_LotFlg			= iCurColumnPos(43)
			C_SFStockQty		= iCurColumnPos(44)
			C_ReOrderPnt		= iCurColumnPos(45)
			C_InvFlg			= iCurColumnPos(46)
			C_OverRcptFlg		= iCurColumnPos(47)
			C_OverRcptRate		= iCurColumnPos(48)
			C_CycleCntPerd		= iCurColumnPos(49)
			C_ABCFlg			= iCurColumnPos(50)
			C_InvMgr			= iCurColumnPos(51)
			C_PurInspType		= iCurColumnPos(52)
			C_MfgInspType		= iCurColumnPos(53)
			C_FinalInspType		= iCurColumnPos(54)
			C_IssueInspType		= iCurColumnPos(55)
			C_MfgIssueLT		= iCurColumnPos(56)
			C_PurIssueLT		= iCurColumnPos(57)
			C_InspecMgr			= iCurColumnPos(58)
			C_PrcCtrlInd	    = iCurColumnPos(59)
			C_StdPrice		    = iCurColumnPos(60)
			C_MoveAvgPrice		= iCurColumnPos(61)
			C_LineNo			= iCurColumnPos(62)
			C_OrderFrom			= iCurColumnPos(63)
			C_AtpLt				= iCurColumnPos(64)
			C_CalType			= iCurColumnPos(65)
			C_ItemSpec			= iCurColumnPos(66)
			C_TrackingNo			= iCurColumnPos(67)
    End Select    
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'===========================================  2.3.1 Tab Click ó��  =====================================
'=	Name : Tab Click																					=
'=	DeScription : Tab Click�� �ʿ��� ����� �����Ѵ�.													=
'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)
	
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	
	gSelframeFlg = TAB2
End Function

Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
	
	Call changeTabs(TAB3)
	
	gSelframeFlg = TAB3
End Function

'------------------------------------------  OpenConPlant()  ---------------------------------------------
'	Name : OpenConPlant()
'	DeScription : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenConItemCd()  --------------------------------------------
'	Name : OpenConItemCd()
'	DeScription : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(frm1.txtItemCd.value) 						
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"

	iCalledAspName = AskPRAspName("B1B11PA4")
	    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	DeScription : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	DeScription : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
		
	End With
End Function

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If frm1.vspdData.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	
	Call vspdData_Click(NewCol, NewRow)
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000111111")

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
        Exit Sub
    End If
	
	If lgOldRow <> Row Then
		
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row
	
		lgOldRow = Row
		  		
		With frm1
		
			.vspdData.Row = Row
			.vspdData.Col = C_ItemCd
			
	
			.txtItemCd1.value = .vspdData.Text
	
			.vspdData.Col = C_ItemNm
			.txtItemNm1.value = .vspdData.Text

			.vspdData.Col = C_Unit
			 .txtBasicUnit.value = .vspdData.Text
			 
			.vspdData.Col = C_Account
			.txtAccount.value = .vspdData.Text
			 
			.vspdData.Col = C_ProcType
			.txtProcType.value = .vspdData.Text
	
			.vspdData.Col = C_ProdEnv
			.txtProdEnv.value  = .vspdData.Text
			 
			.vspdData.Col = C_ItemClass
			.txtItemClass.value = .vspdData.Text
			 
			.vspdData.Col = C_MPSItemFlg
			If  .vspdData.Text = "Y" Then
				.rdoMPSItem1.checked = True
			Else 
				.rdoMPSItem2.checked = True
			End IF
	
			.vspdData.Col = C_TrackingFlg
			If .vspdData.Text = "Y" Then
				.rdoTrackingItem1.checked = True
			Else
				.rdoTrackingItem2.checked = True
			End If

			.vspdData.Col = C_CollectFlg
			If .vspdData.Text = "Y" Then
				.rdoCollectFlg1.checked = True
			Else
				.rdoCollectFlg2.checked = True
			End If
			 
			.vspdData.Col = C_WcCd
			.txtWorkCenter.value  = .vspdData.Text	
	
			.vspdData.Col = C_AvailableFlg
			If .vspdData.Text = "Y" Then
				.rdoAvailable1.checked = True
			Else
				.rdoAvailable2.checked = True
			End IF
			 
			.vspdData.Col = C_MPSMgr
			.txtMPSMgr.value = .vspdData.Text	
			 
			 .vspdData.Col = C_StdTime
			 .txtStdTime.value = .vspdData.Text
	
			.vspdData.Col = C_ValidFromDt
			.txtValidFromDt.text = .vspdData.Text
			 
			.vspdData.Col = C_ValidToDt
			.txtValidToDt.text = .vspdData.Text
			 
			.vspdData.Col = C_MrpFlg
			IF .vspdData.Text = "Y" Then
				.rdoMRPFlg1.checked = True 
			Else 
				.rdoMRPFlg2.checked = True 
			End IF	
	
			.vspdData.Col = C_LotSizing
			.TXTLotSizing.value = .vspdData.Text
			 
			.vspdData.Col = C_VarLT
			.txtVarLT.value = .vspdData.Text
 
			.vspdData.Col = C_RoundPeriod
			.txtRoundPeriod.value = .vspdData.Text
			 
			.vspdData.Col = C_FixOrderQty
			.txtFixOrderQty.value = .vspdData.Text
			 
			.vspdData.Col = C_MinOrderQty
			.txtMinOrderQty.value = .vspdData.Text
	
			.vspdData.Col = C_MaxOrderQty
			.txtMaxOrderQty.value = .vspdData.Text
			 
			.vspdData.Col = C_RoundQty
			.txtRoundQty.value = .vspdData.Text
			 
			.vspdData.Col = C_RoundFlg
			If .vspdData.Text = "Y" Then
				.rdoRoundFlg1.checked = True
			Else
				.rdoRoundFlg2.checked = True
			End IF
			 
			.vspdData.Col = C_MRPMgr
			.txtMRPMgr.value = .vspdData.Text
			 
			.vspdData.Col = C_DamperFlg
			If .vspdData.Text = "Y" Then
				.rdoDamperFlg1.checked = True
			Else
				.rdoDamperFlg2.checked = True 
			End IF 		
			 
			.vspdData.Col = C_DamperMaxQty
			.txtOffsetLt.value = .vspdData.Text
			 
			.vspdData.Col = C_DamperMinQty
			.txtDamperMinQty.value = .vspdData.Text

			.vspdData.Col = C_MfgOrderUnit	
			.txtMfgOrderUnit.value = .vspdData.Text
	
			.vspdData.Col = C_MfgOrderLT	
			.txtMfgOrderLT.value = .vspdData.Text
	
			.vspdData.Col = C_ProdMgr	
			.txtProdMgr.value = .vspdData.Text
			 
			.vspdData.Col = C_MfgScrapRate	
			.txtMfgScrapRate.value = .vspdData.Text
			 	 
			.vspdData.Col = C_PurOrderUnit	
			.txtPurOrderUnit.value = .vspdData.Text
	
			.vspdData.Col = C_PurOrderLT	
			.txtPurOrderLT.value = .vspdData.Text
			 
			.vspdData.Col = C_PurOrg	
			.txtPurOrg.value = .vspdData.Text
			 
			.vspdData.Col = C_PurScrapRate	
			.txtPurScrapRate.value = .vspdData.Text
			 
			.vspdData.Col = C_SLCd	
			.txtSLCd.value = .vspdData.Text
	
			.vspdData.Col = C_IssueType
			.txtIssueType.value = .vspdData.Text

			.vspdData.Col = C_IssueSLCd
			.txtIssueSLCd.value = .vspdData.Text
			 
			.vspdData.Col = C_IssueUnit
			.txtIssueUnit.value = .vspdData.Text
			 
			.vspdData.Col = C_LotFlg
			If .vspdData.Text = "Y" Then
				.rdoLotNoFlg1.checked = True
			Else
				.rdoLotNoFlg2.checked = True
			End If
	
			.vspdData.Col = C_SFStockQty
			.txtSFStockQty.value = .vspdData.Text
			 
			.vspdData.Col = C_ReOrderPnt
			.txtReorderPnt.value = .vspdData.Text
			 
			.vspdData.Col = C_InvFlg
			IF .vspdData.Text = "Y" Then
				.rdoInvCheckFlg1.checked = True
			Else
				.rdoInvCheckFlg2.checked = True 
			End IF 
	
			.vspdData.Col = C_OverRcptFlg
			IF .vspdData.Text = "Y" Then
				.rdoOverRcptFlg1.checked = True
			Else
				.rdoOverRcptFlg2.checked = True
			End If
	
			.vspdData.Col = C_OverRcptRate
			.txtOverRcptRate.value = .vspdData.Text

			.vspdData.Col = C_CycleCntPerd
			.txtCycleCntPerd.value = .vspdData.Text
			 
			.vspdData.Col = C_ABCFlg
			.txtABCFlg.value = .vspdData.Text
	
			.vspdData.Col = C_InvMgr
			.txtInvMgr.value = .vspdData.Text
	
			.vspdData.Col = C_PurInspType
			IF .vspdData.Text = "Y" Then
				.rdoPurInspType1.checked = True
			Else 
				.rdoPurInspType2.checked = True
			End IF
	
			.vspdData.Col = C_MfgInspType
			If .vspdData.Text = "Y" Then
				.rdoMfgInspType1.checked = True		
			Else
				.rdoMfgInspType2.checked = True
			End If
	
			.vspdData.Col = C_FinalInspType
			If .vspdData.Text = "Y" Then
				.rdoFinalInspType1.checked = True
			Else
				.rdoFinalInspType2.checked = True
			End IF
	
			.vspdData.Col = C_IssueInspType
			If  .vspdData.Text = "Y" Then
				.rdoIssueInspType1.checked = True
			Else
				.rdoIssueInspType2.checked = True
			End IF
				 
			.vspdData.Col = C_MfgIssueLT
			.txtMfgInspLT.value = .vspdData.Text
	
			.vspdData.Col = C_PurIssueLT
			.txtPurInspLT.value = .vspdData.Text
			 
			.vspdData.Col = C_InspecMgr
			.txtInspecMgr.value = .vspdData.Text
	
			.vspdData.Col = C_PrcCtrlInd
			.txtPrcCtrlInd.value = .vspdData.Text
	
			.vspdData.Col = C_StdPrice
			.txtStdPrice.value = .vspdData.Text
	
			.vspdData.Col = C_MoveAvgPrice
			.txtMoveAvgPrice.value = .vspdData.Text
		
			.vspdData.Col = C_LineNo
			.txtLineNo.value = .vspdData.Text
	
			.vspdData.Col = C_OrderFrom
			.txtOrderFrom.value = .vspdData.Text
				
			.vspdData.Col = C_AtpLt
			.txtAtpLt.value = .vspdData.Text
		
			.vspdData.Col = C_CalType
			.txtCalType.value = .vspdData.Text
			
			.vspdData.Col = C_ItemSpec
			.txtItemSpec.value = .vspdData.Text
		
		End With   
		
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
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
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtEndDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtStartDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtEndDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'=	Event Name : Form_QueryUnload																		=
'=	Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
	
'=========================================  5.1.1 FncQuery()  ===========================================
'=	Event Name : FncQuery																				=
'=	Event Desc : This function is related to Query Button of Main ToolBar								=
'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													

	Err.Clear															

	'------ Erase contents area ------
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Exit Function
	End If
	
	Call ggoOper.ClearField(Document, "2")								
	'Call SetDefaultVal
	Call InitVariables													
	
	
	'------ Check condition area ------ 
	If Not chkField(Document, "1") Then							
		Exit Function
	End If

	'------ Query function call area ------
	If DbQuery = False Then   
		Exit Function           
    End If 
    
	FncQuery = True							
End Function

'============================================  5.1.9 FncPrint()  ========================================
'=	Event Name : FncPrint																				=
'=	Event Desc : This function is related to Print Button of Main ToolBar								=
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'===========================================  5.1.12 FncExcel()  ========================================
'=	Event Name : FncExcel																				=
'=	Event Desc : This function is related to Excel Button of Main ToolBar								=
'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function

'===========================================  5.1.13 FncFind()  =========================================
'=	Event Name : FncFind																				=
'=	Event Desc :																						=
'========================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, True)
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

'=============================================  5.2.1 DbQuery()  ========================================
'=	Event Name : DbQuery																				=
'=	Event Desc : This function is data query and display												=
'========================================================================================================
Function DbQuery()
	Dim strAvailableItem
	
	Err.Clear															

	DbQuery = False														

	LayerShowHide(1)
	
	Dim strVal
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)	
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)		
		strVal = strVal & "&cboAccount=" & Trim(frm1.hItemAccunt.value)
		strVal = strVal & "&cboProcType=" & Trim(frm1.hProcType.value)
		strVal = strVal & "&txtStartDt=" & Trim(frm1.hStartDt.value)
		strVal = strVal & "&txtEndDt=" & Trim(frm1.hEndDt.value)
		strVal = strVal & "&rdoAvailableItem=" & frm1.hAvailableItem.value	
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)	
		strVal = strVal & "&cboAccount=" & Trim(frm1.cboAccount.value)
		strVal = strVal & "&cboProcType=" & Trim(frm1.cboProcType.value)
		strVal = strVal & "&txtStartDt=" & Trim(frm1.txtStartDt.text)
		strVal = strVal & "&txtEndDt=" & Trim(frm1.txtEndDt.text)
		If frm1.rdoAvailableItem1.checked = True then
			strAvailableItem = "A"
		ElseIf frm1.rdoAvailableItem2.checked = True then
			strAvailableItem = "Y"
		Else			
			strAvailableItem = "N"
		End IF
		strVal = strVal & "&rdoAvailableItem=" & strAvailableItem
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		
	End If	
	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function

'=============================================  5.2.4 DbQueryOk()  ======================================
'=	Event Name : DbQueryOk																				=
'=	Event Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű�	=
'========================================================================================================
Function DbQueryOk()												
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		Call vspdData_Click(1,1)
	End If	
	
	lgIntFlgMode = parent.OPMD_UMODE										

	Call ggoOper.LockField(Document, "Q")							
	Call SetToolbar("11000000000111")	
	
	lgBlnFlgChgValue = False
	
	
	
End Function