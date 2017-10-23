<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2341ma1.asp
'*  4. Program Name         : MRP전개결과조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : ?
'*  8. Modified date(Last)  : 2003/12/06
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'**********************************************************************************************-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
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

'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p2341mb1.asp"

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_TrackingNo
Dim C_StartDt
Dim C_EndDt
Dim C_PlanQty
Dim C_Unit 
Dim C_ProcType
Dim C_MRPController
Dim C_ProdMgr		
Dim C_PurOrg
Dim C_PurOrg_Nm		
Dim C_PurGrp
Dim C_PurGrp_Nm
Dim C_Suppl		
Dim C_Suppl_Nm
Dim C_ItemGroupCd
Dim C_ItemGroupNm

'==========================================================================================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgStrPrevKey5

'==========================================================================================================
Dim ihGridCnt
Dim intItemCnt
Dim IsOpenPop

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,parent.gDateFormat)

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCd        =  1
	C_ItemNm        =  2
	C_Spec			=  3
	C_TrackingNo    =  4
	C_StartDt		=  5
	C_EndDt			=  6
	C_PlanQty		=  7
	C_Unit			=  8
	C_ProcType		=  9
	C_MRPController	= 10
    C_ProdMgr		= 11
    C_PurOrg		= 12
    C_PurOrg_Nm		= 13
    C_PurGrp		= 14
    C_PurGrp_Nm		= 15
    C_Suppl			= 16
    C_Suppl_Nm		= 17
	C_ItemGroupCd	= 18
	C_ItemGroupNm	= 19
End Sub
'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgStrPrevKey5 = ""
    lgLngCurRows = 0
	lgSortKey    = 1
End Sub


'==========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromPlanDt.text	= StartDate
	frm1.txtToPlanDt.text	= LastDate
	frm1.txtPlantCd.focus
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()    
	
    With frm1
    
    ggoSpread.Source = .vspdData
    ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
    
    .vspdData.Redraw = False
    
    .vspdData.MaxCols = C_ItemGroupNm + 1
    .vspdData.MaxRows = 0
    
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_ItemCd, 		"품목"			, 18
    ggoSpread.SSSetEdit 	C_ItemNm,		"품목명"		, 25
	ggoSpread.SSSetEdit 	C_Spec,         "규격"			, 25
    ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No."	, 25
    ggoSpread.SSSetDate 	C_StartDt,		"시작일"		, 11, 2, parent.gDateFormat    
    ggoSpread.SSSetDate 	C_EndDt, 		"완료일"		, 11, 2, parent.gDateFormat    
    ggoSpread.SSSetFloat	C_PlanQty, 		"계획수량"		, 15, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetEdit 	C_Unit, 		"단위"			, 7
    ggoSpread.SSSetEdit 	C_ProcType,		"조달구분"		, 10
    ggoSpread.SSSetEdit		C_MRPController,"MRP 담당자"	, 12
	ggoSpread.SSSetEdit		C_ProdMgr,		"생산담당자"	, 12
	ggoSpread.SSSetEdit		C_PurOrg,		"구매조직"		, 12
	ggoSpread.SSSetEdit		C_PurOrg_Nm,	"구매조직명"	, 12
	ggoSpread.SSSetEdit		C_PurGrp,		"구매그룹"		, 12
	ggoSpread.SSSetEdit		C_PurGrp_Nm,	"구매그룹명"	, 12
	ggoSpread.SSSetEdit		C_Suppl,		"공급처"		, 12
	ggoSpread.SSSetEdit		C_Suppl_Nm,		"공급처명"		, 12
	ggoSpread.SSSetEdit 	C_ItemGroupCd,	"품목그룹"		, 15
	ggoSpread.SSSetEdit		C_ItemGroupNm,	"품목그룹명"	, 30
    
    ggoSpread.SSSetSplit2(1)
    
    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols,True)
    
    ggoSpread.Source = .vspdData
    
    .vspdData.Redraw = True
    
    End With
    
    Call SetSpreadLock()
    
End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
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

'==========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_TrackingNo	= iCurColumnPos(4)
			C_StartDt		= iCurColumnPos(5)
			C_EndDt			= iCurColumnPos(6)
			C_PlanQty		= iCurColumnPos(7)
			C_Unit			= iCurColumnPos(8)
			C_ProcType		= iCurColumnPos(9)
			C_MRPController	= iCurColumnPos(10)
			C_ProdMgr		= iCurColumnPos(11)
			C_PurOrg		= iCurColumnPos(12)
			C_PurOrg_Nm		= iCurColumnPos(13)
			C_PurGrp		= iCurColumnPos(14)
			C_PurGrp_Nm		= iCurColumnPos(15)
			C_Suppl			= iCurColumnPos(16)
			C_Suppl_Nm		= iCurColumnPos(17)
			C_ItemGroupCd	= iCurColumnPos(18)
			C_ItemGroupNm	= iCurColumnPos(19)
			
    End Select    

End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""			' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
    
    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPlant(arrRet)
	End If	
End Function

'--------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0):"ITEM_CD"
    arrField(1) = 2 							' Field명(1):"ITEM_NM"
    
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

'-----------------------------------------  OpenTrackingInfo()  ------------------------------------------
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

'------------------------------------------  OpenItemGroup()  -------------------------------------------------
'	Name : OpenItemGroup()
'	Description : Item Group Reference PopUp
'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
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
 
		.vspdData.Focus
		Set gActiveElement = document.activeElement
    
		ggoSpread.Source = .vspdData

		If .vspdData.ActiveRow < 1 Then
			Call DisplayMsgBox("202250", "X", "X", "X")
			Exit Function
		End If
		
		Call .vspdData.GetText(C_ItemCd, .vspdData.ActiveRow, strItemCd)
			
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
		frm1.txtTrackingNo.focus
		Set gActiveElement = document.activeElement		
	End With
End Function

 '-------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement		
End Function
'=========================================================================================================
Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement
End Function
'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
	Call SetDefaultVal
	Call InitVariables
	Call InitComboBox
	
	Call SetToolBar("11000000000011")

    If parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement
End Sub
'========================================================================================================
'   Event Name : txtFromPlanDt_DblClick
'   Event Desc :
'=========================================================================================================
Sub txtFromPlanDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromPlanDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromPlanDt.Focus
	End if
End Sub

'========================================================================================================
'   Event Name : txtToPlanDt_DblClick
'   Event Desc :
'========================================================================================================
Sub txtToPlanDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPlanDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToPlanDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtFromPlanDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromPlanDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToPlanDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtToPlanDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
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

    If frm1.vspdData.MaxRows = 0 Then
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
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       Exit Sub
    End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
		End If
    End if
    
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()

    Dim IntRetCD 

    FncQuery = False

    Err.Clear
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    Call ggoOper.ClearField(Document, "2")  
    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If ValidDateCheck(frm1.txtFromPlanDt, frm1.txtToPlanDt)  = False Then		
		Exit Function
	End If  

    If DbQuery = False Then
		Exit Function
	End If	

    FncQuery = True	
    
   
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

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
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
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

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

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
		strVal = strVal & "&lgStrPrevKey5=" & lgStrPrevKey5
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtTrackingNo=" & frm1.hTrackingNo.value 
		strVal = strVal & "&rdoProcType=" & frm1.hProcType.value
		strVal = strVal & "&txtFromPlanDt=" & Trim(frm1.hFromPlanDt.value)
		strVal = strVal & "&txtToPlanDt=" & Trim(frm1.hToPlanDt.value)	
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&cboMrpMgr=" & Trim(frm1.hMrpMgr.value)
		strVal = strVal & "&cboProdMgr=" & Trim(frm1.hProdMgr.value)
		strVal = strVal & "&txtPurOrg=" & Trim(frm1.hPurOrg.value)
		strVal = strVal & "&txtPurGrp=" & Trim(frm1.hPurGrp.value)
		strVal = strVal & "&txtSuppl=" & Trim(frm1.hSuppl.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtTrackingNo=" & frm1.txtTrackingNo.value    
		If frm1.rdoProcType(0).checked Then
			strVal = strVal & "&rdoProcType=A"
		ElseIf frm1.rdoProcType(1).checked Then
			strVal = strVal & "&rdoProcType=M"
		ElseIf frm1.rdoProcType(2).checked Then
			strVal = strVal & "&rdoProcType=P"
		ElseIf frm1.rdoProcType(3).checked Then
			strVal = strVal & "&rdoProcType=O"	
		End If
		strVal = strVal & "&txtFromPlanDt=" & Trim(frm1.txtFromPlanDt.Text)
		strVal = strVal & "&txtToPlanDt=" & Trim(frm1.txtToPlanDt.Text)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&cboMrpMgr=" & Trim(frm1.cboMrpMgr.value)
		strVal = strVal & "&cboProdMgr=" & Trim(frm1.cboProdMgr.value)
		strVal = strVal & "&txtPurOrg=" & Trim(frm1.txtPurOrg.value)
		strVal = strVal & "&txtPurGrp=" & Trim(frm1.txtPurGrp.value)
		strVal = strVal & "&txtSuppl=" & Trim(frm1.txtSuppl.value)
		
	End If
	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolbar("11000000000111")

    lgIntFlgMode = parent.OPMD_UMODE
    lgBlnFlgChgValue = False		
    Call ggoOper.LockField(Document, "Q")
    
    frm1.vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP전개결과조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>계획일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p2341ma1_OBJECT1_txtFromPlanDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p2341ma1_OBJECT2_txtToPlanDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>				
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>MRP담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboMrpMgr" ALT="MRP담당자" STYLE="Width: 98px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=13 MAXLENGTH=4 tag="11XXXU" ALT="구매조직"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurOrg()">&nbsp;<INPUT TYPE= TEXT NAME="txtPurOrgNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>생산담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdMgr" ALT="생산담당" STYLE="Width: 98px;" tag="11XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=13 MAXLENGTH=4 tag="11XXXU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurGrp()">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>공급처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSuppl" SIZE=13 MAXLENGTH=10 tag="11XXXU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSuppl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSuppl()">&nbsp;<INPUT TYPE=TEXT NAME="txtSupplNm" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>조달구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType4" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoProcType4">전체</LABEL>
														 <INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType1" CLASS="RADIO" tag="1X"><LABEL FOR="rdoProcType1">제조</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType2" CLASS="RADIO" tag="1X"><LABEL FOR="rdoProcType2">구매</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType3" CLASS="RADIO" tag="1X"><LABEL FOR="rdoProcType3">외주</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p2341ma1_I663016325_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProcType" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdMgr" tag="24"><INPUT TYPE=HIDDEN NAME="hMrpMgr" tag="24">
<INPUT TYPE=HIDDEN NAME="hPurOrg" tag="24"><INPUT TYPE=HIDDEN NAME="hPurGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hSuppl" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromPlanDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToPlanDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
