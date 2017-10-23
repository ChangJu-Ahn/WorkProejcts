<%@ LANGUAGE="VBSCRIPT" %>
<!--'******************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : MRP Partial Conversion
'*  3. Program ID           : p2346ma1.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Hyun Jae
'* 10. Modifier (Last)      : Jung Yu Kyung
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

Const BIZ_PGM_QRY_ID		= "p2346mb1.asp"
Const BIZ_PGM_CONVPAR_ID	= "p2346mb2.asp"

'==========================================================================================================

Dim C_Select     
Dim C_ItemCd	
Dim C_ItemNm	
Dim C_Spec 	    
Dim C_TrackingNo
Dim C_StartDt	
Dim C_EndDt		
Dim C_PlanQty	
Dim C_Unit		
Dim C_ProcType	
Dim C_PlanOrderNo
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
Dim C_SelectForPurQty

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,Parent.gDateFormat)

'==========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim IsOpenPop          
Dim lgSelRows
Dim lgButtonSelection

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_Select        = 1
    C_ItemCd		= 2
    C_ItemNm		= 3
    C_Spec 			= 4
    C_TrackingNo	= 5
    C_StartDt		= 6
    C_EndDt			= 7
    C_PlanQty		= 8
    C_Unit			= 9
    C_ProcType		= 10
    C_PlanOrderNo	= 11
    C_MRPController	= 12
    C_ProdMgr		= 13
    C_PurOrg		= 14
    C_PurOrg_Nm		= 15
    C_PurGrp		= 16
    C_PurGrp_Nm		= 17
    C_Suppl	    	= 18
    C_Suppl_Nm		= 19
    C_ItemGroupCd	= 20
    C_ItemGroupNm	= 21
    C_SelectForPurQty = 22
End Sub

'==========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0
    lgSelRows = 0
    lgSortKey    = 1
	lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
	
End Sub

'==========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()	
	frm1.txtStartDt.text  = StartDate
	frm1.txtEndDt.text	  = LastDate
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
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
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
    With frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20031210",,parent.gAllowDragDropSpread    
    
    .Redraw = False
	
    .MaxCols = C_SelectForPurQty + 1
	.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetCheck	C_Select,		"", 2,,,1
	ggoSpread.SSSetEdit		C_ItemCd,		"품목"		, 18
	ggoSpread.SSSetEdit		C_ItemNm,		"품목명"	, 25
	ggoSpread.SSSetEdit		C_Spec,			"규  격", 25
	ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25	
	ggoSpread.SSSetDate		C_StartDt,		"시작일"	, 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_EndDt,		"완료일"	, 11, 2, parent.gDateFormat
	ggoSpread.SSSetFloat	C_PlanQty,		"계획수량"	, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	ggoSpread.SSSetEdit		C_Unit,			"단위"		, 7
	ggoSpread.SSSetEdit		C_ProcType,		"조달구분"	, 10
	ggoSpread.SSSetEdit		C_PlanOrderNo,	"계획오더번호", 18
	ggoSpread.SSSetEdit		C_MRPController,"MRP 담당자", 12
	ggoSpread.SSSetEdit		C_ProdMgr,		"생산담당자", 12
	ggoSpread.SSSetEdit		C_PurOrg,		"구매조직", 12
	ggoSpread.SSSetEdit		C_PurOrg_Nm,	"구매조직명", 12
	ggoSpread.SSSetEdit		C_PurGrp,		"구매그룹", 12
	ggoSpread.SSSetEdit		C_PurGrp_Nm,	"구매그룹명", 12
	ggoSpread.SSSetEdit		C_Suppl,		"공급처", 12
	ggoSpread.SSSetEdit		C_Suppl_Nm,		"공급처명", 12
	ggoSpread.SSSetEdit 	C_ItemGroupCd,	"품목그룹",		15
	ggoSpread.SSSetEdit		C_ItemGroupNm,	"품목그룹명",	30
	ggoSpread.SSSetCheck	C_SelectForPurQty,	"", 2,,,1
	
	Call ggoSpread.SSSetColHidden(C_SelectForPurQty,C_SelectForPurQty,True)
	Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
	ggoSpread.SSSetSplit2(2)
	
	.ReDraw = true

	Call SetSpreadLock 

    End With
    
End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
	
	ggoSpread.SpreadLock -1, -1
	ggoSpread.SpreadUnLock C_Select, -1, C_Select

	.vspdData.ReDraw = True
	
	End With
End Sub

'==========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'==========================================================================================================
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
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Select		= iCurColumnPos(1)
			C_ItemCd		= iCurColumnPos(2)
			C_ItemNm		= iCurColumnPos(3)
			C_Spec			= iCurColumnPos(4)    
			C_TrackingNo	= iCurColumnPos(5)
			C_StartDt		= iCurColumnPos(6)
			C_EndDt			= iCurColumnPos(7)
			C_PlanQty		= iCurColumnPos(8)
			C_Unit			= iCurColumnPos(9)
			C_ProcType		= iCurColumnPos(10)
			C_PlanOrderNo	= iCurColumnPos(11)
			C_MRPController	= iCurColumnPos(12)
			C_ProdMgr		= iCurColumnPos(13)
			C_PurOrg		= iCurColumnPos(14)
			C_PurOrg_Nm		= iCurColumnPos(15)
			C_PurGrp		= iCurColumnPos(16)
			C_PurGrp_Nm		= iCurColumnPos(17)
			C_Suppl	    	= iCurColumnPos(18)
			C_Suppl_Nm		= iCurColumnPos(19)
			C_ItemGroupCd	= iCurColumnPos(20)
			C_ItemGroupNm	= iCurColumnPos(21)
			C_SelectForPurQty = iCurColumnPos(22)
    End Select    

End Sub
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
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
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function
'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()

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
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)			' Item Code
	arrParam(2) = "" 									' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""									' Default Value
	
    arrField(0) = 1 									' Field명(0)
    arrField(1) = 2										' Field명(1)
        
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
		Call SetItemCd(arrRet)
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

'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function
'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(ByRef arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
		.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End With
End Function
'===========================================================================================================
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
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function


Function btnAutoSel_onClick()

	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	Dim index, Count

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
	Else
		lgButtonSelection = "SELECT"
	End If
	
	frm1.vspdData.ReDraw = False
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index

		frm1.vspdData.Col = C_Select	
		
		If lgButtonSelection = "SELECT" Then
			frm1.vspdData.Value = 1
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData.Value = 0
		End if

	Next 
	
	frm1.vspdData.ReDraw = True

End Function

'------------------------------------------  OpenErrorList()  -------------------------------------------------
'	Name : OpenErrorList()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenErrorList()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value  = "" Then
		call DisplayMsgBox("220705", "X","X","X")			'⊙: 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)		
	arrParam(1) = Trim(frm1.hmrpno.value)			
	
	iCalledAspName = AskPRAspName("P2345RA1")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P2345RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName , Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

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
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
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

	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		Call SetPopupMenuItemInf("0001111111") 
	End If	
    
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
  
    With frm1.vspdData
		
		If Row <= 0 Then Exit Sub

		If Col = C_Select Then
			.Row = Row
			.Col = C_Select
		
			If Buttondown = 1 Then
				lgSelRows = lgSelRows + 1
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			Else
				If lgSelRows - 1 < 0 Then
					lgSelRows = 0 
				Else
					lgSelRows = lgSelRows - 1
				End If
				.Col = C_SelectForPurQty
				If .value <> 1 Then
					ggoSpread.Source = frm1.vspdData
					ggoSpread.SSDeleteFlag Row,Row
				End If	
			End If

		End If
	End With
		
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If 
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" or lgStrPrevKey2 <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 			
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
	
	With frm1.vspdData 

		Select Case Col
			
		    Case C_PlanQty
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
				.Row = Row
				.Col = C_PlanQty
				If .Value <= 0 Then
					Call DisplayMsgBox("169918", "x", "x", "x")
					.Value = ""
					.Focus
					Set gActiveElement = document.activeElement 
					Exit Sub
				End If
				
				.Col = C_SelectForPurQty
				.Value = 1
							
		End Select
    
   End With

End Sub

'==========================================================================================
'   Event Name : btnConvPar_onClick()
'   Event Desc : 선택전환(50개단위)
'==========================================================================================
Function btnConvPar_onClick()
  
    Dim lRow         
	Dim strVal					' for partly MRP
	Dim strVal2					' for changing purchase plan qty.
	Dim IntRetCD
	
	'------------------------------------
	' Previous Check
	'------------------------------------
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If	
	        
	Call LayerShowHide(1)

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		strVal = ""
		strVal2 = ""
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = C_Select
		    
		    Select Case .vspdData.value

		        Case 1									
		            									 
		            .vspdData.Col = C_PlanOrderNo						 
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep		 
		            
		    End Select
		    
		    .vspdData.Col = C_SelectForPurQty
		    
		    Select Case .vspdData.value
				
				Case 1
				
					.vspdData.Col = C_PlanOrderNo
					strVal2 = strVal2 & Trim(.vspdData.Text) & parent.gColSep	
					.vspdData.Col = C_PlanQty
					strVal2 = strVal2 & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
				
			End Select	
		            
		Next
		
		.txtSpread.value = strVal
		.txtSpread2.value = strVal2
	
	Call ExecMyBizASP(frm1, BIZ_PGM_CONVPAR_ID)										'☜: 비지니스 ASP 를 가동 

	End With
End Function

Function MRPConvOK()

	Call InitVariables
	Call MainQuery
	
End Function

'=======================================================================================================
'   Event Name : txtStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStartDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtStartDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtStartDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtEndDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtEndDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtEndDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtStartDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtEndDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

Sub rdoProcType1_OnClick()
	frm1.txtPurOrg.value = ""
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"D")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"Q")
End Sub

Sub rdoProcType2_OnClick()
	frm1.cboProdMgr.value = ""
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"Q")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"D")
End Sub

Sub rdoProcType3_OnClick()
	frm1.cboProdMgr.value = ""
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"Q")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"D")
End Sub

Sub rdoProcType4_OnClick()
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"D")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"D")
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False 

    Err.Clear
	
    Call ggoOper.ClearField(Document, "2")  
    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt)  = False Then		
		Exit Function
	End If 

	Call SetToolbar("11000000000011")

    If DbQuery = False Then
		Exit Function
	End If

    FncQuery = True
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False
    
    Err.Clear
    
    Call btnConvPar_onClick
	
    FncSave = True
    
End Function
'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function	
	
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	
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
	Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
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
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strVal
    
    DbQuery = False

    Call LayerShowHide(1)
    
    Err.Clear

    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & .hPlantCd.value
	    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	    strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	    strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	    strVal = strVal & "&txtItemCd=" & .hItemCd.value
		strVal = strVal & "&txtStartDt=" & Trim(.hStartDt.value)
		strVal = strVal & "&txtEndDt=" & Trim(.hEndDt.value)
	    strVal = strVal & "&txtTrackingNo=" & .hTrackingNo.value
		strVal = strVal & "&txtItemGroupCd=" & .hItemGroupCd.value
    	strVal = strVal & "&rdoProcType=" & .hProcType.value
    	strVal = strVal & "&cboProdMgr=" & Trim(.hProdMgr.value)
    	strVal = strVal & "&cboMrpMgr=" & Trim(.hMrpMgr.value)
		strVal = strVal & "&txtPurOrg=" & Trim(.hPurOrg.value)
		strVal = strVal & "&txtPurGrp=" & Trim(.hPurGrp.value)
		strVal = strVal & "&txtSuppl=" & Trim(.hSuppl.value)
    Else
	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	    strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	    strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtStartDt=" & Trim(.txtStartDt.text)
		strVal = strVal & "&txtEndDt=" & Trim(.txtEndDt.text)
		strVal = strVal & "&cboProdMgr=" & Trim(.cboProdMgr.value)
		strVal = strVal & "&cboMrpMgr=" & Trim(.cboMrpMgr.value)
	    strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
	
		If frm1.rdoProcType(0).checked Then
			strVal = strVal & "&rdoProcType=A"
		ElseIf frm1.rdoProcType(1).checked Then
			strVal = strVal & "&rdoProcType=M"
		ElseIf frm1.rdoProcType(2).checked Then
			strVal = strVal & "&rdoProcType=P"
		ElseIf frm1.rdoProcType(3).checked Then
			strVal = strVal & "&rdoProcType=O"	
		End If
		
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
Function DbQueryOk(ByVal LngMaxRow)
	
	Dim strProctype
    Dim lRow
	
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
    With frm1.vspdData 
    
        .ReDraw = False
      
		For lRow = LngMaxRow To .MaxRows

			.Row = lRow
			.Col =  C_ProcType
			strProctype = .Value
			
			If strProctype = "구매" Then
				ggoSpread.source = frm1.vspddata
				ggoSpread.SpreadUnLock C_PlanQty, lRow, C_PlanQty,lRow
				ggoSpread.SSSetRequired C_PlanQty,		lRow, lRow
			End If
			
		Next
		
		.ReDraw = True
		
	End With
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE
	Call SetToolbar("11001001000111")
	frm1.btnAutoSel.disabled = False
    frm1.vspdData.Focus
			
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP부분전환</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="VBScript:OpenPeggInfo()">Pegging정보</A>&nbsp;|&nbsp;<A href="vbscript:OpenErrorList()">ERROR내역리스트</A></TD>				
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>시작예정일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtStartDt CLASSID=<%=gCLSIDFPDT%> ALT="시작일" tag="11X1" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtEndDt CLASSID=<%=gCLSIDFPDT%> ALT="완료일" tag="11X1" VIEWASTEXT id=OBJECT2> </OBJECT>');</SCRIPT>									
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
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
								</TR>
									<TD CLASS=TD5 NOWRAP>조달구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType4" CLASS="RADIO" tag="1X" Value="A" CHECKED><LABEL FOR="rdoProcType4">전체</LABEL>
									                <INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType1" CLASS="RADIO" tag="1X" Value="M"><LABEL FOR="rdoProcType1">제조</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType2" CLASS="RADIO" tag="1X" Value="P"><LABEL FOR="rdoProcType2">구매</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoProcType" ID="rdoProcType3" CLASS="RADIO" tag="1X" Value="O"><LABEL FOR="rdoProcType3">외주</LABEL></TD>																			
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD><BUTTON NAME="btnAutoSel" CLASS="CLSMBTN">전체선택</BUTTON></TD></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hmrpno" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24"><INPUT TYPE=HIDDEN NAME="hSuppl" tag="24">
<INPUT TYPE=HIDDEN NAME="hPurOrg" tag="24"><INPUT TYPE=HIDDEN NAME="hPurGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdMgr" tag="24"><INPUT TYPE=HIDDEN NAME="hMrpMgr" tag="24">
<INPUT TYPE=HIDDEN NAME="hProcType" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
