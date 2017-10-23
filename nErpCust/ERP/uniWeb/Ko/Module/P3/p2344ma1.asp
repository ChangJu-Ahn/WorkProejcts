<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2344ma1
'*  4. Program Name         :
'*  5. Program Desc         : 계획오더조회 
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

Const BIZ_PGM_QRY1_ID	= "p2344mb1.asp"
Const BIZ_PGM_QRY2_ID	= "p2344mb2.asp"
Const BIZ_PGM_QRY3_ID	= "p2344mb3.asp"

'==========================================================================================================

 ' Grid 1(vspdData1) - Operation 
Dim C_ItemCd1	
Dim C_ItemNm1
Dim C_ItemSpec1	
Dim C_TrackingNo1
Dim C_PlanQty1	
Dim C_Unit1		
Dim C_StartDt1		
Dim C_EndDt1		
Dim C_Status1		
Dim C_ProdOrderNo	
Dim C_ProdMgr		
Dim C_ItemGroupCd1
Dim C_ItemGroupNm1

 ' Grid 2(vspdData2) - Operation 
Dim C_ItemCd2		
Dim C_ItemNm2		
Dim C_ItemSpec2
Dim C_TrackingNo2	
Dim C_PlanQty2		
Dim C_Unit2			
Dim C_StartDt2		
Dim C_EndDt2		
Dim C_Status2		
Dim C_PurOrderNo	
Dim C_PurOrg		
Dim C_ItemGroupCd2
Dim C_ItemGroupNm2

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,Parent.gDateFormat)

'=========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgStrPrevKey11
Dim lgStrPrevKey12
Dim lgStrPrevKey21
Dim lgStrPrevKey22

'=========================================================================================================
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgItemCd
         
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	 ' Grid 1(vspdData1) - Operation 
    C_ItemCd1		= 1
    C_ItemNm1		= 2
    C_ItemSpec1		= 3
    C_TrackingNo1	= 4
    C_PlanQty1		= 5
    C_Unit1			= 6
    C_StartDt1		= 7
    C_EndDt1		= 8
    C_Status1		= 9
    C_ProdOrderNo	= 10
    C_ProdMgr		= 11
	C_ItemGroupCd1	= 12
	C_ItemGroupNm1	= 13

 ' Grid 2(vspdData2) - Operation 
    C_ItemCd2		= 1
    C_ItemNm2		= 2
    C_ItemSpec2		= 3
    C_TrackingNo2	= 4
    C_PlanQty2		= 5
    C_Unit2			= 6
    C_StartDt2		= 7
    C_EndDt2		= 8
    C_Status2		= 9
    C_PurOrderNo	= 10
    C_PurOrg		= 11
	C_ItemGroupCd2	= 12
	C_ItemGroupNm2	= 13
End Sub 

'=========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    
    lgStrPrevKey11 = ""
    lgStrPrevKey12 = ""
    lgStrPrevKey21 = ""
    lgStrPrevKey22 = ""
    lgLngCurRows = 0
    lgAfterQryFlg = False 
    lgLngCnt = 0
    lgOldRow = 0 
    lgSortKey    = 1
End Sub

'=========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtStartDt.text	= StartDate
	frm1.txtEndDt.text	= LastDate
End Sub

'=========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub


'=========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call initSpreadPosVariables()    
	
	If pvSpdNo = "A" Then									'☜: 대상이 vspdData1일때 
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
    
		.ReDraw = false
    
		.MaxCols = C_ItemGroupNm1 +1	
		.MaxRows = 0
  	
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd1, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm1,		"품목명"		, 25
		ggoSpread.SSSetEdit 	C_ItemSpec1,	"규격"			, 25
		ggoSpread.SSSetEdit		C_TrackingNo1,	"Tracking No."	, 25
		ggoSpread.SSSetFloat	C_PlanQty1, 	"수량"			, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_Unit1,		"단위"			, 7
		ggoSpread.SSSetDate 	C_StartDt1, 	"시작일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetDate 	C_EndDt1, 		"종료일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetEdit		C_Status1,	 	"Status"		, 10
		ggoSpread.SSSetEdit		C_ProdOrderNo, 	"제조오더번호"	, 18
		ggoSpread.SSSetEdit		C_ProdMgr, 		"생산담당자"	, 15
		ggoSpread.SSSetEdit 	C_ItemGroupCd1,	"품목그룹",		15
		ggoSpread.SSSetEdit		C_ItemGroupNm1,	"품목그룹명",	30

		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		ggoSpread.SSSetSplit2(1)
	
		.ReDraw = true
    
		End With
	
	ElseIf pvSpdNo = "B" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------

		With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
	
		.ReDraw = false
	
		.MaxCols = C_ItemGroupNm2+1
		.MaxRows = 0
	
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit		C_ItemCd2, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm2,      "품목명"		, 25
		ggoSpread.SSSetEdit 	C_ItemSpec2,	"규격"			, 25
		ggoSpread.SSSetEdit		C_TrackingNo2,	"Tracking No."	, 25
		ggoSpread.SSSetFloat	C_PlanQty2, 	"수량"			, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_Unit2, 		"단위"			, 7
		ggoSpread.SSSetDate 	C_StartDt2, 	"시작일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetDate 	C_EndDt2, 		"종료일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetEdit		C_Status2,	 	"Status"		, 10
		ggoSpread.SSSetEdit		C_PurOrderNo, 	"구매오더번호"	, 18
		ggoSpread.SSSetEdit		C_PurOrg, 		"구매조직"		, 15
		ggoSpread.SSSetEdit 	C_ItemGroupCd2,	"품목그룹",		15
		ggoSpread.SSSetEdit		C_ItemGroupNm2,	"품목그룹명",	30
	
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		ggoSpread.SSSetSplit2(1)
		.ReDraw = true
	
		End With
	
	Else													'☜: 대상이 모든 Spread일때    
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------

		With frm1.vspdData1 
	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
    
		.ReDraw = false
    
		.MaxCols = C_ItemGroupNm1 +1
		.MaxRows = 0
  	
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd1, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm1,		"품목명"		, 25
		ggoSpread.SSSetEdit 	C_ItemSpec1,	"규격"			, 25
		ggoSpread.SSSetEdit		C_TrackingNo1,	"Tracking No."	, 25
		ggoSpread.SSSetFloat	C_PlanQty1, 	"수량"			, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_Unit1,		"단위"			, 7
		ggoSpread.SSSetDate 	C_StartDt1, 	"시작일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetDate 	C_EndDt1, 		"종료일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetEdit		C_Status1,	 	"Status"		, 10
		ggoSpread.SSSetEdit		C_ProdOrderNo, 	"제조오더번호"	, 18
		ggoSpread.SSSetEdit		C_ProdMgr, 		"생산담당자"	, 15
		ggoSpread.SSSetEdit 	C_ItemGroupCd1,	"품목그룹",		15
		ggoSpread.SSSetEdit		C_ItemGroupNm1,	"품목그룹명",	30

		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		ggoSpread.SSSetSplit2(1)
	
		.ReDraw = true
    
		End With
		
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
	
		.ReDraw = false
	
		.MaxCols = C_ItemGroupNm2+1
		.MaxRows = 0
	
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit		C_ItemCd2, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm2,      "품목명"		, 25
		ggoSpread.SSSetEdit 	C_ItemSpec2,	"규격"			, 25
		ggoSpread.SSSetEdit		C_TrackingNo2,	"Tracking No."	, 25
		ggoSpread.SSSetFloat	C_PlanQty2, 	"수량"			, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_Unit2, 		"단위"			, 7
		ggoSpread.SSSetDate 	C_StartDt2, 	"시작일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetDate 	C_EndDt2, 		"종료일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetEdit		C_Status2,	 	"Status"		, 10
		ggoSpread.SSSetEdit		C_PurOrderNo, 	"구매오더번호"	, 18
		ggoSpread.SSSetEdit		C_PurOrg, 		"구매조직"		, 15
		ggoSpread.SSSetEdit 	C_ItemGroupCd2,	"품목그룹",		15
		ggoSpread.SSSetEdit		C_ItemGroupNm2,	"품목그룹명",	30
	
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		ggoSpread.SSSetSplit2(1)
		.ReDraw = true
		
		End With

	End If
	
	Call SetSpreadLock 
    
End Sub

'=========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
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

 '=========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6, i
	Dim iProdMgrArr, iProdMgrNmArr
	
	
    On Error Resume Next
    Err.Clear

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
End Sub

'=========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd1		= iCurColumnPos(1)
			C_ItemNm1		= iCurColumnPos(2)
			C_ItemSpec1		= iCurColumnPos(3)
			C_TrackingNo1	= iCurColumnPos(4)
			C_PlanQty1		= iCurColumnPos(5)    
			C_Unit1			= iCurColumnPos(6)
			C_StartDt1		= iCurColumnPos(7)
			C_EndDt1		= iCurColumnPos(8)
			C_Status1		= iCurColumnPos(9)    
			C_ProdOrderNo	= iCurColumnPos(10)
			C_ProdMgr		= iCurColumnPos(11)
			C_ItemGroupCd1	= iCurColumnPos(12)
			C_ItemGroupNm1	= iCurColumnPos(13)

		Case "B"
            ggoSpread.Source = frm1.vspdData2
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd2		= iCurColumnPos(1)
			C_ItemNm2		= iCurColumnPos(2)
			C_ItemSpec2		= iCurColumnPos(3)
			C_TrackingNo2	= iCurColumnPos(4)    
			C_PlanQty2		= iCurColumnPos(5)
			C_Unit2			= iCurColumnPos(6)
			C_StartDt2		= iCurColumnPos(7)
			C_EndDt2		= iCurColumnPos(8)    
			C_Status2		= iCurColumnPos(9)
			C_PurOrderNo	= iCurColumnPos(10)
			C_PurOrg		= iCurColumnPos(11)
			C_ItemGroupCd2	= iCurColumnPos(12)
			C_ItemGroupNm2	= iCurColumnPos(13)
			
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

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

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

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

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
'	arrParam(3) = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")'frm1.txtPlanStartDt.Text
'	arrParam(4) = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")'frm1.txtPlanEndDt.Text
	
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

'--------------------------------------------  OpenStockRef()  -------------------------------------------
'	Name : OpenStockRef()
'	Description : Stock Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenStockRef()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002","X", "X","X")
		
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

    arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(lgItemCd)
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("P4212RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4212RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'----------------------------------------  OpenPurOrg()  -------------------------------------------------
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

'------------------------------------------  SetItemInfo()  ----------------------------------------------
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

Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetPurOrg()  ------------------------------------------------
'	Name : SetPurOrg()
'	Description : PurOrg Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPurOrg(ByRef arrRet)
	frm1.txtPurOrg.Value    = arrRet(0)
	frm1.txtPurOrgNm.Value  = arrRet(1)
	frm1.txtPurOrg.focus
	Set gActiveElement = document.activeElement	
End Function
'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
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
    
    Call InitSpreadSheet("*")
    
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
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	
	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then
       Exit Sub
   	End If
   	
   	If Col < 0 Then	Exit Sub
   	
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
	
	lgItemCd = GetSpreadText(frm1.vspdData1,C_ItemCd1,frm1.vspdData1.ActiveRow,"X","X")
	
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
   	
   	If Col < 0 Then	Exit Sub
   	
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
    
	lgItemCd = GetSpreadText(frm1.vspdData2,C_ItemCd2,frm1.vspdData2.ActiveRow,"X","X")
	
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
		If lgStrPrevKey11 <> "" OR lgStrPrevKey12 <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery2 = False Then 
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
		If lgStrPrevKey21 <> "" Or lgStrPrevKey22 <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery3 = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
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
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt)  = False Then		
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim strVal
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear
        
    With frm1
    
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey11=" & lgStrPrevKey11
		strVal = strVal & "&lgStrPrevKey12=" & lgStrPrevKey12

		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtStartDt=" & Trim(frm1.txtStartDt.Text)
		strVal = strVal & "&txtEndDt=" & Trim(frm1.txtEndDt.Text)
		If frm1.rdoConvType(0).checked Then
			strVal = strVal & "&rdoConvType=A"
		ElseIf frm1.rdoConvType(1).checked Then
			strVal = strVal & "&rdoConvType=NL"
		ElseIf frm1.rdoConvType(2).checked Then
			strVal = strVal & "&rdoConvType=P"
		End if		
		strVal = strVal & "&cboProdMgr=" & Trim(.cboProdMgr.value)
		strVal = strVal & "&txtPurOrg=" & Trim(.txtPurOrg.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)

		strVal = strVal & "&txtMaxRows1=" & .vspdData1.MaxRows
		strVal = strVal & "&txtMaxRows2=" & .vspdData2.MaxRows
        
    End With

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
   
	If frm1.vspdData1.ActiveRow <= 0 and frm1.vspdData2.ActiveRow > 0 Then 
		lgItemCd = GetSpreadText(frm1.vspdData2,C_ItemCd2,frm1.vspdData2.ActiveRow,"X","X")
	ElseIF frm1.vspdData2.ActiveRow <= 0 Then
		lgItemCd = ""
	Else
		lgItemCd = GetSpreadText(frm1.vspdData1,C_ItemCd1,frm1.vspdData1.ActiveRow,"X","X")	
	End IF 

End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery2() 
    Dim strVal

    DbQuery2 = False
    
    Call LayerShowHide(1)
    
    Err.Clear 
        
    With frm1
    
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey11=" & lgStrPrevKey11
		strVal = strVal & "&lgStrPrevKey12=" & lgStrPrevKey12
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode

		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtStartDt=" & Trim(.hStartDt.value)
		strVal = strVal & "&txtEndDt=" & Trim(.hEndDt.value)
		strVal = strVal & "&cboProdMgr=" & Trim(.hProdMgr.value)
		strVal = strVal & "&rdoConvType=" & Trim(.hConvType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		
		strVal = strVal & "&txtMaxRows1=" & .vspdData1.MaxRows
		
    End With

	Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery2 = True

End Function

'========================================================================================
' Function Name : DbQuery1
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery3() 
    Dim strVal

    DbQuery3 = False
    
    Call LayerShowHide(1)
    
    Err.Clear
        
    With frm1
    
		strVal = BIZ_PGM_QRY3_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey21=" & lgStrPrevKey21
		strVal = strVal & "&lgStrPrevKey22=" & lgStrPrevKey22
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode

		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtStartDt=" & Trim(.hStartDt.value)
		strVal = strVal & "&txtEndDt=" & Trim(.hEndDt.value)		
		strVal = strVal & "&txtPurOrg=" & Trim(.hPurOrg.value)
		strVal = strVal & "&rdoConvType=" & Trim(.hConvType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		
		strVal = strVal & "&txtMaxRows2=" & .vspdData2.MaxRows		
        
    End With

	Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery3 = True

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계획오더조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD WIDTH=* align=right><A href="vbscript:OpenStockRef()">재고현황</A></TD>				
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
									<TD CLASS=TD5 NOWRAP>시작일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p2344ma1_OBJECT1_txtStartDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p2344ma1_OBJECT2_txtEndDt.js'></script>	
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
			 						<TD CLASS="TD5">Tracking No.</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo()"></TD>									

								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>생산담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdMgr" ALT="생산담당" STYLE="Width: 98px;" tag="11XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Convert 여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoConvType" ID="rdoConvType1" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoConvType1">전체</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoConvType" ID="rdoConvType2" CLASS="RADIO" tag="1X"><LABEL FOR="rdoConvType2">전환이전</LABEL>
													<INPUT TYPE="RADIO" NAME="rdoConvType" ID="rdoConvType3" CLASS="RADIO" tag="1X"><LABEL FOR="rdoConvType3">전환완료</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="구매조직"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurOrg()">&nbsp;<INPUT TYPE=HIDDEN NAME="txtPurOrgNm" SIZE=25 tag="X4"></TD>
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
							<TR HEIGHT="100%">
								<TD WIDTH="50%">
									<script language =javascript src='./js/p2344ma1_A_vspdData1.js'></script>
								</TD>							
								<TD WIDTH="50%">
									<script language =javascript src='./js/p2344ma1_B_vspdData2.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows1" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hConvType" tag="24"><INPUT TYPE=HIDDEN NAME="hProdMgr" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hPurOrg" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
