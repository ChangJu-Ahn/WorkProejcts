<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2212ma1.asp
'*  4. Program Name         : MPS생성근거조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/12/12
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

'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p2212mb1.asp"								
Const BIZ_PGM_QRY2_ID	= "p2212mb2.asp"								

 ' Grid 1(vspdData1) - Operation 
Dim C_ItemCd		
Dim C_ItemNm
Dim C_Spec				
Dim C_TrackingNo	

 ' Grid 2(vspdData2) - Operation 
Dim C_ReqDt
Dim C_CoQty
Dim C_SpQty
Dim C_MpsQty
Dim C_ResrvQty
Dim C_SchRcptQty
Dim C_InvldQty	
Dim C_PlanMpsQty

'========================================================================================================= 
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgStrPrevKey11
Dim lgStrPrevKey12
Dim lgStrPrevKey2

Dim IsOpenPop 
Dim lgOldRow
         
'========================================================================================================= 
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey11 = ""
    lgStrPrevKey12 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0
    lgSortKey    = 1
    
    lgOldRow = 0
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

	If pvSpdNo = "A" Then
		 ' Grid 1(vspdData1) - Operation 
		C_ItemCd		= 1
		C_ItemNm		= 2
		C_Spec			= 3
		C_TrackingNo	= 4
	
	ElseIf pvSpdNo = "B" Then
		 ' Grid 2(vspdData2) - Operation 
		C_ReqDt			= 1
		C_CoQty			= 2
		C_SpQty			= 3
		C_MpsQty		= 4
		C_ResrvQty		= 5
		C_SchRcptQty	= 6	
		C_InvldQty		= 7
		C_PlanMpsQty	= 8
	Else
		 ' Grid 1(vspdData1) - Operation 
		C_ItemCd		= 1
		C_ItemNm		= 2
		C_Spec			= 3
		C_TrackingNo	= 4
	 ' Grid 2(vspdData2) - Operation 
		C_ReqDt			= 1
		C_CoQty			= 2
		C_SpQty			= 3
		C_MpsQty		= 4
		C_ResrvQty		= 5
		C_SchRcptQty	= 6	
		C_InvldQty		= 7
		C_PlanMpsQty	= 8
	End If
	
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
	
	Call initSpreadPosVariables(pvSpdNo)  
	
	If pvSpdNo = "A" Then
	
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
    
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    
		.ReDraw = false
	
		.MaxCols = C_TrackingNo +1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit	C_ItemCd, 		"품목"			, 15
		ggoSpread.SSSetEdit C_ItemNm,       "품목명"		, 20
		ggoSpread.SSSetEdit	C_Spec,			"규격"			, 20
		ggoSpread.SSSetEdit	C_TrackingNo,	"Tracking No."	, 25
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

		.ReDraw = true
    
		End With
	
	ElseIF	pvSpdNo = "B" Then
	
		With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread   
    
		.ReDraw = false
	
		.MaxCols = C_PlanMpsQty + 1
 
 		.Col = .MaxCols
		.ColHidden = True   
	
		.MaxRows = 0
    
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetDate 	C_ReqDt,	 	"일자"			, 11, 2, gDateFormat    
		ggoSpread.SSSetFloat	C_CoQty, 		"수주량"		, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_SpQty, 		"판매계획량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_MpsQty, 		"MPS량"			, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_ResrvQty,		"출고예정량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_SchRcptQty,	"입고예정량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_InvldQty, 	"가용재고량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PlanMpsQty,	"계획MPS량"		, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
    
		ggoSpread.SSSetSplit2(1)
				  
		.ReDraw = true
	
		End With
		
	Else
		
		With frm1.vspdData1 
    
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    
		.ReDraw = false
	
		.MaxCols = C_TrackingNo +1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit	C_ItemCd, 		"품목"			, 15
		ggoSpread.SSSetEdit C_ItemNm,       "품목명"		, 20
		ggoSpread.SSSetEdit	C_Spec,			"규격"			, 20
		ggoSpread.SSSetEdit	C_TrackingNo,	"Tracking No."	, 25
    
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

		.ReDraw = true
    
		End With
		
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread   
    
		.ReDraw = false
	
		.MaxCols = C_PlanMpsQty + 1
 
 		.Col = .MaxCols
		.ColHidden = True   
	
		.MaxRows = 0
    
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetDate 	C_ReqDt,	 	"일자"			, 11, 2, gDateFormat    
		ggoSpread.SSSetFloat	C_CoQty, 		"수주량"		, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_SpQty, 		"판매계획량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_MpsQty, 		"MPS량"			, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_ResrvQty,		"출고예정량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_SchRcptQty,	"입고예정량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_InvldQty, 	"가용재고량"	, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PlanMpsQty,	"계획MPS량"		, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
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
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================================= 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd          = iCurColumnPos(1)
			C_ItemNm          = iCurColumnPos(2)
			C_Spec		      = iCurColumnPos(3)    
			C_TrackingNo      = iCurColumnPos(4)    
		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ReqDt         = iCurColumnPos(1)
			C_CoQty         = iCurColumnPos(2)
			C_SpQty			= iCurColumnPos(3)
			C_MpsQty        = iCurColumnPos(4)
			C_ResrvQty      = iCurColumnPos(5)
			C_SchRcptQty    = iCurColumnPos(6)    
			C_InvldQty      = iCurColumnPos(7)
			C_PlanMpsQty    = iCurColumnPos(8)
    End Select    
End Sub    

'========================================================================================================= 
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================================= 
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================= 
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================================= 
Sub PopRestoreSpreadColumnInf()
	Dim pvSpdNo
	
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)  
    
    If pvSpdNo = "A" Then
		ggoSpread.Source = frm1.vspdData1
	Else
		ggoSpread.Source = frm1.vspdData2
	End If
	
	Call ggoSpread.ReOrderingSpreadData()

End Sub

'-----------------------------------  OpenConItemInfo()  -------------------------------------------------
'	Name : OpenConItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
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

'------------------------------------------  OpenConPlant()  -----------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
	
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
'	arrParam(3) = UniConvYYYYMMDDToDate(gDateFormat, "1900", "01", "01")'frm1.txtPlanStartDt.Text
'	arrParam(4) = UniConvYYYYMMDDToDate(gDateFormat, "2999", "12", "31")'frm1.txtPlanEndDt.Text
	
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

'------------------------------------------  SetItemInfo()  --------------------------------------------------
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

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetTrackingNo()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement 
End Function

'========================================================================================================= 
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
  
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")    
    Call InitSpreadSheet("*")
    
    Call InitVariables
    Call SetToolBar("11000000000011")
    
    If Parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If		
	
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopupMenuItemInf("0000111111")

	If Col < 0 Then
		Exit Sub
	End If

	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	If lgOldRow <> Row Then
		
		frm1.vspdData1.Col = C_ItemCd
		frm1.vspdData1.Row = Row
	
		lgOldRow = Row
		
		frm1.vspdData2.MaxRows = 0
		
		Call LayerShowHide(1)
		  		
		Call DisableToolBar(Parent.TBC_QUERY)   ': Query 버튼을 disable 시킴.
		
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
	
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
	Call SetPopupMenuItemInf("0000111111")

	If Col < 0 Then
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
			Call DisableToolBar(Parent.TBC_QUERY)
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
			Call DisableToolBar(Parent.TBC_QUERY)
            If DbDtlQuery = False Then 
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

Sub txtRoutNo_OnChange()
	If frm1.txtRoutNo.value = "" Then
		frm1.txtRoutNm.value = ""
	End If	
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
     Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
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
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&lgStrPrevKey11=" & lgStrPrevKey11
		strVal = strVal & "&lgStrPrevKey12=" & lgStrPrevKey12
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
    Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&lgStrPrevKey11=" & lgStrPrevKey11
		strVal = strVal & "&lgStrPrevKey12=" & lgStrPrevKey12
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
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
	Call SetToolBar("11000000000111")

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisableToolBar(Parent.TBC_QUERY)
        If DbDtlQuery = False Then 
           Call RestoreToolBar()
           Exit Function
        End If 
	End If
	
	frm1.vspdData1.focus
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 
    Dim strVal
    Dim strItemCd
    Dim strTrackingNo
    
	frm1.vspdData1.Col = C_ItemCd
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	
	strItemCd = Trim(frm1.vspdData1.Text)
	
	frm1.vspdData1.Col = C_TrackingNo
	
	strTrackingNo = Trim(frm1.vspdData1.Text)

    DbDtlQuery = False
    
    Err.Clear
        
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtItemCd=" & Trim(strItemCd)
		strVal = strVal & "&txtTrackingNo=" & Trim(strTrackingNo)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	Else
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtItemCd=" & Trim(strItemCd)
		strVal = strVal & "&txtTrackingNo=" & Trim(strTrackingNo)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbDtlQuery = True

End Function

Function DbDtlQueryOk()
    
    lgIntFlgMode = Parent.OPMD_UMODE
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MPS생성근거조회</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
			 						<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
										<TD CLASS="TD5">Tracking No.</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo()"></TD>
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
								<TD WIDTH="30%">
									<script language =javascript src='./js/p2212ma1_A_vspdData1.js'></script>
								</TD>							
								<TD WIDTH="70%">
									<script language =javascript src='./js/p2212ma1_B_vspdData2.js'></script>
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
