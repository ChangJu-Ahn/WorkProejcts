<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name   : Inventory																		*
'*  2. Function Name  : Reference Popup Component List													*
'*  3. Program ID   : i2312ra1																			*
'*  4. Program Name   : 재고현황조회																	*
'*  5. Program Desc   : Reference Popup																	*
'*  7. Modified date(First) : 2000/04/06																*
'*  8. Modified date(Last) :  2003/03/19																*
'*  9. Modifier (First)     : Kim, Gyoung-Don															*
'* 10. Modifier (Last)  : Ahn, Jung-Je																	*
'* 11. Comment     :																					*
'********************************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ==================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit
'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID = "i2312rb1.asp"        

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID = "i2312rb2.asp"        

'==========================================  1.2.1 Global 상수 선언  ======================================
' Grid 1(vspdData1) - Operation
Dim C_SlCd
Dim C_SlNm
Dim C_GoodOnHandQty
Dim C_SchdRcptQty
Dim C_SchdIssueQty
Dim C_AvailQty
Dim C_TrnsQty
Dim C_AllocationQty

' Grid 2(vspdData2) - Operation
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_GoodOnHandQty1
Dim C_TrnsQty1
Dim C_PickingQty
Dim C_BlockIndicator

'==========================================  1.2.2 Global 변수 선언  ==================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevKey2
Dim IsOpenPop 

Dim lgPlantCD
Dim lgItemCD
Dim lgItemNm
Dim lgSlCD
Dim lgSlNm

'==========================================  1.2.3 Global Variable값 정의  ===============================
Dim lgOldRow

'*********************************************  1.3 변 수 선 언  ****************************************
'* 설명: Constant는 반드시 대문자 표기.                *
'********************************************************************************************************
Dim arrParam     
  
arrParam		= window.dialogArguments
Set PopupParent = arrParam(0)

lgPlantCD	= arrParam(1)
lgItemCD	= arrParam(2)

top.document.title = PopupParent.gActivePRAspName

'==========================================  2.1.1 InitVariables()  =====================================
'= Name : InitVariables()                    =
'= Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)    =
'========================================================================================================
Function InitVariables()
	lgIntGrpCount		= 0       
	lgStrPrevKey		= ""      
	lgStrPrevKey2		= ""
	Self.Returnvalue	= Array("")
End Function


'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub


'==========================================   2.1.2 InitSetting()   =====================================
'= Name : InitSetting()                    =
'= Description : Passed Parameter를 Variable에 Setting한다.           =
'========================================================================================================
Function InitSetting()
	txtPlantCd.value = lgPlantCD
	txtItemCd.value  = lgItemCD
End Function

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	
	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "" OR pvSpdNo = "A" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		ggoSpread.Source = vspdData1
		ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

		With vspdData1 
			.ReDraw = false
			.MaxCols = C_AllocationQty +1           
			.MaxRows = 0
			 
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit  C_SlCd,           "창고",    10
			ggoSpread.SSSetEdit  C_SlNm,          "창고명",   15
			ggoSpread.SSSetFloat C_GoodOnHandQty, "양품재고", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_SchdRcptQty,   "입고예정", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_SchdIssueQty,  "출고예정", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_AvailQty,      "가용재고", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_TrnsQty,		  "이동중재고",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_AllocationQty, "재고할당량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			 
			ggoSpread.SSSetSplit2(2)
			'ggoSpread.MakePairsColumn()
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.ReDraw = true
		End With
	 
	End If
    
	If pvSpdNo = "" OR pvSpdNo = "B" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		ggoSpread.Source = vspdData2
		ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

		With vspdData2
			.ReDraw = false
			.MaxCols = C_BlockIndicator +1           
			.MaxRows = 0

			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit  C_TrackingNo,     "Tracking No.", 20
			ggoSpread.SSSetEdit  C_LotNo,          "Lot No.",      18
			ggoSpread.SSSetEdit  C_LotSubNo,       "Lot Sub No.",  15
			ggoSpread.SSSetFloat C_GoodOnHandQty1, "양품재고",     20, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"      
			ggoSpread.SSSetFloat C_TrnsQty1,	   "이동중재고",   20, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_PickingQty,	   "PICKING수량",  20, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit  C_BlockIndicator, "Block",        15
			 
			ggoSpread.SSSetSplit2(2)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.ReDraw = true
		End With
	End If
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables(ByVal pvSpdNo)
If pvSpdNo = "" OR pvSpdNo = "A" Then
' Grid 1(vspdData1) - Operation
	C_SlCd          = 1
	C_SlNm          = 2
	C_GoodOnHandQty = 3
	C_SchdRcptQty   = 4
	C_SchdIssueQty  = 5
	C_AvailQty      = 6
	C_TrnsQty		= 7
	C_AllocationQty	= 8
End If
If pvSpdNo = "" OR pvSpdNo = "B" Then
' Grid 2(vspdData2) - Operation
	C_TrackingNo     = 1
	C_LotNo          = 2
	C_LotSubNo       = 3
	C_GoodOnHandQty1 = 4
	C_TrnsQty1		 = 5
	C_PickingQty	 = 6
	C_BlockIndicator = 7
End If	
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData1 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_SlCd          = iCurColumnPos(1)
		C_SlNm          = iCurColumnPos(2)
		C_GoodOnHandQty = iCurColumnPos(3)
		C_SchdRcptQty   = iCurColumnPos(4)
		C_SchdIssueQty  = iCurColumnPos(5)
		C_AvailQty      = iCurColumnPos(6)
		C_TrnsQty		= iCurColumnPos(7)
		C_AllocationQty = iCurColumnPos(8)
	Case "B"
		ggoSpread.Source = vspdData2 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_TrackingNo     = iCurColumnPos(1)
		C_LotNo          = iCurColumnPos(2)
		C_LotSubNo       = iCurColumnPos(3)
		C_GoodOnHandQty1 = iCurColumnPos(4)
		C_TrnsQty1		 = iCurColumnPos(5)
		C_PickingQty	 = iCurColumnPos(6)
		C_BlockIndicator = iCurColumnPos(7)		
	End Select

End Sub

'=========================================  2.3.2 CancelClick()  ========================================
'= Name : CancelClick()                    =
'= Description : Return Array to Opener Window for Cancel button click         =
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function


'------------------------------------------  OpenItemInfo()  ---------------------------------------------
' Name : OpenItemInfo()
' Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)

	If IsOpenPop = True Then Exit Function

	If txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		txtPlantCd.focus 
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True
	 
	arrParam(0) = Trim(txtPlantCd.value) 
	arrParam(1) = strCode      
	arrParam(2) = ""      
	arrParam(3) = ""      
	 
	arrField(0) = 1 
	arrField(1) = 2 

	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","x")
		IsOpenPop = False
		Exit Function
	End If
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	
		 
	If arrRet(0) = "" Then
		txtItemCd.focus
		Exit Function
	Else
		txtItemCd.Value    = arrRet(0)  
		txtItemNm.Value    = arrRet(1)
		txtItemCd.focus
	End If
		 
End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
' Name : OpenSLCd()
' Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		txtPlantCd.focus 
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "창고팝업"				
	arrParam(1) = "B_STORAGE_LOCATION"          
	arrParam(2) = Trim(txtSLCd.Value)        
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(txtPlantCd.value, "''", "S")
	arrParam(5) = "창고"           
	 
	arrField(0) = "SL_CD"             
	arrField(1) = "SL_NM"             
	    
	arrHeader(0) = "창고"         
	arrHeader(1) = "창고명"       
	    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		txtSLCd.focus
		Exit Function
	Else
		txtSLCd.Value    = arrRet(0)  
		txtSLNm.Value    = arrRet(1) 
		txtSLCd.focus
	End If
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
'= Name : Form_Load()                     =
'= Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분    =
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029          
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call InitVariables           
	Call InitSpreadSheet("")
	Call InitSetting()
	Call FncQuery()
End Sub


Function FncQuery
	FncQuery = False
	ggoSpread.Source = vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = vspdData2
    ggoSpread.ClearSpreadData
    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then Exit Function      


	If DbQuery = False Then Exit Function
	
	FncQuery = False
End Function

'*********************************************  3.3 Object Tag 처리  ************************************
'* Object에서 발생 하는 Event 처리                  *
'********************************************************************************************************
'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
  
	Set gActiveSpdSheet = vspdData1
  
	If vspdData1.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = vspdData1 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		
			lgSortKey = 1
		End If
	End If
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP2C"   
   
	Set gActiveSpdSheet = vspdData2
   
	If vspdData2.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = vspdData2 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col					
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		
			lgSortKey = 1
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
	If Row <= 0 Then Exit Sub
	If vspdData1.MaxRows = 0 Then Exit Sub
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)
	If Row <= 0 Then Exit Sub
	If vspdData2.MaxRows = 0 Then Exit Sub
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
   ggoSpread.Source = vspdData1
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = vspdData2
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = vspdData1
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = vspdData2
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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()

    Select Case gActiveSpdSheet.id
		Case "vaSpread1"
			Call InitSpreadSheet("A")
		Case "vaSpread2"
			Call InitSpreadSheet("B")      		
	End Select 

   Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, ByVal Cancel)
	
	If NewRow <= 0 Or Row = NewRow Then
		Exit Sub
	End If
	ggoSpread.Source = vspdData2
	ggoSpread.ClearSpreadData
	
	lgStrPrevKey = ""
	If DbDtlQuery(NewRow) = False Then	
		Exit Sub
	End If
	
End Sub


'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If vspdData1.MaxRows < NewTop + VisibleRowCnt(vspdData1, NewTop) Then
		If lgStrPrevKey <> "" Then
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
    
	If vspdData2.MaxRows < NewTop + VisibleRowCnt(vspdData2, NewTop) Then
		If lgStrPrevKey2 <> "" Then
			If DbDtlQuery(vspdData1.ActiveRow) = False Then
				Exit Sub
			End if
		End if
	End if
End Sub



Sub vspdData1_KeyPress(keyAscii)
	If keyAscii=27 Then
		Call CancelClick()
		Exit Sub
	End If
End Sub 

Sub vspdData2_KeyPress(keyAscii)
	If keyAscii=27 Then
		Call CancelClick()
		Exit Sub
	End If
End Sub 

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery                    *
' Function Desc : This function is data query and display            *
'********************************************************************************************************
Function DbQuery()
 
    Err.Clear					
     
    DbQuery = False				
     
    Call LayerShowHide(1)
     
    Dim strVal

    strVal =  BIZ_PGM_QRY1_ID & "?txtMode="			& PopupParent.UID_M0001	& _
								"&txtPlantCd="		& txtPlantCd.value		& _
								"&txtItemCd="		& txtItemCd.value		& _
								"&txtSlCd="			& txtSLCd.value			& _
								"&lgStrPrevKey="	& lgStrPrevKey
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
	Call DbDtlQuery(1)
	vspdData1.Focus
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal Row) 

	Dim strVal
	   
	DbDtlQuery = False   
	    
	Call LayerShowHide(1)

	strVal = BIZ_PGM_QRY2_ID &	"?txtMode="			& PopupParent.UID_M0001 & _
								"&txtPlantCd="		& txtPlantCd.value		& _
								"&txtItemCd="		& Trim(txtItemCd.value)
	vspdData1.Row = Row
	vspdData1.Col = C_SLCd
	strVal = strVal          &	"&txtSlCd="			& Trim(vspdData1.Text)	& _
								"&lgStrPrevKey2="	& lgStrPrevKey2   
	Call RunMyBizASP(MyBizASP, strVal)        

	DbDtlQuery = True

End Function

Function DbDtlQueryOk()          
	Call ggoOper.LockField(Document, "Q")
	vspdData1.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
 <TR>
  <TD HEIGHT=40>
   <FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>     
     <TR>
      <TD CLASS=TD5 NOWRAP>공장</TD>
      <TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="14xxxU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
      <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
      <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
     </TR>
     <TR>
      <TD CLASS=TD5 NOWRAP>품목</TD>
      <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="12xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
      <TD CLASS=TD5 NOWRAP>창고</TD>
      <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=15 tag="14" ALT="창고명"></TD>
     </TR>
     <TR>
      <TD CLASS=TD5 NOWRAP>규격</TD>
      <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=40 tag="14" ALT="규격"></TD>
      <TD CLASS=TD5 NOWRAP>안전재고</TD>
      <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/i2312ra1_fpDoubleSingle1_txtSafetyStock.js'></script></TD>
     </TR> 
    </TABLE>
   </FIELDSET>
  </TD>
 </TR>
 <TR>
  <TR HEIGHT="40%">
   <TD WIDTH="100%" colspan=4>
    <script language =javascript src='./js/i2312ra1_vaSpread1_vspdData1.js'></script>
   </TD>
  </TR>
  <TR HEIGHT="60%">
   <TD WIDTH="100%" colspan=4>
    <script language =javascript src='./js/i2312ra1_vaSpread2_vspdData2.js'></script>
   </TD>
  </TR> 
 </TR>
 <TR>
  <TD HEIGHT=30>
   <TABLE CLASS="basicTB" CELLSPACING=0>
    <TR>
     <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
     <TD WIDTH=30% ALIGN=RIGHT>
     <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hSlCd" tag="24" TABINDEX="-1">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
