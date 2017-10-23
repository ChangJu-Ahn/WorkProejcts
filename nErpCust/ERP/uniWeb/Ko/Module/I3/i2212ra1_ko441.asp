<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I2212ra1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 현재고 상세 조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/01
'*  8. Modified date(Last)  : 2003/05/12
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : SeungWook Lee
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "i2212rb1_ko441.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_ItemCode        				
Dim C_ItemName        
Dim C_TrackingNo      
Dim C_LotNo           
Dim C_LostSubNo       

'2008-05-19 4:53오후 :: hanc
Dim C_EXT2_CD   				
Dim C_BP_CD    				
Dim C_BP_NM

Dim C_GoodQty         
Dim C_ConvUnitGoodQty 
Dim C_BadQty          
Dim C_InspQty         
Dim C_TrnsQty         
Dim C_PickingQty      
Dim C_PrevGoodQty     
Dim C_PrevBadQty      
Dim C_PrevInspQty     
Dim C_PrevTrnsQty     

Dim arrReturn
Dim arrParam
Dim arrSL_Cd
Dim arrSL_Nm
Dim arrItem_Cd
Dim arrItem_Nm
Dim arrTracking_No
Dim arrPlant_Cd
Dim arrUserFlag
Dim arrLotNo
Dim arrTrns_Unit
		
'------ Set Parameters from Parent ASP ------
arrParam 		= window.dialogArguments
Set PopupParent = arrParam(0)

arrSL_Cd 		= arrParam(1)
arrItem_Cd 		= arrParam(2)
arrTracking_No  = arrParam(3)
arrPlant_Cd		= arrParam(4)
arrUserFlag		= arrParam(5)
arrLotNo		= arrParam(6)
arrSL_Nm 		= arrParam(7)
arrItem_Nm 		= arrParam(8)
arrTrns_Unit    = arrParam(9)

top.document.title = PopupParent.gActivePRAspName
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevSubKey
Dim lgUserFlag				

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop          
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntGrpCount = 0                       
    '---- Coding part--------------------------------------------------------------------    
    lgLngCurRows = 0                          
    Self.Returnvalue = Array("")
    lgStrPrevKey	=	""
    lgStrPrevSubKey =	""
    lgIntFlgMode	=	PopupParent.OPMD_CMODE	
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
 	txtItem_Cd.value 	 = arrItem_Cd	
	txtItem_Nm.value 	 = arrItem_Nm	
	txtSL_Cd.value 		 = arrSL_Cd
	txtSL_Nm.value 		 = arrSL_Nm
	txtTracking_No.value = arrTracking_No
	txtTrns_Unit.value   = arrTrns_Unit
	lgUserFlag			 = arrUserFlag
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

	With  vspdData
	    .ReDraw = false
	    .MaxCols = C_PrevTrnsQty+1												
	    .MaxRows = 0
		Call GetSpreadColumnPos("A")
	    Call AppendNumberPlace("6", "3", "0")    
	    ggoSpread.SSSetEdit C_ItemCode, "품목", 18
	    ggoSpread.SSSetEdit C_ItemName, "품목명", 25
	    ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20				
		ggoSpread.SSSetEdit C_LotNo, "Lot No.", 12
		ggoSpread.SSSetFloat C_LostSubNo,       "순번",            6,"6",                  ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, , ,"Z"
	    ggoSpread.SSSetEdit C_EXT2_CD, "STOCK TYPE", 10
	    ggoSpread.SSSetEdit C_BP_CD, "공급처", 10
	    ggoSpread.SSSetEdit C_BP_NM, "공급처명", 15
		
		ggoSpread.SSSetFloat C_GoodQty,         "양품재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_ConvUnitGoodQty, "환산재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_BadQty,          "불량재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_InspQty,         "검사중수량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_TrnsQty,         "이동재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PickingQty,      "PICKING수량",    15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevGoodQty,     "전월양품재고량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevBadQty,      "전월불량재고량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevInspQty,     "전월검사중수량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevTrnsQty,     "전월이동중수량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec

		'ggoSpread.MakePairsColumn()
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(2)
	    .ReDraw = true

        Call SetSpreadLock 
    End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode        = 1
	C_ItemName        = 2
	C_TrackingNo      = 3
	C_LotNo           = 4
	C_LostSubNo       = 5
    C_EXT2_CD   	  = 6  		
    C_BP_CD    		  = 7  
    C_BP_NM           = 8  
	C_GoodQty         = 9  
	C_ConvUnitGoodQty = 10 
	C_BadQty          = 11 
	C_InspQty         = 12 
	C_TrnsQty         = 13
	C_PickingQty      = 14
	C_PrevGoodQty     = 15
	C_PrevBadQty      = 16
	C_PrevInspQty     = 17
	C_PrevTrnsQty     = 18
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCode        = iCurColumnPos(1)
		C_ItemName        = iCurColumnPos(2)
		C_TrackingNo      = iCurColumnPos(3)
		C_LotNo           = iCurColumnPos(4)
		C_LostSubNo       = iCurColumnPos(5)
        C_EXT2_CD   	  = iCurColumnPos(6)  		
        C_BP_CD    		  = iCurColumnPos(7)  
        C_BP_NM           = iCurColumnPos(8)  
		C_GoodQty         = iCurColumnPos(9) 
		C_ConvUnitGoodQty = iCurColumnPos(10)
		C_BadQty          = iCurColumnPos(11)
		C_InspQty         = iCurColumnPos(12)
		C_TrnsQty         = iCurColumnPos(13)
		C_PickingQty      = iCurColumnPos(14)
		C_PrevGoodQty     = iCurColumnPos(15)
		C_PrevBadQty      = iCurColumnPos(16)
		C_PrevInspQty     = iCurColumnPos(17)
		C_PrevTrnsQty     = iCurColumnPos(18)		
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    vspdData.ReDraw = False    
    ggoSpread.SpreadLockWithOddEvenRowColor()
    vspdData.ReDraw = True
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt
		
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
					
		vspdData.Col = C_ItemCode
		arrReturn(0) = vspdData.Text
		vspdData.Col = C_ItemName
		arrReturn(1) = vspdData.Text
		vspdData.Col = C_TrackingNo
		arrReturn(2) = vspdData.Text
		vspdData.Col = C_LotNo
		arrReturn(3) = vspdData.Text
		vspdData.Col = C_LostSubNo
		arrReturn(4) = vspdData.Text
		vspdData.Col = C_EXT2_CD
		arrReturn(5) = vspdData.Text
		vspdData.Col = C_BP_CD
		arrReturn(6) = vspdData.Text
		vspdData.Col = C_BP_NM
		arrReturn(7) = vspdData.Text


'		vspdData.Col = C_GoodQty
'		arrReturn(5) = vspdData.Text
'		vspdData.Col = C_ConvUnitGoodQty
'		arrReturn(6) = vspdData.Text
'		vspdData.Col = C_BadQty
'		arrReturn(7) = vspdData.Text
'		vspdData.Col = C_InspQty
'		arrReturn(8) = vspdData.Text
'		vspdData.Col = C_TrnsQty
'		arrReturn(9) = vspdData.Text
'		vspdData.Col = C_PickingQty
'		arrReturn(10) = vspdData.Text
'		vspdData.Col = C_PrevGoodQty
'		arrReturn(11) = vspdData.Text
'		vspdData.Col = C_PrevBadQty
'		arrReturn(12) = vspdData.Text
'		vspdData.Col = C_PrevInspQty
'		arrReturn(13) = vspdData.Text
'		vspdData.Col = C_PrevTrnsQty
'		arrReturn(14) = vspdData.Text
			
		Self.Returnvalue = arrReturn
	End If
		
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
	Function CancelClick()
		Self.Close()
	End Function



'------------------------------------------  OpenTrnsUnit()  -------------------------------------------------
'	Name : OpenTrnsUnit()
'	Description : Entry Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrnsUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위 팝업"						
	arrParam(1) = "B_UNIT_OF_MEASURE"					
	arrParam(2) = txtTrns_Unit.value					
	arrParam(3) = ""			 						
	arrParam(4) = ""									
	arrParam(5) = "단위"							
	
	arrField(0) = "UNIT"	
	arrField(1) = "UNIT_NM"	
	
	arrHeader(0) = "단위"		
	arrHeader(1) = "단위명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		txtTrns_Unit.focus
		Exit Function
	Else
		txtTrns_Unit.value = arrRet(0)	
		txtTrns_Unit.focus
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029													
    Call ggoOper.LockField(Document, "N")                               
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables                                                  
    Call SetDefaultVal()
    Call DbQuery()

End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = vspdData
   
	If vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <= 0 Then
		ggoSpread.Source = vspdData 
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
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If vspdData.MaxRows = 0 Then
		Exit Sub
	End If

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   parent.ggoSpread.Source = vspdData
   Call parent.ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   ggoSpread.Source = vspdData
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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 



'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If CheckRunningBizProcess = True Then
			Exit Sub
		End If

    if  vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then	
		If lgStrPrevKey <> "" and lgStrPrevSubKey <> "" Then				
			DbQuery
		End If
	End if 

End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'   Event Desc : 
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
	On error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	End if
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
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                      
    
    Err.Clear                                                             

    '-----------------------
    'Erase contents area
    '-----------------------
    Call InitVariables 													

    ggoSpread.Source = vspdData
    ggoSpread.ClearSpreadData

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Exit Function
	End if
       
    FncQuery = True														
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call PopupParent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    On Error Resume Next                                                   
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      
    Call LayerShowHide(1)
    DbQuery = False
    
    Err.Clear                                                             
    
    Dim strVal
    Dim strFlag
    
	If RadioOutputType.rdoCase1.Checked Then
		strFlag = "Y"
	Else
		strFlag = "N"
	End if
    
    if lgIntFlgMode = PopupParent.OPMD_UMODE Then		
    	
    	strVal = BIZ_PGM_ID		& "?txtSL_Cd="		& Trim(txtSL_Cd.value) & _			
								"&txtItem_Cd="      & Trim(txtItem_Cd.value) & _
								"&txtTracking_No="	& Trim(arrTracking_No) & _
								"&txtTrns_Unit="    & Trim(hTrns_Unit.value) & _
								"&txtFlag="			& strFlag & _
								"&txtPlant_Cd="     & Trim(arrPlant_Cd) & _
								"&lgStrUserFlag="   & lgUserFlag & _
								"&txtLotNo="        & Trim(arrLotNo) & _
								"&lgStrPrevKey="    & lgStrPrevKey & _
								"&lgStrPrevSubKey=" & lgStrPrevSubKey & _
								"&txtMaxRows="		& vspdData.MaxRows
		
    Else
    
		strVal = BIZ_PGM_ID		& "?txtSL_Cd="		& Trim(txtSL_Cd.value) & _			
								"&txtItem_Cd="      & Trim(txtItem_Cd.value) & _
								"&txtTracking_No="	& Trim(arrTracking_No) & _
								"&txtTrns_Unit="    & Trim(txtTrns_Unit.value) & _
								"&txtFlag="			& strFlag & _
								"&txtPlant_Cd="     & Trim(arrPlant_Cd) & _
								"&lgStrUserFlag="   & lgUserFlag & _
								"&txtLotNo="        & Trim(arrLotNo) & _
								"&lgStrPrevKey="    & lgStrPrevKey & _
								"&lgStrPrevSubKey=" & lgStrPrevSubKey & _
								"&txtMaxRows="		& vspdData.MaxRows
    
    End if
'    msgbox strval    
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
	vspdData.focus
	
	lgIntFlgMode	=	PopupParent.OPMD_UMODE	
End Function

 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
			     <TD CLASS="TD5" NOWRAP>품목</TD>
				 <TD CLASS="TD6" NOWRAP><input NAME="txtItem_Cd" TYPE="Text" MAXLENGTH="18" tag="14XXXU" ALT = "품목" size=15>&nbsp;<input NAME="txtItem_Nm" TYPE="Text" MAXLENGTH="40" tag="14N"></TD>
				 <TD CLASS="TD5" NOWRAP>창고</TD>
			     <TD CLASS="TD6" NOWRAP><input NAME="txtSL_Cd" TYPE="Text" MAXLENGTH="7" tag="14XXXU" ALT = "창고" size=10>&nbsp;<input NAME="txtSL_Nm" TYPE="Text" MAXLENGTH="40" tag="14N"></TD>
			</TR>
			<TR>
				 <TD CLASS="TD5" NOWRAP>Tracking No</TD>
			     <TD CLASS="TD6" NOWRAP><input NAME="txtTracking_No" TYPE="Text" MAXLENGTH="25" tag="14" ALT = "Tracking No" size=25></TD>
				 <TD CLASS="TD5" NOWRAP>환산단위</TD>
				 <TD CLASS="TD6" NOWRAP><input NAME="txtTrns_Unit" TYPE="Text" MAXLENGTH="3" tag="11XXXU" ALT = "환산단위" size=8><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenTrnsUnit()"></TD>
		    </TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>수량유무여부</TD>
				<TD CLASS="TD6">
				<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X" ><LABEL FOR="rdoCase1">수량있음</LABEL>
				<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">전품목</LABEL>
				</TD>
				<TD CLASS="TD5" NOWRAP></TD>
				<TD CLASS="TD6" NOWRAP></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
	</TD></TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
					                          <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
				    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTrns_Unit" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTracking_No" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hSL_Cd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItem_Cd" tag="24" TABINDEX="-1">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


