
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List ROP
'*  3. Program ID           : I2411ma1.asp
'*  4. Program Name         : ROP 정보 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'
'*  7. Modified date(First) : 2000/05/03
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/18 : ..........
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID			= "i2411mb1.asp"          
Const BIZ_PGM_PURREQSAVE_ID = "i2411mb2.asp"

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_Check        
Dim C_ItemCd
Dim C_ItemName
Dim C_ItemSpec
Dim C_BaseUnit
Dim C_TrackingNo
Dim C_AvlQty
Dim C_Rop
Dim C_Qty
Dim C_ReqDt
Dim C_PurReqDt
Dim C_OnhandQty
Dim C_SchdRcptQty
Dim C_SchdIssueQty
Dim C_SafetyStock
Dim C_FixedOrderQty
Dim C_PurLT

'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgEndDate
 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

lgIntFlgMode = Parent.OPMD_CMODE 
lgIntGrpCount = 0                         
lgBlnFlgChgValue = False                  
'---- Coding part--------------------------------------------------------------------
lgStrPrevKey = ""
lgLngCurRows = 0                          

End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal() 
	lgEndDate = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
	lgBlnFlgChgValue = False             

	With frm1
	
		.btnRun.Disabled = True 
		
		If .txtPlantCd.value = "" Then .txtPlantNm.value = ""
		If .txtItemCd.value = "" Then .txtItemNm.value = ""
		 
		If Parent.gPlant <> "" Then
			.txtPlantCd.value = UCase(Parent.gPlant)
			.txtPlantNm.value = Parent.gPlantNm
			.txtPlantCd.focus
		Else
			.txtPlantCd.focus
		End If     
	
	End With
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData  
	ggoSpread.Spreadinit "V20051101", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_PurLT+1        
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		  
		ggoSpread.SSSetCheck C_Check, "", 4,,,1
		ggoSpread.SSSetEdit C_ItemCd, "품목", 18
		ggoSpread.SSSetEdit C_ItemName, "품목명", 25  
		ggoSpread.SSSetEdit C_ItemSpec, "규격", 10
		ggoSpread.SSSetEdit C_BaseUnit, "재고단위", 10
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No.", 25
		ggoSpread.SSSetFloat C_Rop, "발주점", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_Qty, "수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_OnhandQty, "현재고량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_SchdRcptQty, "입고예정량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_SchdIssueQty, "출고예정량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_SafetyStock, "안전재고량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_FixedOrderQty, "고정발주량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_AvlQty, "가용재고", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetDate C_ReqDt, "필요일", 10,2,Parent.gDateFormat
		ggoSpread.SSSetDate C_PurReqDt, "구매요청일", 10, 2,Parent.gDateFormat  
		ggoSpread.SSSetEdit C_PurLT, "구매L/T", 10,2
		  
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	  
		Call SetSpreadLock

		.ReDraw = true
		
		ggoSpread.SSSetSplit2(3)
		
	End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_Check = 1
	C_ItemCd = 2      
	C_ItemName = 3
	C_ItemSpec = 4
	C_BaseUnit = 5
	C_TrackingNo = 6
	C_AvlQty = 7
	C_Rop = 8
	C_Qty = 9
	C_ReqDt = 10
	C_PurReqDt = 11
	C_OnhandQty = 12
	C_SchdRcptQty = 13
	C_SchdIssueQty = 14
	C_SafetyStock = 15
	C_FixedOrderQty = 16
	C_PurLT = 17
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_Check         = iCurColumnPos(1)
		C_ItemCd        = iCurColumnPos(2)
		C_ItemName      = iCurColumnPos(3)
		C_ItemSpec      = iCurColumnPos(4)
		C_BaseUnit      = iCurColumnPos(5)
		C_TrackingNo	= iCurColumnPos(6)
		C_AvlQty        = iCurColumnPos(7)
		C_Rop           = iCurColumnPos(8)
		C_Qty           = iCurColumnPos(9)
		C_ReqDt         = iCurColumnPos(10)
		C_PurReqDt      = iCurColumnPos(11)
		C_OnhandQty     = iCurColumnPos(12)
		C_SchdRcptQty   = iCurColumnPos(13)
		C_SchdIssueQty  = iCurColumnPos(14)
		C_SafetyStock   = iCurColumnPos(15)
		C_FixedOrderQty = iCurColumnPos(16)
		C_PurLT         = iCurColumnPos(17)
	End Select
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock C_ItemCd, -1
	'ggoSpread.SpreadLockWithOddEvenRowColor()
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
	 
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantCd.focus
		Exit function
	End If
	  
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
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","x")
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
    
    arrHeader(0) = "Tracking_No"		
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

 '------------------------------------------  JumpPurReq()  -------------------------------------------------
' Name : JumpPurReq()
' Description : Pur Requistion Registration Reference Windows
'--------------------------------------------------------------------------------------------------------- 
Function JumpPurReq()
 Dim intRet
 If lgBlnFlgChgValue = True Then  
  
  IntRet = DisplayMsgBox("900017",Parent.VB_YES_NO,"X", "X")
  If intRet = vbNo Then     
   Exit Function
  End If
 End If
 
 WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
 WriteCookie "txtItemNm", frm1.txtItemCd.value 
 
 PgmJump(BIZ_PGM_JUMPPURREQ_ID)

End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029           
    Call ggoOper.LockField(Document, "N")                                       
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call InitSpreadSheet          
    Call InitVariables            
    Call SetDefaultVal
	Call SetToolbar("11000000000011")

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
	 
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	 
	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
		If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
			Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 
	Dim IPurLT
	Dim IAvlQty
	Dim IRop
	Dim IFixedOrderQty
	Dim lVlaidDt 
	Dim intRet
  
 
  
	ggoSpread.Source = frm1.vspdData
	     
	With frm1.vspdData
		.ReDraw = False 
		.Col = Col
		.Row = Row
		IF Col <> 1 Then     
			Exit sub
		Else
			IF Row > 0 And Col = C_Check Then
				.Col = Col
				.Row = Row         
				IF .Text = 1 Then
					     
					.Col = 0
					.Text = ggoSpread.UpdateFlag
					    
					.Col = C_PurLT
					IPurLT = .Text 
					.Col = C_Rop
					IRop = .Text
					.Col = C_AvlQty
					IAvlQty = .Text
					.Col = C_FixedOrderQty
					IFixedOrderQty = .Text      
					.Col = C_Qty
					.Text = IFixedOrderQty
					'.Text = IRop - IAvlQty
					.Col = C_PurReqDt     
					.Text = lgEndDate 
					.Col = C_ReqDt
					     
					.Text = UNIDateAdd("d", IPurLT, lgEndDate, Parent.gDateFormat)
					
					.Col = C_ItemCd
					'	IItem = Trim(.Text)     
					CALL CommonQueryRs(" VALID_TO_DT "," B_ITEM_BY_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")  & "AND ITEM_CD = " & FilterVar(.Text, "''", "S"), _
																lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					lVlaidDt = split(lgF0, chr(11))		
					
					IF lgEndDate > lVlaidDt(0) Then	
						IntRet = DisplayMsgBox("169982",Parent.VB_YES_NO,"X", "X")
						If intRet = vbNo Then  
							.Col =  C_Check
							.Text = 0  
							Exit Sub
						End If
					END IF	
  									
					ggoSpread.SpreadUnLock C_Qty,Row,C_Qty,Row
					ggoSpread.SpreadUnLock C_ReqDt,Row,C_ReqDt,Row
					ggoSpread.SpreadUnLock C_PurReqDt,Row,C_PurReqDt,Row
					ggoSpread.SsSetRequired C_Qty, Row, Row 
					ggoSpread.SsSetRequired C_PurReqDt, Row, Row 
					ggoSpread.SsSetRequired C_ReqDt, Row, Row 
					lgBlnFlgChgValue = True 
				ELSEIF .Text = 0 Then
					.Col = 0
					.Text = ""
					     
					.Col = C_Qty
					.Text = ""
					.Col = C_PurReqDt     
					.Text = ""
					.Col = C_ReqDt     
					.Text = ""
					ggoSpread.EditUndo                  
					ggoSpread.SpreadLock C_Qty,Row,C_Qty,Row
					ggoSpread.SpreadLock C_ReqDt,Row,C_ReqDt,Row
					ggoSpread.SpreadLock C_PurReqDt,Row,C_PurReqDt,Row
					lgBlnFlgChgValue = False
				END IF         
			END If     
		END if

		.ReDraw = True 
	         
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	 
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then 
		If lgStrPrevKey <> "" Then       
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End if 
	End if   
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
	If Row <= 0 Then Exit Sub
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
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
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	If NewCol = C_Check Or Col = C_Check Then
		Cancel = True
		Exit Sub
	End If

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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                   
    
    Err.Clear                                                             
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X", "X", "X")                         
        Exit Function
    End If
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                         
       IntRetCD = DisplayMsgBox("900005","X", "X", "X")                      
       Exit Function
    End If
    
  '-----------------------
    'Save function call area
    '-----------------------
     IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")
	If IntRetCD = vbNo Then Exit Function
                                                   
	frm1.btnRun.Disabled = True

	If DbSave = False Then Exit Function
    
    FncSave = True                                                    
    
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()    
    Dim IntRetCD 

	FncQuery = False                                               
	 
	Err.Clear                                                              
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X") 
		If IntRetCD = vbNo Then Exit Function
	End If

	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then       
		Call SetDefaultVal
		Exit Function
	End If

	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")        
	Call InitVariables              
	frm1.btnRun.Disabled = True

	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0) 

	'-----------------------
	'Check Plant CODE  '품목코드가 있는 지 체크 
	'-----------------------
	If frm1.txtItemCd.Value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			  
			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItemNm.value = lgF0(0)
		Else
			frm1.txtItemNm.value = ""
		End If
	End If

	   
	'-----------------------
	'Query function call area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_CMODE  
	Call SetToolbar("11000000000111")     
	
	If DbQuery = False Then	Exit Function
		 
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
	call parent.FncExport(Parent.C_MULTI)                         
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , True)                                               
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X", "X")  
		If IntRetCD = vbNo Then Exit Function
	End If
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

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery()
 Dim LngLastRow
 Dim LngMaxRow
 Dim StrNextKey
 
 DbQuery = False
 
 Call LayerShowHide(1)
 
 Err.Clear                                                           
 
 Dim strVal
 
 With frm1
 
  if lgIntFlgMode = Parent.OPMD_UMODE Then   
   strVal = BIZ_PGM_ID & "?txtPlantCd="		& Trim(.hPlantCd.value)			& _
						"&txtItemCd="		& Trim(.hItemCd.value)			& _
						"&txtTrackingNo="	& Trim(.hTrackingNo.value)		& _	 
						"&hDate="			& Trim(.hDate.value)			& _      
						"&lgStrPrevKey="	& lgStrPrevKey					& _
						"&txtMaxRows="		& .vspdData.MaxRows			
  Else
   strVal = BIZ_PGM_ID & "?txtPlantCd="		& Trim(.txtPlantCd.value)		& _  
						"&txtItemCd="		& Trim(.txtItemCd.value)		& _
						"&txtTrackingNo="	& Trim(.txtTrackingNo.value)	 & _
						"&txtQryFrItemCd="	& Trim(.txtItemCd.value)		& _
						"&hDate="			& Trim(.hDate.value)			& _      
						"&lgStrPrevKey="	& lgStrPrevKey					& _
						"&txtMaxRows="		& .vspdData.MaxRows
  End if 
  
  Call RunMyBizASP(MyBizASP, strVal)     
 End With
 
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
	Dim i
	Dim Avl_Qty
	Dim Rop_Qty

	lgIntFlgMode = Parent.OPMD_UMODE 

	lgBlnFlgChgValue = False

	With frm1.vspdData

		frm1.hRunChk.Value = 0 
		
		.ReDraw = False  

		For i = 1  to .MaxRows 
			.Row = i 
		    .Col = C_AvlQty 
		    Avl_Qty = UNICDbl(.Text) 
		    .Col = C_Rop 
		    Rop_Qty = UNICDbl(.Text) 
		    
		    IF Avl_Qty > Rop_Qty Then 
				ggoSpread.SpreadLock C_Check , i , C_Check, i
			Else 
		        frm1.hRunChk.Value = 1 
		    End if 
		    
		Next  
		frm1.vspdData.focus
	    .ReDraw = True  

	End With
	frm1.vspdData.focus
	IF frm1.hRunchk.value = 1 Then	frm1.btnRun.Disabled = False

End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 

 Dim lRow        
    Dim lGrpCnt     
    Dim IItem      
    Dim IQty   
    Dim IUnit   
    Dim lReqDt    
    Dim lPurReqDt
    '********** tracking no add LSW 2005-11-01 ******************
    Dim lTrackingNo
    Dim strVal
 
   'Show Processing Bar
    
        
        Err.Clear 
 
    DbSave = False                                                      
    
    On Error Resume Next                                                 

	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtUpdtUserId.value = Parent.gUsrID
	frm1.txtInsrtUserId.value = Parent.gUsrID
 
	 With frm1.vspdData
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
    
		For lRow = 1 To .MaxRows
    
		    .Row = lRow
		    .Col = 0
		    
		    Select Case .Text
		                        
		        Case ggoSpread.UpdateFlag      
    
					.Col = C_Check
    				If Trim(.Text) <> "0" then   
						.Col = C_ItemCd
						IItem = Trim(.Text)
						.Col = C_Qty
						IQty = Trim(.Text)
						.Col = C_BaseUnit
						IUnit = Trim(.Text)
						.Col = C_ReqDt
						lReqDt = Trim(.Text)
						.Col = C_PurReqDt
						lPurReqDt = Trim(.Text)
						.Col = C_TrackingNo
						lTrackingNo = Trim(.Text)
						
						If  UniCDbl(IQty) = 0 Then
							Call  DisplayMsgBox("169918","X", "X", "X")
							frm1.btnRun.Disabled = False 
							Exit Function
						End if

						strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep & _
										IItem & Parent.gColSep & IQty & Parent.gColSep & IUnit & Parent.gColSep & _
										lReqDt & Parent.gColSep & lPurReqDt & Parent.gColSep & _
										lTrackingNo & parent.gColSep & Parent.gUsrID & Parent.gRowSep
						
						lGrpCnt = lGrpCnt + 1

		            End if
		                         
		    End Select
		            
		Next
 
		frm1.txtMaxRows.value = lGrpCnt-1
		frm1.txtSpread.value = strVal
 
		If lGrpCnt <= 1 then    
			Call DisplayMsgBox("169911","X", "X", "X")   
		End if
		
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_PURREQSAVE_ID)  
	
	End With
 
    DbSave = True   
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()             

	Dim PrNo

    Call DisplayMsgBox("169914","X", "X", "X")   
    
   	Call ggoOper.ClearField(Document, "2")        
    Call InitVariables    
    Call MainQuery()

End Function

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
						<TABLE <%=LR_SPACE_TYPE_10%>>
							<TR>
								<TD WIDTH=10>&nbsp;</TD>
								<TD CLASS="CLSMTABP">
									<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
											<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ROP</font></td>
											<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
										</TR>
									</TABLE>
								</TD>
								<TD WIDTH=* align=right>&nbsp;</TD>     
								<TD WIDTH=10>&nbsp;</TD>
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
										<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>공장</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="13XXXU" ALT = "공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="14"></TD>
												<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()"></TD>
											</TR>
											<TR>      
												<TD CLASS="TD5" NOWRAP>품목</TD>      
												<TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT = "품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="14"></TD>
												<TD CLASS="TD5" NOWRAP></TD>
												<TD CLASS="TD6" NOWRAP></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=* valign=top>
									<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
										<TR>
											<TD HEIGHT="100%">
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
				<TR HEIGHT=20>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_30%> >
							<TR>
								<TD WIDTH=10></TD>
								<TD><BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:FncSave()" Flag=2>구매요청등록</BUTTON>
								</TD>
								<TD WIDTH=10></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>        
				<TR>
					<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
					</TD>
				</TR>
			</TABLE>

		<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
			<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hDate" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hRunChk" tag="24" TABINDEX="-1">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

