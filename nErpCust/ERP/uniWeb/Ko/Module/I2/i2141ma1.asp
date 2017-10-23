<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 실사Posting Manual작업 
'*  3. Program ID           : I21411Post phy inv Svr
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*			       I21411Post Phy Inv Svr
'*			       I21119Lookup Phy inv Svr
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2006/08/29
'*  9. Modifier (First)     : Mr  Kim 
'* 10. Modifier (Last)      : LEE SEUNG WOOK
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit										

Const BIZ_PGM_ID = "i2141mb2.asp"								
Const BIZ_LOOKUP_ID = "i2141mb1.asp"							

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgCheckall
Dim lgPrevMaxRows

Dim C_Check 
Dim C_ItemCd									
Dim C_ItemNm
Dim C_Spec 
Dim C_ItemUnit
Dim C_DiffQty 
Dim C_DiffAmount
Dim C_InvQty
Dim C_Qty
Dim C_InvAmount
Dim C_Amount
Dim C_TrackingNo 
Dim C_LotNo 
Dim C_LotSubNo
Dim C_SeqNo
Dim C_SubCheck 

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

	lgIntFlgMode = Parent.OPMD_CMODE         	         
	lgBlnFlgChgValue = False     	               	
	lgIntGrpCount = 0                           	
	lgStrPrevKey1 = "" 
	lgStrPrevKey2 = "" 
	                          	
	lgLngCurRows = 0            
	lgCheckall = 0
    Call SetToolbar("11000000000011")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.btnRun.Disabled = True
	If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = ""
	End if	
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtCondPhyInvNo.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
 		ggoSpread.Spreadinit "V20030425", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		.MaxCols = C_SubCheck + 1						
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")
		
		ggoSpread.SSSetCheck C_Check, "", 4,,,1
		ggoSpread.SSSetEdit C_ItemCd, "품목", 18
		ggoSpread.SSSetEdit C_ItemNm, "품목명", 25
		ggoSpread.SSSetEdit C_Spec, "규격", 20
		ggoSpread.SSSetEdit C_ItemUnit, "단위", 10,2
	 	ggoSpread.SSSetFloat C_DiffQty, "차이수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec   '수량'
	 	ggoSpread.SSSetFloat C_DiffAmount, "차이금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec     '금액'
	 	ggoSpread.SSSetFloat C_InvQty, "실사수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_Qty, "전산수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_InvAmount, "실사금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_Amount, "전산금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20		
		ggoSpread.SSSetEdit C_LotNo, "LOT NO", 12
		ggoSpread.SSSetFloat C_LotSubNo, "Lot No.순번", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
		ggoSpread.SSSetEdit C_SeqNo,"",5
		ggoSpread.SSSetEdit C_SubCheck, "",4
		
 		Call ggoSpread.SSSetColHidden(C_SeqNo, C_SubCheck, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		ggoSpread.SSSetSplit2(3)  
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SpreadUnLock C_Check, -1, C_Check
		.ReDraw = true
    End With
    
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_Check			= 1
	C_ItemCd		= 2									
	C_ItemNm		= 3
	C_Spec			= 4
	C_ItemUnit		= 5
	C_DiffQty		= 6
	C_DiffAmount	= 7
	C_InvQty		= 8
	C_Qty			= 9
	C_InvAmount		= 10
	C_Amount		= 11
	C_TrackingNo	= 12
	C_LotNo			= 13
	C_LotSubNo		= 14
	C_SeqNo			= 15
	C_SubCheck		= 16
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_Check		= iCurColumnPos(1)
		C_ItemCd	= iCurColumnPos(2)									
		C_ItemNm	= iCurColumnPos(3)
		C_Spec		= iCurColumnPos(4)
		C_ItemUnit	= iCurColumnPos(5)
		C_DiffQty	= iCurColumnPos(6)
		C_DiffAmount= iCurColumnPos(7)
		C_InvQty	= iCurColumnPos(8)
		C_Qty		= iCurColumnPos(9)
		C_InvAmount = iCurColumnPos(10)
		C_Amount	= iCurColumnPos(11)
		C_TrackingNo= iCurColumnPos(12)
		C_LotNo		= iCurColumnPos(13)
		C_LotSubNo	= iCurColumnPos(14)
		C_SeqNo		= iCurColumnPos(15)
		C_SubCheck	= iCurColumnPos(16)
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
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

'------------------------------------------  OpenPhyInvNo()  --------------------------------------------
Function OpenPhyInvNo()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam1,arrParam2,arrParam3,arrParam4
    
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus  
		Exit Function
	Else
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	iCalledAspName = AskPRAspName("i2111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "i2111pa1", "X")
		IsOpenPop = False
		Exit Function
    End If

	IsOpenPop = True

	arrParam1 = frm1.txtCondPhyInvNo.value
	arrParam2 = "PD"
	arrParam3 = frm1.txtPlantCd.value
	arrParam4 = ""
        	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam1,arrParam2,arrParam3,arrParam4), _
 		 "dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCondPhyInvNo.focus
		Exit Function
	Else
		frm1.txtCondPhyInvNo.Value  = arrRet(0)		
		frm1.txtInspDt.Value		= arrRet(1)	
		frm1.txtSLCd.Value 			= arrRet(2)
		frm1.txtSLNm.Value			= arrRet(3)
	
		If arrRet(4) = "Y" Then
			frm1.rdoDisplayFlg1.checked = True
		End If
			
		frm1.txtCondPhyInvNo.focus
	End If	
End Function

'------------------------------------------  OpenCostCd()  -------------------------------------------------
  Function OpenCostCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtCostCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
	    frm1.txtPlantCd.Focus
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "Cost Center 팝업"			
	arrParam(1) = "B_COST_CENTER A,B_PLANT B"
	arrParam(2) = Trim(frm1.txtCostCd.Value)		
	arrParam(3) = ""								
	arrParam(4) = "A.BIZ_AREA_CD = B.BIZ_AREA_CD AND B.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "Cost Center"					
	
	arrField(0) = "COST_CD"
	arrField(1) = "COST_NM"
    
	arrHeader(0) = "Cost Center"			    	
	arrHeader(1) = "Cost Center 명"				

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		frm1.txtCostCd.value = arrRet(0)
		frm1.txtCostNm.value = arrRet(1)
		frm1.txtCostCd.focus
	End If	
End Function

'------------------------------------------  OpenItemOrigin()  --------------------------------------------------
Function OpenItemCd()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus  
		Exit Function
	Else
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("b1b11pa3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
    End If
    
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	       
	arrParam(1) = Trim(frm1.txtItemCd.Value)	  
	arrParam(2) = ""				              
	arrParam(3) = ""				              
	
	arrField(0) = 1 
    arrField(1) = 2 
    arrField(2) = 9 
    arrField(3) = 6 
    arrField(4) = 45
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenMvmtListRef()  -------------------------------------------------
' Name : OpenMvmtListRef()
' Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenMvmtListRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1 
	Dim Param2
	Dim Param3
	Dim Param4
	Dim Param5
	Dim Param6
	 
	If IsOpenPop = True Then Exit Function

	Param1 = Trim(frm1.txtPlantCd.value)
	Param2 = Trim(frm1.txtPlantNm.value)
	Param3 = Trim(frm1.txtCondPhyInvNo.value)
	Param4 = Trim(frm1.txtInspDt.value)
	
	If Param1 = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	If Param3 = "" then
		Call DisplayMsgBox("169971","X", "X", "X")    
		frm1.txtCondPhyInvNo.focus
		Exit Function
	End If
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("800167","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData    
	With frm1.vspdData     
		.Col = C_ItemCd
		.Row = .ActiveRow
		Param5 = Trim(.Text )
		
		.Col = C_ItemNm
		.Row = .ActiveRow
		Param6 = Trim(.Text)
	End With
	
	If Param1 = "" then
		Call DisplayMsgBox("169903","X", "X", "X")    
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2141RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2141RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")      
	     
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	End If 
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    If GetSetupMod(Parent.gSetupMod, "A") = "Y" Then
		txtCostTitle.style.display = ""
	Else
		frm1.txtCostCd.tag = "25"
		ggoOper.SetReqAttr frm1.txtCostCd, "Q"
	End if
	
    Call LoadInfTB19029                                                     				
    Call ggoOper.LockField(Document, "N")                                   				
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)	
    
    Call InitSpreadSheet                                                    				
    Call InitVariables                                                      				
    
    Call SetDefaultVal
    Call InitComboBox
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_DblClick(Button)
'=======================================================================================================
Sub txtInspDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtInspDt.Focus
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
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
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_Check Then
			.Col = Col
			.Row = Row									

			IF .Text = "1" Then
				
				.Col = C_SubCheck
				IF .Text = "0" Then
					.Col = 0
					.Text = ggoSpread.UpdateFlag
					lgBlnFlgChgValue = True
				else
					.Col = 0
					.Text = ""
				End If
	
			Elseif .Text = "0" Then
				
				.Col = C_SubCheck
				IF .Text = "0" Then
					.Col = 0
					.Text = ""
				else
					.Col = C_Check
					.Text = "1"
				End If
				
				lgBlnFlgChgValue = False
				
			End if  			
		End If	
	End With
End Sub

'=========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey1 <> "" and lgStrPrevKey2 <> "" Then		'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End if
	End if 
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    If lgIntFlgMode = Parent.OPMD_CMODE Then
 		Call SetPopupMenuItemInf("0000111111") 
 	Else
 	 	Call SetPopupMenuItemInf("0001111111")
 	End If
 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
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
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	If NewCol = C_Check or Col = C_Check Then
		Cancel = True
		Exit Sub
	End If
	    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData

    Dim lRow 
    With frm1.vspdData
		.ReDraw = False
		For lRow = 1 To .MaxRows 
			.Col = C_Check
			.Row = lRow									
			IF .Text = "1" Then
				ggoSpread.SpreadLock C_Check , lRow, C_Check, lRow
			End	If	
		Next 
		.ReDraw = True
	End With	
End Sub 

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                      
    Err.Clear                                                             

    '-----------------------
    'Check condition area
    '-----------------------
    If Not ChkField(Document, "1") Then Exit Function				
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then	
    	IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")			
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables
  
    If Plant_PhyInvNo_Check = False Then Exit Function
											
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then Exit Function
	   
    FncQuery = True							

End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                     
    Err.Clear                                                          
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then		
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                  
    Call ggoOper.LockField(Document, "N")                  
    Call InitVariables                                     
    Call SetDefaultVal   
    
    FncNew = True                                          

End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
	Dim IntRetCD
	Dim lRow

	FncSave = False                                            
	
	Err.Clear                                                  
	On Error Resume Next                                      
	
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = 0
	
	If frm1.vspdData.Text = ggoSpread.UpdateFlag Then
		If Not ChkField(Document, "2") Then Exit Function	
	End If
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
	    IntRetCD = DisplayMsgBox("900001","X", "X", "X")                       
 		Exit Function
	End If
	
   	frm1.btnRun.Disabled = True
   	
	If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) OR _
  		UCase(Trim(frm1.txtCondPhyInvNo.value)) <> UCase(Trim(frm1.txthCondPhyInvNo.value)) Then

		Call DisplayMsgBox("900002","X","X","X")
	   	frm1.btnRun.Disabled = False
		Exit Function

	End If
	
	
	If Trim(frm1.txtCostCd.Value) <> "" Then
		If 	CommonQueryRs(" COST_NM "," B_COST_CENTER ", " COST_CD = " & FilterVar(frm1.txtCostCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
			Call DisplayMsgBox("124400","X","X","X")
			frm1.txtCostNm.Value = ""
			frm1.txtCostCd.focus
		   	frm1.btnRun.Disabled = False
 			Exit Function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtCostNm.Value = lgF0(0)
	End If
	
	
	'-----------------------
	'Save function call area
	'-----------------------
	If DBSave() = False Then 
	   	frm1.btnRun.Disabled = False
		Exit Function
	End If
	
	FncSave = True                                                    

End Function

'============================================= FncDeleteRow() =========================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 
	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel()
    If frm1.vspdData.maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                         
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)								
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                  
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 

    'Show Processing Bar
    Call LayerShowHide(1)       

    DbQuery = False
    
    Err.Clear                                               

    Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_LOOKUP_ID &	"?txtMode="			& Parent.UID_M0001				& _
									"&txtPlantCd="      & Trim(.txthPlantCd.value)		& _
									"&txtItemCd="       & Trim(.txtItemCd.value)		& _
									"&lgStrPrevKey1="   & lgStrPrevKey1					& _
									"&lgStrPrevKey2="   & lgStrPrevKey2					& _
									"&txtCondPhyInvNo=" & Trim(.txthCondPhyInvNo.value)	& _
									"&rdoDisplayFlg="   & Trim(.txthDisplayFlg.value)	& _
									"&txtMaxRows="      & .vspdData.MaxRows
    else
		strVal = BIZ_LOOKUP_ID &	"?txtMode="			& Parent.UID_M0001				& _
									"&txtPlantCd="      & Trim(.txtPlantCd.value)		& _
									"&txtItemCd="       & Trim(.txtItemCd.value)		& _
									"&txtCondPhyInvNo=" & Trim(.txtCondPhyInvNo.value)	& _
									"&txtMaxRows="      & .vspdData.MaxRows
		If .rdoDisplayFlg1.checked = True Then
			strVal = strVal	& "&rdoDisplayFlg="   & Trim(.rdoDisplayFlg1.value)
		Else
			strVal = strVal	& "&rdoDisplayFlg="   & Trim(.rdoDisplayFlg2.value)
		End If   
    End if
    
	lgPrevMaxRows = .vspdData.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)							
       
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()						
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Dim lRow
    Dim strCanFlag
    
    If 	CommonQueryRs(" REFERENCE "," B_CONFIGURATION ", " MAJOR_CD = " & FilterVar("I0017", "''", "S") & " AND MINOR_CD = " & FilterVar("01","''","S") _
						& " AND SEQ_NO = 1 ", _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
		lgF0 = Split(lgF0, Chr(11))
		strCanFlag = lgF0(0)
	Else
		strCanFlag = "N"	
	End If
	        

	With frm1.vspdData
	
		.Redraw = False
	
		ggoSpread.Source = frm1.vspdData
		
		If frm1.txthPhyInvPosBlk.Value = "N" Then

			ggoSpread.SpreadUnLock C_Check, lgPrevMaxRows + 1, C_Check, .MaxRows 
		
			If frm1.rdoDisplayFlg1.checked = True Then

				For lRow = lgPrevMaxRows + 1 To .MaxRows 
					.Row = lRow									
					.Col = C_Check
					IF .Value = 1 Then
						.Col = 0
						.Text = ""
						ggoSpread.SSSetProtected C_Check , lRow,  lRow
					End If	
				Next 
			End If
		
		    If strCanFlag = "Y" Then
				Call SetToolbar("11101011000111")
			Else
				Call SetToolbar("11101001000111")
			End If
		   	frm1.btnRun.Disabled = False
    	
			If GetSetupMod(Parent.gSetupMod, "A") = "Y" Then  
		 		txtCostTitle.style.display = ""
			End If 

		Else
			ggoSpread.SpreadLock -1, -1

			If strCanFlag = "Y" Then
				Call SetToolbar("11101011000111")
			Else
				Call SetToolbar("11101001000111")
			End If
			txtCostTitle.style.display = "none"
		  	frm1.btnRun.Disabled = True
		End If
		
		Call SetFocusToDocument("M")        
        .Focus

		.Redraw = True
		
	End With	
	
	lgIntFlgMode = Parent.OPMD_UMODE         	         
    Call ggoOper.LockField(Document, "Q")				

End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
    
    Dim lRow        
    Dim strVal
    Dim ColSep, RowSep     

	Dim strCUTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount

    Call LayerShowHide(1)
        
    Err.Clear		
	DbSave = False                                                   
    
	frm1.txtMode.value = Parent.UID_M0002

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0
	
	With frm1.vspdData
		ColSep = Parent.gColSep
		RowSep = Parent.gRowSep
	
		For lRow = 1 To .MaxRows
    
		    .Row = lRow
		    .Col = 0
		    
			Select Case .Text
				Case ggoSpread.UpdateFlag			
					
				   .Col = C_Check
					
					If .Text <> "0" Then		
						.Col = C_SeqNo	
						
						strVal = "U" & ColSep 
						strVal = strVal & lRow & ColSep
						strVal = strVal & Trim(.Text) & RowSep
			
						If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then
						                            
							Set objTEXTAREA = document.createElement("TEXTAREA")
							objTEXTAREA.name = "txtCUSpread"
							objTEXTAREA.value = Join(iTmpCUBuffer,"")
							divTextArea.appendChild(objTEXTAREA)     
							 
							iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT   
							ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
							iTmpCUBufferCount = -1
							strCUTotalvalLen  = 0
									
						End If
						       
						iTmpCUBufferCount = iTmpCUBufferCount + 1
						      
						If iTmpCUBufferCount > iTmpCUBufferMaxCount Then    
							iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
							ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
						End If   
									
						iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
						strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
											
					End if
				
				Case ggoSpread.DeleteFlag	
					.Col = C_Check
					
					If .Text <> "0" Then
						.Col = C_SeqNo
						
						strVal = "D" & ColSep
						strVal = strVal & lRow & ColSep
						strVal = strVal & Trim(.Text) & RowSep
						
						If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then
						                            
							Set objTEXTAREA = document.createElement("TEXTAREA")
							objTEXTAREA.name = "txtCUSpread"
							objTEXTAREA.value = Join(iTmpCUBuffer,"")
							divTextArea.appendChild(objTEXTAREA)     
							 
							iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT   
							ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
							iTmpCUBufferCount = -1
							strCUTotalvalLen  = 0
									
						End If
						       
						iTmpCUBufferCount = iTmpCUBufferCount + 1
						      
						If iTmpCUBufferCount > iTmpCUBufferMaxCount Then    
							iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
							ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
						End If   
									
						iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
						strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
					End If
			End Select
		Next
	End With

	If iTmpCUBufferCount > -1 Then 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	Else
		Call DisplayMsgBox("169909","X", "X", "X")   
	   	frm1.btnRun.Disabled = False
		Call LayerShowHide(0)
		Exit function
	End If  
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	
    DbSave = True                                                

End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()						
    Dim ItemDocNo
    
    ItemDocNo = frm1.txthItemDocumentNo.value

    If  Trim(ItemDocNo) <> "" Then
        Call DisplayMsgBox("169910","X",ItemDocNo, "X")   
	End if
	Call InitVariables	
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

'------------------------------------------  ()  --------------------------------------------------
'	Name : Checkall()
'--------------------------------------------------------------------------------------------------------- 
Function Checkall()
	
	Dim IRowCount 
	Dim IClnCount
	 
	ggoSpread.Source = frm1.vspdData
	 
	With frm1.vspdData    
	  
		IF lgCheckall = 0 Then 
			
			For IRowCount = 1 to .MaxRows
				.Row = IRowCount 
				.Col = C_Check	 
				.text = 1     
			Next    

			lgCheckall = 1

		Else
			   
			For IRowCount = 1 to .MaxRows
				.Row = IRowCount 
				.Col = C_Check	 
				.text = 0     
			Next    
			   
			lgCheckall = 0
		  
		End If
	End With
End Function

'========================================================================================
' Function Name : Plant_PhyInvNo_Check
'========================================================================================
Function Plant_PhyInvNo_Check()
	
	Plant_PhyInvNo_Check = False

    If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus 
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	
	If 	CommonQueryRs(" A.POS_BLK_INDCTR, A.DOC_STS_INDCTR, CONVERT(CHAR(10), A.REAL_INSP_DT, 21), A.SL_CD, B.PLANT_CD "," I_PHYSICAL_INVENTORY_HEADER	A, I_PHYSICAL_INVENTORY_DETAIL B ", _
	    " A.PHY_INV_NO = " & FilterVar(frm1.txtCondPhyInvNo.Value, "''", "S") & " AND  A.PHY_INV_NO = B.PHY_INV_NO ", _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

		Call DisplayMsgBox("160301","X","X","X")
		frm1.txtCondPhyInvNo.focus
		Exit function
	
	End If
	lgF0 = Split(lgF0,Chr(11))
	lgF1 = Split(lgF1,Chr(11))
	lgF2 = Split(lgF2,Chr(11))
	lgF3 = Split(lgF3,Chr(11))
	lgF4 = Split(lgF4,Chr(11))

	If UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(lgF4(0))) Then
		Call DisplayMsgBox("169943","X","X","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit function
	End If

	frm1.txthPhyInvPosBlk.Value = UCase(Trim(lgF0(0)))

	If UCase(Trim(lgF1(0))) <> "PD" Then
		Call DisplayMsgBox("169908","X","X","X")
		frm1.txtCondPhyInvNo.focus
		Set gActiveElement = document.activeElement
		Exit function
	ElseIf Trim(lgF0(0)) = "Y" and frm1.rdoDisplayFlg2.checked = True Then
		Call DisplayMsgBox("169907","X","X","X")
		frm1.txtCondPhyInvNo.focus
		Set gActiveElement = document.activeElement
		Exit function
	End If 

	frm1.txtInspDt.value = UniConvDateAToB(lgF2(0),Parent.gServerDateFormat,Parent.gDateFormat)	
	frm1.txtSLCd.value  = UCase(Trim(lgF3(0)))

	If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125700","X","X","X")
		frm1.txtSLNm.Value = ""
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtSLNm.Value = lgF0(0)

    If Trim(frm1.txtItemCd.Value) <> "" Then
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)
		Else
			frm1.txtItemNm.Value = ""
		End If
	End If

    Plant_PhyInvNo_Check = True
		    
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	Dim pvRow, pvTempRow
	
	
	pvTempRow= 0
	
	With frm1.vspdData
		For pvRow = 1 to frm1.vspdData.MaxRows
			.Row = pvRow
			.Col = 0
			Select Case .Text
				Case ggoSpread.UpdateFlag	
					.Col = C_Check
							
					If .Text <> "0" Then		
						pvTempRow = pvTempRow + 1
						If CLng(pvTempRow) = CLng(lRow) Then	
							frm1.vspdData.focus
							frm1.vspdData.Row = pvRow
							frm1.vspdData.Col = lCol
							frm1.vspdData.Action = 0
							frm1.vspdData.SelStart = 0
							frm1.vspdData.SelLength = len(frm1.vspdData.Text)
							Exit For
						End If	
					End If	
			End Select
		Next
	End With
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사조정(Manual)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMvmtListRef()">실사선별후 수불참조</A></TD>
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
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>실사번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCondPhyInvNo" SIZE=20 MAXLENGTH=16 tag="12XXXU" ALT="실사번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPhyInvNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPhyInvNo()"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 MAXLENGTH=30 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>조회선택</TD>												
								<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDisplayFlg" ID="rdoDisplayFlg1"  tag="11" Value="Y" CHECKED><LABEL FOR="rdoDisplayFlg1">전체</LABEL>
							    	<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDisplayFlg" ID="rdoDisplayFlg2"  tag="11" Value="N"><LABEL FOR="rdoDisplayFlg2">미조정</LABEL></TD>												
							</TR>
						</TABLE>
					</FIELDSET>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% VALIGN=TOP>
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="24" ALT="창고">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=27 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>실사일</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspDt" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: center" tag="24X1" ALT="실사일"></TD>
						</TR>
						<TR ID="txtCostTitle" STYLE="DISPLAY: none">
								<TD CLASS="TD5" NOWRAP>Cost Center</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCD()">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=23 MAXLENGTH=23 tag="24">
								</TD>
								<TD CLASS="TD5"></TD>
								<TD CLASS="TD6"></TD>
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
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
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
		
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:Checkall()">전체 선택/취소</BUTTON></TD>		
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthCondPhyInvNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPhyInvPosBlk" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthPhyInvDocSts" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthItemDocumentNo" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthDisplayFlg" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

