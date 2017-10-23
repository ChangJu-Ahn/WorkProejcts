<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 실사등록....
'*  3. Program ID           : I2131ma1.asp
'*  4. Program Name         : 실사Counitng등록 
'*  5. Program Desc         : 실사된 내용을 등록한다.
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2003/06/02
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : Kim Nam Hoon
'* 11. Comment              :
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit										

Const BIZ_PGM_ID = "i2131mb2.asp"								
Const BIZ_LOOKUP_ID = "i2131mb1.asp"							

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim CheckedNoall()
Dim lgCheckall
Dim lgPrevMaxRows


Dim C_ZeroInd
Dim C_ItemCd 								
Dim C_ItemNm 
Dim C_Spec   
Dim C_BaseUnit
Dim C_InvGoodQty
Dim C_InvBadQty 
Dim C_GoodQty   
Dim C_BadQty    
Dim C_TrackingNo
Dim C_LotNo     
Dim C_LotSubNo  
Dim C_SeqNo     
Dim C_ZeroChk   

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
    BtnRunTitle1.style.DISPLAY = ""
	BtnRunTitle2.style.DISPLAY = "none"
    Call SetToolbar("11000000000011")	
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()

    frm1.btnRun1.Disabled = True
    If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = ""
	End if	

	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
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
' Function Name : InitSpreadSheet
'========================================================================================
Sub InitSpreadSheet()

 	Call InitSpreadPosVariables()

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
 		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		
		.MaxCols = C_ZeroChk + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("A")
 		Call AppendNumberPlace("6", "3", "0")

		ggoSpread.SSSetCheck C_ZeroInd, "Zero IND. ", 5,,,1
		ggoSpread.SSSetEdit C_ItemCd, "품목", 18
		ggoSpread.SSSetEdit C_ItemNm, "품목명", 25
		ggoSpread.SSSetEdit C_Spec, "규격", 20
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20				
		ggoSpread.SSSetEdit C_LotNo, "LOT NO", 12

		ggoSpread.SSSetFloat C_LotSubNo, "Lot No.순번", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_InvGoodQty, "실사양품수", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat C_InvBadQty, "실사불량수", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"		
		ggoSpread.SSSetFloat C_GoodQty, "양품수", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_BadQty, "불량수", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit C_BaseUnit, "단위",8,2
		ggoSpread.SSSetEdit C_SeqNo,"",5
		ggoSpread.SSSetEdit C_ZeroChk, "", 5
		
 		Call ggoSpread.SSSetColHidden(C_SeqNo, .MaxCols, True)
 		
		ggoSpread.SpreadLock -1, -1
		ggoSpread.SpreadUnLock C_ZeroInd, -1, C_ZeroInd
		ggoSpread.SpreadUnLock C_InvGoodQty, -1, C_InvGoodQty
		ggoSpread.SpreadUnLock C_InvBadQty, -1, C_InvBadQty
		ggoSpread.SSSetRequired  C_InvGoodQty, -1, -1
		ggoSpread.SSSetRequired  C_InvBadQty, -1, -1
		
		ggoSpread.SSSetSplit2(3)
		.ReDraw = true

		End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.SSSetProtected  C_InvGoodQty, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_InvBadQty, pvStartRow, pvEndRow
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_ZeroInd    = 1
	C_ItemCd     = 2									
	C_ItemNm     = 3
	C_Spec       = 4
	C_BaseUnit   = 5
	C_InvGoodQty = 6
	C_InvBadQty  = 7
	C_GoodQty    = 8
	C_BadQty     = 9
	C_TrackingNo = 10
	C_LotNo      = 11
	C_LotSubNo   = 12
	C_SeqNo      = 13
	C_ZeroChk    = 14
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
 		C_ZeroInd		= iCurColumnPos(1)
		C_ItemCd		= iCurColumnPos(2)	
		C_ItemNm		= iCurColumnPos(3)
		C_Spec			= iCurColumnPos(4)
		C_BaseUnit		= iCurColumnPos(5)
		C_InvGoodQty	= iCurColumnPos(6)
		C_InvBadQty		= iCurColumnPos(7)
		C_GoodQty		= iCurColumnPos(8)
		C_BadQty		= iCurColumnPos(9)
		C_TrackingNo	= iCurColumnPos(10)
		C_LotNo			= iCurColumnPos(11)
		C_LotSubNo		= iCurColumnPos(12)
		C_SeqNo			= iCurColumnPos(13)
		C_ZeroChk		= iCurColumnPos(14)

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
	arrParam2 = "EC"
	arrParam3 = frm1.txtPlantCd.value
	arrParam4 = ""

  	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam1,arrParam2,arrParam3,arrParam4), _
 		 "dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCondPhyInvNo.focus
		Exit Function
	Else
		frm1.txtCondPhyInvNo.Value  = arrRet(0)
		frm1.txtInspDt.value	    = arrRet(1)	
		frm1.txtSLCd.Value 			= arrRet(2)
		frm1.txtSLNm.Value			= arrRet(3)
		frm1.txtCondPhyInvNo.focus
	End If	

End Function

'------------------------------------------  OpenItemOrigin()  --------------------------------------------------
Function OpenItemCd()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)

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

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
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

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	Call LoadInfTB19029               
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)	                                	
    Call InitSpreadSheet              
    Call InitVariables            
    Call SetDefaultVal    
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
	ggoSpread.UpdateRow Row

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
	
	If 	lgCheckall = 5 or lgCheckall = 6 Then Exit Sub	 

	With frm1.vspdData

		.ReDraw = False
 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_ZeroInd Then
			.Col = Col
			.Row = Row									
			
			IF .Text = "1" Then
				.col	=	C_InvGoodQty
				.text	=	"0"
				.col	=	C_InvBadQty
				.text	=	"0"
				ggoSpread.SSSetProtected  C_InvGoodQty, Row, Row
				ggoSpread.SSSetProtected  C_InvBadQty, Row, Row
				.Col = C_ZeroChk
				IF .Text = "0" Then			
					.Col = 0
					.Text = ggoSpread.UpdateFlag
					CheckedNoall(Row) = True
					lgBlnFlgChgValue = True
				else
					ggoSpread.Source = frm1.vspdData	
					ggoSpread.EditUndo                                          
					.Col = 0
					.Text = ""
					CheckedNoall(Row) = False
					lgBlnFlgChgValue = False
				End If
				
			Elseif .Text = "0" Then

				.Col = C_ZeroChk
				IF .Text = "0" Then
					ggoSpread.Source = frm1.vspdData	

					ggoSpread.SpreadUnLock C_InvGoodQty, Row, C_InvGoodQty, Row
					ggoSpread.SpreadUnLock C_InvBadQty, Row, C_InvBadQty, Row
	  				ggoSpread.SSSetRequired  C_InvGoodQty, Row, Row
					ggoSpread.SSSetRequired  C_InvBadQty, Row, Row
					ggoSpread.EditUndo                                          
					
					.Col = 0
					.Text = ""
					CheckedNoall(Row) = False
					lgBlnFlgChgValue = False
				Else

					ggoSpread.SpreadUnLock C_InvGoodQty, Row, C_InvGoodQty, Row
					ggoSpread.SpreadUnLock C_InvBadQty, Row, C_InvBadQty, Row
					ggoSpread.SSSetRequired  C_InvGoodQty, Row, Row
					ggoSpread.SSSetRequired  C_InvBadQty, Row, Row
					
					.Col = 0
					.Text = ggoSpread.UpdateFlag
					CheckedNoall(Row) = True
					lgBlnFlgChgValue = True
				End If
				
			End if 		
		End If
		.ReDraw = True
	
	End With

End Sub

'==========================================================================================
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
    If NewCol = C_ZeroInd or Col = C_ZeroInd Then
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
			.Col = C_ZeroInd
			.Row = lRow									
			IF .Text = "1" Then
				ggoSpread.SSSetProtected  C_InvGoodQty, lRow, lRow
				ggoSpread.SSSetProtected  C_InvBadQty, lRow, lRow
			Else
				ggoSpread.SSSetRequired  C_InvGoodQty, lRow, lRow
				ggoSpread.SSSetRequired  C_InvBadQty, lRow, lRow			
			End If	
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
    If Not chkField(Document, "1") Then	Exit Function					
       
    '-----------------------
    'Check previous data area
    '-----------------------
	If ggoSpread.SSCheckChange = True Then 
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")	
		If IntRetCD = vbNo Then Exit Function
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables	  										
  
    If Plant_PhyInvNo_Check = False Then 
        Call BtnDisabled(1)
		Exit Function
	End If
	
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
    If ggoSpread.SSCheckChange = True Then 
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If

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

	FncSave = False                                                      

	Err.Clear                                                            
	On Error Resume Next                                   
	
	'-----------------------
	'Check content area
	'-----------------------
   	If Not chkField(Document,"1") Then Exit Function		
	
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then 
	    IntRetCD = DisplayMsgBox("900001","X", "X", "X")                     
 		Exit Function
	End If
	
	If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) OR _
  		UCase(Trim(frm1.txtCondPhyInvNo.value)) <> UCase(Trim(frm1.txthCondPhyInvNo.value)) Then

		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	If DbSave() = False Then Exit Function                                      

	FncSave = True                                                         

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

	If ggoSpread.SSCheckChange = True Then 
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

    DbQuery = False

	Call LayerShowHide(1)    

    Err.Clear                                                        

	Dim strVal

    With frm1
 
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_LOOKUP_ID &	"?txtMode="			& Parent.UID_M0001				& _
										"&txtPlantCd="      & Trim(.txthPlantCd.value)		& _
										"&txtItemCd="       & Trim(.txtItemCd.value)		& _
										"&txtCondPhyInvNo=" & Trim(.txthCondPhyInvNo.value)	& _
										"&lgStrPrevKey1="   & lgStrPrevKey1					& _
										"&lgStrPrevKey2="   & lgStrPrevKey2					& _
										"&rdoDisplayFlg="   & Trim(.txthDisplayFlg.value)	& _
										"&txtMaxRows="      & .vspdData.MaxRows
		Else
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
		End If

		lgPrevMaxRows = .vspdData.MaxRows
		
		Call RunMyBizASP(MyBizASP, strVal)					

    End With

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()							
	Dim lRow
    '-----------------------
    'Reset variables area
    '-----------------------
	Call SetToolbar("11101001000111")	

    lgIntFlgMode = Parent.OPMD_UMODE						

    Call ggoOper.LockField(Document, "Q")					

	With frm1.vspdData

		ReDim Preserve CheckedNoall(.MaxRows)
		
		.Redraw = False
		
		If frm1.rdoDisplayFlg1.checked = True Then
	
			For lRow = lgPrevMaxRows + 1 To .MaxRows 
				
				CheckedNoall(lRow) = False
				
				.Col = C_ZeroInd
				.Row = lRow									
				IF .Text = "1" Then
					Call SetSpreadColor(lRow, lRow)
				End If	
		
			Next 
		
		End If
		
		Call SetFocusToDocument("M")        
        .Focus
		.Redraw = True
	End With
	Call SetActiveCell(frm1.vspdData,C_InvGoodQty,frm1.vspdData.ActiveRow,"M","X","X")	

	lgBlnFlgChgValue = False

End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim strVal
	Dim lBadQty
	Dim lGoodQty
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

	 '-----------------------
	 'Data manipulate area
	 '-----------------------
 		ColSep		= Parent.gColSep
		RowSep		= Parent.gRowSep
				
		For lRow = 1 To .MaxRows
		    .Row = lRow
		    .Col = 0
			
		    Select Case .Text

		        Case ggoSpread.UpdateFlag						
		
					strVal = "U" & ColSep
					strVal = strVal & lRow & ColSep
					.Col = C_SeqNo	'8
					strVal = strVal & Trim(.Text) & ColSep
 		            .Col = C_InvBadQty
 		           	lBadQty = UNIConvNum(Trim(.Text), 0) 
					strVal = strVal & lBadQty & ColSep
					.Col = C_InvGoodQty
 		           	lGoodQty = UNIConvNum(Trim(.Text), 0) 
					strVal = strVal & lGoodQty & ColSep

		            .Col = C_ZeroInd	'1
		            if Trim(.Text) = "1" Then
		               strVal = strVal & "Y" & RowSep
		            else
   						
   						If lBadQty = 0 and lGoodQty = 0 Then
							Call DisplayMsgBox("160415","X","X","X")
							Call SetActiveCell(frm1.vspdData,C_ItemCd,lRow,"M","X","X")
							Call LayerShowHide(0)
       						Call RestoreToolBar()
						    Call SetFocusToDocument("M")        
						    .Focus
							Exit Function
						End If

		               strVal = strVal & "N" & RowSep
		            end if
					
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
						
		    End Select

		Next
	End With
	
	If iTmpCUBufferCount > -1 Then 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)				

    DbSave = True                                      

End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()							
	Call InitVariables
    Call ggoOper.ClearField(Document, "2")
    Call FncQuery()
End Function

'========================================================================================
' Function Name : Checkall()
'========================================================================================
Function Checkall()

	Dim lngGoodQty    
	Dim lngBadQty     
	Dim lngInvGoodQty 
	Dim lngInvBadQty  
	Dim lngZeroInd	
	Dim lngZeroChk	
	Dim lngChk		
	Dim lRow

	With frm1.vspdData
		.ReDraw = False
        Call BtnDisabled(1)

		IF lgCheckall = 0 Then                              
			
			lgCheckall = 1                                  
 			
 			For lRow = 1 To .MaxRows
 				.Row = lRow
 					
 				.Col = 0
 				If .Text = "" Then                         

 					.Col = C_GoodQty
 					lngGoodQty = .Text					

 					.Col = C_InvGoodQty	
 					lngInvGoodQty = .Text				

 					.Col = C_BadQty
 					lngBadQty = .Text						

 					.Col = C_InvBadQty
 					lngInvBadQty = .Text				

					.Col = C_ZeroInd
 					If (UNICDbl(lngInvGoodQty) = 0 And UNICDbl(lngInvBadQty) = 0) And .Text = "0" Then 
 						.Col = C_InvGoodQty
 						.Text = lngGoodQty						
 						
 						.Col = C_InvBadQty
 						.Text = lngBadQty						

 						Call vspdData_Change(C_GoodQty,lRow)
 						Call vspdData_Change(C_BadQty,lRow)

						.Col = C_ZeroChk							
 						IF UNICDbl(lngGoodQty) = 0 and UNICDbl(lngBadQty) = 0 and .Text = "0" Then
 							lgCheckall = 6
 							.Col = C_ZeroInd
 							.Text = 1
							ggoSpread.SSSetProtected  C_InvGoodQty, lRow, lRow
							ggoSpread.SSSetProtected  C_InvBadQty, lRow, lRow

 						End If

 						lgBlnFlgChgValue = True					
 						lgCheckall = 4							
 						
					End If
				End If 	
			Next
 		    
 		    If lgCheckall = 1 Then   
 				lgCheckall = 0
 				Call DisplayMsgBox("900001","X", "X", "X")                               
			Else	
	 			lgCheckall = 2        
	 			BtnRunTitle1.style.DISPLAY = "none"
	 			BtnRunTitle2.style.DISPLAY = ""
	 		End If
						
 		Else						
			lgCheckall = 3          

  			For lRow = 1 To .MaxRows
 				
 				.Row = lRow

				.Col = 0
 				lngChk = .Text
 				
 				.Col = C_GoodQty
 				lngGoodQty = .Text

 				.Col = C_InvGoodQty
 				lngInvGoodQty = .Text

 		 		.Col = C_BadQty
 				lngBadQty = .Text

 				.Col = C_InvBadQty
 				lngInvBadQty = .Text
 				
 				.Col = C_ZeroInd
				lngZeroInd = .Text
				
				.Col = C_ZeroChk
				lngZeroChk = .Text
				
 				If lngChk = ggoSpread.UpdateFlag And (lngGoodQty = lngInvGoodQty And lngBadQty	= lngInvBadQty) And CheckedNoall(lRow) = False Then 'lngZeroChk = "0" Then
 					
 					.Col = 0
					.Text = ""				

					.col	=	C_ZeroInd
					If .text =	1 Then

						lgCheckall = 5			
						.text =	0				
						ggoSpread.SpreadUnLock C_InvGoodQty, lRow, C_InvGoodQty, lRow
						ggoSpread.SpreadUnLock C_InvBadQty, lRow, C_InvBadQty,  lRow						
		  				ggoSpread.SSSetRequired  C_InvGoodQty, lRow, lRow
						ggoSpread.SSSetRequired  C_InvBadQty, lRow, lRow
					Else
						.col	=	C_InvGoodQty  
						.text	=	"0"
						.col	=	C_InvBadQty
						.text	=	"0"

					End If
					
					lgBlnFlgChgValue = False 			

				End If
			                      
 			Next
 			
 			lgCheckall = 0
	 		BtnRunTitle1.style.DISPLAY = ""
	 		BtnRunTitle2.style.DISPLAY = "none"
		
		End if

		Call BtnDisabled(0)

		.ReDraw = True

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
	
		'-----------------------
		'Check PhyInvNo CODE	 
		'-----------------------
		If 	CommonQueryRs(" A.POS_BLK_INDCTR, CONVERT(CHAR(10), A.REAL_INSP_DT, 21), A.SL_CD, B.PLANT_CD "," I_PHYSICAL_INVENTORY_HEADER	A, I_PHYSICAL_INVENTORY_DETAIL B ", _
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

		If UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(lgF3(0))) Then
			Call DisplayMsgBox("169943","X","X","X")
			frm1.txtPlantCd.focus
			Exit function
		End If

		If UCase(Trim(lgF0(0))) = "Y" Then
			Call DisplayMsgBox("169907","X","X","X")
			frm1.txtCondPhyInvNo.focus
			Exit function
		End If 

		frm1.txtInspDt.value = UniConvDateAToB(lgF1(0),Parent.gServerDateFormat,Parent.gDateFormat)	
		frm1.txtSLCd.value  = lgF2(0)
			
		If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("125700","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
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
 		Set gActiveElement = document.activeElement
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사등록</font></td>
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
							<TD CLASS="TD5">공장</TD>
							<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
							<TD CLASS="TD5">실사번호</TD>
							<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCondPhyInvNo" SIZE=20 MAXLENGTH=16 tag="12XXXU" ALT="실사번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPhyInvNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPhyInvNo()"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 MAXLENGTH=30 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>조회선택</TD>												
								<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDisplayFlg" ID="rdoDisplayFlg1"  tag="11" Value="Y" CHECKED><LABEL FOR="rdoDisplayFlg1">전체</LABEL>
							    	<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDisplayFlg" ID="rdoDisplayFlg2"  tag="11" Value="N"><LABEL FOR="rdoDisplayFlg2">미등록</LABEL></TD>												
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
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<script language =javascript src='./js/i2131ma1_I518043088_vspdData.js'></script>
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
					<TR  ID="BtnRunTitle1" STYLE="DISPLAY: none">
						<TD WIDTH=10>&nbsp;</TD>
						<TD><BUTTON NAME="btnRun1" CLASS="CLSMBTN" ONCLICK="vbscript:Checkall()" Flag=1>전산재고수량복사</BUTTON></TD>		
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
					<TR  ID="BtnRunTitle2" STYLE="DISPLAY: none">
						<TD WIDTH=10>&nbsp;</TD>
						<TD><BUTTON NAME="btnRun2" CLASS="CLSMBTN" ONCLICK="vbscript:Checkall()" Flag=1>복사 취소</BUTTON></TD>		
						<TD WIDTH=10>&nbsp;</TD>
					</TR>					
				</TABLE>
			</TD>
		</TR>
		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthCondPhyInvNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPhyInvDocSts" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthPostFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthDisplayFlg" tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
