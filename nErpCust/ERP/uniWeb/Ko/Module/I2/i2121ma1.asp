<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 실사선별조정 
'*  3. Program ID           : i2121ma1 
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2000/04/07
'*  8. Modified date(Last)  : 2003/06/02
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Lee Seung Wook
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit											

Const BIZ_LOOKUP_PGM_ID = "i2121mb1.asp"								
Const BIZ_PGM_ID        = "i2121mb2.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim  C_ItemCd            	
Dim  C_ItemPopup
Dim  C_ItemNm
Dim  C_ItemSpec
Dim  C_BasicUnit
Dim  C_TrackingNo
Dim  C_LotNo
Dim  C_LotSubNo
Dim  C_ABCFlag
Dim  C_STS_Indctr
Dim  C_SeqNo

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE         	      
	lgBlnFlgChgValue = False     	               	
	lgIntGrpCount = 0                           	
	
	lgStrPrevKey1 = ""                           	
	lgStrPrevKey2 = ""
	lgLngCurRows = 0                            	
    Call SetToolbar("11000000000011")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()	
	
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
	Set gActiveElement = document.activeElement 
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
 		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		
		.MaxCols = C_SeqNo + 1						
		.Col = .MaxCols								
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")

		ggoSpread.SSSetEdit C_ItemCd, "품목",				20,,,, 2
		ggoSpread.SSSetButton C_ItemPopup
		ggoSpread.SSSetEdit C_ItemNm, "품목명",				25
		ggoSpread.SSSetEdit C_ItemSpec, "규격",				20
		ggoSpread.SSSetEdit C_BasicUnit, "단위",			8,2				
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No",	20		
		ggoSpread.SSSetEdit C_LotNo, "LOT NO",				12
		ggoSpread.SSSetFloat C_LotSubNo, "Lot No.순번",		8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit C_ABCFlag, "ABC",				8,2
		ggoSpread.SSSetEdit C_STS_Indctr, "실사등록",		8,2
		ggoSpread.SSSetEdit C_SeqNo, "SEQ",					5,2
		
 		Call ggoSpread.MakePairsColumn(C_ItemCd, C_ItemPopup)
 		Call ggoSpread.SSSetColHidden(C_SeqNo, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		
		.ReDraw = true
		
		ggoSpread.SSSetSplit2(3)
       End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.SpreadLock -1, pvStartRow, -1, pvEndRow
	ggoSpread.SpreadUnLock	C_ItemCd, pvStartRow, C_ItemCd, pvEndRow
	ggoSpread.SpreadUnLock	C_ItemPopup, pvStartRow, C_ItemPopup, pvEndRow
	ggoSpread.SSSetRequired	C_ItemCd,		pvStartRow, pvEndRow
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_ItemCd		= 1									
	C_ItemPopup		= 2	
	C_ItemNm		= 3
	C_ItemSpec		= 4
	C_BasicUnit		= 5
	C_TrackingNo	= 6
	C_LotNo			= 7
	C_LotSubNo		= 8
	C_ABCFlag		= 9
	C_STS_Indctr	= 10
	C_SeqNo			= 11
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	
	Dim iCurColumnPos
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_ItemCd		= iCurColumnPos(1)									
		C_ItemPopup		= iCurColumnPos(2)	
		C_ItemNm		= iCurColumnPos(3)
		C_ItemSpec		= iCurColumnPos(4)
		C_BasicUnit		= iCurColumnPos(5)
		C_TrackingNo	= iCurColumnPos(6)
		C_LotNo			= iCurColumnPos(7)
		C_LotSubNo		= iCurColumnPos(8)
		C_ABCFlag		= iCurColumnPos(9)
		C_STS_Indctr	= iCurColumnPos(10)
		C_SeqNo			= iCurColumnPos(11)

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
	arrParam2 = "DC"
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
		frm1.txtInspDt.value		= arrRet(1)	
		frm1.txtSLCd.Value 			= arrRet(2)
		frm1.txtSLNm.Value			= arrRet(3)
		frm1.txtCondPhyInvNo.focus	
	End If	
	
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	
	If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) OR _
  		UCase(Trim(frm1.txtCondPhyInvNo.value)) <> UCase(Trim(frm1.txthCondPhyInvNo.value)) Then
		Call DisplayMsgBox("900002","X","X","X")
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("i2121pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "i2121pa1", "X")
		IsOpenPop = False
		Exit Function
    End If

	IsOpenPop = True
	Param1 = Trim(frm1.txtPlantCd.value)	
	Param2 = strCode
	Param3 = Trim(frm1.txtSLCd.value)
	Param4 = Trim(frm1.txtPlantNm.value)
	Param5 = Trim(frm1.txtSLNm.value)	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, Param1, Param2, Param3,Param4, Param5), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		With frm1.vspdData
			Call .SetText(C_ItemCd,	   .ActiveRow, arrRet(0))
			Call .SetText(C_ItemNm,	   .ActiveRow, arrRet(1))
			Call .SetText(C_ItemSpec,  .ActiveRow, arrRet(2))
			Call .SetText(C_BasicUnit, .ActiveRow, arrRet(3))
			Call .SetText(C_TrackingNo,.ActiveRow, arrRet(4))
			Call .SetText(C_LotNo,	   .ActiveRow, arrRet(5))
			Call .SetText(C_LotSubNo,  .ActiveRow, arrRet(6))
			Call .SetText(C_ABCFlag,   .ActiveRow, arrRet(7))
			Call .SetText(C_STS_Indctr,.ActiveRow, "N")
			Call vspdData_Change(C_ItemCd, .ActiveRow)		 
			.Col = C_ItemCd
			.Action = 0
		End With
	End If	

End Function

'------------------------------------------  OpenItemOrigin()  --------------------------------------------------
Function OpenItemCd()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then 
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
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
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
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_ItemPopUp Then
			.Col = C_ItemCd
			.Row = Row
			Call OpenItem(.Text)
			
		Elseif Row > 0 And Col = C_TrackingNoPopUp Then
			.Col = C_TrackingNo
			.Row = Row			
			Call OpenTrackingNo(.text)
		
		Elseif Row > 0 And Col = C_LotNoPopUp Then
			.Col = C_LotNo
			.Row = Row			
			Call OpenLotNo()
		End If
	
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
 		Call SetPopupMenuItemInf("1101111111") 
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

    If Not chkField(Document, "1") Then Exit Function				
    
    ggoSpread.Source = frm1.vspdData  
    If  lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")			   
    	If IntRetCD = vbNo Then Exit Function
    End If
 
    Call ggoOper.ClearField(Document, "2")						
    
    Call InitVariables
    
    If Plant_PhyInvNo_Check(1) = False Then Exit Function

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
    ggoSpread.Source = frm1.vspdData	   
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then	Exit Function
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

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001","X", "X", "X")
		Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then Exit Function
    
    If Not chkField(Document,"1") Then Exit Function	
    
	If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) OR _
  		UCase(Trim(frm1.txtCondPhyInvNo.value)) <> UCase(Trim(frm1.txthCondPhyInvNo.value)) Then
		Call DisplayMsgBox("900002","X","X","X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

    If Plant_PhyInvNo_Check(2) = False Then Exit Function

    If DbSave = False Then Exit Function
    
    FncSave = True                                                  
    
End Function

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 
    If frm1.vspdData.maxrows < 1 then exit function
    frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    lgBlnFlgChgValue = True
    
	frm1.vspdData.ReDraw = True
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
' Function Name : FncInsertRow
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = False
	
    If frm1.vspdData.maxrows < 1 then
	   call DisplayMsgBox("900002","X", "X", "X")
	   Exit function
    End If

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If 
	
    With frm1.vspdData	
		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow .ActiveRow, imRow
		SetSpreadColor .ActiveRow, .ActiveRow + imRow -1
		.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
   
End Function

'========================================================================================
' Function Name : FncDeleteRow
'========================================================================================
Function FncDeleteRow() 
	
	Dim lDelRows 
	Dim lTempRows 

	If frm1.vspdData.maxrows < 1 then exit function
	
	ggoSpread.Source = frm1.vspdData	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
	
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
	If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
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

    Call LayerShowHide(1)  

    DbQuery = False
    
    Err.Clear                                                           
    Dim strVal
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_LOOKUP_PGM_ID &	"?txtMode="				& Parent.UID_M0001				& _				
											"&txtPlantCd="			& Trim(.txthPlantCd.value)		& _		
											"&txtItemCd="			& Trim(.txtItemCd.value)		& _
											"&lgStrPrevKey1="       & lgStrPrevKey1					& _
											"&lgStrPrevKey2="       & lgStrPrevKey2					& _
											"&txtCondPhyInvNo="		& Trim(.txthCondPhyInvNo.value)	& _
											"&txtMaxRows="			& .vspdData.MaxRows
		Else
			strVal = BIZ_LOOKUP_PGM_ID &	"?txtMode="				& Parent.UID_M0001				& _					
											"&txtPlantCd="			& Trim(.txtPlantCd.value)		& _			
											"&txtItemCd="			& Trim(.txtItemCd.value)		& _
											"&lgStrPrevKey1="       & lgStrPrevKey1					& _
											"&lgStrPrevKey2="       & lgStrPrevKey2					& _
											"&txtCondPhyInvNo="		& Trim(.txtCondPhyInvNo.value)	& _
											"&txtMaxRows="			& .vspdData.MaxRows
		End if    
		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()							
    lgIntFlgMode = Parent.OPMD_UMODE			
    Call SetToolbar("11101111001111")	
    Call ggoOper.LockField(Document, "Q")		
    frm1.vspdData.focus
End Function


'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
    Dim lRow        
	Dim strVal, strDel
    Dim ColSep, RowSep     

	Dim strCUTotalvalLen
	Dim strDTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount
	Dim iTmpDBuffer
	Dim iTmpDBufferCount
	Dim iTmpDBufferMaxCount

    Call LayerShowHide(1)
        
    Err.Clear	
	
    DbSave = False                                                      
    
	frm1.txtMode.value = Parent.UID_M0002

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	strCUTotalvalLen = 0
	strDTotalvalLen = 0
			
	With frm1.vspdData

		'-----------------------
		'Data manipulate area
		'-----------------------
		ColSep = Parent.gColSep
		RowSep = Parent.gRowSep
				
		For lRow = 1 To .MaxRows

		    .Row = lRow
		    .Col = 0
		    
		    Select Case .Text

		        Case ggoSpread.InsertFlag					
					
					strVal = "C" & ColSep
					strVal = strVal & lRow & ColSep
		            .Col = C_ItemCd
					strVal = strVal & Trim(.Text) & ColSep
		            .Col = C_TrackingNo
					strVal = strVal & Trim(.Text) & ColSep
		            .Col = C_LotNo		
					strVal = strVal & Trim(.Text) & ColSep
		            .Col = C_LotSubNo
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

		         Case ggoSpread.DeleteFlag					
					
					strDel = "D" & ColSep
					strDel = strDel & lRow & ColSep
				    .Col = C_SeqNo		
					strDel = strDel & Trim(.Text) & RowSep
  
					If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
					   Set objTEXTAREA   = document.createElement("TEXTAREA")
					   objTEXTAREA.name  = "txtDSpread"
					   objTEXTAREA.value = Join(iTmpDBuffer,"")
					   divTextArea.appendChild(objTEXTAREA)     
					 
					   iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
					   ReDim iTmpDBuffer(iTmpDBufferMaxCount)
					   iTmpDBufferCount = -1
					   strDTotalvalLen = 0 
					End If
       
					iTmpDBufferCount = iTmpDBufferCount + 1

					If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
					   iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
					   ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
					End If   
         
					iTmpDBuffer(iTmpDBufferCount) =  strDel         
					strDTotalvalLen = strDTotalvalLen + Len(strDel)

			End Select
		Next
	End With
	
	If iTmpCUBufferCount > -1 Then 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  

	If iTmpDBufferCount > -1 Then   
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
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
	lgBlnFlgChgValue = False
    frm1.vspdData.MaxRows = 0 		
    Call MainQuery()
End Function

'========================================================================================
' Function Name : Plant_PhyInvNo_Check
'========================================================================================
Function Plant_PhyInvNo_Check(ByVal ChkIndex)

	Plant_PhyInvNo_Check = False

	Select Case ChkIndex

		Case 1

			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus 
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtPlantNm.Value = lgF0(0)


			If 	CommonQueryRs(" A.DOC_STS_INDCTR, CONVERT(CHAR(10), A.REAL_INSP_DT, 21), A.POS_BLK_INDCTR, A.SL_CD, B.SL_NM "," I_PHYSICAL_INVENTORY_HEADER A, B_STORAGE_LOCATION B", _
							  " A.SL_CD = B.SL_CD AND A.PHY_INV_NO = " & FilterVar(frm1.txtCondPhyInvNo.Value, "''", "S"), _
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
			
			If UCase(Trim(lgF0(0))) = "PD" AND UCase(Trim(lgF2(0))) = "Y" Then
				Call DisplayMsgBox("169907","X","X","X")
				Exit function
			End If 

			frm1.txtInspDt.value = UniConvDateAToB(lgF1(0),Parent.gServerDateFormat,Parent.gDateFormat)
			frm1.txtSLCd.value = lgF3(0)
			frm1.txtSLNm.Value = lgF4(0)

		    If Trim(frm1.txtItemCd.Value) <> "" Then
				If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					
					lgF0 = Split(lgF0, Chr(11))
					frm1.txtItemNm.Value = lgF0(0)
				Else
					frm1.txtItemNm.Value = ""
				End If
			End If

		Case 2
			If 	CommonQueryRs(" DOC_STS_INDCTR, POS_BLK_INDCTR  "," I_PHYSICAL_INVENTORY_HEADER ", _
							  " PHY_INV_NO = " & FilterVar(frm1.txtCondPhyInvNo.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("160301","X","X","X")
				frm1.txtCondPhyInvNo.focus
				Exit function
			End If
			lgF0 = Split(lgF0,Chr(11))
			lgF1 = Split(lgF1,Chr(11))
			
				
			If UCase(Trim(lgF0(0))) = "PD" AND UCase(Trim(lgF1(0))) = "Y" Then
				Call DisplayMsgBox("169907","X","X","X")
				Exit function
			End If 
		
	End Select		  

	Plant_PhyInvNo_Check = True

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사선별조정</font></td>
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
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>실사번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCondPhyInvNo" SIZE=20 MAXLENGTH=16 tag="12XXXU" ALT="실사번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPhyInvNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPhyInvNo()"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 MAXLENGTH=30 tag="14"></TD>
							    <TD CLASS="TD5" NOWRAP></TD>
							    <TD CLASS="TD6" NOWRAP></TD>
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
							<script language =javascript src='./js/i2121ma1_I377213275_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthCondPhyInvNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPhyInvDocSts" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

