<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4214ma1.asp																*
'*  4. Program Name         : Container 배정 
'*  5. Program Desc         : 																*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2005/01/24																*
'*  8. Modified date(Last)  : 																*
'*  9. Modifier (First)     : HJO																*
'* 10. Modifier (Last)      : 																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/04/17 : Coding Start												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                       

Dim C_CC_SEQ	
Dim C_PLANT_CD
Dim C_PLANT_POP
Dim C_PLANT_NM
Dim C_ITEM_CD
Dim C_ITEM_POP
Dim C_ITEM_NM
Dim C_DESC
Dim C_UNIT
Dim C_QTY
Dim C_PRICE
Dim C_DOC_AMT
Dim C_REMARK

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent					
Dim IsOpenPop

Const BIZ_PGM_QRY_ID 	= "s4214mb1_KO441.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "s4214mb1_KO441.asp"			'☆: 비지니스 로직 ASP명 
Const EXCC_HEADER_ENTRY_ID = "s4211ma1_KO441"			'☆: 이동할 ASP명: 통관등록 
Const EXCC_DETAIL_ENTRY_ID = "s4212ma1_KO441"			'☆: 이동할 ASP명 : 통관내역등록 
Const EXCC_LAN_ENTRY_ID = "s4213ma1"			'☆: 이동할 ASP명 : 통관란등록 
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_CC_SEQ		= 1
	C_PLANT_CD	= 2
	C_PLANT_POP	= 3
	C_PLANT_NM	= 4
	C_ITEM_CD		= 5
	C_ITEM_POP	= 6
	C_ITEM_NM		= 7
	C_DESC			= 8
	C_UNIT			= 9
	C_QTY				= 10
	C_PRICE			= 11
	C_DOC_AMT		= 12
	C_REMARK		= 13

End Sub
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE					
	lgBlnFlgChgValue = False					
	lgIntGrpCount = 0							
	lgStrPrevKey = ""							
	lgLngCurRows = 0 							
		
	gblnWinEvent = False
End Function
'========================================================================================================
Sub SetDefaultVal()
	lgBlnFlgChgValue = False
End Sub
'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %> 
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	    
    Call initSpreadPosVariables()
	    
    With frm1

		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20081127",,parent.gAllowDragDropSpread    
			
		.vspdData.ReDraw = False
			
		.vspdData.MaxCols = C_REMARK+1
		.vspdData.MaxRows = 0
			
		Call GetSpreadColumnPos("A")	
			
		ggoSpread.SSSetEdit		C_CC_SEQ, 			"순번", 	15, 		0
	  ggoSpread.SSSetEdit		C_PLANT_CD,			"공장",		8,		,					,	  4,	  2
	  ggoSpread.SSSetButton	C_PLANT_POP		
	  ggoSpread.SSSetEdit		C_PLANT_NM,			"공장명",	15, 		0
	  ggoSpread.SSSetEdit		C_ITEM_CD,			"품목",		15,		,					,	  18,	  2
	  ggoSpread.SSSetButton	C_ITEM_POP		
	  ggoSpread.SSSetEdit		C_ITEM_NM,			"품목명",	15, 		0
	  ggoSpread.SSSetEdit		C_DESC,				  "규격",	 	15, 		0
	  ggoSpread.SSSetEdit		C_UNIT,				  "단위",	 	10,		,					,	  3,	  2
		ggoSpread.SSSetFloat	C_Qty,					"수량",		15,		Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,""
	  ggoSpread.SSSetFloat	C_PRICE,				"단가",		15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_DOC_AMT,			"금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetEdit		C_REMARK,			  "비고",		60,		,					,	  120
			
		call ggoSpread.SSSetColHidden(C_CC_SEQ,C_CC_SEQ,True)
		call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols,True)
		call SetSpreadLock()
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
			
		ggoSpread.SpreadUnLock C_PLANT_CD, -1, -1
		ggoSpread.SSSetRequired C_PLANT_CD, -1, -1
		ggoSpread.SpreadLock C_PLANT_NM, -1, -1
		ggoSpread.SpreadUnLock C_ITEM_CD, -1, -1
		ggoSpread.SSSetRequired C_ITEM_CD, -1, -1
		ggoSpread.SpreadLock C_ITEM_NM, -1, -1
		ggoSpread.SpreadLock C_DESC, -1, -1
		ggoSpread.SpreadLock C_UNIT, -1, -1
		ggoSpread.SpreadUnLock C_QTY, -1, -1
		ggoSpread.SSSetRequired C_QTY, -1, -1
		ggoSpread.SpreadUnLock C_PRICE, -1, -1
		ggoSpread.SSSetRequired C_PRICE, -1, -1
		ggoSpread.SpreadUnLock C_DOC_AMT, -1, -1
		ggoSpread.SSSetRequired C_DOC_AMT, -1, -1

		
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadColor(ByVal fRow, byVal tRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		.redraw = false
		ggoSpread.SSSetRequired C_PLANT_CD, fRow, tRow
		ggoSpread.SpreadLock 		C_PLANT_NM, fRow, C_PLANT_NM, tRow
		ggoSpread.SSSetRequired C_ITEM_CD, 	fRow, tRow
		ggoSpread.SpreadLock 		C_ITEM_NM, 	fRow, C_ITEM_NM, tRow
		ggoSpread.SpreadLock 		C_DESC, 		fRow, C_DESC, tRow
		ggoSpread.SpreadLock 		C_UNIT, 		fRow, C_UNIT, tRow
		ggoSpread.SSSetRequired C_QTY, 			fRow, tRow
		ggoSpread.SSSetRequired C_PRICE, 		fRow, tRow
		ggoSpread.SSSetRequired C_DOC_AMT, 	fRow, tRow		
		
		.redraw = true
	End With
End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
   
    Select Case UCase(pvSpdNo)
       Case "A"
            
          ggoSpread.Source = frm1.vspdData
          Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

					C_CC_SEQ		= iCurColumnPos(1)
					C_PLANT_CD	= iCurColumnPos(2)
					C_PLANT_POP	= iCurColumnPos(3)
					C_PLANT_NM	= iCurColumnPos(4)
					C_ITEM_CD		= iCurColumnPos(5)
					C_ITEM_POP	= iCurColumnPos(6)
					C_ITEM_NM		= iCurColumnPos(7)
					C_DESC			= iCurColumnPos(8)
					C_UNIT			= iCurColumnPos(9)
					C_QTY				= iCurColumnPos(10)
					C_PRICE			= iCurColumnPos(11)
					C_DOC_AMT		= iCurColumnPos(12)
					C_REMARK		= iCurColumnPos(13)

    End Select    
End Sub

'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtCCNo.value

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
				
		If strTemp = "" then Exit Function
				
		frm1.txtCCNo.value =  strTemp
			
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
			
		Call MainQuery()
						
		WriteCookie CookieSplit , ""
			
	End If

End Function
'===========================================================================
Function JumpChgCheck(ByVal IWhere)

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Select Case IWhere
	Case 0		'통관등록 
		Call CookiePage(1)
		Call PgmJump(EXCC_HEADER_ENTRY_ID)
	Case 1		'통관내역등록 
		Call CookiePage(1)				
		Call PgmJump(EXCC_DETAIL_ENTRY_ID)
	Case 2		'통관란등록 
		Call CookiePage(1)
		Call PgmJump(EXCC_LAN_ENTRY_ID)
	End Select		
End Function
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	'Call HideNonRelGrid()
	
End Sub
'====================================================================================================

'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'총통관금액 
		'ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_PRICE,-1, .txtCur.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_DOC_AMT,-1, .txtCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub
'========================================================================================================
Sub Form_Load()

	Call GetGlobalVar												
	Call LoadInfTB19029												
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")							

	Call InitSpreadSheet											

	Call SetDefaultVal

	Call CookiePage(0)	

	Call InitVariables

		
	Call SetToolbar("1110000000011111")								

	If UCase(Trim(frm1.txtCCNo.value)) <> "" Then
		Call MainQuery
	End If

	frm1.txtCCNo.focus
	Set gActiveElement = document.activeElement 

End Sub

'========================================================================================================
Sub btnCCNoOnClick()
	Call OpenExCCNoPop()
End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim iQty, iPrice
	
	ggoSpread.Source = frm1.vspdData
	with frm1.vspdData
	
	Select Case Col
		Case C_PLANT_CD, C_ITEM_CD
			call SetDefaultUnit(Row)
		Case C_QTY, C_PRICE
			iQty = UniConvNum(GetSpreadText(frm1.vspdData,C_QTY,Row,"X","X"),0)
			iPrice = UniConvNum(GetSpreadText(frm1.vspdData,C_PRICE,Row,"X","X"),0)
			call frm1.vspdData.SetText(C_DOC_AMT,	Row, iQty*iPrice)
		Case Else
	End Select
	
	End With
					
	ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True
End sub

sub SetDefaultUnit(byVal pRow)
	Dim iStrWhere
	
		if CommonQueryRs(" a.ISSUED_UNIT,b.ITEM_NM ", " b_item_by_plant a inner join b_item b on (a.item_cd=b.item_cd) ", " a.PLANT_CD = " & FilterVar(GetSpreadText(frm1.vspdData,C_PLANT_CD,pRow,"X","X"), "''", "S") & " AND a.ITEM_CD = " & FilterVar(GetSpreadText(frm1.vspdData,C_ITEM_CD,pRow,"X","X"), "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			call frm1.vspdData.SetText(C_UNIT,	pRow, replace(lgF0,chr(11),""))
			call frm1.vspdData.SetText(C_ITEM_NM,	pRow, replace(lgF1,chr(11),""))
		End if

		iStrWhere =  iStrWhere & " BP_CD					= " & FilterVar(frm1.txtApplicant.value,"''","S")
		iStrWhere =  iStrWhere & " and ITEM_CD		= " & FilterVar(GetSpreadText(frm1.vspdData,C_ITEM_CD,pRow,"X","X"), "''", "S")
		iStrWhere =  iStrWhere & " and DEAL_TYPE	= 'INVOICE'"
		iStrWhere =  iStrWhere & " and SALES_UNIT	= " & FilterVar(GetSpreadText(frm1.vspdData,C_UNIT,pRow,"X","X"), "''", "S")
		iStrWhere =  iStrWhere & " and CURRENCY		= " & FilterVar(frm1.txtCur.value,"''","S")
		iStrWhere =  iStrWhere & " and VALID_FROM_DT = "
		iStrWhere =  iStrWhere & " 	( "
		iStrWhere =  iStrWhere & " 	select max(VALID_FROM_DT) "
		iStrWhere =  iStrWhere & " 	from s_bp_item_price "
		iStrWhere =  iStrWhere & " 	where BP_CD		=  " & FilterVar(frm1.txtApplicant.value,"''","S")
		iStrWhere =  iStrWhere & " 	and ITEM_CD		=  " & FilterVar(GetSpreadText(frm1.vspdData,C_ITEM_CD,pRow,"X","X"), "''", "S")
		iStrWhere =  iStrWhere & " 	and DEAL_TYPE	= 'INVOICE' "
		iStrWhere =  iStrWhere & " 	and SALES_UNIT=  " & FilterVar(GetSpreadText(frm1.vspdData,C_UNIT,pRow,"X","X"), "''", "S")
		iStrWhere =  iStrWhere & " 	and CURRENCY	=  " & FilterVar(frm1.txtCur.value,"''","S")
		iStrWhere =  iStrWhere & " 	) "

		if CommonQueryRs(" ITEM_PRICE ", " s_bp_item_price ", iStrWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			call frm1.vspdData.SetText(C_PRICE,	pRow, replace(lgF0,chr(11),""))
		End if
	
End sub
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	Exit Sub

	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							
				DbQuery
			End If
		End If
	End With
End Sub
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           Call DbQuery
        End If
    End if
End Sub
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("0111111111")
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If
    	frm1.vspdData.Row = Row

End Sub
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
       
    If Row <= 0 Then
    End If
	
End sub
'================ vspdData_ButtonClicked() ================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 Then
        .Col = Col
        .Row = Row
        
				Select Case Col 			
				Case C_PLANT_POP
					Call OpenPlant()
				Case C_ITEM_POP
					call OpenItem(GetSpreadText(frm1.vspdData,C_ITEM_CD,Row,"X","X"))
				End Select
        
    End If
    
    End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenExCCNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD	
		
		
	If gblnWinEvent = True Or UCase(frm1.txtCCNo.className) = "PROTECTED" Then Exit Function		
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("S4211PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4211PA1", "X")			
		gblnWinEvent = False
		Exit Function
	End If						
			
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetExCCNo(strRet)
	End If	
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetExCCNo(strRet)
	frm1.txtCCNo.value = strRet(0)
	frm1.txtCCNo.focus
End Function

'===========================================================================
Function OpenItem(ByVal strCode)

	Dim arrParam(1)
	Dim strRet
	Dim iCalledAspName
	
	arrParam(0) = strCode
	frm1.vspdData.Col = C_PLANT_CD
	arrParam(1) = frm1.vspdData.text 

	If IsOpenPop = True Then Exit Function
	  
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s3112pa2")
	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112pa2", "x")
		IsOpenPop = False
		exit Function
	End if
	 
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
	 "dialogWidth=820px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_ITEM_CD
		frm1.vspdData.Text = strRet(0)
		frm1.vspdData.Col = C_ITEM_NM
		frm1.vspdData.Text = strRet(1)
		frm1.vspdData.Col = C_PLANT_CD
		frm1.vspdData.Text = strRet(2)
		call SetDefaultUnit(frm1.vspdData.ActiveRow)
	End If 

End Function 

'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False											

	Err.Clear													

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")	         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call InitVariables											

	If Not chkField(Document, "1") Then							
		Exit Function
	End If

	Call DbQuery()												

	FncQuery = True												
End Function
'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False												
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")								
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call ggoOper.LockField(Document, "N")						
	Call SetDefaultVal
	call InitVariables											

	FncNew = True												

End Function
'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False										
		
	If lgIntFlgMode <> Parent.OPMD_UMODE Then						
		Call DisplayMsgBox("900002", "x", "x")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "x", "x")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call DbDelete											

	FncDelete = True										
End Function
'========================================================================================================
Function FncSave()
	Dim IntRetCD
		
	FncSave = False											
		
	Err.Clear												
		
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
	    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")   
	    Exit Function
	End If

	ggoSpread.Source = frm1.vspdData

	If Not chkField(Document, "2") Then	
		Exit Function
	End If

	If Not ggoSpread.SSDefaultCheck Then		
		Exit Function
	End If

	Call DbSave	
		
	FncSave = True								
End Function
'========================================================================================================
Function FncCopy()
	frm1.vspdData.ReDraw = False

	ggoSpread.Source = frm1.vspdData	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
End Function
'========================================================================================================
Function FncCancel() 
	ggoSpread.Source = frm1.vspdData
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	ggoSpread.EditUndo							
End Function

'================== FncInsertRow() ===========================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
    Dim IntRetCD
    Dim imRow
    Dim inti
    inti=1
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If
    
	With frm1
	        
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		
		For inti= .vspdData.ActiveRow  to .vspdData.ActiveRow +imRow-1
			.Row=inti
			
			'공장 기본값 추가 
			Call .vspdData.SetText(C_PLANT_CD,	inti, Parent.gPlant)
			Call .vspdData.SetText(C_QTY,	inti, "0")
			Call .vspdData.SetText(C_PRICE,	inti, "0")
			Call .vspdData.SetText(C_DOC_AMT,	inti, "0")

			SetSpreadColor inti, inti
		Next
		
	
	End with

	If Err.number = 0 Then FncInsertRow = True                                                          '☜: Processing is OK
    
	Set gActiveElement = document.ActiveElement   
        
End Function

Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"
	frm1.vspdData.Col=C_PLANT_CD
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)
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
		frm1.vspdData.Col = C_PLANT_CD
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_PLANT_NM
		frm1.vspdData.Text = arrRet(1)		
	End If	
		
End Function

'========================================================================================================
Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
	
	With frm1.vspdData 
		If .MaxRows = 0 Then
			Exit Function
		End If
	
		.focus
		ggoSpread.Source = frm1.vspdData

		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
	End With
End Function
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function
'========================================================================================================
Function FncPrev() 
	On Error Resume Next						
End Function
'========================================================================================================
Function FncNext()
	On Error Resume Next						
End Function
'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function
'========================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")	

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function
'========================================================================================================
Function DbQuery()
	Err.Clear													

	DbQuery = False												

	Dim strVal

	call SetDefaultCur

	If   LayerShowHide(1) = False Then

         Exit Function 
    End If

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001				
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtHCCNo.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If

	Call RunMyBizASP(MyBizASP, strVal)						
	
	DbQuery = True													
End Function
'========================================================================================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDel
	Dim intInsrtCnt

	DbSave = False													
    
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID

		lGrpCnt = 1

		strVal = ""
		intInsrtCnt = 1

		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0

			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag
				
					strVal = strVal & "C" & Parent.gColSep 

					.vspdData.Col = C_CC_SEQ
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_PLANT_CD							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_ITEM_CD								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_UNIT								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_QTY								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_PRICE								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_DOC_AMT								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_REMARK								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					strVal = strVal & lRow & Parent.gRowSep	

					lGrpCnt = lGrpCnt + 1
					intInsrtCnt = intInsrtCnt + 1

				Case ggoSpread.UpdateFlag						
								
					strVal = strVal & "U" & Parent.gColSep 

					.vspdData.Col = C_CC_SEQ
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_PLANT_CD							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_ITEM_CD								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_UNIT								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_QTY								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_PRICE								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_DOC_AMT								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_REMARK								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					strVal = strVal & lRow & Parent.gRowSep	
	

					lGrpCnt = lGrpCnt + 1
					'intInsrtCnt = intInsrtCnt + 1


				Case ggoSpread.DeleteFlag							
				
					strVal = strVal & "D" & Parent.gColSep 

					.vspdData.Col = C_CC_SEQ
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_PLANT_CD							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_ITEM_CD								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_UNIT								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_QTY								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_PRICE								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_DOC_AMT								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_REMARK								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					strVal = strVal & lRow & Parent.gRowSep	


					lGrpCnt = lGrpCnt + 1
			End Select
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)						
	End With

	DbSave = True													
End Function
'========================================================================================================
Function DbDelete()
End Function
'========================================================================================================
Function DbQueryOk()												
	lgIntFlgMode = Parent.OPMD_UMODE									

	Call ggoOper.LockField(Document, "Q")						
	Call SetToolbar("11101111001111")							
	'Call HideNonRelGrid()
	lgBlnFlgChgValue = False
		
    frm1.vspdData.Focus     
	call SetDefaultCur
End Function

'========================================================================================================
sub SetDefaultCur()
		if CommonQueryRs(" CUR ", " s_cc_hdr ", " CC_NO = " & FilterVar(frm1.txtCCNo.value, "''", "S") , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			frm1.txtCur.value = replace(lgF0,chr(11),"")
		End if
End sub

'========================================================================================================
Function CCHdrQueryOk()												
'	Call HideNonRelGrid()
	Call SetToolbar("11101111000011")
		
'	If frm1.txtRefFlg.value = "M" Then
'		Call ggoSpread.SSSetColHidden(C_HsPopup,C_HsPopup,False)
'	Else
		'Call ggoSpread.SSSetColHidden(C_HsPopup,C_HsPopup,True)
'	End IF								
End Function
'========================================================================================================
Function DbSaveOk()													
	Call InitVariables
	frm1.txtCCNo.value = frm1.txtHCCNo.value  
	Call ggoOper.ClearField(Document, "2")	         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call MainQuery()
End Function
'========================================================================================================
Function DbDeleteOk()												
End Function
'========================================================================================================
Function checkCtNo(iRow)
	Dim intStartCtNo, intEndCtNo
	Dim i 
	checkCtNo = True
	
	With frm1.vspdData
		If iRow ="A" then
			For i=1 to .MaxRows 
				.Row = i
				.Col = C_StartCtNo
				intStartCtNo = .Text
				.row = i
				.Col = C_EndCtNo
				intEndCtNo = .Text
				If intStartCtNo>intEndCtNo then						
						checkCtNo= False
						Exit function
				End If
			Next
		Else
			.Row = iRow
			.Col = C_StartCtNo
			intStartCtNo = .Text
			.row = irow
			.Col = C_EndCtNo
			intEndCtNo = .Text
			If intStartCtNo>intEndCtNo then					
					checkCtNo= False
					Exit function
			End If
		End If
	
	End With

End Function 

Function BtnPreview() 
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	Dim ObjName

  If frm1.txtCCNo.value = "" Then
		Call DisplayMsgBox("173133",  "x", "통관번호", "x")
		Exit Function
	End If

	strUrl = strUrl & "CCNo|" & frm1.txtCCNo.value 
    
		ObjName = AskEBDocumentName("s4211oa1_ko441", "ebr")
		Call FncEBRPreview(ObjName, strUrl)		
		
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관Invoice추가품목등록</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="통관관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnCCNoOnClick()"></TD>
										<TD CLASS=TDT NOWRAP></TD>
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
						<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>수입자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>송장번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" MAXLENGTH=18 SIZE=20 TAG="24XXXU" ALT="송장번호"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>작성일</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtIvDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="작성일"></OBJECT></TD>
									<TD CLASS=TD5 NOWRAP>선적일</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtShipDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="선적일"></OBJECT></TD>
																			
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Carton수</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtCarton" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="총포장개수" Title="FPDOUBLESINGLE"></OBJECT></TD>
									<TD CLASS=TD5 NOWRAP>총 포장개수</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtPacking" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="총포장개수" Title="FPDOUBLESINGLE"></OBJECT>
									</td>												
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총중량</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtGrossW" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="총중량" Title="FPDOUBLESINGLE"></OBJECT></TD>										
									<TD CLASS=TD5 NOWRAP>총순중량</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtNetW" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="총순중량" Title="FPDOUBLESINGLE"></OBJECT></TD>
								</TR>		
								<TR>
									<TD CLASS=TD5 NOWRAP>통화</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCur" MAXLENGTH=18 SIZE=20 TAG="24XXXU" ALT="통화"></TD>										
									<TD CLASS=TD5 NOWRAP>총용적</TD>
									<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtMsmnt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="총용적" Title="FPDOUBLESINGLE"></OBJECT></TD>
								</TR>	
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD">
											<PARAM NAME="MaxRows" Value=0>
											<PARAM NAME="MaxCols" Value=0>
											<PARAM NAME="ReDraw" VALUE=0>
										</OBJECT>
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
					<TD WIDTH=10>&nbsp;</TD>
						<TD>    
						    <BUTTON NAME="btnRun" CLASS="CLSLBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>Commercial Invoice</BUTTON>&nbsp;
						</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(0)">통관등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">통관내역등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(2)">통관란등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SCROLLING=NO noresize  FRAMEBORDER=0  framespacing=0 TABINDEX="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHCCNo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRefFlg" TAG="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHXchRate" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHXchRateOp" TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>
