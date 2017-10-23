<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4213ma1.asp																*
'*  4. Program Name         : 통관란등록																*
'*  5. Program Desc         : 통관란등록																*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/05/04 : Coding Start												*
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                       
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim C_LanNo								
Dim C_HsCd			
Dim C_FOBDocAmt		
Dim C_FOBLocAmt		
Dim C_TotPackingCnt	
Dim C_GrossWeight	
Dim C_Measurement
Dim C_NetWeight		
Dim C_Qty			
Dim C_DocAmt		
Dim C_ChgFlg		

Dim gblnWinEvent					

Const BIZ_PGM_QRY_ID = "s4213mb1.asp"		
Const BIZ_PGM_SAVE_ID = "s4213mb1.asp"		
Const EXCC_DETAIL_ENTRY_ID = "s4212ma1"		
Const EXCC_HEADER_ENTRY_ID = "s4211ma1"
Const EXCC_ASSIGN_ENTRY_ID ="s4214ma1"		'☆: 이동할 ASP명 : container 배정 

'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_LanNo				= 1				
	C_HsCd				= 2
	C_FOBDocAmt			= 3
	C_FOBLocAmt			= 4
	C_TotPackingCnt		= 5
	C_GrossWeight			= 6
	C_Measurement			= 7
	C_NetWeight			= 8
	C_Qty				= 9
	C_DocAmt			= 10
	C_ChgFlg			= 11

End Sub
'========================================================================================================
Function InitVariables()	
	lgIntGrpCount = 0						
	lgStrPrevKey = ""						
	lgLngCurRows = 0 						

	frm1.txtLocCurrency.value = Parent.gCurrency
	frm1.txtLocCCCurrency.value = Parent.gCurrency
	frm1.txtCurrencyUSD.value = "USD"

	lgBlnFlgChgValue = False	
	lgIntFlgMode = Parent.OPMD_CMODE			
	gblnWinEvent = False
End Function
'========================================================================================================
Sub SetDefaultVal()
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
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
			
		.vspdData.MaxCols = C_ChgFlg
		.vspdData.MaxRows = 0
			
		.vspdData.ReDraw = False
									
		Call GetSpreadColumnPos("A")	

		ggoSpread.SSSetEdit		C_LanNo, "란번호", 10, 0
		ggoSpread.SSSetEdit		C_HsCd, "HS부호", 20, 0
		ggoSpread.SSSetFloat	C_FOBDocAmt,"FOB금액(US $)",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_FOBLocAmt,"FOB자국금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    Call AppendNumberPlace("6", "10", "0")
	    ggoSpread.SSSetFloat	C_TotPackingCnt,"포장갯수" ,15, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_GrossWeight,"총중량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    ggoSpread.SSSetFloat	C_Measurement,"총용적" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat	C_NetWeight,"순중량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat	C_Qty,"총수량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_DocAmt,"총금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_ChgFlg, "Chgfg", 6
			
		SetSpreadLock "", 0, -1, ""

		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
			
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
			
		ggoSpread.SpreadLock C_LanNo, lRow, -1
		ggoSpread.SpreadLock C_HsCd, lRow, -1
		ggoSpread.SpreadUnLock C_FOBDocAmt,lRow, -1
		ggoSpread.SSSetRequired C_FOBDocAmt, lRow, -1 
		ggoSpread.SSSetRequired C_FOBLocAmt, lRow, -1 		
		ggoSpread.SpreadLock C_TotPackingCnt, lRow, -1  
		ggoSpread.SpreadLock C_GrossWeight, lRow, -1
		ggoSpread.SpreadLock C_Measurement, lRow, -1
		ggoSpread.SpreadLock C_NetWeight, lRow, -1
		ggoSpread.SpreadLock C_Qty, lRow, -1
		ggoSpread.SpreadLock C_DocAmt, lRow, -1
		ggoSpread.SpreadLock C_ChgFlg, lRow, -1
			
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadColor(ByVal lRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
	    
		.Redraw = False

		ggoSpread.SSSetProtected C_LanNo, lRow, lRow
		ggoSpread.SSSetProtected C_HsCd, lRow, lRow
		ggoSpread.SSSetRequired C_FOBDocAmt, lRow, -1 
		ggoSpread.SSSetRequired C_FOBLocAmt, lRow, -1 
		ggoSpread.SSSetProtected C_TotPackingCnt, lRow, lRow
		ggoSpread.SSSetProtected C_Qty, lRow, lRow
		ggoSpread.SSSetProtected C_DocAmt, lRow, lRow
		ggoSpread.SSSetProtected C_GrossWeight, lRow, lRow
		ggoSpread.SSSetProtected C_Measurement, lRow, lRow
		ggoSpread.SSSetProtected C_NetWeight, lRow, lRow
		ggoSpread.SSSetProtected C_ChgFlg, lRow, lRow

		.Col = 1
		.Row = .ActiveRow
		.Action = 0
		.EditMode = True

		.ReDraw = True
	End With
End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
   
    Select Case UCase(pvSpdNo)
       Case "A"
            
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_LanNo				= iCurColumnPos(1)  				
			C_HsCd				= iCurColumnPos(2)  
			C_FOBDocAmt			= iCurColumnPos(3)  
			C_FOBLocAmt			= iCurColumnPos(4)  
			C_TotPackingCnt		= iCurColumnPos(5)  
			C_GrossWeight			= iCurColumnPos(6)
			C_Measurement			= iCurColumnPos(7)  
			C_NetWeight			= iCurColumnPos(8)    
			C_Qty				= iCurColumnPos(9)  
			C_DocAmt			= iCurColumnPos(10)  
			C_ChgFlg			= iCurColumnPos(11)  
			
    End Select    
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

'========================================================================================================
Function LoadCCDtl()
	Dim strDtlOpenParam

	WriteCookie "txtCCNo", UCase(Trim(frm1.txtCCNo.value))
		
	strDtlOpenParam = EXCC_DETAIL_ENTRY_ID

	document.location.href = GetUserPath & strDtlOpenParam
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						
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
'========================================================================================================
Sub SumDocAmt()
	With frm1
		Dim strVal
		Dim dblTotDocAmt
		Dim intCnt
			
		ggoSpread.Source = .vspdData
			
		For intCnt=1 to .vspdData.MaxRows
				
			.vspdData.Col = C_FOBDocAmt
			.vspdData.Row = intCnt
				
			If .vspdData.text <>"" Then
				dblTotDocAmt = dblTotDocAmt + UNICDbl(.vspdData.text)		
			End If	

		Next
			
		.txtFobDocAmt.Text = UNIFormatNumberByCurrecny(dblTotDocAmt,frm1.txtCurrencyUSD.value,Parent.ggAmtOfMoneyNo)
	End With

End Sub		
'========================================================================================================
Sub SumLocAmt()
	With frm1
		Dim strVal
		Dim dblTotLocAmt
		Dim intCnt
			
		ggoSpread.Source = .vspdData
			
		For intCnt=1 to .vspdData.MaxRows
			.vspdData.Col = C_FOBLocAmt
			.vspdData.Row = intCnt
				
			If .vspdData.text <>"" Then
				dblTotLocAmt = dblTotLocAmt + UNICDbl(.vspdData.text)
			End If	
		Next
			
		.txtFobLocAmt.text = UNIFormatNumber(dblTotLocAmt, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	End With

End Sub		
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
	Case 0 
		Call CookiePage(1)
		Call PgmJump(EXCC_HEADER_ENTRY_ID)
	Case 1
		Call CookiePage(1)
		Call PgmJump(EXCC_DETAIL_ENTRY_ID)
	Case 2
		Call CookiePage(1)
		Call PgmJump(EXCC_ASSIGN_ENTRY_ID)
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
	
End Sub
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'통관FOB금액 
		ggoOper.FormatFieldByObjectOfCur .txtFobDocAmt, .txtCurrencyUSD.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		
		' 통관금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCCCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		'FOB금액 
		ggoSpread.SSSetFloatByCellOfCur C_FOBDocAmt,-1, .txtCurrencyUSD.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'총금액 
		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCCCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029													
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")								
	Call InitSpreadSheet												

	Call SetDefaultVal
	Call CookiePage(0)
	Call InitVariables
	Call SetToolbar("1110000000001111")				

	frm1.txtCCNo.focus
	Set gActiveElement = document.activeElement 
End Sub

'========================================================================================================
Sub btnCCNoOnClick()
	Call OpenExCCNoPop()
End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

	End With
End Sub
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim strVal
	Dim dblFobAmt
	Dim dblLocAmt
		
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.row = Row
		If Col = C_FOBDocAmt Then
			.Col = Col
			dblFobAmt = UNICDbl(.Text)

			If frm1.txtHExchRateOp.value = "*" then
				dblLocAmt = dblFobAmt * UNICDbl(frm1.txtUsdXchRate.text)
			ElseIf frm1.txtHExchRateOp.value = "/" then
				dblLocAmt = dblFobAmt / UNICDbl(frm1.txtUsdXchRate.text)
			End If
				
			.Col = C_FOBLocAmt
			.Text = UNIFormatNumberByCurrecny(dblLocAmt,parent.gCurrency,Parent.ggAmtOfMoneyNo)
		End If
	End With
		
	Call SumDocAmt()
	Call SumLocAmt()
			
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
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
	
	Call SetPopupMenuItemInf("0000111111")	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                  
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
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
	
End Sub
'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False											

	Err.Clear												
		
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
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "x", "x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")		
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call ggoOper.LockField(Document, "N")
	Call SetDefaultVal
		
	Call SetToolbar("1110000000001111")
	Call InitVariables
		
	FncNew = True										

End Function
'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False										
		
	<% '------ Precheck area ------ %>
	If lgIntFlgMode <> Parent.OPMD_UMODE Then						
		Call DisplayMsgBox("900002", "x", "x", "x")
'			Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
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
	If Not chkField(Document, "2") Then Exit Function
	If ggoSpread.SSDefaultCheck = False Then Exit Function

	Call DbSave							
		
	FncSave = True						
End Function
'========================================================================================================
Function FncCopy()
	frm1.vspdData.ReDraw = False

	ggoSpread.Source = frm1.vspdData	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
End Function
'========================================================================================================
Function FncCancel() 
	ggoSpread.Source = frm1.vspdData
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	ggoSpread.EditUndo	
End Function
'========================================================================================================
Function FncInsertRow()
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False
		ggoSpread.InsertRow
		.vspdData.ReDraw = True

		SetSpreadColor .vspdData.ActiveRow
    End With
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

'			IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
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
				Case ggoSpread.UpdateFlag								
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep	

					.vspdData.Col = C_LanNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_HsCd								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						
					.vspdData.Col = C_Qty								
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					.vspdData.Col = C_DocAmt							
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					' Setting LocAmt Value							
					strVal = strVal & Trim(UNICDbl(0)) & Parent.gColSep
												
					.vspdData.Col = C_FOBDocAmt							
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					.vspdData.Col = C_FOBLocAmt							
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					.vspdData.Col = C_TotPackingCnt						
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					.vspdData.Col = C_NetWeight							
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep
						
					' unit							
					strVal = strVal & "" & Parent.gColSep

					' cc_seq						
					strVal = strVal & 0 & Parent.gColSep

					' ext_qty1							
					strVal = strVal & 0 & Parent.gColSep

					'' ext_qty1						
					strVal = strVal & 0 & Parent.gColSep

					'.' ext_qty1							
					strVal = strVal & 0 & Parent.gColSep

					'' ext_amt1							
					strVal = strVal & 0 & Parent.gColSep

					'' ext_amt1								
					strVal = strVal & 0 & Parent.gColSep

					'' ext_amt1									
					strVal = strVal & 0 & Parent.gColSep

					' ext_cd1							
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					' ext_cd1
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gColSep

					' ext_cd1
					strVal = strVal & Trim(UNICDbl(.vspdData.Text)) & Parent.gRowSep
	
					lGrpCnt = lGrpCnt + 1

			End Select
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
			
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
	Call SetToolbar("1110100100011111")			
		
	lgBlnFlgChgValue = False
End Function
'========================================================================================================
Function DbSaveOk()								
	Call InitVariables
	Call SetDefaultVal()
	frm1.txtCCNo.value = frm1.txtHCCNo.value  
	Call ggoOper.ClearField(Document, "2")	         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call MainQuery()
End Function
'========================================================================================================
Function DbDeleteOk()							
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관란정보</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>USD환율</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4213ma1_fpDoubleSingle3_txtUsdXchRate.js'></script>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>통관금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s4213ma1_fpDoubleSingle1_txtDocAmt.js'></script></TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCCCurrency" ALT="통관금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU">
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>통관자국금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s4213ma1_fpDoubleSingle2_txtLocAmt.js'></script></TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCCCurrency" ALT="통관자국금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>통관FOB금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s4213ma1_fpDoubleSingle1_txtFobDocAmt.js'></script>
												</TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrencyUSD" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐" VALUE = "USD"></TD>
											</TR>
										</TABLE>
									</TD>	
									<TD CLASS=TD5 NOWRAP>통관FOB원화금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s4213ma1_fpDoubleSingle2_txtFobLocAmt.js'></script>
												</TD>
												<TD>
													&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU">
												</TD>												
											</TR>
										</TABLE>
									</TD>			
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<script language =javascript src='./js/s4213ma1_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(2)">Container배정</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(0)">통관등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">통관내역등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHCCNo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHOpenDt" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHExchRateOp" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>
