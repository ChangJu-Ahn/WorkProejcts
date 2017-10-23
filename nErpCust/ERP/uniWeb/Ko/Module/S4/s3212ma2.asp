<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3212ma2.asp																*
'*  4. Program Name         : Local L/C 내역등록														*
'*  5. Program Desc         : Local L/C 내역등록														*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/20																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/23 : Coding Start												*
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim C_LCSeq			
Dim C_ItemCd		
Dim C_ItemNm		
Dim C_Unit			
Dim C_RemainSoQty	
Dim C_LCQty			
Dim C_Price			
Dim C_DocAmt		
Dim C_OverTolerance	
Dim C_UnderTolerance
Dim C_ClsFlg		
Dim C_HsCd			
Dim C_SoNo			
Dim C_SoSeq			
Dim C_DNQty			
Dim C_CCQty			
Dim C_BLQty			
Dim C_DnNo			
Dim C_DnSeq
Dim C_TrackingNo			
Dim C_ChgFlg
Dim C_ItemSpec		

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent					
	
Const BIZ_PGM_QRY_ID		= "s3212mb2.asp"		
Const BIZ_PGM_SAVE_ID		= "s3212mb2.asp"		
Const LC_HEADER_ENTRY_ID	= "s3211ma2"

'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_LCSeq			= 1
	C_ItemCd		= 2		
	C_ItemNm		= 3
        C_ItemSpec		= 4
	C_Unit			= 5
	C_RemainSoQty	= 6
	C_LCQty			= 7
	C_Price			= 8
	C_DocAmt		= 9
	C_OverTolerance	= 10
	C_UnderTolerance= 11
	C_ClsFlg		= 12
	C_HsCd			= 13
	C_SoNo			= 14
	C_SoSeq			= 15
	C_DNQty			= 16
	C_CCQty			= 17
	C_BLQty			= 18
	C_DnNo			= 19
	C_DnSeq			= 20
	C_TrackingNo	= 21
	C_ChgFlg		= 22

End Sub

'========================================================================================================
 Function InitVariables()
 	lgIntFlgMode = parent.OPMD_CMODE			
 	lgBlnFlgChgValue = False			
 	lgIntGrpCount = 0					
 	lgStrPrevKey = ""					
 	lgLngCurRows = 0 					
		
 	gblnWinEvent = False
 End Function

'========================================================================================================
 Sub SetDefaultVal()
 	lgBlnFlgChgValue = False
 	frm1.txtLcNo.focus
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

 		ggoSpread.Spreadinit "V20021123",,parent.gAllowDragDropSpread    
			
 		.vspdData.MaxCols = C_ChgFlg
 		.vspdData.MaxRows = 0
			
 		.vspdData.ReDraw = False
						
 		Call GetSpreadColumnPos("A")	
			
 		ggoSpread.SSSetEdit		C_LcSeq, "L/C순번", 10, 1
 		ggoSpread.SSSetEdit		C_ItemCd, "품목", 18, 0
 		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 40, 0
                ggoSpread.SSSetEdit		C_ItemSpec, "규격", 20
 		ggoSpread.SSSetEdit		C_Unit, "단위", 10, 2
 		ggoSpread.SSSetFloat	C_RemainSoQty,"수주잔량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat	C_LCQty,"L/C수량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat	C_Price,"단가",15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
 		ggoSpread.SSSetFloat	C_DocAmt,"금액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat	C_OverTolerance,"과부족허용율(+)",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
 	    ggoSpread.SSSetFloat	C_UnderTolerance,"과부족허용율(-)",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"        
 		ggoSpread.SSSetCheck	C_ClsFlg, "중단", 12,,,true
 		ggoSpread.SSSetEdit		C_HsCd, "HS부호", 20, 0
 		ggoSpread.SSSetEdit		C_SoNo, "수주번호", 18, 0
 		ggoSpread.SSSetEdit		C_SoSeq, "수주순번", 10, 1
         ggoSpread.SSSetFloat	C_DNQty,"출하수량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
         ggoSpread.SSSetFloat	C_CCQty,"통관수량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
         ggoSpread.SSSetFloat	C_BLQty,"매출수량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
 		ggoSpread.SSSetEdit		C_DnNo, "출하번호", 18, 0
 		ggoSpread.SSSetEdit		C_DnSeq, "출하순번", 10, 1
                ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No", 25, 0
			
 		ggoSpread.SSSetEdit		C_ChgFlg, "Chgfg", 1, 2
			
 		SetSpreadLock "", 0, -1, ""

		 
 		 Call ggoSpread.SSSetColHidden(C_LcSeq,C_LcSeq,True)
 		 Call ggoSpread.SSSetColHidden(C_DNQty,C_DNQty,True)
 		 Call ggoSpread.SSSetColHidden(C_CCQty,C_CCQty,True)
 		 Call ggoSpread.SSSetColHidden(C_BLQty,C_BLQty,True)
 		 Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
			
 		.vspdData.ReDraw = True
 	End With
 End Sub

'========================================================================================================
 Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
     With frm1
 		ggoSpread.Source = .vspdData
			
 		.vspdData.ReDraw = False
			
 		ggoSpread.SpreadLock C_ItemCd, lRow, -1
 		ggoSpread.SpreadLock C_ItemNm, lRow, -1
 		ggoSpread.SpreadLock C_Unit, lRow, -1
                ggoSpread.SpreadLock C_ItemSpec, lRow, -1
 		ggoSpread.SpreadLock C_RemainSoQty, lRow, -1
 		ggoSpread.SSSetRequired	C_LCQty, lRow, -1
 		ggoSpread.SSSetRequired	C_Price, lRow, -1
 		ggoSpread.SpreadLock C_DocAmt, lRow, -1
 		ggoSpread.SpreadLock C_HsCd, lRow, -1
 		ggoSpread.SpreadLock C_SoNo, lRow, -1
 		ggoSpread.SpreadLock C_SoSeq, lRow, -1
 		ggoSpread.SpreadLock C_DnNo, lRow, -1
 		ggoSpread.SpreadLock C_DnSeq, lRow, -1
 		ggoSpread.SpreadLock C_DNQty, lRow, -1
 		ggoSpread.SpreadLock C_CCQty, lRow, -1
 		ggoSpread.SpreadLock C_BLQty, lRow, -1
                ggoSpread.SpreadLock C_TrackingNo, lRow, -1
 		ggoSpread.SpreadLock C_ChgFlg, lRow, -1
			
 		.vspdData.ReDraw = True
 	End With
 End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
 	ggoSpread.Source = frm1.vspdData
		
     With frm1.vspdData

 		ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
                ggoSpread.SSSetProtected C_ItemSpec, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_RemainSoQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_LCQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_Price, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_DocAmt, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_HsCd, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LcSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_SoNo, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_SoSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_DnNo, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_DnSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_DNQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_CCQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_BLQty, pvStartRow, pvEndRow
                ggoSpread.SSSetProtected C_TrackingNo, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ChgFlg, pvStartRow, pvEndRow
		
 	End With
 End Sub

'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 Dim iCurColumnPos
    
 Select Case UCase(pvSpdNo)
    Case "A"
         ggoSpread.Source = frm1.vspdData
         Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
							
 		C_LCSeq			=iCurColumnPos(1)			
 		C_ItemCd		=iCurColumnPos(2) 					
 		C_ItemNm		=iCurColumnPos(3)
                C_ItemSpec		= iCurColumnPos(4)						
 		C_Unit			= iCurColumnPos(5)
		C_RemainSoQty		= iCurColumnPos(6)			
		C_LCQty			= iCurColumnPos(7)			
		C_Price			= iCurColumnPos(8)			
		C_DocAmt		= iCurColumnPos(9)			
		C_OverTolerance		= iCurColumnPos(10)			
		C_UnderTolerance	= iCurColumnPos(11)			
		C_ClsFlg		= iCurColumnPos(12)			
		C_HsCd			= iCurColumnPos(13)			
		C_SoNo			= iCurColumnPos(14)			
		C_SoSeq			= iCurColumnPos(15)			
		C_DNQty			= iCurColumnPos(16)			
		C_CCQty			= iCurColumnPos(17)			
		C_BLQty			= iCurColumnPos(18)			
		C_DnNo			= iCurColumnPos(19)			
		C_DnSeq			= iCurColumnPos(20)			
		C_TrackingNo		= iCurColumnPos(21)			
		C_ChgFlg		= iCurColumnPos(22)				
			
 End Select    
End Sub

'========================================================================================================
 Function OpenLCNoPop()
 	Dim strRet
 	Dim iCalledAspName
 	Dim IntRetCD
		
 	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = "PROTECTED" Then Exit Function
		
 	gblnWinEvent = True
		
 	iCalledAspName = AskPRAspName("s3211pa2")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3211pa2", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
 	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
 			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False
		
 	If strRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetLCNo(strRet)
 	End If	
 End Function

'========================================================================================================
 Function OpenSODtlRef()
 	Dim arrRet
 	Dim strSONo
 	Dim arrParam(10)
 	Dim iCalledAspName
 	Dim IntRetCD		
		
 	If UCase(Trim(frm1.txtHLCNo.value)) = "" Then
 		Call DisplayMsgBox("900002", "x", "x", "x")
 		Exit Function
 	End If
		
 	arrParam(0) = Trim(frm1.txtSONo.value)					
 	arrParam(1) = Trim(frm1.txtApplicant.value)	
 	arrParam(2) = Trim(frm1.txtApplicantNm.value)					
 	arrParam(3) = Trim(frm1.txtSalesGroup.value)	
 	arrParam(4) = Trim(frm1.txtSalesGroupNm.value)			
 	arrParam(5) = Trim(frm1.txtPayTerms.value)	
 	arrParam(6) = Trim(frm1.txtPayTermsNm.value)				
 	arrParam(7) = Trim(frm1.txtCurrency.value) 									

 	iCalledAspName = AskPRAspName("s3112ra2")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3112ra2", "X") 		
 		Exit Function
 	End If

 	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
 	If arrRet(0, 0) = "" Then
 		Exit Function
 	Else
 		Call SetSODtlRef(arrRet)
 	End If	
 End Function

'========================================================================================================
 Function OpenDNDtlRef()
 	Dim arrRet
 	Dim strSONo
 	Dim arrParam(10)
 	Dim iCalledAspName
 	Dim IntRetCD
				
 	If UCase(Trim(frm1.txtHLCNo.value)) = "" Then
 		Call DisplayMsgBox("900002", "x", "x", "x")
 		Exit Function
 	End If
		
 	arrParam(0) = Trim(frm1.txtSONo.value)					
 	arrParam(1) = Trim(frm1.txtApplicant.value)	
 	arrParam(2) = Trim(frm1.txtApplicantNm.value)					
 	arrParam(3) = Trim(frm1.txtSalesGroup.value)	
 	arrParam(4) = Trim(frm1.txtSalesGroupNm.value)			
 	arrParam(5) = Trim(frm1.txtPayTerms.value)	
 	arrParam(6) = Trim(frm1.txtPayTermsNm.value)				
 	arrParam(7) = Trim(frm1.txtCurrency.value)
 	'2002-10-10추가 									
 	arrParam(8) = Trim(frm1.txtHLCNo.value)

 	iCalledAspName = AskPRAspName("s4112ra4")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s4112ra4", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
 	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
 	If arrRet(0, 0) = "" Then
 		Exit Function
 	Else
 		Call SetDNDtlRef(arrRet)
 	End If	
 End Function	

'========================================================================================================
 Function SetLCNo(strRet)
 	frm1.txtLCNo.value = strRet(0)
 	frm1.txtLcNo.focus
 End Function

'========================================================================================================
 Function SetDNDtlRef(arrRet)
 	Dim intRtnCnt, strData
 	Dim TempRow, I, j
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	'2002-08-12 메세지 추가			
 	Dim strtemp1, strtemp2, strMessage
		
 	With frm1
 		.vspdData.focus
 		ggoSpread.Source = .vspdData
 		.vspdData.ReDraw = False	

 		TempRow = .vspdData.MaxRows					
 		intLoopCnt = Ubound(arrRet, 1)				

 		For intCnt = 1 to intLoopCnt + 1
 			blnEqualFlg = False

 			If TempRow <> 0 Then
 				For j = 1 To TempRow
 					.vspdData.Row = j
 					.vspdData.Col = C_DnNo
 					strtemp1 = .vspdData.text
						
 					If Trim(.vspdData.Text) = Trim(arrRet(intCnt - 1, 0)) Then
 						.vspdData.Row = j
 						.vspdData.Col = C_DnSeq
 						strtemp2 = .vspdData.text
							
 						If Trim(.vspdData.Text) = Trim(arrRet(intCnt - 1, 1)) Then
 							strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
 							blnEqualFlg = True
 							Exit For
 						End If
 					End If
 				Next
 			End If

 			If blnEqualFlg = False Then
		
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
 				.vspdData.Col = 0
 				.vspdData.Text = ggoSpread.InsertFlag

 				.vspdData.Col = C_SoNo				
 				.vspdData.text = arrRet(intCnt - 1, 13)
 				.vspdData.Col = C_SoSeq				
 				.vspdData.text = arrRet(intCnt - 1, 14)
 				.vspdData.Col = C_ItemCd			
 				.vspdData.text = arrRet(intCnt - 1, 2)
 				.vspdData.Col = C_ItemNm			
 				.vspdData.text = arrRet(intCnt - 1, 3)
                                .vspdData.Col = C_ItemSpec			
				.vspdData.text = arrRet(intCnt - 1, 4)
 				.vspdData.Col = C_Unit				
 				.vspdData.text = arrRet(intCnt - 1, 5)
 				.vspdData.Col = C_RemainSoQty				
 				.vspdData.text = arrRet(intCnt - 1, 6)
 				.vspdData.Col = C_LCQty				
 				.vspdData.text = arrRet(intCnt - 1, 6)
 				.vspdData.Col = C_Price				
 				.vspdData.text = arrRet(intCnt - 1, 7)
 				.vspdData.Col = C_DocAmt			
 				.vspdData.text = arrRet(intCnt - 1, 8)
 				.vspdData.Col = C_OverTolerance		
 				.vspdData.text = arrRet(intCnt - 1, 10)
 				.vspdData.Col = C_UnderTolerance	
 				.vspdData.text = arrRet(intCnt - 1, 11)
 				.vspdData.Col = C_HsCd				
 				.vspdData.text = arrRet(intCnt - 1, 12)
 				.vspdData.Col = C_DnNo				
 				.vspdData.text = arrRet(intCnt - 1, 0)
 				.vspdData.Col = C_DnSeq				
 				.vspdData.text = arrRet(intCnt - 1, 1)
                                .vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 15)
 				.vspdData.Col = C_ChgFlg			
 				.vspdData.text = .vspdData.Row

 				SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)

 				lgBlnFlgChgValue = True
 			End If
 		Next
 		Call SumLCAmt()
 		'2002-08-12 메세지 추가			
 		If strMessage <> "" Then
 			Call DisplayMsgBox("17a005", "X",strmessage,"출하번호" & "," & "출하순번")
 			.vspdData.ReDraw = True
 		End If
			
 		.vspdData.ReDraw = True

 	End With
 End Function

'========================================================================================================
 Function SetSODtlRef(arrRet)
 	Dim intRtnCnt, strData
 	Dim TempRow, I, j
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	Dim strtemp1, strtemp2, strMessage

 	With frm1
 		.vspdData.focus
 		ggoSpread.Source = .vspdData
 		.vspdData.ReDraw = False	

 		TempRow = .vspdData.MaxRows								
 		intLoopCnt = Ubound(arrRet, 1)							
			
 		For intCnt = 1 to intLoopCnt + 1
 			blnEqualFlg = False

 			If TempRow <> 0 Then
 				For j = 1 To TempRow
 					.vspdData.Row = j
 					.vspdData.Col = C_SoNo
 					strtemp1 = .vspdData.text
							
 					If Trim(.vspdData.Text) = Trim(arrRet(intCnt - 1, 0)) Then
 						.vspdData.Row = j
 						.vspdData.Col = C_SoSeq
 						strtemp2 = .vspdData.text
	
 						If Trim(.vspdData.Text) = Trim(arrRet(intCnt - 1, 1)) Then
 							strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
 							blnEqualFlg = True
 							Exit For
 						End If
 					End If	
 				Next
 			End If

 			If blnEqualFlg = False Then
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = CLng(TempRow) + CLng(intCnt)

 				.vspdData.Col = 0
 				.vspdData.Text = ggoSpread.InsertFlag
 				.vspdData.Col = C_SoNo							
 				.vspdData.text = arrRet(intCnt - 1, 0)
 				.vspdData.Col = C_SoSeq							
 				.vspdData.text = arrRet(intCnt - 1, 1)
 				.vspdData.Col = C_ItemCd						
 				.vspdData.text = arrRet(intCnt - 1, 2)
 				.vspdData.Col = C_ItemNm						
 				.vspdData.text = arrRet(intCnt - 1, 3)
                                .vspdData.Col = C_ItemSpec						
				.vspdData.text = arrRet(intCnt - 1, 4)
 				.vspdData.Col = C_Unit							
 				.vspdData.text = arrRet(intCnt - 1, 5)
 				.vspdData.Col = C_RemainSoQty							
 				.vspdData.text = arrRet(intCnt - 1, 6)
 				.vspdData.Col = C_LCQty							
 				.vspdData.text = arrRet(intCnt - 1, 6)
 				.vspdData.Col = C_Price							
 				.vspdData.text = arrRet(intCnt - 1, 7)
 				.vspdData.Col = C_DocAmt						
 				.vspdData.text = arrRet(intCnt - 1, 8)
 				.vspdData.Col = C_OverTolerance					
 				.vspdData.text = arrRet(intCnt - 1, 10)
 				.vspdData.Col = C_UnderTolerance				
 				.vspdData.text = arrRet(intCnt - 1, 11)
 				.vspdData.Col = C_HsCd							
 				.vspdData.text = arrRet(intCnt - 1, 12)
                                .vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 13)
 				.vspdData.Col = C_ChgFlg							
 				.vspdData.text = .vspdData.Row
															
 				SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)

 				lgBlnFlgChgValue = True
 			End If
 		Next
 		Call SumLCAmt()
			
 		If strMessage <> "" Then
 			Call DisplayMsgBox("17a005", "X",strmessage,"수주번호" & "," & "수주순번")
 			.vspdData.ReDraw = True
 		End If

 		.vspdData.ReDraw = True

 	End With
 End Function	

'========================================================================================================
 Function CookiePage(ByVal Kubun)

 	On Error Resume Next

 	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
 	Dim strTemp, arrVal

 	If Kubun = 1 Then

 		WriteCookie CookieSplit , frm1.txtLCNo.value

 	ElseIf Kubun = 0 Then

 		strTemp = ReadCookie(CookieSplit)
				
 		If strTemp = "" then Exit Function
				
 		frm1.txtLCNo.value =  strTemp
			
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
 Function OpenCookie()
 	frm1.txtLCNo.Value = ReadCookie("txtLCNo")
 	WriteCookie "txtLCNo", ""
 End Function

'========================================================================================================
 Sub SumLCAmt()
 	Dim dblQty
 	Dim dblPrice
 	Dim dblAmt
 	Dim dblTotAmt
 	Dim intCnt
		
 	With frm1
 		ggoSpread.Source = .vspdData

 		For intCnt=1 to .vspdData.MaxRows
				
 			.vspdData.Col = C_DocAmt
 			.vspdData.Row = intCnt
				
 			If .vspdData.text <>"" Then
 				dblTotAmt = dblTotAmt + UNICDbl(.vspdData.text)
 			End If	
 		Next
			
 		.txtTotItemAmt.Text = UNIFormatNumberByCurrecny(dblTotAmt,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
	
 	End With
 End Sub		

'========================================================================================================
 Function JumpChgCheck()
 	Dim IntRetCD

 	ggoSpread.Source = frm1.vspdData	
 	If ggoSpread.SSCheckChange = True Then
 		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

 	Call CookiePage(1)
 	Call PgmJump(LC_HEADER_ENTRY_ID)
 End Function	

'========================================================================================================
 Sub CurFormatNumericOCX()
 	With frm1
 		'총개설금액 
 		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
 		'총품목금액 
 		ggoOper.FormatFieldByObjectOfCur .txtTotItemAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
 	End With
 End Sub

'========================================================================================================
 Sub CurFormatNumSprSheet()
 	With frm1
 		ggoSpread.Source = frm1.vspdData
 		'단가 
 		ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
 		'금액 
 		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
 	End With
 End Sub

'========================================================================================================
 Sub Form_Load()
 	Call LoadInfTB19029														
 	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
     Call ggoOper.LockField(Document, "N")									
 	Call InitSpreadSheet													

 	Call SetDefaultVal
 	Call CookiePage(0)
 	Call InitVariables
		
 	Call SetToolBar("11100000000011")										

 	frm1.txtLcNo.focus
 End Sub
	
	
'========================================================================================================
 Sub vspdData_Change(ByVal Col, ByVal Row )
 	Dim dblQty
 	Dim dblPrice
 	Dim dblAmt

 	ggoSpread.Source = frm1.vspdData

 	Select Case Col
 		Case C_LCQty
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = Col

 			dblQty = frm1.vspdData.Text

 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_Price

 			dblPrice = frm1.vspdData.Text

 			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_DocAmt
			
 			frm1.vspdData.Text = UNIFormatNumberByCurrecny(dblAmt,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
				
 			Call SumLCAmt()
					
 		Case C_Price 
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = Col

 			dblPrice = frm1.vspdData.Text

 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_LCQty

 			dblQty = frm1.vspdData.Text

 			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_DocAmt
 			
 			frm1.vspdData.Text = UNIFormatNumberByCurrecny(dblAmt,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)

				
 			Call SumLCAmt()
 		Case Else

 	End Select
 	ggoSpread.UpdateRow Row

 	lgBlnFlgChgValue = True
 End Sub

'========================================================================================================
 Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 	With frm1.vspdData 
 		ggoSpread.Source = frm1.vspdData

 		If Row > 0 And Col = C_ClsFlg Then
 			lgBlnFlgChgValue = False	
 		End If
 	End With
 End Sub		

'========================================================================================================
 Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
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

 '------ Developer Coding part (Start ) -------------------------------------------------------------- 
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

 If Button = 2 And gMouseClickStatus = "SPC" Then
    gMouseClickStatus = "SPCR"
 End If
    
 '------ Developer Coding part (Start ) -------------------------------------------------------------- 
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
 ggoSpread.Source = frm1.vspdData
 Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
 Call GetSpreadColumnPos("A")
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

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
        If DbQuery = False Then
           Call RestoreToolBar()
           Exit Sub
        End if
     End If
 End if
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
 Call SetPopupMenuItemInf("0111111111")
    
 gMouseClickStatus = "SPC"
 Set gActiveSpdSheet = frm1.vspdData
    
 If frm1.vspdData.MaxRows = 0 Then 
 	Exit Sub
 End If  
	   
 If Row <= 0 Then
 	ggoSpread.Source = frm1.vspdData
		
 	If lgSortKey = 1 Then
 		ggoSpread.SSSort Col				'Sort in Ascending
 		lgSortkey = 2
 	Else
 		ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 		lgSortkey = 1
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
 '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
 End If
	
End Sub

'========================================================================================================
 Function FncQuery()
 	Dim IntRetCD

 	FncQuery = False										

 	Err.Clear												

 	ggoSpread.Source = frm1.vspdData

 	If ggoSpread.SSCheckChange = True Then
 		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

 	Call ggoOper.ClearField(Document, "2")			
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
 		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")

 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

 	Call ggoOper.ClearField(Document, "A")				
 	Call ggoOper.LockField(Document, "N")		
 	Call SetDefaultVal
 	Call SetToolBar("11100000000011")			
 	Call InitVariables							
		
 	FncNew = True								

 End Function
	
'========================================================================================================
 Function FncDelete()
 	Dim IntRetCD

 	FncDelete = False							
		
 	If lgIntFlgMode <> parent.OPMD_UMODE Then						
 		Call DisplayMsgBox("900002", "x", "x", "x")
'			Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
 		Exit Function
 	End If

 	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

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

 	If Not chkField(Document, "2")  Then		
 		Exit Function
 	End If
		
 	If ggoSpread.SSDefaultCheck = False Then	
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
 	SetSpreadColor frm1.vspdData.ActiveRow

 	frm1.vspdData.ReDraw = True
 End Function

'========================================================================================================
 Function FncCancel() 
 	ggoSpread.Source = frm1.vspdData
 	If frm1.vspdData.MaxRows < 1 Then Exit Function
 	ggoSpread.EditUndo									
 	Call SumLCAmt()
 End Function

'========================================================================================================
 Function FncInsertRow()
 End Function
 
'========================================================================================================
 Function FncDeleteRow()
 	Dim lDelRows
 	Dim iDelRowCnt, i
	
 	With frm1.vspdData 
	
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
 	Call parent.FncExport(parent.C_SINGLEMULTI)
 End Function

'========================================================================================================
 Function FncFind() 
 	Call parent.FncFind(parent.C_SINGLEMULTI, False)
 End Function

'========================================================================================================
 Function FncExit()
 	Dim IntRetCD

 	FncExit = False

 	ggoSpread.Source = frm1.vspdData
 	If ggoSpread.SSCheckChange = True Then
 		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	

'			IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

 	FncExit = True
 End Function
'========================================================================================================
Sub FncSplitColumn()

 If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
    Exit Sub
 End If

 ggoSpread.Source = gActiveSpdSheet
 ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Sub PopSaveSpreadColumnInf()
 ggoSpread.Source = gActiveSpdSheet
 Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
 ggoSpread.Source = gActiveSpdSheet
 Call ggoSpread.RestoreSpreadInf()
 Call InitSpreadSheet()      
 Call ggoSpread.ReOrderingSpreadData()

End Sub
'========================================================================================================
 Function DbQuery()
 	Err.Clear													

 	DbQuery = False												
		
					
 	If   LayerShowHide(1) = False Then
 	         Exit Function 
 	End If


		
 	Dim strVal

 	If lgIntFlgMode = parent.OPMD_UMODE Then
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				
 		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)		
 		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
 	Else
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				
 		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)		
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
 	Dim TotDocAmt, dblQty, dblPrice, dblOldQty

 	DbSave = False														
		
					
 	If   LayerShowHide(1) = False Then
 	         Exit Function 
 	End If



 	TotDocAmt = UNICDbl(frm1.txtTotDocAmt.value)
		
 	With frm1
 		.txtMode.value = parent.UID_M0002
 		.txtUpdtUserId.value = parent.gUsrID
 		.txtInsrtUserId.value = parent.gUsrID

 		lGrpCnt = 1

 		strVal = ""
 		intInsrtCnt = 1

 		For lRow = 1 To .vspdData.MaxRows
 			.vspdData.Row = lRow
 			.vspdData.Col = 0

 			Select Case .vspdData.Text
 				Case ggoSpread.InsertFlag								
 					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep	

 					.vspdData.Col = C_Unit								
 					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

 					.vspdData.Col = C_LCQty								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_Price								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
 					dblPrice = UNICDbl(Trim(.vspdData.Text))

 					.vspdData.Col = C_DocAmt							 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
							
 					.vspdData.Col = C_OverTolerance						 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_UnderTolerance					 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_ClsFlg							
						
 			        If Trim(.vspdData.Text) = "1" then	            
 						strVal = strVal & "Y" & parent.gColSep
 			        Else		            
 						strVal = strVal & "N" & parent.gColSep
 			        End If
				        
 					.vspdData.Col = C_SoNo								
 					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

 					.vspdData.Col = C_SoSeq								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_DnNo								
 					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_DnSeq								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gRowSep
 					
 					lGrpCnt = lGrpCnt + 1
 					intInsrtCnt = intInsrtCnt + 1

 					TotDocAmt = UNICDbl(TotDocAmt + (dblQty*dblPrice))

 				Case ggoSpread.UpdateFlag								
 					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	

 					.vspdData.Col = C_LcSeq								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_Unit								
 					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

 					.vspdData.Col = C_LCQty								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_Price								 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
 					dblPrice = UNICDbl(Trim(.vspdData.Text))

 					.vspdData.Col = C_DocAmt							 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_OverTolerance						 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_UnderTolerance					
 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
						
 					.vspdData.Col = C_ClsFlg							
						
 			        If Trim(.vspdData.Text) = "1" then	            
 						strVal = strVal & "Y" & parent.gColSep
 			        Else		            
 						strVal = strVal & "N" & parent.gColSep
 			        End If
						
 					.vspdData.Col = C_SoNo								
 					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

 					.vspdData.Col = C_SoSeq								
 					
 					strVal = strVal & UNICDbl(.vspdData.Text) & parent.gRowSep
						
 					lGrpCnt = lGrpCnt + 1

 				Case ggoSpread.DeleteFlag								
 					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep	

 					.vspdData.Col = C_LcSeq								
 					
 					strDel = strDel & UNICDbl(.vspdData.Text) & parent.gRowSep
						
 					lGrpCnt = lGrpCnt + 1

 					TotDocAmt = UNICDbl(TotDocAmt - (dblOldQty*dblPrice))
 			End Select
 		Next

 		If UNICDbl(frm1.txtDocAmt.value) < UNICDbl(TotDocAmt) Then
 			Call DisplayMsgBox("203523", "x", "x", "x")					
 			Call LayerShowHide(0)
 			Exit Function
 		End If

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
 	lgIntFlgMode = parent.OPMD_UMODE											
 	lgBlnFlgChgValue = False
		
 	Call ggoOper.LockField(Document, "Q")								
 	Call SetToolBar("11101011000111")
		
 	If frm1.vspdData.MaxRows > 0 Then
 		frm1.vspdData.Focus
 	Else
 		frm1.txtLcNo.focus
 	End If
		
 End Function

'========================================================================================================
 Function LCHrdQueryOk()													
 	Call SetToolBar("11101011000011")									
 End Function
	

'========================================================================================================
 Function DbSaveOk()														
 	Call InitVariables
 	frm1.txtLcNo.value = frm1.txtHLCNo.value
 	Call ggoOper.ClearField(Document, "2")	
 	Call FncQuery()							
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
			<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Local L/C내역</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenSODtlRef">수주내역참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNDtlRef">출하내역참조</A></TD>
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
										<TD CLASS="TD5" NOWRAP>LOCAL L/C 관리번호</TD>
										<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLcNo" TYPE="Text" MAXLENGTH=18 TAG="12XXXU" ALT="LOCAL L/C 관리번호" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenLCNoPop"></TD>
										<TD CLASS="TD6" NOWRAP></TD>
										<TD CLASS="TD6" NOWRAP></TD>
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
									<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSONo" ALT="수주번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="24XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설신청인</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="개설신청인">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="영업그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설일</TD>
									<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s3212ma2_fpDateTime2_txtOpenDt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>결제방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="결재방법">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
								</TR>   
								<TR>
									<TD CLASS=TD5 NOWRAP>총개설금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s3212ma2_fpDoubleSingle1_txtDocAmt.js'></script></TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency"SIZE=10 MAXLENGTH=3  TAG="24XXXU" ALT="통화"></TD>
											</TR>
										</TABLE>	
									</TD>
									<TD CLASS=TD5 NOWRAP>총품목금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s3212ma2_fpDoubleSingle1_txtTotItemAmt.js'></script></TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency1"SIZE=10 MAXLENGTH=3  TAG="24XXXU" ALT="통화"></TD>
											</TR>
										</TABLE>	
									</TD>	
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<script language =javascript src='./js/s3212ma2_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">LOCAL L/C등록</A></TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtXchRate" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtTotDocAmt" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
