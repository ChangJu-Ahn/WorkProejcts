<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3222ma2.asp																*
'*  4. Program Name         : Local L/C Amend 내역등록													*
'*  5. Program Desc         : Local L/C Amend 내역등록													*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/24																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : 화면 design												*
'*							  2. 2000/04/03 : Coding Start												*
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        

Dim C_LCAmdFlg		
Dim C_LCAmdFlgDtl		
Dim C_ItemCd							
Dim C_ItemNm			
Dim C_Unit			
Dim C_BeQty			
Dim C_AtQty			
Dim C_BePrice			
Dim C_AtPrice			
Dim C_BeDocAmt		
Dim C_AtDocAmt		
Dim C_BeLocAmt		
Dim C_AtLocAmt		
Dim C_HsCd			
Dim C_OverTolerance		
Dim C_UnderTolerance		
Dim C_LCNoSeq			
Dim C_SoNo			
Dim C_SoSeq			
Dim C_LCAmdSeq		
Dim C_ChgFlg		
Dim C_TrackingNo
Dim C_ItemSpec	

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent					

Const BIZ_PGM_QRY_ID			= "s3222mb2.asp"		
Const BIZ_PGM_SAVE_ID			= "s3222mb2.asp"		
Const BIZ_PGM_CAL_AMT_ID		= "s3211mb6.asp"
Const LCAMEND_HEADER_ENTRY_ID	= "s3221ma2"
	
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_LCAmdFlg			= 1
	C_LCAmdFlgDtl		= 2
	C_ItemCd			= 3
	C_ItemNm			= 4  
    C_ItemSpec			= 5
    C_Unit				= 6
	C_BeQty				= 7  
	C_AtQty				= 8  
	C_BePrice			= 9  
	C_AtPrice			= 10  
	C_BeDocAmt			= 11 
	C_AtDocAmt			= 12 
	C_BeLocAmt			= 13 
	C_AtLocAmt			= 14 
	C_HsCd				= 15 
	C_OverTolerance		= 16 
	C_UnderTolerance	= 17 
	C_LCNoSeq			= 18 
	C_SoNo				= 19 
	C_SoSeq				= 20 
	C_LCAmdSeq			= 21 
	C_ChgFlg			= 22 
	C_TrackingNo		= 23
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
		frm1.txtLocCurrency.value = parent.gCurrency
		frm1.txtTotLocCurrency.value = parent.gCurrency
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
			
			ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
			
			.vspdData.MaxCols = C_TrackingNo
			.vspdData.MaxRows = 0
			
			.vspdData.ReDraw = False
						
			Call GetSpreadColumnPos("A")		
			
			ggoSpread.SSSetEdit		C_LCAmdSeq, "AMEND순번", 10, 0
			ggoSpread.SSSetCombo	C_LCAmdFlg, "변경구분", 10, 2, True 
			ggoSpread.SSSetEdit		C_LCAmdFlgDtl, "변경내용", 10, 0
			ggoSpread.SSSetEdit		C_ItemCd, "품목", 18, 0
			ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25, 0
            ggoSpread.SSSetEdit		C_ItemSpec, "규격", 20
			ggoSpread.SSSetEdit		C_Unit, "단위", 10, 0
	        ggoSpread.SSSetFloat	C_BeQty,"변경전수량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat	C_AtQty,"변경후수량" ,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	        ggoSpread.SSSetFloat	C_BePrice,"변경전단가",15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat	C_AtPrice,"변경후단가",15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BeDocAmt,"변경전금액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_AtDocAmt,"변경후금액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BeLocAmt,"변경전자국금액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_AtLocAmt,"변경후자국금액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_HsCd, "HS부호", 20, 0
			ggoSpread.SSSetEdit		C_LCNoSeq, "L/C순번", 10, 1
			ggoSpread.SSSetEdit		C_SoNo, "수주번호", 18, 0
			ggoSpread.SSSetEdit		C_SoSeq, "수주순번", 10, 1
	        ggoSpread.SSSetFloat	C_OverTolerance,"과부족허용율(+)",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		    ggoSpread.SSSetFloat	C_UnderTolerance,"과부족허용율(-)",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ChgFlg, "Chgfg", 1, 2
			ggoSpread.SetCombo		"U" & vbTab & "D", C_LCAmdFlg
            ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No", 25,,,25,2
			
			SetSpreadLock "", 0, -1, ""
	 
			Call ggoSpread.SSSetColHidden(C_LCAmdSeq,C_LCAmdSeq,True)
			Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
			
			.vspdData.ReDraw = True
			
		End With
	End Sub

'========================================================================================================
	Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
	    With frm1
			ggoSpread.Source = .vspdData
			
			.vspdData.ReDraw = False
			
			ggoSpread.SSSetRequired  C_LCAmdFlg, lRow, lRow
			ggoSpread.SpreadLock C_LCAmdFlgDtl, lRow, -1
			ggoSpread.SpreadLock C_ItemCd, lRow, -1
			ggoSpread.SpreadLock C_ItemNm , lRow, -1
            ggoSpread.SpreadLock C_ItemSpec, lRow, -1
			ggoSpread.SpreadLock C_Unit , lRow, -1
			ggoSpread.SpreadLock C_BeQty , lRow, -1
			ggoSpread.SpreadUnLock C_AtQty, lRow, -1 
			ggoSpread.SSSetRequired  C_AtQty, lRow, lRow
			ggoSpread.SpreadLock C_BePrice, lRow, -1
			ggoSpread.SpreadUnLock C_AtPrice, lRow, -1
			ggoSpread.SSSetRequired  C_AtPrice, lRow, lRow
			ggoSpread.SpreadLock C_BeDocAmt, lRow, -1
			ggoSpread.SpreadUnLock C_AtDocAmt, lRow, -1
			ggoSpread.SSSetRequired  C_AtDocAmt, lRow, lRow
			ggoSpread.SpreadLock C_BeLocAmt, lRow, -1
			ggoSpread.SpreadUnLock C_AtLocAmt, lRow, -1
			ggoSpread.SSSetRequired  C_AtLocAmt, lRow, lRow
			ggoSpread.SpreadLock C_HsCd, lRow, -1
			ggoSpread.SpreadLock C_OverTolerance, lRow, -1
			ggoSpread.SpreadLock C_UnderTolerance, lRow, -1
			ggoSpread.SpreadLock C_LCNoSeq, lRow, -1
			ggoSpread.SpreadLock C_SoNo, lRow, -1
			ggoSpread.SpreadLock C_SoSeq, lRow, -1
            ggoSpread.SpreadLock C_TrackingNo, lRow, -1
			
			.vspdData.ReDraw = True
		End With
	End Sub

'========================================================================================================
	Sub SetSpreadColor(ByVal lRow)
		ggoSpread.Source = frm1.vspdData
		
	    With frm1.vspdData
			ggoSpread.SSSetRequired  C_LCAmdFlg, lRow, .MaxRows
			ggoSpread.SSSetProtected C_LCAmdFlgDtl, lRow, .MaxRows
			ggoSpread.SSSetProtected C_ItemCd, lRow, .MaxRows
			ggoSpread.SSSetProtected C_ItemNm, lRow, .MaxRows
            ggoSpread.SSSetProtected C_ItemSpec, lRow, .MaxRows
			ggoSpread.SSSetProtected C_Unit, lRow, .MaxRows
			ggoSpread.SSSetProtected C_BeQty, lRow, .MaxRows
			ggoSpread.SSSetRequired  C_AtQty, lRow, .MaxRows
			ggoSpread.SSSetProtected C_BePrice, lRow, .MaxRows
			ggoSpread.SSSetRequired  C_AtPrice, lRow, .MaxRows
			ggoSpread.SSSetProtected C_BeDocAmt, lRow, .MaxRows
			ggoSpread.SSSetRequired  C_AtDocAmt, lRow, .MaxRows
			ggoSpread.SSSetProtected C_BeLocAmt, lRow, .MaxRows
			ggoSpread.SSSetRequired  C_AtLocAmt, lRow, .MaxRows
			ggoSpread.SSSetProtected C_HsCd, lRow, .MaxRows
			ggoSpread.SSSetProtected C_LcNoSeq, lRow, .MaxRows
			ggoSpread.SSSetProtected C_SoNo, lRow, .MaxRows
			ggoSpread.SSSetProtected C_SoSeq, lRow, .MaxRows
                        ggoSpread.SSSetProtected C_TrackingNo, lRow, .MaxRows
			ggoSpread.SSSetProtected C_OverTolerance, lRow, .MaxRows
			ggoSpread.SSSetProtected C_UnderTolerance, lRow, .MaxRows
		
		End With
	End Sub

'========================================================================================================
	Sub SetSpreadDeleteRow(ByVal lRow)
	   With frm1
			ggoSpread.Source = .vspdData			
			ggoSpread.SSSetProtected C_AtQty, lRow, lRow
			ggoSpread.SSSetProtected C_AtPrice, lRow, lRow
			ggoSpread.SSSetProtected C_AtDocAmt, lRow, lRow
			ggoSpread.SSSetProtected C_AtLocAmt, lRow, lRow

		End With
	End Sub

'========================================================================================================
	Sub SetReleaseDeleteRow(ByVal lRow)
	   With frm1
			ggoSpread.Source = .vspdData
			
			.vspdData.ReDraw = False

			ggoSpread.SpreadUnLock C_AtQty, lRow, C_AtQty, lRow 

			ggoSpread.SpreadUnLock C_AtQty, lRow, C_AtQty, lRow 
			ggoSpread.SpreadUnLock C_AtPrice, lRow, C_AtPrice, lRow  
			ggoSpread.SpreadUnLock C_AtDocAmt, lRow, C_AtDocAmt, lRow  
			ggoSpread.SpreadUnLock C_AtLocAmt, lRow, C_AtLocAmt, lRow 
			ggoSpread.SSSetRequired C_AtQty, lRow, lRow
			ggoSpread.SSSetRequired C_AtPrice, lRow, lRow
			ggoSpread.SSSetRequired C_AtDocAmt, lRow, lRow
			ggoSpread.SSSetRequired C_AtLocAmt, lRow, lRow
	
			.vspdData.ReDraw = True
		End With
	End Sub		

'========================================================================================================
Sub SetQuerySpreadColor()

    Dim lRow
    With frm1

    .vspdData.ReDraw = False

	ggoSpread.source = frm1.vspdData
	
	
		For lRow = 1 To .vspdData.MaxRows 
			
			ggoSpread.SSSetProtected  C_LCAmdFlg, lRow, lRow
							
			.vspdData.Col = C_LCAmdFlg
			If .vspdData.text = "U" Then
				ggoSpread.SSSetRequired C_AtQty, lRow, lRow	
			Else 
				ggoSpread.SSSetProtected C_AtQty, lRow, lRow
			End if	
									
			.vspdData.Col = C_LCAmdFlg
			If .vspdData.text = "U" Then
				ggoSpread.SSSetRequired C_AtPrice, lRow, lRow
			Else 
				ggoSpread.SSSetProtected C_AtPrice, lRow, lRow
			End if	
									
			.vspdData.Col = C_LCAmdFlg
			If .vspdData.text = "U" Then
				ggoSpread.SSSetRequired C_AtDocAmt, lRow, lRow
			Else 
				ggoSpread.SSSetProtected C_AtDocAmt, lRow, lRow
			End if
			
			.vspdData.Col = C_LCAmdFlg
			If .vspdData.text = "U" Then
				ggoSpread.SSSetRequired C_AtLocAmt, lRow, lRow
			Else 
				ggoSpread.SSSetProtected C_AtLocAmt, lRow, lRow
			End if
			
		Next

    .vspdData.ReDraw = True
    
    End With

End Sub
	
'============================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
   
    Select Case UCase(pvSpdNo)
       Case "A"
            
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_LCAmdFlg			= iCurColumnPos(1) 
			C_LCAmdFlgDtl		= iCurColumnPos(2)
			C_ItemCd			= iCurColumnPos(3)			
			C_ItemNm			= iCurColumnPos(4) 
			C_ItemSpec			= iCurColumnPos(5)	 
			C_Unit				= iCurColumnPos(6)
			C_BeQty				= iCurColumnPos(7) 
			C_AtQty				= iCurColumnPos(8) 
			C_BePrice			= iCurColumnPos(9) 
			C_AtPrice			= iCurColumnPos(10) 
			C_BeDocAmt			= iCurColumnPos(11)
			C_AtDocAmt			= iCurColumnPos(12)
			C_BeLocAmt			= iCurColumnPos(13)
			C_AtLocAmt			= iCurColumnPos(14)
			C_HsCd				= iCurColumnPos(15)
			C_OverTolerance		= iCurColumnPos(16)
			C_UnderTolerance	= iCurColumnPos(17)
			C_LCNoSeq			= iCurColumnPos(18)
			C_SoNo				= iCurColumnPos(19)
			C_SoSeq				= iCurColumnPos(20)
			C_LCAmdSeq			= iCurColumnPos(21)
			C_ChgFlg			= iCurColumnPos(22)
			C_TrackingNo		= iCurColumnPos(23)		

    End Select    
End Sub
'============================================================================================================
	Function OpenLCAmdNoPop()
		Dim strRet
		Dim iCalledAspName
		Dim IntRetCD
		
		If gblnWinEvent = True Or UCase(frm1.txtLCAmdNo.className) = "PROTECTED" Then Exit Function
		
		gblnWinEvent = True
		
		iCalledAspName = AskPRAspName("s3221pa2")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3221pa2", "X")
			gblnWinEvent = False
			Exit Function
		End If
		
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		gblnWinEvent = False
		
		If strRet = "" Then
			Exit Function
		Else
			Call SetLCAmdNo(strRet)
		End If	
	End Function

'============================================================================================================
	Function OpenLCDtlRef()
		Dim arrRet
		Dim arrParam(2)
		Dim strRet
		Dim iCalledAspName
		Dim IntRetCD

		If Trim(frm1.txtLCAmdNo.value) = "" Then
			Call DisplayMsgBox("900002", "x", "x", "x")			
			Exit Function
		End If

		arrParam(0) = Trim(frm1.txtLCNo.value)					
		arrParam(1) = Trim(frm1.txtCurrency.value)
		arrParam(2) = Trim(frm1.txtHLCAmdNo.value) 	
		
		iCalledAspName = AskPRAspName("s3212ra2")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3212ra2", "X")
			gblnWinEvent = False
			Exit Function
		End If
		
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
		
		If arrRet(0, 0) = "" Then
			Exit Function
		Else
			Call SetLCDtlRef(arrRet)
		End If
	End Function


'============================================================================================================
	Function OpenSODtlRef()
		Dim arrRet
		Dim arrParam(10)
		Dim strSONo
		Dim iCalledAspName
		Dim IntRetCD

		If Trim(frm1.txtLCAmdNo.value) = "" Then
			Call DisplayMsgBox("900002", "x", "x", "x")			
			Exit Function
		End If
		
		arrParam(0) = Trim(frm1.txtHSONo.value)					
		arrParam(1) = Trim(frm1.txtApplicant.value)	
		arrParam(2) = Trim(frm1.txtApplicantNm.value)					
		arrParam(3) = Trim(frm1.txtHSalesGroup.value)	
		arrParam(4) = Trim(frm1.txtHSalesGroupNm.value)	
		arrParam(5) = Trim(frm1.txtHPayTerms.value)	
		arrParam(6) = Trim(frm1.txtHPayTermsNm.value)				
		arrParam(7) = Trim(frm1.txtCurrency.value) 
		
		iCalledAspName = AskPRAspName("s3112ra2")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3112ra2", "X")
			gblnWinEvent = False
			Exit Function
		End If
		
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
		If arrRet(0, 0) = "" Then
			Exit Function
		Else
			Call SetSODtlRef(arrRet)
		End If	
	End Function

'============================================================================================================
	Function SetLCAmdNo(strRet)
		frm1.txtLCAmdNo.value = strRet
		frm1.txtLCAmdNo.focus
	End Function

'============================================================================================================
	Function SetLCDtlRef(arrRet)
		Dim intRtnCnt, strData
		Dim TempRow, I, j, AdjustNo
		Dim blnEqualFlg
		Dim intLoopCnt
		Dim intCnt
		Dim dblAmt, dblLocAmt
		Dim strtemp1, strtemp2, strMessage

		AdjustNo = 0
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
						.vspdData.Col = C_LCNoSeq
						
						If .vspdData.Text = arrRet(intCnt - 1, 0) Then
							strtemp1 = .txtLCNo.value
							strtemp2 = .vspdData.text
							strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
							AdjustNo = AdjustNo + 1
							blnEqualFlg = True
							Exit For
						End If
					Next
				End If

				If blnEqualFlg = False Then
					.vspdData.MaxRows = .vspdData.MaxRows + 1
					.vspdData.Row = CLng(TempRow) + CLng(intCnt)

					.vspdData.Col = 0
					.vspdData.Text = ggoSpread.InsertFlag

					.vspdData.Col = C_LCAmdFlg			
					.vspdData.text = "U"
					.vspdData.Col = C_LCAmdFlgDtl		
					.vspdData.text = "내용변경"
					.vspdData.Col = C_ItemCd			
					.vspdData.text = arrRet(intCnt - 1, 0)'배열 순서가 전면수정 arrRet(intCnt - 1, 1)
					.vspdData.Col = C_ItemNm			
					.vspdData.text = arrRet(intCnt - 1, 1)
                    .vspdData.Col = C_ItemSpec			
					.vspdData.text = arrRet(intCnt - 1, 2)
					.vspdData.Col = C_Unit				
					.vspdData.text = arrRet(intCnt - 1, 3)
					.vspdData.Col = C_BeQty				
					.vspdData.text = arrRet(intCnt - 1, 4)
					.vspdData.Col = C_AtQty				
					.vspdData.text = arrRet(intCnt - 1, 4)
										
					.vspdData.Col = C_BePrice			
					.vspdData.text = arrRet(intCnt - 1, 5)
					.vspdData.Col = C_AtPrice			
					.vspdData.text = arrRet(intCnt - 1, 5)
					
					.vspdData.Col = C_BeDocAmt			
					.vspdData.text = arrRet(intCnt - 1, 6)
					.vspdData.Col = C_AtDocAmt			
					.vspdData.text = arrRet(intCnt - 1, 6)		
					
					.vspdData.Col = C_BeLocAmt			
					'.vspdData.text = UNIFormatNumberByCurrecny(UNIConvNum(Trim(frm1.txtXchRate.text),0) * CDBL(arrRet(intCnt - 1, 6)),parent.gCurrency,parent.ggAmtOfMoneyNo)
					.vspdData.text = arrRet(intCnt - 1, 7)		
					
					.vspdData.Col = C_AtLocAmt								
					'.vspdData.text = UNIFormatNumberByCurrecny(UNIConvNum(Trim(frm1.txtXchRate.text),0) * CDBL(arrRet(intCnt - 1, 6)),parent.gCurrency,parent.ggAmtOfMoneyNo)					
					.vspdData.text = arrRet(intCnt - 1, 7)		
									
					.vspdData.Col = C_HsCd				
					.vspdData.text = arrRet(intCnt - 1, 8)
					.vspdData.Col = C_OverTolerance		
					.vspdData.text = arrRet(intCnt - 1, 9)
					.vspdData.Col = C_UnderTolerance						
					.vspdData.text = arrRet(intCnt - 1, 10)
					.vspdData.Col = C_LCNoSeq			
					.vspdData.text = arrRet(intCnt - 1, 11)
					.vspdData.Col = C_SoNo				
					.vspdData.text = arrRet(intCnt - 1, 12)
					.vspdData.Col = C_SoSeq				
					.vspdData.text = arrRet(intCnt - 1, 13)
                    .vspdData.Col = C_TrackingNo
					.vspdData.text = arrRet(intCnt - 1, 14)
					.vspdData.Col = C_ChgFlg			
					.vspdData.text = .vspdData.Row
					
					SetSpreadColor CLng(TempRow) + CLng(intCnt) - CLng(AdjustNo)

					lgBlnFlgChgValue = True
				End If
			Next

			Call SumItemVal() 

			If strMessage <> "" Then
				Call DisplayMsgBox("17a005", "X",strmessage,"L/C번호" & "," & "L/C순번")
				.vspdData.ReDraw = True
				Exit Function
			End If

			.vspdData.ReDraw = True
			
		End With
	End Function

'============================================================================================================
	Function SetSODtlRef(arrRet)
		Dim intRtnCnt, strData
		Dim TempRow, I, j, AdjustNo
		Dim blnEqualFlg
		Dim intLoopCnt
		Dim intCnt
		Dim dblAmt, dblLocAmt
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

						If .vspdData.Text = arrRet(intCnt - 1, 0) Then
							.vspdData.Row = j
							.vspdData.Col = C_SoSeq
							strtemp2 = .vspdData.text

							If .vspdData.Text = arrRet(intCnt - 1, 1) Then
								strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
								AdjustNo = AdjustNo + 1 
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

					.vspdData.Col = C_LCAmdFlg										
					.vspdData.text = "C"
					.vspdData.Col = C_LCAmdFlgDtl										
					.vspdData.text = "품목추가"
					.vspdData.Col = C_ItemCd					
					.vspdData.text = arrRet(intCnt - 1, 2)
					.vspdData.Col = C_ItemNm					
					.vspdData.text = arrRet(intCnt - 1, 3)
                    .vspdData.Col = C_ItemSpec					
					.vspdData.text = arrRet(intCnt - 1, 4)
					.vspdData.Col = C_Unit						
					.vspdData.text = arrRet(intCnt - 1, 5)
					.vspdData.Col = C_BeQty						
					.vspdData.text = 0 
					.vspdData.Col = C_AtQty						
					.vspdData.text = arrRet(intCnt - 1, 6)
					.vspdData.Col = C_BePrice					
					.vspdData.text = 0 
					.vspdData.Col = C_AtPrice					
					.vspdData.text = arrRet(intCnt - 1, 7)
					.vspdData.Col = C_BeDocAmt 
					.vspdData.text = 0 
					.vspdData.Col = C_AtDocAmt					
					.vspdData.text = arrRet(intCnt - 1, 8)
					.vspdData.Col = C_BeLocAmt 
					.vspdData.text = 0 
					.vspdData.Col = C_HsCd						
					.vspdData.text = arrRet(intCnt - 1, 12)
					.vspdData.Col = C_OverTolerance				
					.vspdData.text = arrRet(intCnt - 1, 9)
					.vspdData.Col = C_UnderTolerance			
					.vspdData.text = arrRet(intCnt - 1, 10)
					.vspdData.Col = C_SoNo						
					.vspdData.text = arrRet(intCnt - 1, 0)
					.vspdData.Col = C_SoSeq						
					.vspdData.text = arrRet(intCnt - 1, 1)
                    .vspdData.Col = C_TrackingNo
					.vspdData.text = arrRet(intCnt - 1, 13)
					
					.vspdData.Col = C_ChgFlg					
					.vspdData.text = .vspdData.Row
					
					SetSpreadColor CLng(TempRow) + CLng(intCnt) - CLng(AdjustNo)

					ggoSpread.SSSetProtected C_LCAmdFlg, TempRow, TempRow + intCnt - AdjustNo

					Call vspdData_Change(C_AtQty, TempRow+intCnt-AdjustNo)	
					lgBlnFlgChgValue = True
				End If
			Next

			If strMessage <> "" Then
				Call DisplayMsgBox("17a005", "X",strmessage,"수주번호" & "," & "수주순번")
				.vspdData.ReDraw = True
			End If

			Call SumItemVal() 

			.vspdData.ReDraw = True

		End With
	End Function
	
'============================================================================================================
	Function LocCurrencySet()
		frm1.txtLocCurrency.value = parent.gCurrency
		frm1.txtTotLocCurrency.value = parent.gCurrency 
	End Function

'============================================================================================================
	Function CookiePage(ByVal Kubun)

		On Error Resume Next

		Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
		Dim strTemp, arrVal

		If Kubun = 1 Then

			WriteCookie CookieSplit , frm1.txtLCAmdNo.value

		ElseIf Kubun = 0 Then

			strTemp = ReadCookie(CookieSplit)
				
			If strTemp = "" then Exit Function
				
			frm1.txtLCAmdNo.value =  strTemp
			
			If Err.number <> 0 Then
				Err.Clear
				WriteCookie CookieSplit , ""
				Exit Function 
			End If
			
			Call MainQuery()
						
			WriteCookie CookieSplit , ""
			
		End If

	End Function
'============================================================================================================
	Sub SumLCAmt()
		Dim dblQty
		Dim dblPrice
		Dim dblAmt
		Dim dblTotAmt
		Dim intCnt
		
		With frm1
			ggoSpread.Source = .vspdData

			For intCnt=1 to .vspdData.MaxRows
				
				.vspdData.Col = C_AtDocAmt
				.vspdData.Row = intCnt
				
				If .vspdData.text <>"" Then
					dblTotAmt = dblTotAmt + UNICDbl(.vspdData.text)
				End If	
			Next
			
			.txtTotItemAmt.Text = UNIFormatNumberByCurrecny(dblTotAmt, frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
	
		End With
	End Sub	
	
'============================================================================================================
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
		Call PgmJump(LCAMEND_HEADER_ENTRY_ID)

	End Function
'============================================================================================================
	Function SumItemVal()
		Dim dblDifferece, dblBeAmt, dblAtAmt, i
		Dim dblLocDifferece, dblBeLocAmt, dblAtLocAmt
		dblDifferece = 0
		dblBeAmt = 0
		dblAtAmt = 0

		ggoSpread.Source = frm1.vspdData

		If frm1.vspdData.MaxRows = 0 Then Exit Function
		
		For i=1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = i
			frm1.vspdData.Col = 0

			If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
				frm1.vspdData.Row = i
				frm1.vspdData.Col = C_BeDocAmt
				dblBeAmt = UNICDbl(dblBeAmt) + UNICDbl(frm1.vspdData.Text)

				frm1.vspdData.Row = i
				frm1.vspdData.Col = C_BeLocAmt
				dblBeLocAmt = dblBeLocAmt + UNICDbl(frm1.vspdData.Text)

				frm1.vspdData.Row = i
				frm1.vspddata.Col = C_AtDocAmt
				dblAtAmt = UNICDbl(dblAtAmt) + UNICDbl(frm1.vspdData.Text)

				frm1.vspdData.Row = i
				frm1.vspddata.Col = C_AtLocAmt
				dblAtLocAmt = dblAtLocAmt + UNICDbl(frm1.vspdData.Text)

			End If
		Next
		
		dblDifferece = dblAtAmt - dblBeAmt 
		dblLocDifferece = dblAtLocAmt - dblBeLocAmt 
				
		frm1.txtTotItemAmt.Text = UNIFormatNumberByCurrecny(UNICDbl(frm1.txtHBeDocAmt.value) + dblDifferece,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		frm1.txtTotItemLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtHBeLocAmt.value) + dblLocDifferece, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	
	End Function				

'============================================================================================================
Function CalLocAmt(ByVal Col, ByVal Row)
	Dim dblQty, dblAmt, dblPrice, dblDocAmt
		
	ggoSpread.Source = frm1.vspdData

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_AtQty
	dblQty = frm1.vspdData.Text

	frm1.vspdData.Row = Row
	frm1.vspddata.Col = C_AtPrice
	dblPrice = frm1.vspdData.Text

	dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_AtDocAmt		
	frm1.vspdData.Text = UNIFormatNumberByCurrecny(dblAmt, frm1.txtCurrency.value, parent.ggAmtOfMoneyNo)
		
	'추가 20030814 by inkuk
	dblDocAmt = UNIFormatNumberByCurrecny(dblAmt, frm1.txtCurrency.value, parent.ggAmtOfMoneyNo) 
				
	If frm1.txtCurrency.value = parent.gCurrency Then
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_AtLocAmt
		frm1.vspdData.Text = UNIFormatNumber(dblDocAmt, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	Else
		If frm1.txtExchRateOp.value = "*" Then
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_AtLocAmt
			frm1.vspdData.text = UNIFormatNumber(dblDocAmt * UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit) 
		ElseIf frm1.txtExchRateOp.value = "/" Then 	
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_AtLocAmt
			frm1.vspdData.text = UNIFormatNumber(dblDocAmt / UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit) 
		End If
	End If				
	Call SumItemVal()
End Function

'============================================================================================================
	Sub CurFormatNumericOCX()
		With frm1
			'총개설금액 
			ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
			'총품목금액 
			ggoOper.FormatFieldByObjectOfCur .txtTotItemAmt, .txtTotCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		End With
	End Sub

'============================================================================================================
	Sub CurFormatNumSprSheet()
		With frm1
			ggoSpread.Source = frm1.vspdData
			'변경전단가 
			ggoSpread.SSSetFloatByCellOfCur C_BePrice,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
			'변경후단가 
			ggoSpread.SSSetFloatByCellOfCur C_AtPrice,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
			'변경전금액 
			ggoSpread.SSSetFloatByCellOfCur C_BeDocAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
			'변경후금액 
			ggoSpread.SSSetFloatByCellOfCur C_AtDocAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
			
		End With
	End Sub
	
'============================================================================================================
	Sub Form_Load()
		Call LoadInfTB19029													
		Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
		Call ggoOper.LockField(Document, "N")								
		Call InitSpreadSheet												
		
		Call SetDefaultVal
		Call CookiePage(0)	
		
		Call InitVariables
		Call SetToolBar("11100000000011")									

		If UCase(Trim(frm1.txtLCAmdNo.value)) <> "" Then
			Call MainQuery
		End If

		frm1.txtLCAmdNo.focus
	End Sub
	
'============================================================================================================
	Sub btnLCAmdNoOnClick()
		Call OpenLCAmdNoPop()
	End Sub

'============================================================================================================
	Sub vspdData_Change(ByVal Col, ByVal Row )
		Dim dblQty
		Dim dblPrice
		Dim dblAmt
		Dim dblAtAmt, dblBeAmt 
		Dim dblAtLocAmt, dblBeLocAmt
		Dim dblDifference, dblLocDifference
		Dim dblBeVal
		Dim strVal
		Dim iwhere, i

		ggoSpread.Source = frm1.vspdData

		Select Case Col
			Case C_AtQty, C_AtPrice
				Call CalLocAmt(Col, Row)

			Case C_AtDocAmt, C_AtLocAmt
				Call SumItemVal() 

			Case C_LCAmdFlg
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = Col
				
				iwhere = frm1.vspdData.text 
			
				Select Case iwhere												'AmdFlg 변경시 FlgDtl 변경 
					Case "U"	
						frm1.vspdData.Row = Row
						frm1.vspdData.Col = C_LCAmdFlgDtl
								
						frm1.vspdData.text = "내용변경"
						frm1.vspdData.Col = C_BeQty
						dblBeVal = frm1.vspdData.text 
						
						frm1.vspdData.Col = C_AtQty
						frm1.vspdData.text = dblBeVal
						
						frm1.vspdData.Col = C_BePrice
						dblBeVal = frm1.vspdData.text 
						
						frm1.vspdData.Col = C_AtPrice
						frm1.vspdData.text = dblBeVal

						frm1.vspdData.Col = C_BeDocAmt
						dblBeVal = frm1.vspdData.text 
						
						frm1.vspdData.Col = C_AtDocAmt
						frm1.vspdData.text = dblBeVal

						frm1.vspdData.Col = C_BeLocAmt
						dblBeVal = frm1.vspdData.text 
						
						frm1.vspdData.Col = C_AtLocAmt
						frm1.vspdData.text = dblBeVal

						Call SetReleaseDeleteRow(Row)	
						Call SumItemVal() 
					Case "D"
						frm1.vspdData.Row = Row
						frm1.vspdData.Col = C_LCAmdFlgDtl
						frm1.vspdData.text = "품목삭제"
						
						frm1.vspdData.Col = C_AtQty
						frm1.vspdData.text = 0
						
						frm1.vspdData.Col = C_AtPrice
						frm1.vspdData.text = 0
						
						frm1.vspdData.Col = C_AtDocAmt
						frm1.vspdData.text = 0
						
						frm1.vspdData.Col = C_AtLocAmt
						frm1.vspdData.text = 0
						
						Call SetSpreadDeleteRow(Row)
						Call SumItemVal() 
					Case Else
						frm1.vspdData.Row = Row
						frm1.vspdData.Col = C_LCAmdFlg
						frm1.vspdData.text = ""
		
						frm1.vspdData.Row = Row
						frm1.vspdData.Col = C_LCAmdFlgDtl
						frm1.vspdData.text = ""
				End Select
			
			Case Else

		End Select

		ggoSpread.UpdateRow Row

		lgBlnFlgChgValue = True
	End Sub
	
'============================================================================================================
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
	
'============================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'============================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'============================================================================================================
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

'============================================================================================================
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

'============================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'============================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
        
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'============================================================================================================
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
		Call SetDefaultVal
		Call InitVariables										

		If Not chkField(Document, "1") Then						
			Exit Function
		End If

		Call DbQuery()											

		FncQuery = True											
				
	End Function
	
'============================================================================================================
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
		Call InitVariables										
		Call SetToolBar("11100000000011")						
		Call SetDefaultVal

		FncNew = True											

	End Function
	
'============================================================================================================
	Function FncDelete()
		Dim IntRetCD

		FncDelete = False										
				
		If lgIntFlgMode <> parent.OPMD_UMODE Then						
			Call DisplayMsgBox("900002", "x", "x", "x")
			Exit Function
		End If

		IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

		If IntRetCD = vbNo Then
			Exit Function
		End If

		Call DbDelete											

		FncDelete = True										
	End Function
	
'============================================================================================================
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
		If ggoSpread.SSDefaultCheck = False Then
			Exit Function
		End If
		
		Call DbSave							
		
		FncSave = True						
	End Function

'============================================================================================================
	Function FncCopy()
		frm1.vspdData.ReDraw = False

		ggoSpread.Source = frm1.vspdData	
		If frm1.vspdData.MaxRows < 1 Then Exit Function
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow

		frm1.vspdData.ReDraw = True
	End Function

'============================================================================================================
Function FncCancel() 
    Dim iDx

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncCancel = False                                                             

    ggoSpread.Source = Frm1.vspdData
    If frm1.vspdData.MaxRows < 1 Then Exit Function	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Call SumItemVal()
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'============================================================================================================
	Function FncInsertRow()
	End Function
'============================================================================================================
	Function FncDeleteRow()
		Dim lDelRows
		Dim iDelRowCnt, i
	
		With frm1.vspdData 
	
			.focus
			ggoSpread.Source = frm1.vspdData

			lDelRows = ggoSpread.DeleteRow

			Call SumItemVal() 

			lgBlnFlgChgValue = True
		End With
	End Function

'============================================================================================================
	Function FncPrint()
		Call parent.FncPrint()
	End Function

'============================================================================================================
	Function FncPrev() 
		On Error Resume Next				
	End Function

'============================================================================================================
	Function FncNext()
		On Error Resume Next				
	End Function

'============================================================================================================
	Function FncExcel() 
		Call parent.FncExport(parent.C_SINGLEMULTI)
	End Function

'============================================================================================================
	Function FncFind() 
		Call parent.FncFind(parent.C_SINGLEMULTI, False)
	End Function

'============================================================================================================
	Function FncExit()
		Dim IntRetCD

		FncExit = False
		
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		FncExit = True
	End Function
'============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'============================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'============================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor()
End Sub

'============================================================================================================
	Function DbQuery()
		Err.Clear													

		DbQuery = False												

		Dim strVal

					
		If   LayerShowHide(1) = False Then
		         Exit Function 
		End If



		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001		
			strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtHLCAmdNo.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				
			strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo.value)	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)									
	
		DbQuery = True														
	End Function
	
'============================================================================================================
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

						.vspdData.Col = C_LCAmdFlg							
						strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
						
						.vspdData.Col = C_LCAmdSeq							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
						
						.vspdData.Col = C_AtQty								
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
					
						.vspdData.Col = C_AtPrice							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
						
						.vspdData.Col = C_AtDocAmt							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
										
						.vspdData.Col = C_LCNoSeq							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
						
						.vspdData.Col = C_SoNo								
						strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

						.vspdData.Col = C_SoSeq								
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep

						.vspdData.Col = C_AtLocAmt							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep 
						
						lGrpCnt = lGrpCnt + 1
						intInsrtCnt = intInsrtCnt + 1

						TotDocAmt = UNICDbl(TotDocAmt + (dblQty*dblPrice))

					Case ggoSpread.UpdateFlag								
						strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	

						.vspdData.Col = C_LCAmdFlg							
						strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
						
						.vspdData.Col = C_LCAmdSeq							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
						
						.vspdData.Col = C_AtQty								
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
	
						.vspdData.Col = C_AtPrice							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
				
						.vspdData.Col = C_AtDocAmt							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
												
						.vspdData.Col = C_LCNoSeq							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
						
						.vspdData.Col = C_SoNo								
						strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

						.vspdData.Col = C_SoSeq								
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep

						.vspdData.Col = C_AtLocAmt							
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep 
						
						lGrpCnt = lGrpCnt + 1

					Case ggoSpread.DeleteFlag								
						strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep	

						.vspdData.Col = C_LCAmdSeq							
						strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep

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
		lgIntFlgMode = parent.OPMD_UMODE											
		lgBlnFlgChgValue = False
		
		Call ggoOper.LockField(Document, "Q")		
		Call SetToolBar("11101011000111")			
	
		With frm1
			ggoSpread.Source = .vspdData
		
			Dim TempRow, i
		
			TempRow = .vspdData.MaxRows				
			
			For i = 1 to TempRow 
				.vspdData.Row = i
				.vspdData.Col = 0
				
				if .vspdData.Text <> ggoSpread.InsertFlag then
					ggoSpread.SSSetProtected C_LCAmdFlg, i, i	
				End if
				
				'.vspdData.Row = i
				.vspdData.Col = C_LCAmdFlg
				
				If .vspdData.text = "D" Then
					Call SetSpreadDeleteRow(i) 
				End If
			Next
		End With
		
		If frm1.vspdData.MaxRows > 0 Then
			frm1.vspdData.Focus
		Else
			frm1.txtLCAmdNo.focus
		End If
		
	End Function

'========================================================================================================
	Function LCAmendQueryOk()						
		Call LocCurrencySet()
		Call SetToolBar("111010110000111")			
	End Function
	
	
'========================================================================================================
	Function DbSaveOk()								
		Call InitVariables
		frm1.txtLCAmdNo.value = frm1.txtHLCAmdNo.value   
		Call ggoOper.ClearField(Document, "2")					
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
			<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
		</TR>
		<TR HEIGHT=23>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_10%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSLTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>Local L/C Amend 내역정보</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenLCDtlRef">LOCAL L/C내역참조</A>&nbsp;|&nbsp;<A><A href="vbscript:OpenSODtlRef">수주내역참조</A></TD>
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
										<TD CLASS=TD5 NOWRAP>LOCAL L/C AMEND관리번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="LOCAL L/C AMEND관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCAmdNoOnClick()"></TD>
										<TD CLASS=TD6 NOWRAP></TD>
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
									<TD CLASS=TD5 NOWRAP>LOCAL L/C 번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C 번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" ALT="LOCAL L/C관리번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="24XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설신청인</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" ALT="개설신청인" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3222ma2_fpDoubleSingle3_txtXchRate.js'></script>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총개설금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s3222ma2_fpDoubleSingle1_txtDocAmt.js'></script></TD>
												<TD>&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>총품목금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s3222ma2_fpDoubleSingle1_txtTotItemAmt.js'></script></TD>
												<TD>&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtTotCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
											</TR>
										</TABLE>	
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총개설원화금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s3222ma2_fpDoubleSingle1_txtLocAmt.js'></script></TD>
												<TD>&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="자국화폐"></TD>
											</TR>
										</TABLE>	
									</TD>
									<TD CLASS=TD5 NOWRAP>총품목원화금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/s3222ma2_fpDoubleSingle1_txtTotItemLocAmt.js'></script></TD>
												<TD>&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtTotLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="자국화폐"></TD>
											</TR>
										</TABLE>	
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<script language =javascript src='./js/s3222ma2_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">L/C AMEND등록</A></TD>
					<!--<TD WIDTH=50>&nbsp;</TD>-->
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHOpenDt" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHLCAmdNo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroupNm" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPayTermsNm" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHXchRate" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHTotDocAmt" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtExchRateOp" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHBeDocAmt" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHBeLocAmt" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
