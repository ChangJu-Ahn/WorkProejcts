<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4212ma1.asp																*
'*  4. Program Name         : 통관내역등록																*
'*  5. Program Desc         : 통관내역등록																*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2000/04/11																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                       

Dim C_ItemCd		
Dim C_ItemNm
Dim C_Unit			
Dim C_Qty			
Dim C_Price			
Dim C_DocAmt		
Dim C_NetWeight	
Dim C_PackingQty	
Dim C_HsCd			
Dim C_HsPopup       
Dim C_LanNo			
Dim C_Plant			
Dim C_DNNo		
Dim C_DNSeq			
Dim C_SoNo			
Dim C_SoSeq			
Dim C_SOISeq		
Dim C_LCNo			
Dim C_LCDocNo		
Dim C_LCSeq		
Dim C_MvmtNo	
Dim C_PONo			
Dim C_POSeq			
Dim C_CCSeq	
Dim C_TrackingNo									
Dim C_Spec		
Dim C_ChgFlg			

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent					
Dim IsOpenPop

Const BIZ_PGM_QRY_ID = "s4212mb1.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "s4212mb1.asp"			'☆: 비지니스 로직 ASP명 
Const EXCC_HEADER_ENTRY_ID = "s4211ma1"			'☆: 이동할 ASP명 
Const EXCC_LAN_ENTRY_ID = "s4213ma1"			'☆: 이동할 ASP명 
Const EXCC_ASSIGN_ENTRY_ID ="s4214ma1"		'☆: 이동할 ASP명 : container 배정 
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_ItemCd		= 1								
	C_ItemNm		= 2
	C_Unit			= 3
	C_Qty			= 4
	C_Price			= 5
	C_DocAmt		= 6
	C_NetWeight		= 7
	C_PackingQty	=8
	C_HsCd			= 9
	C_HsPopup       = 10
	C_LanNo			= 11
	C_Plant			= 12
	C_DNNo			= 13
	C_DNSeq			= 14
	C_SoNo			= 15
	C_SoSeq			= 16
	C_SOISeq		= 17
	C_LCNo			= 18
	C_LCDocNo		= 19
	C_LCSeq			= 20
	C_MvmtNo		= 21	
	C_PONo			= 22
	C_POSeq			= 23
	C_CCSeq			= 24								
	C_TrackingNo	= 25									
	C_Spec			= 26	
	C_ChgFlg		= 27
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
	frm1.txtDocAmt.Text = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.txtNetWeight.Text = UNIFormatNumber(0, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
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
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
			
		.vspdData.ReDraw = False
			
		.vspdData.MaxCols = C_ChgFlg
		.vspdData.MaxRows = 0
			
		Call GetSpreadColumnPos("A")	
			
		ggoSpread.SSSetEdit		C_CCSeq, "통관순번", 10, 0			
		ggoSpread.SSSetEdit		C_ItemCd, "품목", 18, 0
		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25,,,40
		ggoSpread.SSSetEdit		C_Unit, "단위", 10, 0
        ggoSpread.SSSetFloat	C_Qty,"통관수량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat	C_Price,"단가",15,Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_DocAmt,"금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat	C_NetWeight,"순중량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat	C_PackingQty,"Packing수량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_HsCd, "HS부호", 20, 0
		ggoSpread.SSSetButton	C_HsPopup
		ggoSpread.SSSetEdit		C_LanNo, "란번호", 10, 0
		ggoSpread.SSSetEdit		C_Plant, "공장", 10, 0
		ggoSpread.SSSetEdit		C_DNNo, "출하번호", 18, 0
		ggoSpread.SSSetEdit		C_DNSeq, "출하순번", 10, 1
		ggoSpread.SSSetEdit		C_SoNo, "수주번호", 18, 0
		ggoSpread.SSSetEdit		C_SoSeq, "수주순번", 10, 1
		ggoSpread.SSSetEdit		C_SOISeq, "수주일정순번", 15, 1
		ggoSpread.SSSetEdit		C_LCNo, "L/C관리번호", 18, 0
		ggoSpread.SSSetEdit		C_LCDocNo, "L/C번호", 18, 0
		ggoSpread.SSSetEdit		C_LcSeq, "L/C순번", 10, 1
		ggoSpread.SSSetEdit		C_MvmtNo, "외주출고번호", 18, 0
		ggoSpread.SSSetEdit		C_PONo, "발주번호", 18, 0
		ggoSpread.SSSetEdit     C_POSeq, "발주순번", 10, 1					
		ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 18, 0					
		ggoSpread.SSSetEdit		C_Spec, "규격", 20,,,50
		
		SetSpreadLock "", 0, -1, ""

		Call ggoSpread.SSSetColHidden(C_CCSeq,C_CCSeq,True)
		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
		Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,True)
		Call ggoSpread.SSSetColHidden(C_HsPopup,C_HsPopup,True)
			
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
			
		ggoSpread.SpreadLock C_CCSeq, lRow, -1			
		ggoSpread.SpreadLock C_ItemCd, lRow, -1  
		ggoSpread.SpreadLock C_ItemNm, lRow, -1
		ggoSpread.SpreadLock C_Unit, lRow, -1
		ggoSpread.SpreadUnLock C_Qty, lRow, -1
		ggoSpread.SSSetRequired C_Qty, lRow, -1
		ggoSpread.SpreadLock C_Price, lRow, -1
		ggoSpread.SpreadLock C_DocAmt, lRow, -1
		ggoSpread.SpreadLock C_NetWeight, lRow, -1
		ggoSpread.SpreadLock C_PackingQty, lRow, -1
		ggoSpread.SpreadLock C_HsCd, lRow, -1
		ggoSpread.SpreadLock C_HsPopup, lRow, -1
		ggoSpread.SpreadLock C_LanNo, lRow, -1
		ggoSpread.SpreadLock C_Plant, lRow, -1
		ggoSpread.SpreadLock C_DNNo, lRow, -1
		ggoSpread.SpreadLock C_DNSeq, lRow, -1 
		ggoSpread.SpreadLock C_SoNo, lRow, -1
		ggoSpread.SpreadLock C_SoSeq, lRow, -1
		ggoSpread.SpreadLock C_SOISeq, lRow, -1
		ggoSpread.SpreadLock C_LCNo, lRow, -1
		ggoSpread.SpreadLock C_LCDocNo, lRow, -1
		ggoSpread.SpreadLock C_LCSeq, lRow, -1
		ggoSpread.SpreadLock C_MvmtNo, lRow, -1
		ggoSpread.SpreadLock C_PONo, lRow, -1
		ggoSpread.SpreadLock C_POSeq, lRow, -1				
		ggoSpread.SpreadLock C_Spec, lRow, -1
		ggoSpread.SpreadLock C_TrackingNo, lRow, -1			
		
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadColor(ByVal lRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData

		If UCase(Trim(frm1.txtRefFlg.value)) = "M" Then
			ggoSpread.SpreadUnLock C_Price, lRow, -1
			ggoSpread.SSSetRequired C_Price, lRow, lRow

			ggoSpread.SpreadUnLock C_DocAmt, lRow, -1
			ggoSpread.SSSetRequired C_DocAmt, lRow, lRow

			ggoSpread.SpreadUnLock C_HsCd, lRow, -1
			ggoSpread.SSSetRequired C_HsCd, lRow, lRow
			ggoSpread.SpreadUnLock C_HsPopup, lRow, -1
		Else
			ggoSpread.SSSetProtected C_Price, lRow, lRow
			ggoSpread.SSSetProtected C_DocAmt, lRow, lRow			
			ggoSpread.SSSetProtected C_HsCd, lRow, lRow
			ggoSpread.SSSetProtected C_HsPopup, lRow, lRow
		End If


		ggoSpread.SSSetProtected C_CCSeq, lRow, lRow
		ggoSpread.SSSetProtected C_ItemCd, lRow, lRow
		ggoSpread.SSSetProtected C_Unit, lRow, lRow
		ggoSpread.SSSetRequired C_Qty, lRow, lRow
		ggoSpread.SSSetProtected C_LanNo, lRow, lRow
		ggoSpread.SSSetProtected C_Plant, lRow, lRow
		ggoSpread.SSSetProtected C_DNNo, lRow, lRow
		ggoSpread.SSSetProtected C_DNSeq, lRow, lRow
		ggoSpread.SSSetProtected C_SoNo, lRow, lRow
		ggoSpread.SSSetProtected C_SoSeq, lRow, lRow  
		ggoSpread.SSSetProtected C_SOISeq, lRow, lRow
		ggoSpread.SSSetProtected C_LCNo, lRow, lRow
		ggoSpread.SSSetProtected C_LCDocNo, lRow, lRow
		ggoSpread.SSSetProtected C_LCSeq, lRow, lRow
		ggoSpread.SSSetProtected C_MvmtNo, lRow, lRow	 
		ggoSpread.SSSetProtected C_PONo, lRow, lRow	
		ggoSpread.SSSetProtected C_POSeq, lRow, lRow   
		ggoSpread.SSSetProtected C_Spec, lRow, lRow
		ggoSpread.SSSetProtected C_TrackingNo, lRow, lRow  
		ggoSpread.SSSetProtected C_PackingQty, lRow, lRow  
		ggoSpread.SSSetProtected C_NetWeight, lRow, lRow  
		
		.Col = 1
		.Row = .ActiveRow
		.Action = 0
		.EditMode = True

	End With
End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
   
    Select Case UCase(pvSpdNo)
       Case "A"
            
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ItemCd		= iCurColumnPos(1)  
			C_ItemNm		= iCurColumnPos(2)  
			C_Unit			= iCurColumnPos(3)  
			C_Qty			= iCurColumnPos(4)  
			C_Price			= iCurColumnPos(5)  
			C_DocAmt		= iCurColumnPos(6)  
			C_NetWeight		= iCurColumnPos(7)
			C_PackingQty		= iCurColumnPos(8)    
			C_HsCd			= iCurColumnPos(9)  
			C_HsPopup		= iCurColumnPos(10)  
			C_LanNo			= iCurColumnPos(11) 
			C_Plant			= iCurColumnPos(12) 
			C_DNNo			= iCurColumnPos(13) 
			C_DNSeq			= iCurColumnPos(14) 
			C_SoNo			= iCurColumnPos(15) 
			C_SoSeq			= iCurColumnPos(16) 
			C_SOISeq		= iCurColumnPos(17) 
			C_LCNo			= iCurColumnPos(18) 
			C_LCDocNo		= iCurColumnPos(19) 
			C_LCSeq			= iCurColumnPos(20) 
			C_MvmtNo		= iCurColumnPos(21)
			C_PONo			= iCurColumnPos(22) 
			C_POSeq			= iCurColumnPos(23) 
			C_CCSeq			= iCurColumnPos(24)
			C_TrackingNo	= iCurColumnPos(25)			
			C_Spec			= iCurColumnPos(26)  			
			C_ChgFlg		= iCurColumnPos(27) 

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
Function OpenSODtlRef()
	Dim arrRet
	Dim arrParam(10)
	Dim iCalledAspName
	Dim IntRetCD		

	If Trim(frm1.txtCurrency.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If RefCheckMessage("S") = False Then Exit Function
		
	iCalledAspName = AskPRAspName("S3112RA6")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3112RA6", "X")			
		gblnWinEvent = False
		Exit Function
	End If
			

	arrParam(0) = Trim(frm1.txtApplicant.value)	
	arrParam(1) = Trim(frm1.txtApplicantNm.value)					
	arrParam(2) = Trim(frm1.txtSONo.value)					
	arrParam(3) = Trim(frm1.txtCurrency.value) 
	arrParam(4) = Trim(frm1.txtSalesGroup.value)	
	arrParam(5) = Trim(frm1.txtSalesGroupNm.value)			
	arrParam(6) = Trim(frm1.txtPayTerms.value)	
	arrParam(7) = Trim(frm1.txtPayTermsNm.value)				
	arrParam(8) = Trim(frm1.txtIncoTerms.value)	
	arrParam(9) = Trim(frm1.txtIncoTermsNm.value)									
		
	arrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else
		Call SetSODtlRef(arrRet)
	End If	
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenLCDtlRef()
	Dim strRet
	Dim arrParam(10)
	Dim iCalledAspName
	Dim IntRetCD		
		
	If Trim(frm1.txtCurrency.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
		
	If RefCheckMessage("L") = False Then Exit Function
		
	iCalledAspName = AskPRAspName("S3212RA6")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3212RA6", "X")			
		gblnWinEvent = False
		Exit Function
	End If
		
	arrParam(0) = Trim(frm1.txtApplicant.value)	
	arrParam(1) = Trim(frm1.txtApplicantNm.value)					
	arrParam(2) = Trim(frm1.txtSONo.value)					
	arrParam(3) = Trim(frm1.txtCurrency.value) 
	arrParam(4) = Trim(frm1.txtSalesGroup.value)	
	arrParam(5) = Trim(frm1.txtSalesGroupNm.value)			
	arrParam(6) = Trim(frm1.txtPayTerms.value)	
	arrParam(7) = Trim(frm1.txtPayTermsNm.value)				
	arrParam(8) = Trim(frm1.txtIncoTerms.value)	
	arrParam(9) = Trim(frm1.txtIncoTermsNm.value)
		
	strRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	If strRet(0, 0) = "" Then
		Exit Function
	Else
		Call SetLCDtlRef(strRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenDNDtlRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD		
		
	arrParam(0) = UCase(Trim(frm1.txtApplicant.value))
	arrParam(1) = Trim(frm1.txtApplicantNm.value)

	If Trim(frm1.txtCurrency.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
		
	If RefCheckMessage("M") = False Then Exit Function
	iCalledAspName = AskPRAspName("M4111RA7")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111RA7", "X")			
		gblnWinEvent = False
		Exit Function
	End If
				

	arrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	If arrRet(0, 0) = "" Then
		Exit Function
	Else
		Call SetDNDtlRef(arrRet)
	End If
End Function
'===========================================================================
Function OpenHSCd(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HS부호"
	arrParam(1) = "B_HS_CODE"					
	arrParam(2) = strCode						
	arrParam(3) = ""							
	arrParam(4) = ""							
	arrParam(5) = "HS부호"					
	
	arrField(0) = "HS_CD"						
	arrField(1) = "HS_NM"						

	arrHeader(0) = "HS부호"					
	arrHeader(1) = "HS부호명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetHSCd(arrRet)
	End If	
	
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetExCCNo(strRet)
	frm1.txtCCNo.value = strRet(0)
	frm1.txtCCNo.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
						
					If .vspdData.Text = arrRet(intCnt - 1, 0) Then
						.vspdData.Row = j
						.vspdData.Col = C_SoSeq
						strtemp2 = .vspdData.text
						
						If .vspdData.Text = arrRet(intCnt - 1, 1) Then
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
				.vspdData.Col = C_SOISeq											
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_ItemCd										
				.vspdData.text = arrRet(intCnt - 1, 3)
				.vspdData.Col = C_ItemNm										
				.vspdData.text = arrRet(intCnt - 1, 4)
				.vspdData.Col = C_Spec										
				.vspdData.text = arrRet(intCnt - 1, 5)
				.vspdData.Col = C_Unit											
				.vspdData.text = arrRet(intCnt - 1, 6)
				.vspdData.Col = C_Qty											
				.vspdData.text = arrRet(intCnt - 1, 7)
				.vspdData.Col = C_Price											
				.vspdData.text = arrRet(intCnt - 1, 8)
				.vspdData.Col = C_DocAmt										
				.vspdData.text = arrRet(intCnt - 1, 9)
				.vspdData.Col = C_Plant											
				.vspdData.text = arrRet(intCnt - 1, 10)
				.vspdData.Col = C_DNNo											
				.vspdData.text = arrRet(intCnt - 1, 11)
				.vspdData.Col = C_DNSeq											
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_HsCd								
				.vspdData.text = arrRet(intCnt - 1, 13)
				.vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 14)				
				
				.vspdData.Col = C_ChgFlg							
				.vspdData.text = .vspdData.Row
					


				SetSpreadColor CLng(TempRow) + CLng(intCnt)
					
				lgBlnFlgChgValue = True
			End If
		Next

		If strMessage <> "" Then
			Call DisplayMsgBox("17a005", "X",strmessage,"수주번호" & "," & "수주순번")
			.vspdData.ReDraw = True
		End If

		.vspdData.ReDraw = True

	End With
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetLCDtlRef(arrRet)
		
	Dim intRtnCnt, strData
	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim dblAmt
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
					.vspdData.Col = C_LCNo
					strtemp1 = .vspdData.text

					If .vspdData.Text = arrRet(intCnt - 1, 0) Then
						
						.vspdData.Row = j
						.vspdData.Col = C_LCSeq
						strtemp2 = .vspdData.text

						If .vspdData.Text = arrRet(intCnt - 1, 1) Then
							strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
							blnEqualFlg = True
							Exit For
						End If
					End If
				Next
			End If

			If blnEqualFlg = False Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
'					.vspdData.MaxRows = CLng(TempRow) + CLng(intCnt)
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)					
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag					
				.vspdData.Col = C_LCNo											
				.vspdData.text = arrRet(intCnt - 1, 0)				
				.vspdData.Col = C_LCSeq																
				.vspdData.text = arrRet(intCnt - 1, 1)
				.vspdData.Col = C_SoNo											
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_SoSeq	
				.vspdData.text = arrRet(intCnt - 1, 3)
				.vspdData.Col = C_SOISeq	
				.vspdData.text = arrRet(intCnt - 1, 4)
				.vspdData.Col = C_ItemCd										
				.vspdData.text = arrRet(intCnt - 1, 5)
				.vspdData.Col = C_ItemNm										
				.vspdData.text = arrRet(intCnt - 1, 6)
				.vspdData.Col = C_Spec										
				.vspdData.text = arrRet(intCnt - 1, 7)
				.vspdData.Col = C_Unit											
				.vspdData.text = arrRet(intCnt - 1, 8)
				.vspdData.Col = C_Qty	
				.vspdData.text = arrRet(intCnt - 1, 9)										
				.vspdData.Col = C_Price											
				.vspdData.text = arrRet(intCnt - 1, 10)
				.vspdData.Col = C_DocAmt										
				.vspdData.text = arrRet(intCnt - 1, 11)
				.vspdData.Col = C_Plant										
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_LCDocNo									
				.vspdData.text = arrRet(intCnt - 1, 13)
				.vspdData.Col = C_DNNo										
				.vspdData.text = arrRet(intCnt - 1, 14)
				.vspdData.Col = C_DNSeq										
				.vspdData.text = arrRet(intCnt - 1, 15)
				.vspdData.Col = C_HsCd										
				.vspdData.text = arrRet(intCnt - 1, 16)
				.vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 17)
				.vspdData.Col = C_ChgFlg									
				.vspdData.text = .vspdData.Row
					
				SetSpreadColor CLng(TempRow) + CLng(intCnt)

				lgBlnFlgChgValue = True
			End If
		Next

		If strMessage <> "" Then
			Call DisplayMsgBox("17a005", "X",strmessage,"L/C번호" & "," & "L/C순번")
			.vspdData.ReDraw = True
		End If

		.vspdData.ReDraw = True

	End With
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetDNDtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim strDnNo
	Dim strDnSeq
	Dim strtemp1, strtemp2, strtemp3, strMessage
	Dim intPrice, intQty

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
					.vspdData.Col = C_PoNo
					strtemp1 = .vspdData.text

					If .vspdData.Text = arrRet(intCnt - 1, 7) Then
						
						.vspdData.Row = j
						.vspdData.Col = C_POSeq
						strtemp2 = .vspdData.text

						If .vspdData.Text = arrRet(intCnt - 1, 8) Then


							.vspdData.Row = j
							.vspdData.Col = C_MvmtNo
							strtemp3 = .vspdData.text

							If .vspdData.Text = arrRet(intCnt - 1, 0) Then

								strMessage = strMessage & strtemp1 & "-" & strtemp2 & "-" & strtemp3 & vbCrlf
								blnEqualFlg = True
								Exit For

							End If

						End If

					End If
				Next
			End If

			If blnEqualFlg = False Then

				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag
				.vspdData.Col = C_MvmtNo
				.vspdData.text = arrRet(intCnt - 1, 0)
				.vspdData.Col = C_ItemCd										
				.vspdData.text = arrRet(intCnt - 1, 1)
				.vspdData.Col = C_ItemNm										
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_Spec										
				.vspdData.text = arrRet(intCnt - 1, 3)
				.vspdData.Col = C_Unit											
				.vspdData.text = arrRet(intCnt - 1, 4)
				.vspdData.Col = C_Qty											
				.vspdData.text = arrRet(intCnt - 1, 5)	:	intQty = Trim(.vspdData.text)				
				.vspdData.Col = C_Plant											
				.vspdData.text = arrRet(intCnt - 1, 6)
				.vspdData.Col = C_PONo											
				.vspdData.text = arrRet(intCnt - 1, 8)
				.vspdData.Col = C_POSeq
				.vspdData.text = arrRet(intCnt - 1, 9)
				.vspdData.Col = C_HsCd											
				.vspdData.text = arrRet(intCnt - 1, 10)
				.vspdData.Col = C_ChgFlg
				.vspdData.text = .vspdData.Row

				'# 외주출고참조후 단가를 통관의 화폐와 환율연산자 맞추어 변환	#
				.vspdData.Col = C_Price
				
				.vspdData.text = UNIFormatNumberByCurrecny(arrRet(intCnt - 1, 11),frm1.txtCurrency.value,Parent.ggUnitCostNo)

				intPrice = .vspdData.text

				If Trim(frm1.txtHXchRateOp.value) = "*" then 				
					.vspdData.text = UNIFormatNumberByCurrecny(UNICDbl(intPrice) * UNICDbl(frm1.txtHXchRate.value),frm1.txtCurrency.value,Parent.ggUnitCostNo)
				Else				
					.vspdData.text = UNIFormatNumberByCurrecny(UNICDbl(intPrice) / UNICDbl(frm1.txtHXchRate.value),frm1.txtCurrency.value,Parent.ggUnitCostNo)
				End If  

				intPrice = .vspdData.text

				.vspdData.Col = C_DocAmt
				
				.vspdData.text = UNIFormatNumberByCurrecny(UNICDbl(intPrice) * UNICDbl(intQty),frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)

				SetSpreadColor CLng(TempRow) + CLng(intCnt)

				lgBlnFlgChgValue = True
			End If
		Next

		Call SumAmt()

		If strMessage <> "" Then
			Call DisplayMsgBox("17a005", "X",strmessage,"발주번호" & "," & "발주순번" & "," & "외주출고번호")
			.vspdData.ReDraw = True
		End If

		.vspdData.ReDraw = True

	End With

End Function
'---------------------------------------------------------------------------------------------------------
Function SetHSCd(Byval arrRet)
	With frm1
		.vspdData.Col = C_HsCd
		.vspdData.Text = arrRet(0)

		ggoSpread.Source = .vspdData
		Call vspdData_Change(C_HsCd, .vspdData.ActiveRow )
	End With
End Function
'========================================================================================================
Sub HideNonRelGrid()
	Dim RefFlg
		
	With frm1
		RefFlg = .txtRefFlg.value 
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
			
		Select Case RefFlg
			Case "S"
				Call ggoSpread.SSSetColHidden(C_LCNo,C_LCNo,True)
				Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,True)
				Call ggoSpread.SSSetColHidden(C_LCSeq,C_LCSeq,True)
				Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)
				Call ggoSpread.SSSetColHidden(C_PONo,C_PONo,True)
				Call ggoSpread.SSSetColHidden(C_POSeq,C_POSeq,True)
			Case "L"
				Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)
				Call ggoSpread.SSSetColHidden(C_PONo,C_PONo,True)
				Call ggoSpread.SSSetColHidden(C_POSeq,C_POSeq,True)			

			Case "M"
				Call ggoSpread.SSSetColHidden(C_LCNo,C_LCNo,True)
				Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,True)
				Call ggoSpread.SSSetColHidden(C_LCSeq,C_LCSeq,True)
				Call ggoSpread.SSSetColHidden(C_SONo,C_SONo,True)
				Call ggoSpread.SSSetColHidden(C_SoSeq,C_SoSeq,True)
				Call ggoSpread.SSSetColHidden(C_SOISeq,C_SOISeq,True)
				Call ggoSpread.SSSetColHidden(C_DNNo,C_DNNo,True)
				Call ggoSpread.SSSetColHidden(C_DNSeq,C_DNSeq,True)
					
			Case Else
				
		End Select	
			
		.vspdData.ReDraw = True
			
	End With

End Sub
'========================================================================================================
Sub HideCancelRelGrid()
		
	With frm1
			
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
					
				Call ggoSpread.SSSetColHidden(C_LCNo,C_LCNo,False)
				Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,False)
				Call ggoSpread.SSSetColHidden(C_LCSeq,C_LCSeq,False)
				Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,False)
				Call ggoSpread.SSSetColHidden(C_PONo,C_PONo,False)
				Call ggoSpread.SSSetColHidden(C_POSeq,C_POSeq,False)
				Call ggoSpread.SSSetColHidden(C_SONo,C_SONo,False)
				Call ggoSpread.SSSetColHidden(C_SoSeq,C_SoSeq,False)
				Call ggoSpread.SSSetColHidden(C_DNNo,C_DNNo,False)
				Call ggoSpread.SSSetColHidden(C_DNSeq,C_DNSeq,False)
					
				.vspdData.Col = 1
				.vspdData.Row = .vspdData.ActiveRow
				.vspdData.Action = 0
				.vspdData.EditMode = True
			
		.vspdData.ReDraw = True
	End With

End Sub
'========================================================================================================
Sub SumAmt()
	With frm1
		Dim strVal
		Dim dblTotDocAmt
		Dim dblTotNetweight
		Dim intCnt
			
		ggoSpread.Source = .vspdData
			
		For intCnt = 1 to .vspdData.MaxRows

			.vspdData.Row = intCnt	:	.vspdData.Col = 0
			If Trim(.vspdData.Text) <> ggoSpread.DeleteFlag Then
				
				.vspdData.Col = C_DocAmt	:	.vspdData.Row = intCnt
				If .vspdData.text <>"" Then dblTotDocAmt = dblTotDocAmt + UNICDbl(.vspdData.text)
			
				.vspdData.Col = C_NetWeight	:	.vspdData.Row = intCnt
				If .vspdData.text <>"" Then dblTotNetweight = dblTotNetweight + UNICDbl(.vspdData.text)

			End If

		Next
			
		.txtNetWeight.text = UNIFormatNumber(dblTotNetweight, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
		.txtDocAmt.Text = UNIFormatNumberByCurrecny(dblTotDocAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)

	End With

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
	Case 0 
		Call CookiePage(1)
		Call PgmJump(EXCC_LAN_ENTRY_ID)
	Case 1
		Call CookiePage(1)
		Call PgmJump(EXCC_HEADER_ENTRY_ID)
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
	
	Call HideNonRelGrid()
	
End Sub
'====================================================================================================
Function RefCheckMessage(strRefFlag)

	RefCheckMessage = False
	If strRefFlag <> Trim(frm1.txtRefFlg.value) Then
		Select Case Trim(frm1.txtRefFlg.value)
		Case "L"
			Call DisplayMsgBox("209002", "X", "L/C", "L/C내역참조")
	
			Exit Function
		Case "S"
			Call DisplayMsgBox("209002", "X", "수주", "수주내역참조")
	
			Exit Function
		Case "M"
			Call DisplayMsgBox("209002", "X", "외주출고", "외주출고참조")
	
			Exit Function
		End Select
	End If

	RefCheckMessage = True

End Function
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'총통관금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
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
		
	Call SetToolbar("1110000000001111")								

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
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Row > 0 And Col = C_HsPopup  Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenHSCd(.Text)		
		End If
    
	End With
	Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
End Sub
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim dblQty
	Dim dblPrice
	Dim dblAmt

	ggoSpread.Source = frm1.vspdData

	Select Case Col
		Case C_Qty, C_Price

			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_Qty

			dblQty = frm1.vspdData.Text

			frm1.vspdData.Row = Row
			frm1.vspddata.Col = C_Price

			dblPrice = frm1.vspdData.Text

			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_DocAmt
			
			frm1.vspdData.Text = UNIFormatNumberByCurrecny(dblAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
			
		Case Else
	End Select
		
	Call SumAmt()
		
	ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True
End Sub
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
	
End Sub
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
	Call SetToolbar("1110000000001111")							
	Call InitVariables											

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
	SetSpreadColor frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
End Function
'========================================================================================================
Function FncCancel() 
	ggoSpread.Source = frm1.vspdData
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	ggoSpread.EditUndo							
	Call SumAmt()
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
		
	Call HideCancelRelGrid()
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
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep	

					.vspdData.Col = C_CCSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_ItemCd							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Unit								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Qty								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_Price								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_DocAmt							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_NetWeight							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_HsCd								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Plant								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						
					.vspdData.Col = C_DNNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				
					.vspdData.Col = C_DNSeq								
						
					If Len(.vspdData.Text) Then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & Parent.gColSep
					End If
						
					.vspdData.Col = C_SoNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_SoSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_SOISeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_LCNo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_LCSeq							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_MvmtNo						
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_PONo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_POSeq						

'					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep '2006-03-31 박정순 수정(첨자오류)
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep


					
					'총중량 항목 추가		 2005.10.10		김병상			
					strVal = strVal & "0"								 & Parent.gColSep	'EXT1_QTY
					strVal = strVal & "0"								 & Parent.gColSep	'EXT2_QTY
					strVal = strVal & "0"								 & Parent.gColSep	'EXT3_QTY
					strVal = strVal & "0"								 & Parent.gColSep	'EXT1_AMT
					strVal = strVal & "0"								 & Parent.gColSep	'EXT2_AMT
					strVal = strVal & "0"								 & Parent.gColSep	'EXT3_AMT
					strVal = strVal & ""								 & Parent.gColSep	'EXT1_CD
					strVal = strVal & ""								 & Parent.gColSep	'EXT2_CD
					strVal = strVal & ""								 & Parent.gRowSep	'EXT3_CD

					lGrpCnt = lGrpCnt + 1
					intInsrtCnt = intInsrtCnt + 1

				Case ggoSpread.UpdateFlag								
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep	

					.vspdData.Col = C_CCSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_ItemCd							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Unit								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Qty								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_Price								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_DocAmt							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_NetWeight							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_HsCd								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						
					.vspdData.Col = C_Plant								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_DNNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_DNSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_SoNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_SoSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_SOISeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_LCNo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_LCSeq							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_MvmtNo						
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_PONo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_POSeq						
'					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep '2006-03-31 박정순 수정(첨자오류)
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep 
					
					'총중량 항목 추가		 2005.10.10		김병상			
					strVal = strVal & "0"								 & Parent.gColSep	'EXT1_QTY
					strVal = strVal & "0"								 & Parent.gColSep	'EXT2_QTY
					strVal = strVal & "0"								 & Parent.gColSep	'EXT3_QTY
					strVal = strVal & "0"								 & Parent.gColSep	'EXT1_AMT
					strVal = strVal & "0"								 & Parent.gColSep	'EXT2_AMT
					strVal = strVal & "0"								 & Parent.gColSep	'EXT3_AMT
					strVal = strVal & ""								 & Parent.gColSep	'EXT1_CD
					strVal = strVal & ""								 & Parent.gColSep	'EXT2_CD
					strVal = strVal & ""								 & Parent.gRowSep	'EXT3_CD

					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag							
					strVal = strVal & "D" & Parent.gColSep	& lRow & Parent.gColSep	

					.vspdData.Col = C_CCSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_ItemCd							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Unit								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Qty								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Price								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_DocAmt							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_NetWeight							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_HsCd								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_Plant								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						
					.vspdData.Col = C_DNNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				
					.vspdData.Col = C_DNSeq								
						
					If Len(.vspdData.Text) Then
						strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					Else
						strVal = strVal & "0" & Parent.gColSep
					End If
						
					.vspdData.Col = C_SoNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_SoSeq								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						
					.vspdData.Col = C_SOISeq								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						
					.vspdData.Col = C_LCNo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_LCSeq							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_MvmtNo						
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_PONo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_POSeq						
'					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep  '2006-03-31 박정순 수정(첨자오류)
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					'총중량 항목 추가		 2005.10.10		김병상			
					strVal = strVal & "0"				   & Parent.gColSep	'EXT1_QTY
					strVal = strVal & "0"				   & Parent.gColSep	'EXT2_QTY
					strVal = strVal & "0"				   & Parent.gColSep	'EXT3_QTY
					strVal = strVal & "0"				   & Parent.gColSep	'EXT1_AMT
					strVal = strVal & "0"				   & Parent.gColSep	'EXT2_AMT
					strVal = strVal & "0"				   & Parent.gColSep	'EXT3_AMT
					strVal = strVal & ""				   & Parent.gColSep	'EXT1_CD
					strVal = strVal & ""				   & Parent.gColSep	'EXT2_CD
					strVal = strVal & ""				   & Parent.gRowSep	'EXT3_CD

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
	Call SetToolbar("11101011000111")							
	Call HideNonRelGrid()
	lgBlnFlgChgValue = False
		
    frm1.vspdData.Focus     

End Function
'========================================================================================================
Function CCHdrQueryOk()												
	Call HideNonRelGrid()
	Call SetToolbar("11101011000011")
		
	If frm1.txtRefFlg.value = "M" Then
		Call ggoSpread.SSSetColHidden(C_HsPopup,C_HsPopup,False)
	Else
		Call ggoSpread.SSSetColHidden(C_HsPopup,C_HsPopup,True)
	End IF								
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관내역</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenSODtlRef">수주내역참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCDtlRef">L/C내역참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNDtlRef">외주출고참조</A></TD>
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
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSONo" MAXLENGTH=18 SIZE=20 TAG="24XXXU" ALT="수주번호"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="영업그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LC번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결제방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="결제방법">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>총 통관순중량</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtNetWeight" style="HEIGHT: 20px; WIDTH: 150px" tag="24X3" ALT="통관순중량" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtWeightUnit" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="중량단위"></TD>
											</TR>
										</TABLE>
									</TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>가격조건</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoTerms" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="가격조건">&nbsp;<INPUT TYPE=TEXT NAME="txtIncoTermsNm" SIZE=20 TAG="24"></TD>										
									<TD CLASS=TD5 NOWRAP>총통관금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="통관금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
											</TR>
										</TABLE>										
									</TD>
								</TR>				
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(2)">Container배정</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(0)">통관란등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">통관등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
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
