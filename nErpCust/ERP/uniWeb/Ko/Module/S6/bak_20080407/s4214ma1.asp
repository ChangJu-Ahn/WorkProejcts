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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                       

Dim C_ContNo		'Carton 번호, 컨테이너번호 
Dim C_StartCtNo			'
Dim C_EndCtNo			
Dim C_PackingCnt
Dim C_NetW
Dim C_GrossW
Dim	C_Measure
Dim C_Qty			
Dim C_Unit			
Dim C_HsCd
Dim C_ItemCd	
Dim C_ItemNm			
Dim C_ItemSpec  
Dim C_LanNo     
Dim C_Plant
Dim C_HsPopup		
Dim C_PlantNm		
Dim C_DnNo
Dim C_DnSeq			
Dim C_SoNo			
Dim C_SoSeq			
Dim C_SoISeq
Dim C_CcSeq	
Dim C_ChgFlg			

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent					
Dim IsOpenPop

Const BIZ_PGM_QRY_ID = "s4214mb1.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "s4214mb1.asp"			'☆: 비지니스 로직 ASP명 
Const EXCC_HEADER_ENTRY_ID = "s4211ma1"			'☆: 이동할 ASP명: 통관등록 
Const EXCC_DETAIL_ENTRY_ID = "s4212ma1"			'☆: 이동할 ASP명 : 통관내역등록 
Const EXCC_LAN_ENTRY_ID = "s4213ma1"			'☆: 이동할 ASP명 : 통관란등록 
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_ContNo		=1'Carton 번호, 컨테이너번호 
	C_StartCtNo	=2		'
	C_EndCtNo	=3		
	C_PackingCnt		=4
	C_NetW			=5
	C_GrossW		=6	
	C_Measure	=7
	C_Qty			=8
	C_Unit			=9
	C_HsCd			=10
	'C_HsPopup	=11
	C_ItemCd		=11
	C_ItemNm		=12	
	C_ItemSpec  =13
	C_LanNo		=14
	C_Plant			=15
	'C_DocAmt		=15	
	C_PlantNm		=16
	C_DnNo			=17
	C_DnSeq		=18	
	C_SoNo			=19
	C_SoSeq		=20	
	C_SoISeq	=21
	C_CcSeq		=22
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
	'frm1.txtDocAmt.Text = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.txtNetW.Text = UNIFormatNumber(0, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
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
			
		.vspdData.MaxCols = C_CcSeq
		.vspdData.MaxRows = 0
			
		Call GetSpreadColumnPos("A")	
			
		ggoSpread.SSSetEdit		C_ContNo, "Carton번호", 15, 0
		ggoSpread.SSSetEdit		C_StartCtNo, "CN-No(From)", 15, 0
		ggoSpread.SSSetEdit		C_EndCtNo, "CN-No(End)", 15, 0												
		ggoSpread.SSSetFloat	C_PackingCnt,"포장개수" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,""
		ggoSpread.SSSetFloat	C_NetW,"순중량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,""
		ggoSpread.SSSetFloat	C_GrossW,"총중량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,""
		ggoSpread.SSSetFloat	C_Measure,"용적" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,""
		ggoSpread.SSSetFloat	C_Qty,"Packing수량" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,""
		
		ggoSpread.SSSetEdit		C_Unit, "단위", 10, 0
		ggoSpread.SSSetEdit		C_HsCd, "HS부호", 20, 0		
		ggoSpread.SSSetEdit		C_ItemCd, "품목", 18, 0
		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25,,,40
		ggoSpread.SSSetEdit		C_ItemSpec, "규격", 20,,,50		
		ggoSpread.SSSetEdit		C_LanNo, "란번호", 10, 0
		ggoSpread.SSSetEdit		C_Plant, "공장", 10, 0
		ggoSpread.SSSetEdit		C_PlantNm, "공장명", 10, 0		

		ggoSpread.SSSetEdit		C_DnNo, "출하번호", 18, 0
		ggoSpread.SSSetEdit		C_DnSeq, "출하순번", 10, 1
		ggoSpread.SSSetEdit		C_SoNo, "수주번호", 18, 0
		ggoSpread.SSSetEdit		C_SoSeq, "수주순번", 10, 1
		ggoSpread.SSSetEdit		C_SoISeq, "수주일정순번", 15, 1
		ggoSpread.SSSetEdit		C_CcSeq, "통관순번", 10, 1
			
		SetSpreadLock "", 0, -1, ""

		'Call ggoSpread.SSSetColHidden(C_Plant,C_Plant,True)
		'Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
		
		'Call ggoSpread.SSSetColHidden(C_HsPopup,C_HsPopup,True)
	
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
		
		ggoSpread.SpreadLock C_ContNo, lRow, -1
	
		ggoSpread.SpreadLock C_StartCtNo, lRow, -1

		ggoSpread.SpreadUnLock C_EndCtNo, lRow, -1
		ggoSpread.SSSetRequired C_EndCtNo, lRow, -1
		
		ggoSpread.SpreadUNLock C_Qty, lRow, -1
		ggoSpread.SSSetRequired C_Qty, lRow, -1
		
		ggoSpread.SpreadUnLock C_NetW, lRow, -1		
		ggoSpread.SpreadLock C_HsCd, lRow, -1		
		ggoSpread.SpreadLock C_LanNo, lRow, -1
		ggoSpread.SpreadLock C_Plant, lRow, -1
		ggoSpread.SpreadLock C_PlantNm, lRow, -1
		ggoSpread.SpreadLock C_DNNo, lRow, -1
		ggoSpread.SpreadLock C_DNSeq, lRow, -1 
		ggoSpread.SpreadLock C_SoNo, lRow, -1
		ggoSpread.SpreadLock C_SoSeq, lRow, -1
		ggoSpread.SpreadLock C_SOISeq, lRow, -1
		ggoSpread.SpreadLock C_ItemSpec, lRow, -1

		
		.vspdData.ReDraw = True
	End With
End Sub
'========================================================================================================
Sub SetSpreadColor(ByVal lRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		
		ggoSpread.SSSetProtected C_HsCd, lRow, lRow
		ggoSpread.SSSetProtected C_CCSeq, lRow, lRow
		ggoSpread.SSSetProtected C_ItemCd, lRow, lRow
		ggoSpread.SSSetProtected C_Unit, lRow, lRow
		'ggoSpread.SSSetProtected C_Qty, lRow, lRow
		ggoSpread.SSSetProtected C_LanNo, lRow, lRow
		ggoSpread.SSSetProtected C_Plant, lRow, lRow
		ggoSpread.SSSetProtected C_PlantNm, lRow, lRow
		ggoSpread.SSSetProtected C_DNNo, lRow, lRow
		ggoSpread.SSSetProtected C_DNSeq, lRow, lRow
		ggoSpread.SSSetProtected C_SoNo, lRow, lRow
		ggoSpread.SSSetProtected C_SoSeq, lRow, lRow  
		ggoSpread.SSSetProtected C_SOISeq, lRow, lRow
		
		ggoSpread.SSSetProtected C_ItemSpec, lRow, lRow
		
		ggoSpread.SpreadUnLock C_ContNo, lRow, lRow
		ggoSpread.SSSetRequired C_ContNo, lRow, lRow
		ggoSpread.SpreadUnLock C_StartCtNo, lRow, lRow
		ggoSpread.SSSetRequired C_StartCtNo, lRow, lRow
		ggoSpread.SpreadUnLock C_EndCtNo, lRow, lRow
		ggoSpread.SSSetRequired C_EndCtNo,lRow, lRow	
		
		
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
			
			C_ContNo		= iCurColumnPos(1)  
			C_StartCtNo	=iCurColumnPos(2)		'
			C_EndCtNo	=iCurColumnPos(3)		
			C_PackingCnt		=iCurColumnPos(4)
			C_NetW			=iCurColumnPos(5)
			C_GrossW		=iCurColumnPos(6	)
			C_Measure	=iCurColumnPos(7)
			C_Qty			=iCurColumnPos(8)
			C_Unit			=iCurColumnPos(9)
			C_HsCd			=iCurColumnPos(10)			
			C_ItemCd		=iCurColumnPos(11)
			C_ItemNm		=iCurColumnPos(12)	
			C_ItemSpec  =iCurColumnPos(13)
			C_LanNo		=iCurColumnPos(14)
			C_Plant			=iCurColumnPos(15)			
			C_PlantNm		=iCurColumnPos(16)
			C_DnNo			=iCurColumnPos(17)
			C_DnSeq		=iCurColumnPos(18)	
			C_SoNo			=iCurColumnPos(19)
			C_SoSeq		=iCurColumnPos(20)	
			C_SoISeq	=iCurColumnPos(21)
			C_CcSeq		=iCurColumnPos(22)

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
Function OpenCCDtlRef()
	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD		
		
	If Trim(frm1.txtApplicant.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		frm1.txtCCNo.focus
		Exit Function
	End If
		
'	If RefCheckMessage("L") = False Then Exit Function
		
	iCalledAspName = AskPRAspName("S4213RA9")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4213RA9", "X")					
		gblnWinEvent = False		
		Exit Function
	End If
	arrParam(0) = Trim(frm1.txtCCNo.value)		
	
		
	strRet = window.showModalDialog(iCalledAspName , Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	If strRet(0, 0) = "" Then
		Exit Function
	Else
		Call SetCCDtlRef(strRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetCCDtlRef(arrRet)
		
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

'			If TempRow <> 0 Then
'				For j = 1 To TempRow			
'					.vspdData.Row = j
'					.vspdData.Col = C_CCSeq
'					strtemp2 = .vspdData.text
'
'					If .vspdData.Text = arrRet(intCnt - 1, 14) Then
'						strMessage = strMessage & strtemp2  & vbCrlf
'						blnEqualFlg = True
'						Exit For
'					End If
'				Next
'			End If

			If blnEqualFlg = False Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
'					.vspdData.MaxRows = CLng(TempRow) + CLng(intCnt)
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)					
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag					
				.vspdData.Col = C_HSCd							'hs 						
				.vspdData.text = arrRet(intCnt - 1, 0)				
				.vspdData.Col = C_ItemCd							'품목코드			
				.vspdData.text = arrRet(intCnt - 1, 1)
				.vspdData.Col = C_ItemNm						'품목명				
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_ItemSpec						'규격				
				.vspdData.text = arrRet(intCnt - 1, 3)				
				.vspdData.Col = C_Qty								'packing 잔량  
				.vspdData.text = arrRet(intCnt - 1, 5)								
'				.vspdData.Col = C_PackingCnt							'
'				.vspdData.text = arrRet(intCnt - 1, 5)																
				.vspdData.Col = C_Unit								'단위			
				.vspdData.text = arrRet(intCnt - 1, 6)				
				.vspdData.Col = C_Plant								'공장					
				.vspdData.text = arrRet(intCnt - 1, 7)
				.vspdData.Col = C_PlantNm						'공장명 
				.vspdData.text = arrRet(intCnt - 1, 8)
				.vspdData.Col = C_DNNo							'출하번호			
				.vspdData.text = arrRet(intCnt - 1, 9)
				.vspdData.Col = C_DNSeq							'출하순번			
				.vspdData.text = arrRet(intCnt - 1, 10)
				.vspdData.Col = C_SoNo							'수주번호						
				.vspdData.text = arrRet(intCnt - 1, 11)
				.vspdData.Col = C_SoSeq							'수주순번 
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_SOISeq							'수주일정순번 
				.vspdData.text = arrRet(intCnt - 1, 13)
				.vspdData.Col = C_CcSeq							'통관순번 
				.vspdData.text = arrRet(intCnt - 1, 14)
				.vspdData.Col = C_LanNo							'란번호 
				.vspdData.text = arrRet(intCnt - 1,15)										
				
				
'				.vspdData.Col = C_ChgFlg									
'				.vspdData.text = .vspdData.Row
					
				SetSpreadColor CLng(TempRow) + CLng(intCnt)

				lgBlnFlgChgValue = True
			End If
		Next
		
		Call SumAmt()

		If strMessage <> "" Then
			'Call DisplayMsgBox("17a005", "X",strmessage,"통관순번" & "," & "통관순번")
			Call DisplayMsgBox("17a005", "X",strmessage,"통관순번" )
			.vspdData.ReDraw = True
		End If

		.vspdData.ReDraw = True

	End With
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'---------------------------------------------------------------------------------------------------------
'Function SetHSCd(Byval arrRet)
'	With frm1
'		.vspdData.Col = C_HsCd
'		.vspdData.Text = arrRet(0)
''
'		ggoSpread.Source = .vspdData
'		Call vspdData_Change(C_HsCd, .vspdData.ActiveRow )
'	End With
'End Function
'========================================================================================================
'Sub HideNonRelGrid()
'	Dim RefFlg
'		
'	With frm1
'		RefFlg = .txtRefFlg.value 
'		ggoSpread.Source = .vspdData
'			
'		.vspdData.ReDraw = False
'			
'		Select Case RefFlg
'			Case "S"
'				Call ggoSpread.SSSetColHidden(C_LCNo,C_LCNo,True)
'				Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,True)
'				Call ggoSpread.SSSetColHidden(C_LCSeq,C_LCSeq,True)
'				Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)
'				Call ggoSpread.SSSetColHidden(C_PONo,C_PONo,True)
'				Call ggoSpread.SSSetColHidden(C_POSeq,C_POSeq,True)
'			Case "L"
'				Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)
'				Call ggoSpread.SSSetColHidden(C_PONo,C_PONo,True)
''				Call ggoSpread.SSSetColHidden(C_POSeq,C_POSeq,True)			

'			Case "M"
'				Call ggoSpread.SSSetColHidden(C_LCNo,C_LCNo,True)
'				Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,True)
'				Call ggoSpread.SSSetColHidden(C_LCSeq,C_LCSeq,True)
'				Call ggoSpread.SSSetColHidden(C_SONo,C_SONo,True)
'				Call ggoSpread.SSSetColHidden(C_SoSeq,C_SoSeq,True)
'				Call ggoSpread.SSSetColHidden(C_SOISeq,C_SOISeq,True)
'				Call ggoSpread.SSSetColHidden(C_DNNo,C_DNNo,True)
'				Call ggoSpread.SSSetColHidden(C_DNSeq,C_DNSeq,True)
'					
'			Case Else
'				
'		End Select	
'			
'		.vspdData.ReDraw = True
'			
'	End With
'
'End Sub
'========================================================================================================
'Sub HideCancelRelGrid()
'		
'	With frm1
'			
'		ggoSpread.Source = .vspdData
'			
'		.vspdData.ReDraw = False
'					
'				Call ggoSpread.SSSetColHidden(C_LCNo,C_LCNo,False)
'				Call ggoSpread.SSSetColHidden(C_LCDocNo,C_LCDocNo,False)
'				Call ggoSpread.SSSetColHidden(C_LCSeq,C_LCSeq,False)
'				Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,False)
'				Call ggoSpread.SSSetColHidden(C_PONo,C_PONo,False)
'				Call ggoSpread.SSSetColHidden(C_POSeq,C_POSeq,False)
'				Call ggoSpread.SSSetColHidden(C_SONo,C_SONo,False)
'				Call ggoSpread.SSSetColHidden(C_SoSeq,C_SoSeq,False)
'				Call ggoSpread.SSSetColHidden(C_DNNo,C_DNNo,False)
''				Call ggoSpread.SSSetColHidden(C_DNSeq,C_DNSeq,False)
'					
'				.vspdData.Col = 1
'				.vspdData.Row = .vspdData.ActiveRow
''				.vspdData.Action = 0
'				.vspdData.EditMode = True
'			
'		.vspdData.ReDraw = True
'	End With
'
'End Sub
'========================================================================================================
Sub SumAmt()

		Dim strVal
		Dim dblTotContNo
		Dim dblTotNetw
		Dim dblTotGrossw
		Dim dblTotPacking		
		Dim dblTotMeasure
		Dim intCnt
		
		

			
		With frm1
		ggoSpread.Source = .vspdData
		

		For intCnt = 1 to .vspdData.MaxRows

			.vspdData.Row = intCnt	:	.vspdData.Col = 0

		'	If Trim(.vspdData.Text) <> ggoSpread.DeleteFlag Then
					

				'.vspdData.Col = C_ContNo	:	.vspdData.Row = intCnt
				'If .vspdData.text <>"" Then dblTotContNo = dblTotContNo + UNICDbl(.vspdData.text)
				.vspdData.Col = C_NetW	:	.vspdData.Row = intCnt
				If .vspdData.text <>"" Then dblTotNetw = dblTotNetw + UNICDbl(.vspdData.text)

				.vspdData.Col = C_GrossW	:	.vspdData.Row = intCnt
				If .vspdData.text <>"" Then dblTotGrossw = dblTotGrossw + UNICDbl(.vspdData.text)

				.vspdData.Col = C_Measure	:	.vspdData.Row = intCnt
				If .vspdData.text <>"" Then dblTotMeasure = dblTotMeasure + UNICDbl(.vspdData.text)

				.vspdData.Col = C_PackingCnt	:	.vspdData.Row = intCnt
				If .vspdData.text <>"" Then dblTotPacking = dblTotPacking + UNICDbl(.vspdData.text)

		'	End If

		Next

		.txtCarton.text   = UNIFormatNumber(dblTotContNo, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)		
		.txtNetW.text		= UNIFormatNumber(dblTotNetw, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
		.txtGrossW.text	= UNIFormatNumber(dblTotGrossw, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
		.txtPacking.text	= UNIFormatNumber(dblTotPacking, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
		.txtMsmnt.text	= UNIFormatNumber(dblTotMeasure, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
		
		
		'.txtDocAmt.Text = UNIFormatNumberByCurrecny(dblTotDocAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)

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
		'ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
'		ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'금액 
'		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
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
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'	With frm1.vspdData 
'	
'		ggoSpread.Source = frm1.vspdData
 '  
'		If Row > 0 And Col = C_HsPopup  Then
'		    .Col = Col - 1
'		    .Row = Row
'		    Call OpenHSCd(.Text)		
'		End If
    
'	End With
'	Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
'End Sub
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim intStartCtNo, intEndCtNo
	Dim tmpCal
	Dim dblAmt
	
	
	ggoSpread.Source = frm1.vspdData
	with frm1.vspdData
	
	Select Case Col
		Case C_StartCtNo, C_EndCtNo
			.Row = Row
			.Col = C_StartCtNo
			intStartCtNo = .Text
			.row = row
			.Col = C_EndCtNo
			intEndCtNo = .Text			
			
			If len(intStartCtNo) <>0 and len(intEndCtNo) <>0 then	
				If checkCtNo(Row)	=false then
					Call DisplayMsgBox("800443", "X", "Ct_No(End)", "Ct_No(From)")	
					'msgbox "Ct_No(End)는 Ct_No(From) 보다 작을 수 없습니다."
					Exit Sub
				End If
				tmpCal = intEndCtNo - intStartCtNo +1
				.Row = Row
				.Col = C_PackingCnt
				.Text =tmpCal				
			End If		
		Case Else
	End Select
	
	End With
			
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
	

'	If checkCtNo("A") = False then
'		Call DisplayMsgBox("800443", "X", "Ct_No(End)", "Ct_No(From)")	
'		'msgbox "Ct_No(End)는 Ct_No(From) 보다 작을 수 없습니다."
'		'.focus
'		Exit Function 
'	End If
	

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

					.vspdData.Col = C_ContNo
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_StartCtNo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_EndCtNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep


					.vspdData.Col = C_PackingCnt
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If


					.vspdData.Col = C_NetW
					If len(.vspdData.text) then					
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else					
						strVal = strVal & "0" & parent.gColSep
					End If

				
					.vspdData.Col = C_GrossW
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If
				
				
					.vspdData.Col = C_Measure						
					If Len(.vspdData.Text) Then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & Parent.gColSep
					End If

				
					.vspdData.Col = C_Qty								
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If

								
					.vspdData.Col = C_Unit								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep						
					

				
					.vspdData.Col = C_HsCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					.vspdData.Col = C_ItemNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep										
					.vspdData.Col = C_ItemSpec
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_LanNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_Plant
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					.vspdData.Col = C_PlantNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_DnNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_DnSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep	
					.vspdData.Col = C_SoNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_SoSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep						
					.vspdData.Col = C_SOISeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_CCSeq							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) &  Parent.gRowSep
					'통관관리번호 
					'strVal = strVal & Trim(.txtCCNo.value) & Parent.gColSep					

					lGrpCnt = lGrpCnt + 1
					intInsrtCnt = intInsrtCnt + 1

				Case ggoSpread.UpdateFlag						
								
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep						

					.vspdData.Col = C_ContNo
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_StartCtNo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_EndCtNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep


					.vspdData.Col = C_PackingCnt
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If


					.vspdData.Col = C_NetW
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If

				
					.vspdData.Col = C_GrossW
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If
				
				
					.vspdData.Col = C_Measure						
					If Len(.vspdData.Text) Then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & Parent.gColSep
					End If

				
					.vspdData.Col = C_Qty								
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If

								
					.vspdData.Col = C_Unit								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep						

				
					.vspdData.Col = C_HsCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					.vspdData.Col = C_ItemNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep										
					.vspdData.Col = C_ItemSpec
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_LanNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_Plant
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					.vspdData.Col = C_PlantNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_DnNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_DnSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep	
					.vspdData.Col = C_SoNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_SoSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep						
					.vspdData.Col = C_SOISeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_CCSeq							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) &  Parent.gRowSep
					'통관관리번호 
					'strVal = strVal & Trim(.txtCCNo.value) & Parent.gColSep					

					lGrpCnt = lGrpCnt + 1
					'intInsrtCnt = intInsrtCnt + 1


				Case ggoSpread.DeleteFlag							
				
					strVal = strVal & "D" & Parent.gColSep	& lRow & Parent.gColSep	
					.vspdData.Col = C_ContNo
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

					.vspdData.Col = C_StartCtNo							
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

					.vspdData.Col = C_EndCtNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep


					.vspdData.Col = C_PackingCnt
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If


					.vspdData.Col = C_NetW
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If

				
					.vspdData.Col = C_GrossW
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If
				
				
					.vspdData.Col = C_Measure						
					If Len(.vspdData.Text) Then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & Parent.gColSep
					End If

				
					.vspdData.Col = C_Qty								
					If len(.vspdData.text) then
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					Else
						strVal = strVal & "0" & parent.gColSep
					End If

								
					.vspdData.Col = C_Unit								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep						

				
					.vspdData.Col = C_HsCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					.vspdData.Col = C_ItemNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep										
					.vspdData.Col = C_ItemSpec
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_LanNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_Plant
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					.vspdData.Col = C_PlantNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep					
					
					.vspdData.Col = C_DnNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_DnSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep	
					.vspdData.Col = C_SoNo								
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_SoSeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep						
					.vspdData.Col = C_SOISeq								
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
						
					.vspdData.Col = C_CCSeq							
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) &  Parent.gRowSep
					'통관관리번호 
					'strVal = strVal & Trim(.txtCCNo.value) & Parent.gColSep					

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
	'Call HideNonRelGrid()
	lgBlnFlgChgValue = False
		
    frm1.vspdData.Focus     

End Function
'========================================================================================================
Function CCHdrQueryOk()												
'	Call HideNonRelGrid()
	Call SetToolbar("11101011000011")
		
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Container 배정</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenCCDtlRef">통관참조</A></TD>
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
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP></TD>										
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
