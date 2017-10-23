<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M5113ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 매입 지급내역등록 ASP														*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2001/05/09																*
'*  8. Modified date(Last)  : 2003/06/05																*
'*  9. Modifier (First)     : Ma Jin Ha																			*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/04/11 : Coding Start												*
'********************************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<!--
'============================================  1.1.2 공통 Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
	Option Explicit					 '☜: indicates that All variables must be declared in advance 
	
	Const BIZ_PGM_ID 	 = "m5113mb1.asp"	
	Const BIZ_PGM_JUMP_ID = "M5111MA1"               '매입세금계산서 점프 
	Const BIZ_PGM_JUMP_ID2 = "M5112MA1"              '매입내역등록 점프 
	Const BIZ_PGM_JUMP_ID3 = "M5211MA1"              'B/L 등록 
	Const BIZ_PGM_JUMP_ID4 = "M5212MA1"              'B/L 내역등록 

	Dim C_PayType
	Dim C_PayTypePopup	
	Dim C_PayTypeNm		
	Dim C_PayDocAmt		
	Dim C_PayLocAmt		
	Dim C_ExchRate		
	Dim C_BankAcct		
	Dim C_BankAcctPopup	
	Dim C_BankCd		
	Dim C_BankPopup		
	Dim C_BankNm		
    Dim C_Noteno		
    Dim C_NotenoPopup	
    Dim C_PrepayNo		
    Dim C_PrepayNoPopup
    Dim C_LoanNo
	Dim C_LoanNoPopup
	Dim C_PaySeq		

	Const CID_POST  = 5211    '확정 
	

	Dim lgBlnFlgChgValue					'☜: Variable is for Dirty flag 
	Dim lgIntGrpCount						'☜: Group View Size를 조사할 변수 
	Dim lgIntFlgMode						'☜: Variable is for Operation Status 

	Dim lgStrPrevKey
	Dim lgLngCurRows
	Dim gblnWinEvent						'~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
                    						  '	PopUp Window가 사용중인지 여부를 나타내는 variable 
	Dim IsOpenPop
	Dim lblnWinEvent
	Dim interface_Account

	Dim arrCollectType
	Dim lgSortKey

'=====================================  initSpreadPosVariables()  =======================================
Sub initSpreadPosVariables()  
	C_PayType		= 1                      '지급유형 
	C_PayTypePopup	= 2                      '지급유형 팝업 
	C_PayTypeNm		= 3                      '지급유형명 
	C_PayDocAmt		= 4                      '지급금액 
	C_PayLocAmt		= 5                      '지급자국금액 
	C_ExchRate		= 6                      '환율 
	C_BankAcct		= 7                      '계좌번호 
	C_BankAcctPopup	= 8	                     '계좌번호 팝업 
	C_BankCd		= 9                      '은행 
	C_BankPopup		= 10                     '은행 팝업 
	C_BankNm		= 11                     '은행명 
    C_Noteno		= 12                     '어음번호 
    C_NotenoPopup	= 13                     '어음번호 팝업 
    C_PrepayNo		= 14                     '선급금번호 
    C_PrepayNoPopup	= 15                     '선급금번호 팝업 
    C_LoanNo		= 16					'차입금번호 
	C_LoanNoPopup	= 17					'차입금번호팝업 
	C_PaySeq		= 18                     '지급순번 
End Sub

'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
Dim strTemp
		Dim IntRetCD

	If Kubun = 1 Then                        '매입세금계산서 점프 

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                   'Check if there is retrived data
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
		WriteCookie "txtIvNo" , FilterVar(UCase(Trim(frm1.txtIvNo.value)), "", "SNM")
		
		Call PgmJump(BIZ_PGM_JUMP_ID)
		 
	ElseIf Kubun = 0 Then
		
		strTemp = ReadCookie("txtIvNo")
		
		If strTemp = "" then Exit Function
		
		frm1.txtIvNo.Value = strTemp
		
		WriteCookie "txtIvNo" , ""
		
		if Trim(strTemp) <> "" then
			
			frm1.txtQuerytype.value = "Auto"
			frm1.txthdnIvNo.value = strTemp
			Call dbquery()
		end if
	ElseIf Kubun = 2 Then              '매입내역등록 점프 

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                    'Check if there is retrived data
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie "txtIvNo" , FilterVar(UCase(Trim(frm1.txtIvNo.value)), "", "SNM")
		
		Call PgmJump(BIZ_PGM_JUMP_ID2)		
	ElseIf Kubun = 3 Then             'B/L 등록 
		
	    If lgIntFlgMode <> parent.OPMD_UMODE Then                    'Check if there is retrived data
			
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
		
		If UCase(Trim(frm1.txtBlNo.Value)) = "" then
			Exit Function
		End if

		WriteCookie "BlNo" , FilterVar(UCase(Trim(frm1.txtBlNo.value)), "", "SNM")
		
		Call PgmJump(BIZ_PGM_JUMP_ID3)		
	ElseIf Kubun = 4 Then             'B/L 내역등록점프 

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                    'Check if there is retrived data
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
	    
		if UCase(Trim(frm1.txtBlNo.Value)) = "" then
			Exit Function
		End if

		WriteCookie "BlNo" , UCase(Trim(frm1.txtBlNo.value))
		WriteCookie "PoNo" , UCase(Trim(frm1.hdnPoNo.value))
		
		Call PgmJump(BIZ_PGM_JUMP_ID4)	
	End IF
	
End Function
'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False						'⊙: Indicates that no value changed
	lgIntGrpCount = 0								'⊙: Initializes Group View Size
	lgStrPrevKey = ""								'initializes Previous Key
	lgLngCurRows = 0 								'initializes Deleted Rows Count
	lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent = False
End Function
'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()	
	Call SetToolBar("1110000000001111")							 '⊙: 버튼 툴바 제어 
	frm1.txtPayDocAmt.Text = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
			
	frm1.btnGlSel.Disabled = True
	frm1.btnPosting.Disabled = True
	frm1.btnPosting.value = "확정"
	frm1.txtIvNo.Focus
	Set gActiveElement = document.activeElement
		
	interface_Account = GetSetupMod(parent.gSetupMod, "a")
End Sub
'==========================================  2.2.2 LoadInfTB19029()  ====================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()
    
    With frm1
    
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20030425",,parent.gAllowDragDropSpread    
			
		.vspdData.ReDraw = False
			
		.vspdData.MaxCols = C_PaySeq + 1
		.vspdData.MaxRows = 0
			
		Call AppendNumberPlace("6","4","0")
		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetEdit				C_PayType,		"지급유형 ", 9 ,,,,2
		ggoSpread.SSSetButton 			C_PayTypePopup        
		ggoSpread.SSSetEdit				C_PayTypeNm,	"지급유형명", 12
		ggoSpread.SSSetFloat			C_PayDocAmt,	"지급금액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat			C_PayLocAmt,	"지급자국금액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat			C_ExchRate,		"환율" ,8, parent.ggExchRateNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit				C_BankAcct,		"계좌번호", 16,,,30,2
		ggoSpread.SSSetButton			C_BankAcctPopup
		ggoSpread.SSSetEdit				C_BankCd,		"은행", 6,,,10 ,2
		ggoSpread.SSSetButton 			C_BankPopup
		ggoSpread.SSSetEdit				C_BankNm,		"은행명", 20
		ggoSpread.SSSetEdit				C_Noteno,		"어음번호", 16,,,30,2
		ggoSpread.SSSetButton			C_NotenoPopup
		ggoSpread.SSSetEdit				C_PrepayNo,		"선급금번호", 16,,,18,2
		ggoSpread.SSSetButton			C_PrepayNoPopup
		ggoSpread.SSSetEdit				C_LoanNo,		"차입금번호", 16,,,30,2
		ggoSpread.SSSetButton			C_LoanNoPopup
		ggoSpread.SSSetFloat			C_PaySeq,		"지급순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
				
		Call ggoSpread.MakePairsColumn(C_PayType,C_PayTypePopup)
		Call ggoSpread.MakePairsColumn(C_BankCd,C_BankPopup)
		Call ggoSpread.MakePairsColumn(C_Noteno,C_NotenoPopup)
		Call ggoSpread.MakePairsColumn(C_PrepayNo,C_PrepayNoPopup)
		Call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPopup)
		Call ggoSpread.MakePairsColumn(C_LoanNo,C_LoanNoPopup)
		Call ggoSpread.SSSetColHidden(C_PaySeq + 1, C_PaySeq + 1, True)
			
		.vspdData.ReDraw = True
					
		Call SetSpreadLock
	End With
End Sub
'==========================================  2.2.4 SetSpreadLock()  =====================================
Sub SetSpreadLock()
	With frm1

		.vspdData.ReDraw = False
			 
		ggoSpread.SpreadUnLock 	C_PayType , -1
		ggoSpread.SSSetRequired C_PayType , -1
		ggoSpread.SpreadLock	C_PayTypeNm , -1
		ggoSpread.SpreadUnLock 	C_PayDocAmt , -1
		ggoSpread.SSSetRequired C_PayDocAmt , -1  
		'ggoSpread.SpreadLock 	C_PayLocAmt , -1 
		ggoSpread.SpreadLock 	C_ExchRate, -1 
		ggoSpread.SpreadUnLock 	C_BankAcct , -1
		ggoSpread.SpreadUnLock 	C_BankAcctPopup , -1 
		ggoSpread.SpreadUnLock 	C_BankCd, -1
		ggoSpread.SpreadUnLock 	C_BankPopup, -1  
		ggoSpread.SpreadLock    C_BankNm, -1
		ggoSpread.SpreadUnLock 	C_Noteno , -1
		ggoSpread.SpreadUnLock 	C_NotenoPopup , -1 
		ggoSpread.SpreadUnLock 	C_PrepayNo, -1
		ggoSpread.SpreadUnLock 	C_PrepayNoPopup, -1  	
		ggoSpread.SpreadUnLock 	C_LoanNo, -1
		ggoSpread.SpreadUnLock 	C_LoanNoPopup, -1  	
		ggoSpread.SpreadLock 	C_PaySeq , -1
		ggoSpread.SSSetProtected	C_PaySeq + 1,  -1	
			
		.vspdData.ReDraw = True
	End With
End Sub
'==========================================  2.2.5 SetSpreadColor()  ====================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
	    
		.Redraw = False

		ggoSpread.SSSetRequired	 C_PayType, pvStartRow, pvEndRow                 '지급유형 
		ggoSpread.SSSetProtected C_PayTypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	 C_PayDocAmt, pvStartRow, pvEndRow		       '지급금액 
		'ggoSpread.SSSetProtected C_PayLocAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ExchRate, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BankAcct, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BankAcctPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BankCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BankPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BankNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Noteno, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_NotenoPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PrepayNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PrepayNoPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LoanNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LoanNoPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PaySeq, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PaySeq+1, pvStartRow, pvEndRow
			
		.Col = 1
		.Row = .ActiveRow
		.Action = 0
		.EditMode = True

		.ReDraw = True
	End With
End Sub
'==========================================  2.2.5 SetRdSpreadColor()  ====================================
Sub SetRdSpreadColor(ByVal pvStartRow)
    With frm1
    
		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected C_PayType, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_PayTypePopup, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_PayTypeNm, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_PayDocAmt, pvStartRow, .vspdData.MaxRows 	
		ggoSpread.SSSetProtected C_PayLocAmt, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_ExchRate, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_BankAcct, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_BankAcctPopup, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_BankCd, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_BankPopup, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_BankNm, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_Noteno, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_NotenoPopup, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_PrepayNo, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_PrepayNoPopup, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_LoanNo, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_LoanNoPopup, pvStartRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_PaySeq, pvStartRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PaySeq+1, pvStartRow, .vspdData.MaxRows  
		.vspdData.ReDraw = True
    
    End With
End Sub

'=======================================  GetSpreadColumnPos()  ===================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)

    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"

            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PayType		= iCurColumnPos(1)
			C_PayTypePopup	= iCurColumnPos(2)
			C_PayTypeNm		= iCurColumnPos(3)
			C_PayDocAmt		= iCurColumnPos(4)
			C_PayLocAmt		= iCurColumnPos(5)
			C_ExchRate		= iCurColumnPos(6)
			C_BankAcct		= iCurColumnPos(7)
			C_BankAcctPopup	= iCurColumnPos(8)
			C_BankCd		= iCurColumnPos(9)
			C_BankPopup		= iCurColumnPos(10)
			C_BankNm		= iCurColumnPos(11)
			C_Noteno		= iCurColumnPos(12)
			C_NotenoPopup	= iCurColumnPos(13)
			C_PrepayNo		= iCurColumnPos(14)
			C_PrepayNoPopup	= iCurColumnPos(15)
			C_LoanNo		= iCurColumnPos(16)
			C_LoanNoPopup	= iCurColumnPos(17)
			C_PaySeq		= iCurColumnPos(18)
    End Select    
End Sub
'==========================================  2.2.6 InitCollectType()  =======================================
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("A1006", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iRateArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectType(UBound(iCodeArr) - 1, 1)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectType(i, 0) = iCodeArr(i)
		arrCollectType(i, 1) = iRateArr(i)
	Next

End Sub	

'------------------------------------------  OpenIvNo()  -------------------------------------------------
Function OpenIvNo()
	
	Dim strRet
	Dim arrParam(0)
	Dim iCalledAspName
		
		If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = "PROTECTED" Then Exit Function
		
		lblnWinEvent = True
		arrParam(0) = "ST"  'ivType
		
		iCalledAspName = AskPRAspName("m5111pa1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m5111pa1", "X")
			lgIsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lblnWinEvent = False
	
		If strRet(0) = "" Then
			frm1.txtIvNo.focus
			Exit Function
		Else
			frm1.txtIvNo.value = strRet(0)
			frm1.txtIvNo.focus	
			Set gActiveElement = document.activeElement
		End If	
		
End Function

'------------------------------------------  OpenNoteNo()  -------------------------------------------------
Function OpenNoteNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	if interface_Account = "N" then
		Exit Function
	End if
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "지급어음번호"	
	arrParam(1) = "F_NOTE A, B_BANK B"
	frm1.vspdData.Col = C_Noteno
	arrParam(2) = Trim(frm1.vspdData.text)
	
	arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & " AND A.NOTE_FG = " & FilterVar("D3", "''", "S") & " AND  A.BANK_CD = B.BANK_CD  "
	arrParam(4) = arrParam(4) & " AND BP_CD =  " & FilterVar(frm1.txtPayeeCd.Value, "''", "S") & " "
	frm1.vspdData.Col = C_PayLocAmt		 
 	If Trim(frm1.vspdData.Text) <> "" Then
 		arrParam(4) = arrParam(4) & " AND A.NOTE_AMT = convert(numeric, " & FilterVar(UNICDbl(Trim(frm1.vspdData.Text)), "''", "S") & ")"
 	End If

	arrParam(5) = "지급어음번호"			
	
 	arrField(0) = "A.Note_NO"						' Field명(0)
	arrField(1) = "B.BANK_NM"						' Field명(1)
	arrField(2) = ""	
    
	arrHeader(0) = "지급어음번호"				' Header명(0)
	arrHeader(1) = "발행은행"					' Header명(1)
	arrHeader(2) = ""							' Header명(2)

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_Noteno:		frm1.vspdData.Text = arrRet(0)
	End If	

End Function
'------------------------------------------  OpenPpNo()  -------------------------------------------------
Function OpenPpNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim TmpPayType

	if interface_Account = "N" then
		Exit Function
	End if
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "선급금번호"	
	arrParam(1) = "F_PRPAYM, B_MINOR"
	frm1.vspdData.Col = C_PrepayNo
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	arrParam(2) = Trim(frm1.vspdData.text)
	
	frm1.vspdData.Col = C_PayType
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	TmpPayType = Trim(frm1.vspdData.text)
	

	arrParam(4) = "DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  And BP_CD =  " & FilterVar(frm1.txtPayeeCd.Value, "''", "S") & "  AND BAL_AMT > 0 AND F_PRPAYM.CONF_FG = " & FilterVar("C", "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND B_MINOR.MINOR_CD = F_PRPAYM.CONF_FG AND B_MINOR.MAJOR_CD = " & FilterVar("F1012", "''", "S") & " "
	'arrParam(4) = arrParam(4) & " AND PRPAYM_TYPE = " & FilterVar(TmpPayType, "''", "S") & " "

	arrParam(5) = "선급금번호"			
	
    arrField(0) = "PRPAYM_NO"
    arrField(1) = "F2" & parent.gColSep & "PRPAYM_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "BAL_AMT"
    arrField(4) = "F2" & parent.gColSep & "BAL_LOC_AMT"  

    arrField(5) = "F5" & parent.gColSep & "XCH_RATE" 

    
    arrHeader(0) = "선급금번호"		
    arrHeader(1) = "선급금"		
    arrHeader(2) = "선급금화폐"
    arrHeader(3) = "선급금잔액"
    arrHeader(4) = "선급금잔액(자국)"    
    arrHeader(5) = "환율"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_PrepayNo:		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ExchRate:		frm1.vspdData.Text = arrRet(5)	
		Call ChangeCurOrDt(frm1.vspdData.Row)
		'Call vspdData_Change(C_PrepayNo , frm1.vspdData.Row)
	End If	

End Function
	
'------------------------------------------  OpenAcctNo()  -------------------------------------------------
Function OpenAcctNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	if interface_Account = "N" then
		Exit Function
	End if
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계좌번호"			' 팝업 명칭 
	arrParam(1) = "B_BANK A, F_DPST B, B_BANK_ACCT C"			' TABLE 명칭 
	frm1.vspdData.Col = C_BankAcct
	arrParam(2) = Trim(frm1.vspdData.text)		' Code Condition
	arrParam(3) = ""							' Name Cindition
	
	arrParam(4) = "A.BANK_CD = B.BANK_CD  AND B.DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  " 
	arrParam(4) = arrParam(4) & " AND  B.BANK_CD = C.BANK_CD AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "	' Where Condition
			frm1.vspdData.Col = C_BankCd		 
 			If Trim(frm1.vspdData.Text) <> "" Then
 				arrParam(4) = arrParam(4) & " AND A.BANK_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "  "
 			End If
	arrParam(4) = arrParam(4) & " AND (B.DPST_FG = " & FilterVar("SV", "''", "S") & " OR B.DPST_FG = " & FilterVar("ET", "''", "S") & ") " '예금, 기타 
	arrParam(5) = "계좌번호"				' 조건필드의 라벨 명칭 %>
	
	arrField(0) = "B.BANK_ACCT_NO"				' Field명(0)
	arrField(1) = "B.BANK_CD"					' Field명(1)
	arrField(2) = "A.BANK_NM"					' Field명(2)
    
	arrHeader(0) = "계좌번호"				' Header명(0)
	arrHeader(1) = "은행코드"				' Header명(1)
	arrHeader(2) = "은행명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_BankAcct:			frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_BankCd:			frm1.vspdData.Text = arrRet(1)
		frm1.vspdData.Col = C_BankNm:			frm1.vspdData.Text = arrRet(2)
	End If	
	
End Function

'------------------------------------------  OpenPayType()  -------------------------------------------------
Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급유형"	
	arrParam(1) = "B_CONFIGURATION A,  B_MINOR B"
	arrParam(2) =  Trim(frm1.vspdData.text)
	arrParam(3) = ""
	arrParam(4) = "B.MAJOR_CD  = " & FilterVar("A1006", "''", "S") & " AND B.MAJOR_CD = A.MAJOR_CD AND B.MINOR_CD = A.MINOR_CD "	
	arrParam(4) = arrParam(4) & " AND A.SEQ_NO = " & FilterVar("1", "''", "S") & "  AND A.REFERENCE LIKE " & FilterVar("%P%", "''", "S") & "	"
	arrParam(5) = "지급유형"			
	
	
    arrField(0) = "B.MINOR_CD"	
    arrField(1) = "B.MINOR_NM"	
    
    arrHeader(0) = "지급유형"		
    arrHeader(1) = "지급유형명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_PayType:		    frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_PayTypeNm:		frm1.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_PayType , frm1.vspdData.Row)
	End If	
	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenBank()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	if interface_Account = "N" then
		Exit Function
	End if
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "은행"	
	arrParam(1) = "B_BANK A, F_DPST B"
	arrParam(2) =  Trim(frm1.vspdData.text)
	arrParam(3) = ""
'	arrParam(4) = ""			
	arrParam(5) = "은행"			

	arrParam(4) = "A.BANK_CD = B.BANK_CD  AND B.DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  "  	' Where Condition
	frm1.vspdData.Col = C_BankAcct		 
 	If Trim(frm1.vspdData.Text) <> "" Then
 		arrParam(4) = arrParam(4) & " AND B.BANK_ACCT_NO =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "  "
 	End If
 				
    arrField(0) = "A.BANK_CD"	
    arrField(1) = "A.BANK_NM"	
    
    arrHeader(0) = "은행"		
    arrHeader(1) = "은행명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_BankCd:		    frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_BankNm:		    frm1.vspdData.Text = arrRet(1)
	End If	
	
End Function

'------------------------------------------  OpenLoanNo()  -------------------------------------------------
Function OpenLoanNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	if interface_Account = "N" then
		Exit Function
	End if

	If lblnWinEvent = True Then Exit Function

	if Trim(frm1.txtCur.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "화폐","X")
		Exit Function
	end if
	
	lblnWinEvent = True

	arrParam(0) = "차입금번호"	
	arrParam(1) = "F_LN_INFO"
	
	arrParam(2) = Trim(frm1.vspdData.text)
	
	arrParam(4) = "DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  AND LOAN_BAL_AMT > 0"
	arrParam(5) = "차입금번호"			
	
    arrField(0) = "LOAN_NO"
    arrField(1) = "F2" & parent.gColSep & "LOAN_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "LOAN_BAL_AMT"
    
    arrHeader(0) = "차입금번호"		
    arrHeader(1) = "차입금"		
    arrHeader(2) = "차입금화폐"
    arrHeader(3) = "차입금잔액"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.text = arrRet(0)
		lgBlnFlgChgValue 	= True
		Call vspdData_Change(C_LoanNo , frm1.vspdData.Row)
	End If	

End Function
 '------------------------------------------  OpenGLRef()  -------------------------------------------------
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnGlNo.value)
	'arrParam(1) = Trim(frm1.txtIvNo.value)
	'arrParam(2) = Trim(frm1.txtGrpCd.value)
	'arrParam(3) = Trim(frm1.txtGrpNm.value)

	If frm1.hdnGlType.Value = "A" Then
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "T" Then
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")
    End if
        

	lblnWinEvent = False
	
End Function
'====================================  vspdData_MouseDown()  ===================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'=====================================  FncSplitColumn()  =====================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
    
End Function
'=====================================  vspdData_Click()  =====================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
	If UCase(Trim(frm1.txtPost.Value)) = "Y" Or lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
   
   gMouseClickStatus = "SPC"   
   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
		
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If   
End Sub
'=====================================  SetSpreadLockAfterQuery()  =====================================
Function SetSpreadLockAfterQuery()
 
	Dim index 
    Dim sPayType
    
	    With frm1
		
		.vspdData.ReDraw = False
		
	    For index = 1 to .vspdData.MaxRows 
			
			.vspdData.Row = index
			if frm1.rdoApFlg(0).Checked= true  then 'frm1.hdnPostingFlg.Value = "Y" then    '확정되었으면 
			
				'ggoSpread.SpreadLock		C_PaySeq, index, .vspdData.MaxCols, index
			else

				ggoSpread.SpreadUnLock 	C_PayType , index, C_PayType , index
				ggoSpread.SSSetRequired C_PayType , index, index
    
				ggoSpread.SpreadLock	C_PayTypeNm , index, C_PayTypeNm, index
    
				ggoSpread.SpreadUnLock 	C_PayDocAmt , index, C_PayDocAmt, index
				ggoSpread.SSSetRequired C_PayDocAmt , index, index  
				'ggoSpread.SpreadLock 	C_PayLocAmt ,  index, C_ExchRate , index
				
				ggoSpread.SSSetProtected C_PaySeq, index, index
				ggoSpread.SSSetRequired	 C_PayType, index, index

				ggoSpread.SSSetProtected C_PayTypeNm, index, index
				ggoSpread.SSSetRequired	 C_PayDocAmt, index, index		
				'ggoSpread.SSSetProtected C_PayLocAmt, index, index
				ggoSpread.SSSetProtected C_ExchRate, index, index
				ggoSpread.SSSetProtected C_LoanNo, index, index	'차입금 
				ggoSpread.SSSetProtected C_LoanNoPopup, index, index
				ggoSpread.SSSetProtected C_PaySeq + 1, index, index
				
				frm1.vspdData.Col = C_PayType    '지급유형 
	    	    sPayType = CheckPayType(Frm1.vspdData.text)
				
			    If  sPayType <> "" Then
					sPayType = CheckPayType(Frm1.vspdData.text)
                    '초기화 
					ggoSpread.spreadlock 	C_Noteno, index,C_NotenoPopup,index
					ggoSpread.SSSetProtected C_Noteno, index, index
					ggoSpread.SSSetProtected C_NotenoPopup, index, index

					ggoSpread.spreadlock 	C_BankAcct, index,C_BankAcctPopup,index
					ggoSpread.SSSetProtected C_BankAcct, index, index
					ggoSpread.SSSetProtected C_BankAcctPopup, index, index

					ggoSpread.spreadlock 	C_PrepayNo, index,C_PrepayNoPopup,index
					ggoSpread.SSSetProtected C_PrepayNo, index, index
					ggoSpread.SSSetProtected C_PrepayNoPopup, index, index

					if sPayType = "NO" then   '지급어음 
						ggoSpread.spreadUnlock 	C_Noteno, index, C_NotenoPopup, index
						ggoSpread.SSSetRequired	C_Noteno, index,index

					Elseif sPayType = "DP" then   '예적금 
						ggoSpread.spreadUnlock 	C_BankAcct, index,C_BankAcctPopup, index
						ggoSpread.SSSetRequired	C_BankAcct, index,index
					Elseif sPayType = "PP" then   '선급금경우 
						ggoSpread.spreadUnlock 	C_PrepayNo, index,C_PrepayNoPopup, index
						ggoSpread.SSSetRequired	C_PrepayNo, index,index
					end if
					if sPayType = "DP" then  '예적금일경우 
						ggoSpread.spreadUnlock 	C_BankCd, index,C_BankPopup, index
						ggoSpread.SSSetRequired	C_BankCd, index,index
					else
						ggoSpread.spreadlock 	C_BankCd, index,C_BankPopup, index
					end if
					
					if sPayType <> "PP" and sPayType <> "NO" and sPayType <> "DP" then
						ggoSpread.spreadUnlock 	C_LoanNo , index, C_LoanNoPopup , index  '차입금 
					end if
				else
					ggoSpread.spreadUnlock 	C_Noteno, index, C_NotenoPopup, index
					ggoSpread.spreadUnlock 	C_BankAcct, index,C_BankAcctPopup,index
					ggoSpread.spreadUnlock 	C_PrepayNo, index,C_PrepayNoPopup,index
					ggoSpread.spreadUnlock 	C_BankCd, index, C_BankPopup, index
					ggoSpread.spreadUnlock 	C_LoanNo , index, C_LoanNoPopup , index  '차입금 
				end if
	
			end if    
		Next
	
		.vspdData.ReDraw = True
			
	End With
End Function
'============================================  2.5.1 TotalSum()  ======================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================

Sub TotalSum()

	Dim SumDocAmttotal,SumLocAmttotal, lRow		
	SumDocAmttotal = 0
	SumLocAmttotal = 0
	ggoSpread.source = frm1.vspdData
    For lRow = 1 To frm1.vspdData.MaxRows 		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text <> ggoSpread.DeleteFlag then
			frm1.vspdData.Col = C_PayDocAmt    '지급금액 
			SumDocAmtTotal = SumDocAmtTotal + UNICDbl(frm1.vspdData.Text)
				
			frm1.vspdData.Col = C_PayLocAmt    '지급자국금액 
			SumLocAmtTotal = SumLocAmtTotal + UNICDbl(frm1.vspdData.Text)
		end if
	Next
	frm1.txtPayDocAmt.Text = UNIConvNumPCToCompanyByCurrency(SumDocAmtTotal, frm1.txtCur.value, parent.ggAmtOfMoneyNo,"X","X")
	frm1.txtPayLocAmt.Text =  UNIFormatNumber(SumLocAmtTotal, ggAmtOfMoney.DecPoint, -2, 0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
End Sub

'============================================  2.5.1 CheckPayType()  ======================================
'========================================================================================================
Function CheckPayType(PayType)
    Dim iRow
	For iRow = 0 To UBound(arrCollectType,1)
	    If arrCollectType(iRow,0) = PayType and arrCollectType(iRow,1) <> "" Then
	       CheckPayType = arrCollectType(iRow,1)
	       Exit Function
	    End if
	Next
    CheckPayType = ""
    
 End Function
'=====================================  Posting()  =====================================
Sub Posting()
    Dim IntRetCD 
    
    Err.Clear                                                         '☜: Protect system from crashing
    
    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217","X","X","X")                      '데이타가 변경되었습니다. 저장후 진행하십시오.
		Exit sub
	End if    
    if ggoSpread.SSCheckChange = True then
		Call DisplayMsgBox("189217","X","X","X")
		Exit sub
	End if
	
	if Trim(frm1.txtPostDt.text) = "" then
		Call DisplayMsgBox("17A002","X" , "매입일","X")
		Exit Sub
	End if

    if frm1.rdoApFlg(0).Checked= true  then 'frm1.hdnPostingFlg.Value = "Y" then    '확정되었으면 
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")

		If IntRetCD = vbNo Then
			Exit Sub
		End If
		
		Call DbSave("Posting")

	Elseif frm1.rdoApFlg(1).Checked= true  then  'frm1.hdnPostingFlg.Value = "N" then

		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		
		Call DbSave("Posting")
		
	End if
	
End Sub

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	With frm1   
	
		'매입금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'총지급금액 
		ggoOper.FormatFieldByObjectOfCur .txtPayDocAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec

	End With

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'지급금액 
		ggoSpread.SSSetFloatByCellOfCur C_PayDocAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec

    End With

End Sub

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()

		Call LoadInfTB19029												 '⊙: Load table , B_numeric_format 
	    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,parent.ggStrMinPart,parent.ggStrMaxPart)
		Call ggoOper.LockField(Document, "N")							'⊙: Lock  Suitable  Field 
		Call SetDefaultVal
		Call InitSpreadSheet											'⊙: Setup the Spread sheet 	
		Call InitVariables	
		Call CookiePage(0)
		Call InitCollectType
End Sub
'*********************************************  3.3 Object Tag 처리  ************************************
Sub txtPostDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPostDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPostDt.focus
	End If	
End Sub

Sub txtPostDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'==========================================================================================
'   Event Name : ChangeCurOrDt()
'   Event Desc : 환율,지급자국금액 값변경 (지급금액,선수금변경시 호출)
'==========================================================================================
Function ChangeCurOrDt(Byval LRow)

    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    Dim Cur,ChargeDt,DocAmt,ExchRate
    
    With frm1
			
		.vspdData.Row = LRow '.vspdData.ActiveRow		
		 Cur = Trim(frm1.txtCur.value)	       '화폐 
		 ChargeDt =Trim(frm1.txtIvDt.value)    '매입등록일 

		.vspdData.Col = C_PayDocAmt            '지급금액 
		 DocAmt = .vspdData.Text
		 
		If Cur = "" or ChargeDt = "" then
			Exit Function
		End If
		
		If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then    '화폐가 KRW이면 지급자국금액 = 지급금액 * 1		
			.vspdData.Col = C_ExchRate
			.vspdData.Text = "1"
			.vspdData.Col = C_PayLocAmt                      '지급자국금액 
			.vspdData.Text = DocAmt
			Call TotalSum
			Exit Function
		End If

   		'strVal = BIZ_PGM_ID & "?txtMode=" & "LookupDailyExRt"	
		'strVal = strVal & "&Currency=" & Cur
		'strVal = strVal & "&ChargeDt=" & ChargeDt
		'strVal = strVal & "&LRow=" & LRow
		
		.vspdData.Col = C_ExchRate
		ExchRate = .vspdData.Text
		
		.vspdData.Col = C_PayLocAmt
		If Trim(.hdnDiv.value) = "*" Then
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) * UNICDbl(ExchRate),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		Elseif Trim(frm1.hdnDiv.value) = "/" Then
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) / UNICDbl(ExchRate),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		End If
		Call TotalSum
							
    End With
	
    'if LayerShowHide(1) = false then
	'	exit Function
	'end if
    
	'Call RunMyBizASP(MyBizASP, strVal)
        
End Function
'==========================================================================================
'   Event Name : ChangeCurOrDt2()
'   Event Desc : 환율,지급자국금액 값변경 (지급금액,선수금변경시 호출)
'==========================================================================================	
Function ChangeCurOrDt2(Byval LRow)

    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    Dim Cur,ChargeDt,DocAmt,XchRt
    
    With frm1
			
		.vspdData.Row = LRow '.vspdData.ActiveRow		
		 Cur = Trim(frm1.txtCur.value)	       '화폐 
		 ChargeDt =Trim(frm1.txtIvDt.value)    '매입등록일 
		 XchRt = Trim(frm1.txtXchRt.value)

		.vspdData.Col = C_PayDocAmt            '지급금액 
		 DocAmt = .vspdData.Text
		 
		If Cur = "" or ChargeDt = "" then
			Exit Function
		End If
		
		If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then    '화폐가 KRW이면 지급자국금액 = 지급금액 * 1
						
			.vspdData.Col = C_ExchRate
			.vspdData.Text = "1"
			.vspdData.Col = C_PayLocAmt                      '지급자국금액 
			.vspdData.Text = DocAmt
			Call TotalSum
			Exit Function
		End If
		
		.vspdData.Col = C_ExchRate
		.vspdData.Text = XchRt		
		
		.vspdData.Col = C_PayLocAmt
		If Trim(.hdnDiv.value) = "*" Then
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) * UNICDbl(XchRt),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		Elseif Trim(frm1.hdnDiv.value) = "/" Then
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) / UNICDbl(XchRt),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		End If
		Call TotalSum	

    End With
    
        
End Function
'==========================================  3.3.1 vspdData_Change()  ===================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	
	Dim LocAmt, DocAmt ,SumDocAmt
	Dim sPayType
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
		
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

	lgBlnFlgChgValue = True
    
	Select Case col
	Case C_PayType                     '지급유형 
		frm1.vspdData.ReDraw = false
    	frm1.vspdData.Col = C_PayType

		If 	CommonQueryRs(" B.MINOR_NM ", " B_CONFIGURATION A,  B_MINOR B ", _
								 "B.MAJOR_CD  = " & FilterVar("A1006", "''", "S") & " AND B.MAJOR_CD = A.MAJOR_CD AND B.MINOR_CD = A.MINOR_CD " & _
								 " AND A.SEQ_NO = " & FilterVar("1", "''", "S") & "  AND A.REFERENCE LIKE " & FilterVar("%P%", "''", "S") & " " & _
								 " AND B.MINOR_CD = " & FilterVar(frm1.vspdData.text, "''", "S") & " " , _
								 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("17A003","X","지급유형","X")
			Call frm1.vspdData.SetText(C_PayType, frm1.vspdData.Row, "")
			Call frm1.vspdData.SetText(C_PayTypeNm, frm1.vspdData.Row, "")
			Exit Sub
		End If
		lgF0 = Split(lgF0, Chr(11))
		Call frm1.vspdData.SetText(C_PayTypeNm, frm1.vspdData.Row, lgF0(0))

    	sPayType = CheckPayType(frm1.vspdData.text)

		If  sPayType <> "" Then
			
			if sPayType = "NO"  then   '지급어음인경우 어음번호는 필수입력 
			 	ggoSpread.spreadUnlock 	C_Noteno, frm1.vspdData.ActiveRow,C_NotenoPopup, frm1.vspdData.ActiveRow
				ggoSpread.SSSetRequired	C_Noteno, frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
				frm1.vspdData.col = C_BankAcct  '계좌번호은 ""
				frm1.vspdData.text = ""
				frm1.vspdData.col = C_PrepayNo  '선급금번호은 ""
				frm1.vspdData.text = ""
				ggoSpread.spreadlock 	C_BankAcct, frm1.vspdData.ActiveRow,C_BankAcctPopup,frm1.vspdData.ActiveRow					
				ggoSpread.spreadlock 	C_PrepayNo, frm1.vspdData.ActiveRow,C_PrepayNoPopup,frm1.vspdData.ActiveRow
			elseif sPayType = "CK" then   '수표인경우 
			    ggoSpread.spreadUnlock 	C_Noteno, frm1.vspdData.ActiveRow,C_NotenoPopup, frm1.vspdData.ActiveRow
				frm1.vspdData.col = C_BankAcct  '계좌번호은 ""
				frm1.vspdData.text = ""
				frm1.vspdData.col = C_PrepayNo  '선급금번호은 ""
				frm1.vspdData.text = ""
				ggoSpread.spreadlock 	C_BankAcct, frm1.vspdData.ActiveRow,C_BankAcctPopup,frm1.vspdData.ActiveRow					
				ggoSpread.spreadlock 	C_PrepayNo, frm1.vspdData.ActiveRow,C_PrepayNoPopup,frm1.vspdData.ActiveRow
			    ggoSpread.SSSetProtected	C_NotenoPopup, Row, Row  '지급유형이 수표일때는 어음번호필드는 열리나  번호팝업버튼은 작동하지 않음 
			elseif sPayType = "DP"  then   '예적금인경우는 계좌번호는 필수입력 
			 	ggoSpread.spreadUnlock 	C_BankAcct, frm1.vspdData.ActiveRow,C_BankAcctPopup, frm1.vspdData.ActiveRow
				ggoSpread.SSSetRequired	C_BankAcct, frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
				frm1.vspdData.col = C_Noteno
				frm1.vspdData.text = ""
				frm1.vspdData.col = C_PrepayNo
				frm1.vspdData.text = ""
				ggoSpread.spreadlock 	C_Noteno, frm1.vspdData.ActiveRow,C_NotenoPopup,frm1.vspdData.ActiveRow					
				ggoSpread.spreadlock 	C_PrepayNo, frm1.vspdData.ActiveRow,C_PrepayNoPopup,frm1.vspdData.ActiveRow

			elseif sPayType = "PP" then   '선급금경우인경우 선급금번호 필수입력 
			 	ggoSpread.spreadUnlock 	C_PrepayNo, frm1.vspdData.ActiveRow,C_PrepayNoPopup, frm1.vspdData.ActiveRow
				ggoSpread.SSSetRequired	C_PrepayNo, frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
				frm1.vspdData.col = C_Noteno
				frm1.vspdData.text = ""
				frm1.vspdData.col = C_BankAcct
				frm1.vspdData.text = ""
				ggoSpread.spreadlock 	C_Noteno, frm1.vspdData.ActiveRow,C_NotenoPopup,frm1.vspdData.ActiveRow					
				ggoSpread.spreadlock 	C_BankAcct, frm1.vspdData.ActiveRow,C_BankAcctPopup,frm1.vspdData.ActiveRow					

			ElseIf sPayType = "CS" Then   '현금인경우는 계좌번호,어음번호,선급금번호는 lock

				frm1.vspdData.Col = C_Noteno
				frm1.vspdData.Text = ""
				ggoSpread.spreadlock 	C_Noteno, frm1.vspdData.ActiveRow,C_NotenoPopup, frm1.vspdData.ActiveRow					
					
				frm1.vspdData.Col = C_BankAcct
				frm1.vspdData.Text = ""
				ggoSpread.spreadlock 	C_BankAcct, frm1.vspdData.ActiveRow,C_BankAcctPopup,frm1.vspdData.ActiveRow

				frm1.vspdData.Col = C_PrepayNo
				frm1.vspdData.Text = ""
				ggoSpread.spreadlock 	C_PrepayNo, frm1.vspdData.ActiveRow,C_PrepayNoPopup,frm1.vspdData.ActiveRow

			end if
	
			frm1.vspdData.Col = C_PayType
			if sPayType = "DP" then  '예적금일경우 은행은 필수입력 
				ggoSpread.spreadUnlock 	C_BankCd, frm1.vspdData.ActiveRow,C_BankPopup, frm1.vspdData.ActiveRow
				ggoSpread.SSSetRequired	C_BankCd, frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
			else
				frm1.vspdData.Col = C_BankCd
				frm1.vspdData.Text = ""
				frm1.vspdData.Col = C_BankNm
				frm1.vspdData.Text = ""
				ggoSpread.spreadlock 	C_BankCd, frm1.vspdData.ActiveRow,C_BankPopup, frm1.vspdData.ActiveRow
			end if	
			'***2003.1월 패치************************
			if sPayType <> "PP" and sPayType <> "NO" and sPayType <> "DP" then
				ggoSpread.spreadUnlock 	C_LoanNo, frm1.vspdData.ActiveRow,C_LoanNoPopup,frm1.vspdData.ActiveRow '차입금 
			else
				frm1.vspdData.Col = C_LoanNo
				frm1.vspdData.Text = ""
				ggoSpread.spreadlock 	C_LoanNo, frm1.vspdData.ActiveRow,C_LoanNoPopup,frm1.vspdData.ActiveRow
			end if
			'***************************************
		Else
			frm1.vspdData.Col = C_Noteno
			frm1.vspdData.Text = ""
			ggoSpread.spreadUnlock 	C_Noteno, frm1.vspdData.ActiveRow,C_NotenoPopup, frm1.vspdData.ActiveRow					
					
			frm1.vspdData.Col = C_BankAcct
			frm1.vspdData.Text = ""
			ggoSpread.spreadUnlock 	C_BankAcct, frm1.vspdData.ActiveRow,C_BankAcctPopup,frm1.vspdData.ActiveRow

			frm1.vspdData.Col = C_Noteno
			frm1.vspdData.Text = ""
			ggoSpread.spreadUnlock 	C_PrepayNo, frm1.vspdData.ActiveRow,C_PrepayNoPopup,frm1.vspdData.ActiveRow
					
			frm1.vspdData.Col = C_BankCd
			frm1.vspdData.Text = ""
			frm1.vspdData.Col = C_BankNm
			frm1.vspdData.Text = ""
			ggoSpread.spreadUnlock 	C_BankCd, frm1.vspdData.ActiveRow,C_BankPopup, frm1.vspdData.ActiveRow
			'***2003.1월 패치************************
			frm1.vspdData.Col = C_LoanNo
			frm1.vspdData.Text = ""
			ggoSpread.spreadUnlock 	C_LoanNo, frm1.vspdData.ActiveRow,C_LoanNoPopup,frm1.vspdData.ActiveRow '차입금 
			'***************************************
		End If
		frm1.vspdData.ReDraw = true
	Case C_PayDocAmt  '지급금액 
           
        SumDocAmt = frm1.txtPayDocAmt.Text
            
        frm1.vspdData.Col = C_PayDocAmt
			
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			DocAmt = 0
		Else
			DocAmt = UNICDbl(frm1.vspdData.Text)
		End If  'gCurrency
			
			
		Dim LgPayType
		frm1.vspdData.Col = C_PayType
		LgPayType = frm1.vspdData.Text 
			
		 frm1.vspdData.Col = C_ExchRate

		If  LgPayType <> "PP" Then		'선급금인 경우를 제외하고는 매입헤더의 환율을 가져옴.
		  Call ChangeCurOrDt2(Row)
		ElseIf (LgPayType = "PP" and Trim(frm1.vspdData.Text) <> "") Then  '선급금일경우는 환율이 있어야 함.
		  Call ChangeCurOrDt(Row)
		Else
		  Call TotalSum	
		End If  
		
	Case C_PrepayNo
		frm1.vspdData.Col = C_PayType
		SPayType = Trim(frm1.vspdData.text)
		frm1.vspdData.Col = C_PrepayNo

		If 	CommonQueryRs(" A.XCH_RATE ", " F_PRPAYM A, B_MINOR B ", _
								 " A.DOC_CUR = " & FilterVar(frm1.txtCur.Value, "''", "S") & " AND A.BP_CD = " & FilterVar(frm1.txtPayeeCd.Value, "''", "S") & _
								 " AND A.BAL_AMT > 0 AND A.CONF_FG = " & FilterVar("C", "''", "S") & " AND B.MINOR_CD = A.CONF_FG AND B.MAJOR_CD = " & FilterVar("F1012", "''", "S") & _
								 " AND A.PAYM_TYPE = " & FilterVar(SPayType, "''", "S") & " AND A.PRPAYM_NO = " & FilterVar(Trim(frm1.vspdData.text), "''", "S") & " " , _
								 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("17A003","X","선급금번호","X")
			Call frm1.vspdData.SetText(C_PrepayNo, frm1.vspdData.Row, "")
			Call frm1.vspdData.SetText(C_ExchRate, frm1.vspdData.Row, "")
			Call frm1.vspdData.SetText(C_PayLocAmt, frm1.vspdData.Row, "")
			Exit Sub
		End If

		lgF0 = Split(lgF0, Chr(11))
		Call frm1.vspdData.SetText(C_ExchRate, frm1.vspdData.Row, lgF0(0))
		Call ChangeCurOrDt(Row)
		
	End Select
End Sub
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				DbQuery
			End If
		End If
	End With
End Sub

'===========================================  vspdData_ButtonClicked()  ===========================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
    
     If Row = 0 Or .MaxRows = 0 Then 
          Exit Sub
     End If

   
    If Row > 0 Then
        
        .Col = Col - 1
        .Row = Row
        
        Select Case Col
        
        	Case C_PayTypePopup      '지급유형 팝업 
        		Call OpenPayType()
        	Case C_NotenoPopup       '지급어음 
        	    frm1.vspdData.Col = C_PayType
        	    Call OpenNoteNo()
        	Case C_PrepayNoPopup     '선급금 
        	    frm1.vspdData.Col = C_PayType
        	    Call OpenPpNo()
        	Case C_BankAcctPopup     ' 계좌번호 
        	    frm1.vspdData.Col = C_PayType
        	    Call OpenAcctNo()
        	Case C_BankPopup
        		Call OpenBank()
        	Case C_LoanNoPopup
        		Call OpenLoanNo()	
       End Select
       
    Else
    	Exit sub
    End If
    
    End With
End Sub

'===============================================  vspdData_ScriptDragDropBlock()  ===================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================  vspdData_TopLeftChange()  ==================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey <> "" Then	
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If

		End If
	End if    
End Sub
'==========================================  vspdData_ColWidthChange()  ==========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'=========================================  5.1.1 FncQuery()  ===========================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													 '⊙: Processing is NG 

	Err.Clear															 '☜: Protect system from crashing 

	ggoSpread.Source = frm1.vspdData
		
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	'⊙: "Will you destory previous data" 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	Call InitVariables											'⊙: Initializes local global variables 


	If Not chkField(Document, "1") Then					'⊙: This function check indispensable field 
		Exit Function
	End If
	frm1.txtQueryType.Value = "Query"

	If DbQuery = False Then Exit Function								'☜: Query db data 

	FncQuery = True														'⊙: Processing is OK 
End Function
'===========================================  5.1.2 FncNew()  ===========================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False                                                         '⊙: Processing is NG															<% '☜: Protect system from crashing %>
	ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")	'⊙: "Will you destory previous data" 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")									'⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")									'⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
	'Call SetToolBar("1110100000101")
		
	Call SetDefaultVal

	FncNew = True															'⊙: Processing is OK

End Function
'===========================================  5.1.3 FncDelete()  ========================================
Function FncDelete()
		
	ggoSpread.Source = frm1.vspdData
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then								 'Check if there is retrived data 
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	If DbDelete = False Then Exit Function							 '☜: Delete db data 

	FncDelete = True												'⊙: Processing is OK 
	Call TotalSum
End Function

'===========================================  5.1.4 FncSave()  ==========================================
Function FncSave()
	Dim IntRetCD
	Dim Strval

	FncSave = False		
	Err.Clear	
	ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 	
	
	if CheckRunningBizProcess = true then
		exit function
	end if
		
	If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
		Exit Function
	End If
	    
	ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
	If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
	   Exit Function
	End If	
				
	if UNICDbl(frm1.txtDocAmt.Text) > 0 and UNICDbl(frm1.txtPayDocAmt.Text) < 0 then
	    IntRetCD = DisplayMsgBox("175520","X","X","X")            '⊙: Display Message(There is no changed data.)
	    Exit Function
	end if
	'차입금 금액은 체크하지 않음  LC Usance를 처리하기 위해 Usance Clearence(차입금을 위한 가계정)를 사용하기때문 
	if unicdbl(frm1.txtDocAmt.Value) < 0  and unicdbl(frm1.txtPayDocAmt.Value) < 0 then  '반품일 경우 비교대상들이 음수면 
		if abs(unicdbl(frm1.txtDocAmt.Value)) < abs(unicdbl(frm1.txtPayDocAmt.Value)) then  '총지급금액이 매입금액을 초과한 경우 
			IntRetCD = DisplayMsgBox("175520","X","X","X")            '⊙: Display Message(There is no changed data.)
			Exit Function
		End If
	Else
        if unicdbl(frm1.txtDocAmt.Value) < unicdbl(frm1.txtPayDocAmt.Value) then  '총지급금액이 매입금액을 초과한 경우 
			IntRetCD = DisplayMsgBox("175520","X","X","X")            '⊙: Display Message(There is no changed data.)
			Exit Function
		End If
	End If
	
	If frm1.hdnPostDt.value <> frm1.txtPostDt.value then    '매입일을 변경했을때 저장 
	   	strVal = BIZ_PGM_ID & "?txtMode=" & "PostDtUpdate"	
		strVal = strVal & "&IvNo=" & frm1.txtIvNo.value
		strVal = strVal & "&PostDt=" & frm1.txtPostDt.text

		if LayerShowHide(1) = false then
			exit Function
		end if
    
		Call RunMyBizASP(MyBizASP, strVal)
		If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged 에러메세지 없이 종료 
			Exit Function
		End If
	End If		

	'------ Save function call area ------ 
	If DbSave("ToolBar")  = False Then Exit Function
		
	If frm1.txthdnIvNo.value <> frm1.txtIvNo.value then
			frm1.txtIvNo.value =	frm1.txthdnIvNo.value		
	End If															 '☜: Save db data 
		
	FncSave = True													'⊙: Processing is OK 
    'Call vspdData_Change(C_PayDocAMt , frm1.vspdData.Row)
End Function
'===========================================  5.1.5 FncCopy()  ==========================================
Function FncCopy()
	'메시지 부분 삭제(2003.06.25)		
	if frm1.vspdData.Maxrows < 1	then exit function
	
	ggoSpread.Source = Frm1.vspdData
	ggoSpread.CopyRow
	frm1.vspdData.ReDraw = False
	
	Call SetSpreadColor(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
	
	frm1.vspdData.ReDraw = True
	
	Frm1.VspdData.Focus
		
	If Err.number = 0 Then	
	   FncCopy = True                                                            '☜: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement   
End Function

'===========================================  5.1.6 FncCancel()  ========================================
Function FncCancel() 
    if frm1.vspdData.Maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    Call TotalSum
End Function
'==========================================  5.1.7 FncInsertRow()  ======================================
Function FncInsertRow(ByVal pvRowCnt) 
 	Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End If
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
	    ggoSpread.Source = .vspdData
	    
	    ggoSpread.InsertRow .vspdData.ActiveRow, imRow

	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    .vspdData.ReDraw = True

    End With
	
	Set gActiveElement = document.ActiveElement
	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
		
End Function
'==========================================  5.1.8 FncDeleteRow()  ======================================
Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
	
	if frm1.vspdData.Maxrows < 1	then exit function
		
	With frm1.vspdData 
	
		.focus
		ggoSpread.Source = frm1.vspdData

		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
	End With
    Call TotalSum
End Function
'============================================  5.1.9 FncPrint()  ========================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncPrint()
End Function

'============================================  5.1.10 FncPrev()  ========================================
Function FncPrev() 
	 '------ Precheck area ------ 
	ggoSpread.Source = frm1.vspdData
	If lgIntFlgMode <> parent.OPMD_UMODE Then							'Check if there is retrived data 
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgPrevNo = "" Then								'Check if there is retrived data 
		Call DisplayMsgBox("900011","X","X","X")
	End If
End Function
'============================================  5.1.11 FncNext()  ========================================
Function FncNext()
	 '------ Precheck area ------ 
	ggoSpread.Source = frm1.vspdData
	If lgIntFlgMode <> parent.OPMD_UMODE Then						 'Check if there is retrived data 
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgNextNo = "" Then								 'Check if there is retrived data 
		Call DisplayMsgBox("900012","X","X","X")
	End If
End Function

'===========================================  5.1.12 FncExcel()  ========================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function
'===========================================  5.1.13 FncFind()  =========================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function
'===========================================  PopSaveSpreadColumnInf()  ================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'===========================================  PopRestoreSpreadColumnInf()  ================================
Sub PopRestoreSpreadColumnInf()

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	
	if UCase(Trim(frm1.txtPost.Value)) = "Y" then
		call SetRdSpreadColor(1)
	Else			
		if UNICDbl(Trim(frm1.txtDocAmt.Text)) <> 0 then	'iv detail이 존재하면 확정가능 	    
			Call SetSpreadLockAfterQuery()	
		else
			call SetRdSpreadColor(1)     '전체 lock
		End if	
	End if
	
End Sub

'===========================================  5.1.14 FncExit()  =========================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")		'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function
'=============================================  5.2.1 DbQuery()  ========================================
Function DbQuery()
	Dim strVal

	Err.Clear													'☜: Protect system from crashing

	DbQuery = False												'⊙: Processing is NG

	If lgIntFlgMode = parent.OPMD_UMODE Then		
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtIvNo=" & frm1.txthdnIvNo.value	'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtQuerytype=" & frm1.txtQuerytype.value
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtQuerytype=" & frm1.txtQuerytype.value
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If
		
	if LayerShowHide(1) = false then
		exit Function
	end if

	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True														'⊙: Processing is NG
End Function

'=============================================  5.2.2 DbSave()  =========================================
Function DbSave(byval btnflg) 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	Dim intIndex
	Dim ColSep, RowSep
	Dim PvArr
		
	DbSave = False														 '⊙: Processing is OK 
    
	On Error Resume Next												 '☜: Protect system from crashing 
	Err.Clear
		
	With frm1
		.txtMode.value = parent.UID_M0002
			
		lGrpCnt = 0
		strVal = ""
    
		if btnflg = "Posting" then
			if unicdbl(frm1.txtDocAmt.Value) < unicdbl(frm1.txtPayDocAmt.Value) then  '총지급금액이 매입금액을 초과한 경우 
			IntRetCD = DisplayMsgBox("175520","X","X","X")            '⊙: Display Message(There is no changed data.)
			Exit Function
		End If
			.txtMode.value = "Release" 				'☜: 확정 버튼 
		end if

		ReDim PvArr(.vspdData.MaxRows)
			
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
			    
		    if .vspdData.Text = ggoSpread.InsertFlag or .vspdData.Text = ggoSpread.UpdateFlag or .vspdData.Text = ggoSpread.DeleteFlag then
			    
		        Select Case .vspdData.Text
		            Case ggoSpread.InsertFlag											'☜: 신규 
						strVal = "C" & parent.gColSep				'☜: C=Create
		            Case ggoSpread.UpdateFlag											'☜: 신규 
						strVal = "U" & parent.gColSep				'☜: U=Update
		            Case ggoSpread.DeleteFlag											'☜: 삭제 
						strVal = "D" & parent.gColSep				'☜: D=Delete
		        End Select
			    	
		    	.vspdData.Col = C_PayDocAmt
				if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
					Call DisplayMsgBox("970021","X","지급금액","X")
					Call LayerShowHide(0)
					Exit Function
				End if
				'****2003.1월 패치**********************************
				.vspdData.Col = C_PayType:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PayTypeNm:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PayDocAmt:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep
				.vspdData.Col = C_PayLocAmt:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep
				.vspdData.Col = C_ExchRate:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep
				.vspdData.Col = C_BankAcct:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_BankCd:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_BankNm:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_Noteno:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PrepayNo:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_LoanNo:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PaySeq:		strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				'==================================
				strVal = strVal & lRow & parent.gRowSep
			        
		        PvArr(lGrpCnt) = strVal
		        lGrpCnt = lGrpCnt + 1
			        
			end if        
		Next
			
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = Join(PvArr, "")
			
		if lGrpCnt > 0 or btnflg = "Posting" then
			if LayerShowHide(1) = false then
				exit Function
			end if
		    .hdninterface_Account.value = interface_Account

			Call ExecMyBizASP(frm1, BIZ_PGM_ID)							'☜: 비지니스 ASP 를 가동 

		End if
	
	End With

	DbSave = True														'⊙: Processing is NG 
End Function
'=============================================  5.2.4 DbQueryOk()  ======================================
Function DbQueryOk()													 '☆: 조회 성공후 실행로직 
	Dim index
	
	lgIntFlgMode = parent.OPMD_UMODE											'⊙: Indicates that current mode is Update mode 
	lgBlnFlgChgValue = False

    Call TotalSum					'금액합계 
	
	'**수정(2003.03.26)-회계모듈이 없어도 확정,확정취소 가능하도록 수정함.
	if UCase(Trim(frm1.txtPost.Value)) = "Y" then

		Call SetToolBar("11100000000111")
		call SetRdSpreadColor(1)
		frm1.btnPosting.value = "확정취소"
		if interface_Account <> "N" then
			frm1.btnGlSel.disabled = false
		Else
			frm1.btnGlSel.disabled = True
		End If
		ggoOper.SetReqAttr	frm1.txtPostDt, "Q"     '매입일 수정불가 
		'요기 볼것~
		if UNICDbl(Trim(frm1.txtDocAmt.Text)) <> 0 then	'iv detail이 존재하면 확정가능 
			frm1.btnPosting.Disabled = False
		else
			frm1.btnPosting.Disabled = True
		End if

	Else
		frm1.btnPosting.value = "확정"
	    frm1.btnGlSel.disabled = true
		Call SetToolBar("11101111001111")		
		if UNICDbl(Trim(frm1.txtDocAmt.Text)) <> 0 then	'iv detail이 존재하면 확정가능 
		    ggoOper.SetReqAttr	frm1.txtPostDt, "D"   'N: REQUIRED, D: UNREQUIRED ,Q:PROTECTED
			frm1.btnPosting.Disabled = False
			Call SetSpreadLockAfterQuery()	

		else
			ggoOper.SetReqAttr	frm1.txtPostDt, "Q"
			frm1.btnPosting.Disabled = True
			call SetRdSpreadColor(1)     '전체 lock
		End if	
	End if
    if frm1.hdnGlType.Value = "A" Then
       frm1.btnGlSel.value = "회계전표조회"
    elseif frm1.hdnGlType.Value = "T" Then
       frm1.btnGlSel.value = "결의전표조회"
    end if		
   
End Function
	
'==========================================================================================
'   Event Name : ChkExistIvDtlByIvNo
'   Event Desc : Call at Biz. Logic (2005.04 KJH) 
'==========================================================================================
Function ChkExistIvDtlByIvNo(CurIvNo)
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	Err.Clear 
	
	ChkExistIvDtlByIvNo = False 
	
	If 	CommonQueryRs(" COUNT(IV_NO) ", " M_IV_DTL ", "IV_NO = " & FilterVar(CurIvNo, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		'Call DisplayMsgBox("175200","X","X","X")
		frm1.hdnIvDtlMaxRows.Value = 0
		Exit function
	End If
		
	lgF0 = Split(lgF0, Chr(11))
		
	frm1.hdnIvDtlMaxRows.Value = lgF0(0) 	
	
    ChkExistIvDtlByIvNo = True 
End Function
'=============================================  5.2.5 DbSaveOk()  =======================================
Function DbSaveOk()														'☆: 저장 성공후 실행 로직 
	Call InitVariables
	'frm1.vspdData.MaxRows = 0
	lgBlnFlgChgValue = False
	Call MainQuery()
End Function
'============================================================================================================
' Name : SubGetGlNo
' Desc : Get Gl_no : 2003.03 KJH 전표번호 가져오는 로직 수정 
'============================================================================================================
Sub SubGetGlNo()
	Dim lgStrFrom
	Dim strTempGlNo, strGlNo
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	On Error Resume Next
	Err.Clear 
	
	lgStrFrom =  " ufn_a_GetGlNo( " & FilterVar(frm1.txthdnIvNo.Value, "''", "S") & " )"
	
	Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", lgStrFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 <> "" then 
		strTempGlNo = Split(lgF0, Chr(11))
		strGlNo		= Split(lgF1, Chr(11))
					
		If strGlNo(0) = "" and strTempGlNo(0) = "" then 
			frm1.hdnGlNo.Value		=   ""
			frm1.txtGlNo.Value		=	""
			frm1.hdnGlType.value	=	"B"
		Elseif strGlNo(0) = "" and strTempGlNo(0) <> "" then
			frm1.hdnGlNo.Value		=   strTempGlNo(0) 
			frm1.txtGlNo.Value		=	strTempGlNo(0) 
			frm1.hdnGlType.value	=	"T"
		Elseif strGlNo(0) <> "" then 
			frm1.hdnGlNo.Value		=   strGlNo(0) 
			frm1.txtGlNo.Value		=	strGlNo(0) 
			frm1.hdnGlType.value	=	"A"
		End If
	Else
		frm1.hdnGlNo.Value		=   ""
		frm1.txtGlNo.Value		=	""
		frm1.hdnGlType.value	=	"B"
	End if
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>지급내역</font></TD>
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
										<TD CLASS=TD5 NOWRAP>매입번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=29 MAXLENGTH=18 TAG="12XXXU" ALT="매입번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvNo()"></TD>
										<TD CLASS=TD6>&nbsp;</TD>
										<TD CLASS=TD6>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>공급처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpplCd" SIZE=10 MAXLENGTH=10 TAG="24XXXU">
														 <INPUT TYPE=TEXT NAME="txtSpplNm" SIZE=22 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>B/L번호.IV번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLIvNo" SIZE=34 tag="24XXXU" ALT="B/L번호.IV번호"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>매입형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvTypeCd" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="매입형태">
														 <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=22 MAXLENGTH=20 TAG="24X2" ALT="매입형태" ></TD>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGrpCd" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="구매그룹">
														 <INPUT TYPE=TEXT NAME="txtGrpNm" SIZE=22 TAG="24"></TD>

								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>화폐</TD>
									<TD CLASS=TD6 NOWRAP><!--INPUT TYPE=TEXT NAME="txtCur" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐단위"-->
							                             <!--OBJECT ALT=환율 TYPE=TEXT NAME="txtXchRt"  classid=<%=gCLSIDFPDS%> id=fpDoubleSingle1  STYLE="HEIGHT: 20px; WIDTH: 150px" TAG="24X2" Title="FPDOUBLESINGLE" ></OBJECT-->													
							                             
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP><INPUT TYPE=TEXT NAME="txtCur" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐단위">&nbsp;
												</TD>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환율 TYPE=TEXT NAME="txtXchRt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 STYLE="HEIGHT: 19px; WIDTH: 164px" TAG="24X5" Title="FPDOUBLESINGLE" ></OBJECT>');</SCRIPT>													
												</TD>
											
											</TR>
										</Table>				                             
							        </TD>                     
							                             
									<TD CLASS=TD5 NOWRAP>매입등록일</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvDt" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="매입등록일"></TD>


								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>매입금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="매입금액" TYPE=TEXT NAME="txtDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 STYLE="HEIGHT: 20px; WIDTH: 248px" TAG="24X2" Title="FPDOUBLESINGLE" ></OBJECT>');</SCRIPT></TD>		
		
		
									<TD CLASS=TD5 NOWRAP>매입자국금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="매입자국금액" TYPE=TEXT NAME="txtLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 STYLE="HEIGHT: 20px; WIDTH: 248px" TAG="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>

								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총지급금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="총지급금액" TYPE=TEXT NAME="txtPayDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 STYLE="HEIGHT: 20px; WIDTH: 248px" TAG="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>총지급자국금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="총지급원화금액" TYPE=TEXT NAME="txtPayLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 STYLE="HEIGHT: 20px; WIDTH: 248px" TAG="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>

								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>매입일</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="매입일" NAME="txtPostDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT> 
												</TD>
												<TD NOWRAP>
													&nbsp;<INPUT TYPE=TEXT NAME="txtGlNo"  STYLE="HEIGHT: 21px; WIDTH: 144px " MAXLENGTH=10 TAG="24XXXU" ALT="전표번호">												
												</TD>
											
											</TR>
										</Table>	

									</TD>
									<TD CLASS="TD5" nowrap>확정여부</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=radio NAME="rdoApFlg" ALT="확정여부" CLASS="RADIO" tag="24X"><label for="rdoApFlg">&nbsp;Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio NAME="rdoApFlg" ALT="확정여부" CLASS="RADIO" checked tag="24X"><label for="rdoApFlg">&nbsp;No&nbsp;</label></TD>


								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TD WIDTH=10>&nbsp;</TD>
<!--				<td align="Left"><Div ID="btnintAcc"><button name="btnPostingSel" id="btnPosting" class="clsmbtn" ONCLICK="Posting()">확정</button><Div></td> -->
            <td> 
			   <BUTTON NAME="btnPosting" CLASS="CLSSBTN"  ONCLICK="Posting()">확정처리</BUTTON>&nbsp;
			   <BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
			</td>

					<TD WIDTH="*" ALIGN=RIGHT><a href="VBSCRIPT:CookiePage(1)">매입세금계산서</a>|<a href="VBSCRIPT:CookiePage(2)">매입내역등록</a>|<a href="VBSCRIPT:CookiePage(3)">B/L등록</a>|<a href="VBSCRIPT:CookiePage(4)">B/L내역등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txthdnIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtQuerytype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPost" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvDtlMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPostDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPostingFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLoanAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdninterface_Account" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBlNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPayeeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>

