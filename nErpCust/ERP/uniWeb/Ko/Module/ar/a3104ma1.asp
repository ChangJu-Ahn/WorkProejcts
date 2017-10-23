
<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a3104ma1
'*  4. Program Name         : 가수금정보 등록 
'*  5. Program Desc         : 가수금정보 등록 수정 삭제 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/10/02
'*  8. Modified date(Last)  : 2002/11/26
'*  9. Modifier (First)     : Hee Jung, Kim
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'*********************************************************************************************************
'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">           </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
Const BIZ_PGM_QUERY_ID = "a3104mb1.asp"							'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID  = "a3104mb2.asp"							'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID   = "a3104mb3.asp"							'☆: Head Query 비지니스 로직 ASP명 

Const RcptJnlType = "SR"

Const TAB1 = 1													'☜: Tab의 위치 
Const TAB2 = 2

Dim C_Rcpttype_Cd 
Dim C_Rcpttype_Pb 
Dim C_Rcpttype_Nm 
Dim C_Rcptacct_Cd 
Dim C_Rcptacct_Pb 
Dim C_Rcptacct_Nm 
Dim C_NetRcptAmt  
Dim C_NetRcptLocAmt
Dim C_NoteNo   
Dim C_NoteNoPop
Dim C_BankAcct 
Dim C_BankAcctPop
Dim C_hiddDtlSeq 
Dim C_RcptItem_Desc


Dim IsOpenPop						' Popup
Dim gSelframeFlg
Dim	lgFormLoad
Dim	lgQueryOk						' Queryok여부 (loc_amt =0 check)
Dim lgQueryState					' 조회후 상태 flag
Dim lgstartfnc

<%
Dim dtToday
dtToday = GetSvrDate
%> 


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************





'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables()
	C_Rcpttype_Cd   = 1
	C_Rcpttype_Pb   = 2
	C_Rcpttype_Nm   = 3
	C_Rcptacct_Cd   = 4
	C_Rcptacct_Pb   = 5
	C_Rcptacct_Nm   = 6
	C_NetRcptAmt    = 7
	C_NetRcptLocAmt = 8
	C_NoteNo        = 9
	C_NoteNoPop     = 10
	C_BankAcct      = 11
	C_BankAcctPop   = 12
	C_hiddDtlSeq    = 13
	C_RcptItem_Desc	= 14
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE						'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False								'Indicates that no value changed
    lgIntGrpCount = 0										'initializes Group View Size
    lgStrPrevKey = ""										'initializes Previous Key
    lgLngCurRows = 0										'initializes Deleted Rows Count

	lgstartfnc = False
	lgFormLoad = True
	lgQueryOk  = False
End Sub
 
'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtRcptDt.Text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDocCur.value = parent.gCurrency

	frm1.txtXchRate.text = 1
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
	frm1.txtRcptNo.focus
 
	lgBlnFlgChgValue = False
	lgQueryOk = False	
	lgQueryState = False
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>    
End Sub

'========================================================================================
' Name : InitComboBoxConfFg()
' Description : Combo Display for Confirm status.
'========================================================================================
Sub InitComboBoxConfFg()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
End Sub

'========================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()

    With frm1.vspdData
		.MaxCols = C_RcptItem_Desc + 1								' 마지막 상수명 사용 
		.Col = .MaxCols											'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
		.MaxRows = 0

		ggoSpread.Source = frm1.vspdData
		.Redraw = False		

		ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit     C_Rcpttype_Cd,  "입금유형"       ,10,,,10,2           
		ggoSpread.SSSetButton   C_Rcpttype_Pb
		ggoSpread.SSSetEdit     C_Rcpttype_Nm,  "입금유형명"     ,16
		ggoSpread.SSSetEdit     C_Rcptacct_Cd,  "입금계정코드"   ,12,,,20,2
		ggoSpread.SSSetButton   C_Rcptacct_Pb
		ggoSpread.SSSetEdit     C_Rcptacct_Nm,  "입금계정코드명" ,30
		ggoSpread.SSSetFloat    C_NetRcptAmt,   "입금액"         ,19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat    C_NetRcptLocAmt,"입금액(자국)"   ,19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetEdit     C_NoteNo,       "어음번호"       ,30,,,30,2
		ggoSpread.SSSetButton   C_NoteNoPop
		ggoSpread.SSSetEdit     C_BankAcct,     "은행계좌번호"   ,30,,,30,2
		ggoSpread.SSSetButton   C_BankAcctPop
    	ggoSpread.SSSetEdit     C_Rcptitem_Desc,"비고"           ,20,,,20	

		
		Call ggoSpread.MakePairsColumn(C_Rcpttype_Cd,C_Rcpttype_Pb)
		Call ggoSpread.MakePairsColumn(C_Rcptacct_Cd,C_Rcptacct_Pb)
		Call ggoSpread.MakePairsColumn(C_NetRcptAmt,C_NetRcptLocAmt)				
		Call ggoSpread.MakePairsColumn(C_NoteNo,C_NoteNoPop)
		Call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPop)
		
		Call ggoSpread.SSSetColHidden(C_hiddDtlSeq,C_hiddDtlSeq,True)

 	    Call SetSpreadLock
    End With
End Sub

'========================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    Dim objSpread

    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData

		objSpread.Redraw = False
		    
		ggoSpread.SpreadLock C_Rcpttype_Cd, -1, C_Rcpttype_Cd, -1
		ggoSpread.SpreadLock C_Rcpttype_Nm, -1, C_Rcpttype_Nm, -1				
'		ggoSpread.SpreadLock C_Rcpttype_Pb, -1, C_Rcpttype_Pb, -1
'		ggoSpread.SpreadLock C_Rcptacct_Pb, -1, C_Rcptacct_Pb, -1
		ggoSpread.SpreadLock C_Rcptacct_Nm, -1, C_Rcptacct_Nm, -1                            
		ggoSpread.SpreadLock C_NoteNo     , -1, C_NoteNo, -1                            
		ggoSpread.SpreadLock C_BankAcct   , -1, C_BankAcct, -1                            
		
		ggoSpread.SSSetRequired  C_Rcpttype_Cd, -1, -1 
		ggoSpread.SSSetRequired  C_Rcptacct_Cd, -1, -1
		ggoSpread.SSSetRequired  C_NetRcptAmt , -1, -1		

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'========================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData 
		.Redraw = False
    
		ggoSpread.Source = frm1.vspdData
    
		ggoSpread.SSSetRequired  C_RcptType_Cd, pvStartRow, pvEndRow          
		ggoSpread.SSSetProtected C_RcptType_Nm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_RcptAcct_Cd, pvStartRow, pvEndRow          
		ggoSpread.SSSetProtected C_RcptAcct_Nm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_NetRcptAmt , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_NoteNo     , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BankAcct   , pvStartRow, pvEndRow

		.Redraw = True
    End With
End Sub

'======================================================================================================
' Function Name : GetSpreadColumnPos()
' Function Desc : This method Call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_Rcpttype_Cd   = iCurColumnPos(1)
			C_Rcpttype_Pb   = iCurColumnPos(2)
			C_Rcpttype_Nm   = iCurColumnPos(3)
			C_Rcptacct_Cd   = iCurColumnPos(4) 
			C_Rcptacct_Pb   = iCurColumnPos(5)
			C_Rcptacct_Nm   = iCurColumnPos(6)
			C_NetRcptAmt    = iCurColumnPos(7)
			C_NetRcptLocAmt = iCurColumnPos(8)
			C_NoteNo        = iCurColumnPos(9)
			C_NoteNoPop     = iCurColumnPos(10)
			C_BankAcct      = iCurColumnPos(11)
			C_BankAcctPop   = iCurColumnPos(12)
			C_hiddDtlSeq    = iCurColumnPos(13)
			C_Rcptitem_Desc = iCurColumnPos(14)
	End select
End Sub

'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8) 
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5120ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
 
	arrParam(0) = Trim(frm1.txtGlNo.value)					'회계전표번호 
	arrParam(1) = ""										'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	  
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
End Function

'======================================================================================================
' Function Name : OpenPopupTempGL
' Function Desc : This method Open The Popup window for TempGL
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(8) 
    Dim iCalledAspName
    
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
 
	arrParam(0) = Trim(frm1.txtTempGlNo.value)				'회계전표번호 
	arrParam(1) = ""										'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	  
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtRcptDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDept.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtRcptDt.text = arrRet(3)
				Call txtDept_OnBlur()  
				frm1.txtDept.focus
        End Select
	End With
End Function     
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
		Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If arrRet(0) = "" Then
		Call EscPopup(iWhere)
		Exit Function
	Else  
		Call SetReturnVal(arrRet,iWhere)
	End If 	
End Function
'=========================================================================================================
' Name : Open???()
' Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'      ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopup(Byval strCode, Byval iWhere )
	Dim arrRet
	Dim arrParam(8), arrField(6), arrHeader(6)
	Dim strNoteFg

	If IsOpenPop = True Then Exit Function
 
	Select Case iWhere
		Case 0
		Case 2
			If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function 

			arrParam(0) = "거래처팝업" 
			arrParam(1) = "B_BIZ_PARTNER"
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래처코드"
 
			arrField(0) = "BP_CD" 
			arrField(1) = "BP_NM"
			 
			arrHeader(0) = "거래처코드"  
			arrHeader(1) = "거래처명" 
		Case 3    
			If IsOpenPop = True Or UCase(frm1.txtDocCur.className) = "PROTECTED" Then Exit Function
		 
			arrParam(0) = "거래통화팝업" 
			arrParam(1) = "B_CURRENCY"    
			arrParam(2) = Trim(frm1.txtDocCur.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래통화"
 
			arrField(0) = "CURRENCY" 
			arrField(1) = "CURRENCY_DESC" 
			 
			arrHeader(0) = "거래통화"  
			arrHeader(1) = "거래통화명" 
		Case 4    
			frm1.vspdData.Col = C_Rcpttype_Cd

			Dim strWhere 
			
			strWhere = "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND B_CONFIGURATION.SEQ_NO = 3 AND  B_CONFIGURATION.REFERENCE = " & FilterVar("PR", "''", "S") & "  "
			strWhere = strWhere & "AND  MINOR_CD= " & FilterVar(UCase(frm1.vspdData.Text), "''", "S") & ""


			If CommonQueryRs( "MINOR_CD" , "B_CONFIGURATION" , strWhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
				
				Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
					Case "NR"
						arrParam(0) = "어음번호팝업"													' 팝업 명칭 
						arrParam(1) = "f_note a,b_biz_partner b, b_bank c"									' TABLE 명칭 
						frm1.vspdData.Col = C_NoteNo
						arrParam(2) = Trim(frm1.vspdData.text)												' Code Condition
						arrParam(3) = ""		
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & "  And a.bp_cd = b.bp_cd And a.bank_cd *= c.bank_cd"					' Where Condition										
						arrParam(5) = "어음번호"														' 조건필드의 라벨 명칭 
							
						arrHeader(0) = "어음번호"														' Header명(0)
						arrHeader(1) = "금액"															' Header명(1)
						arrHeader(2) = "발행일"															' Header명(1)     
						arrHeader(3) = "거래처"															' Header명(1)
						arrHeader(4) = "은행"															' Header명(1)						
							
						arrField(0) = "Note_no"																' Field명(0)
						arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"									' Field명(1)
						arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"									' Field명(2)
						arrField(3) = "b.bp_nm"
						arrField(4) = "c.bank_nm"         						
					Case "CR"
						arrParam(0) = "수취구매카드 팝업"											' 팝업 명칭 
						arrParam(1) = "f_note a,b_biz_partner b, b_bank c , b_card_co d "					' TABLE 명칭 
						frm1.vspdData.Col = C_NoteNo					
						arrParam(2) = Trim(frm1.vspdData.text)												' Code Condition
						arrParam(3) = ""						
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CR", "''", "S") & "  And a.bp_cd = b.bp_cd And a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "		' Where Condition   						
						arrParam(5) = "수취구매카드번호"												' 조건필드의 라벨 명칭						               
							
						arrHeader(0) = "수취구매카드번호"												' Header명(0)
						arrHeader(1) = "금액"															' Header명(1)
						arrHeader(2) = "발행일"															' Header명(1)     
						arrHeader(3) = "거래처"															' Header명(1)
						arrHeader(4) = "카드사"															' Header명(1)						
							
						arrField(0) = "Note_no"																' Field명(0)
						arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"									' Field명(1)
						arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"									' Field명(2)
						arrField(3) = "b.bp_nm"
						arrField(4) = "d.card_co_nm"         						
					Case Else
						Exit Function
				End Select
			ENd if				
		Case 5
			arrParam(0) = "계좌번호팝업"
			arrParam(1) = "F_DPST, B_BANK"    
			arrParam(2) = Trim(frm1.vspdData.text)
			arrParam(3) = ""
			   
			arrParam(4) = "F_DPST.BANK_CD = B_BANK.BANK_CD "
			arrParam(5) = "계좌번호"   
		 
			arrField(0) = "F_DPST.BANK_ACCT_NO" 
			arrField(1) = "F_DPST.BANK_CD" 
			arrField(2) = "B_BANK.BANK_NM" 
			   
			arrHeader(0) = "계좌번호"  
			arrHeader(1) = "은행" 
			arrHeader(2) = "은행명"       
		Case 6    
			arrParam(0) = "입금유형"								' 팝업 명칭 
		 
			arrParam(1) = "B_MINOR,B_CONFIGURATION "
			arrParam(2) = Trim(frm1.vspdData.text)
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
			   & "AND B_CONFIGURATION.SEQ_NO = 3 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PR", "''", "S") & " " ' Where Condition        
			arrParam(5) = "입금유형"								' TextBox 명칭 
	 
			arrField(0) = "B_MINOR.MINOR_CD"							' Field명(0)
			arrField(1) = "B_MINOR.MINOR_NM"							' Field명(1)
			  
			arrHeader(0) = "입금유형"								' Header명(0)
			arrHeader(1) = "입금유형명"								' Header명(1) 
		Case 7
			If frm1.txtRcptType.className = parent.UCN_PROTECTED Then Exit Function
		 
			arrParam(0) = frm1.txtRcptType.Alt							' 팝업 명칭 
			arrParam(1) = "a_jnl_item"									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtRcptType.Value)					' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "jnl_type =  " & FilterVar(RcptJnlType  , "''", "S") & ""			' Where Condition
			arrParam(5) = frm1.txtRcptType.Alt							' 조건필드의 라벨 명칭 

			arrField(0) = "JNL_CD"										' Field명(0)
			arrField(1) = "JNL_NM"										' Field명(1)
			 
			arrHeader(0) = frm1.txtRcptType.Alt							' Header명(0)
			arrHeader(1) = frm1.txtRcptTypeNm.Alt						' Header명(1)
		Case 8 '입금계정코드 
			arrParam(0) = "계정코드팝업"							' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"    ' TABLE 명칭 
			arrParam(2) = ""											' Code Condition
			arrParam(3) = ""											' Name Condition
		 
			frm1.vspdData.Col = C_Rcpttype_Cd
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			    " and C.trans_type = " & FilterVar("ar001", "''", "S") & "  and C.jnl_cd = " & FilterVar(frm1.vspdData.Text, "''", "S")         ' Where Condition
			arrParam(5) = "계정코드"								' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
			   arrField(2) = "B.GP_CD"									' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
		 
			arrHeader(0) = "계정코드"								' Header명(0)
			arrHeader(1) = "계정코드명"								' Header명(1)
			arrHeader(2) = "그룹코드"								' Header명(2)
			arrHeader(3) = "그룹명"									' Header명(3)   
	End Select
 
	IsOpenPop = True
 
	If iWhere = 0 Then
	
	
		Dim iCalledAspName
		iCalledAspName = AskPRAspName("a3104ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3104ra1", "X")
			IsOpenPop = False
			Exit Function
		End If

		' 권한관리 추가 
		arrParam(5) = lgAuthBizAreaCd
		arrParam(6) = lgInternalCd
		arrParam(7) = lgSubInternalCd
		arrParam(8) = lgAuthUsrID
			
			arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam), _
		       "dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")     
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   
	End If 

	IsOpenPop = False
 
	If arrRet(0) = "" Then
		Call EscPopup(iWhere)
		Exit Function
	Else  
		Call SetReturnVal(arrRet,iWhere)
	End If 
End Function
'===========================================================================================
' Name : SetReturnVal()
' Description : Plant Popup에서 Return되는 값 setting
'===========================================================================================
Function SetReturnVal(byval arrRet,byval iWhere)
	With frm1 
		Select Case iWhere   
			Case 0
				.txtRcptNo.Value     = arrRet(0)   
				.txtRcptNo.focus  
			Case 2 'OpenBpCd
				.txtBpCd.Value       = arrRet(0)
				.txtBpNm.Value       = arrRet(1)
				.txtBpCd.focus
			Case 3 'OpenCurrency
				.txtDocCur.Value     = arrRet(0)
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 4 
				.vspdData.Col        = C_NoteNo
				.vspdData.Text       = arrRet(0)   
				.vspdData.Col        = C_NetRcptAmt
				.vspdData.Text       = arrRet(1)     
				.vspdData.Col        = C_NetRcptLocAmt
				.vspdData.Text       = arrRet(1)             ' 어음인 경우 금액 => 자국금액 
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
				Call SetActiveCell(frm1.vspdData,C_NoteNo,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 5
				.vspdData.Col        = C_BankAcct
				.vspdData.Text       = arrRet(0) 
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
				Call SetActiveCell(frm1.vspdData,C_BankAcct,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 6
				.vspdData.Col        = C_Rcpttype_Nm
				.vspdData.Text       = arrRet(1)        
				.vspdData.Col        = C_Rcpttype_Cd
				.vspdData.Text		 = arrRet(0)   
				Call subVspdSettingChange(C_Rcpttype_Cd, frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow, arrRet(0) )
				Call vspdData_Change(C_Rcpttype_Cd, .vspdData.Row)
				Call SetActiveCell(frm1.vspdData,C_Rcpttype_Cd,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 7 'OpenBpCd
				.txtRcptType.Value   = arrRet(0)
				.txtRcptTypeNm.Value = arrRet(1)
				.txtRcptType.focus
			Case 8
				.vspdData.Col        = C_Rcptacct_Cd
				.vspdData.Text       = arrRet(0)   
				.vspdData.Col        = C_Rcptacct_Nm
				.vspdData.Text       = arrRet(1)        
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)  
				Call SetActiveCell(frm1.vspdData,C_Rcptacct_Cd,frm1.vspdData.ActiveRow ,"M","X","X")
		End Select 

		If iWhere <> 0 Then lgBlnFlgChgValue = True
	End With
End Function
'===========================================================================================
' Name : EscPopup()
' Description : Plant Popup에서 Return되는 값 setting
'===========================================================================================
Function EscPopup(iWhere)
	With frm1 
		Select Case iWhere   
			Case 0
				.txtRcptNo.focus  
			Case 2 'OpenBpCd
				.txtBpCd.focus
			Case 3 'OpenCurrency
				.txtDocCur.focus
			Case 4 
				Call SetActiveCell(frm1.vspdData,C_NoteNo,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 5
				Call SetActiveCell(frm1.vspdData,C_BankAcct,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 6
				Call SetActiveCell(frm1.vspdData,C_Rcpttype_Cd,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 7 'OpenBpCd
				.txtRcptType.focus
			Case 8
				Call SetActiveCell(frm1.vspdData,C_Rcptacct_Cd,frm1.vspdData.ActiveRow ,"M","X","X")
		End Select 

		If iWhere <> 0 Then lgBlnFlgChgValue = True
	End With
End Function

'===========================================================================================
' 기능: Tab Click
' 설명: Tab Click시 필요한 기능을 수행한다.
'=========================================================================================== 
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function

	If lgQueryState = True Then
		Call SetToolbar("1111110100001111")    
	Else		
		Call SetToolbar("1110110100001111")    
	End If		
	Call changeTabs(TAB1)  
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	If lgQueryState = True Then
		Call SetToolbar("1111111100101111")    
	Else
		Call SetToolbar("1110110100101111")    				
	End If		
	Call changeTabs(TAB2)  
	gSelframeFlg = TAB2
End Function





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************



'=====================================================================================================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=====================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field

    Call InitSpreadSheet()														'⊙: Setup the Spread sheet
	Call InitVariables()   
    Call SetDefaultVal()
	Call SetToolbar("1110110100001111")    
	'Call InitComboBoxConfFg() 
 
	gIsTab = "Y"
	gTabMaxCnt = 2 

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
		
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim var1
    
    FncQuery = False                                                        
    lgstartfnc = True
    Err.Clear                                                               
	'-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
  
    If lgBlnFlgChgValue = True Or var1 = True Then  
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables()														'⊙: Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Query function Call area
    '-----------------------
    Call DbQuery()																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    lgstartfnc = False	 
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
	Dim var1
     
    FncNew = False                                                          
    lgstartfnc = True 
	lgQueryState = False
	    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                               '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                               '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                '⊙: Lock  Suitable  Field

    Call InitVariables()   
    Call InitSpreadsheet()	
	Call ClickTab1()																'sstData.Tab = 1
    
    Call SetDefaultVal()
    Call txtDocCur_OnChange()
        
    lgBlnFlgChgValue = False 
    FncNew = True																'⊙: Processing is OK
    lgFormLoad = True							' tempgldt read
    lgstartfnc = False
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False															'⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                   'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")					'☜ 바뀐부분 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")			'☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If
	'-----------------------
    'Delete function Call area
    '-----------------------
    Call DbDelete()																'☜: Delete db data
    
    FncDelete = True	
    		
	Set gActiveElement = document.activeElement    
															'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
	Dim var1
 
    FncSave = False                                                         
    
    Err.Clear                                                               

    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False Then							'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")					'⊙: Display Message(There is no changed data.)
		Exit Function
    End If
    
    If Not chkField(Document, "2") Then											'⊙: Check required field(Single area)
		Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                 '⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If
 
	If frm1.vspdData.MaxRows < 1 then
		IntRetCD = DisplayMsgBox("112526","x","x","x")					'가수금내역 상세정보가 입력되어 있지 않습니다 
		Exit Function
	End if
	'-----------------------
    'Save function Call area
    '-----------------------
    Call DbSave()																	'☜: Save db data
    
    FncSave = True  
 		
	Set gActiveElement = document.activeElement    
	
 End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False

    If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow , frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_RcptType_Cd
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_RcptType_Nm
    frm1.vspdData.Text = ""
    
    Call Dosum()
	frm1.vspdData.ReDraw = True
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    if frm1.vspdData.MaxRows < 1 then Exit Function

	ggoSpread.Source = frm1.vspdData 
	ggoSpread.EditUndo                                                
 
	Call DoSum()
			
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos

	If gSelframeFlg <> TAB2 Then
		Call ClickTab2()																'sstData.Tab = 1
	End If
	   
	FncInsertRow = False																'☜: Processing is NG	   

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		
	   
	With frm1
		iCurRowPos = .vspdData.ActiveRow	
		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData
	 
		ggoSpread.InsertRow ,imRow
		
		For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
			MaxSpreadVal frm1.vspdData, C_hiddDtlSeq, ii
		Next        
		.vspdData.Col = 1																' 컬럼의 절대 위치로 이동 
		.vspdData.Row = ii - 1
		.vspdData.Action = 0		
		.vspdData.ReDraw = True

		Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)
	End With
	
    If Err.number = 0 Then
       FncInsertRow = True																'☜: Processing is OK
    End If   
    		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 

    If gSelframeFlg <> TAB2 Then
		Call ClickTab2()										'sstData.Tab = 1
    End If

    If frm1.vspdData.MaxRows < 1 then Exit Function
    
	'----------  Coding part  ------------------------------------------------------------- 
    lDelRows = ggoSpread.DeleteRow
    
    Call DoSum()
    		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	parent.FncPrint()   
			
	Set gActiveElement = document.activeElement    
	 
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call FncExport(parent.C_SINGLEMULTI)            '☜: 화면 유형 
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                                                    
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
 
	iColumnLimit = 5
 
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
 
	If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit : frm1.vspdData.Row = 0            
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.Vspddata.text), "X")
		Exit Function
	End If
 
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
 
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	Dim var1
 
	FncExit = False

	ggoSpread.Source = frm1.vspdData
	var1 = ggoSpread.SSCheckChange
	   
	If lgBlnFlgChgValue = True or var1 = True Then											'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")					'데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
			
	Set gActiveElement = document.activeElement    
	
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************





'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    DbDelete = False																		'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)							'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	    
	Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True																			'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()																		'☆: 삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")													'Clear Condition Field
	Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field    
	Call InitVariables()																	'Initializes local global variables
	Call SetDefaultVal()
			       
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	frm1.txtRcptNo.Value = ""
	frm1.txtRcptNo.focus
	
	Call ClickTab1()
	Call SetToolbar("1110110100001111")   
	
	lgBlnFlgChgValue = False    
End Function
 
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '⊙: Processing is NG
	Err.Clear
     
	Call LayerShowHide(1)
 
    Dim strVal

	If lgIntFlgMode = parent.OPMD_UMODE Then                                           
		strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & parent.UID_M0001			'☜: 
'		strVal = strVal & "&txtRcptNo=" & hRcptNo.value						'Hidden의 검색조건으로 Query
		strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)  		'Hidden의 검색조건으로 Query		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & parent.UID_M0001			'☜: 
		strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)		'현재 검색조건으로 Query
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
		   
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
 
    DbQuery = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()  
	Dim varData
	
	lgIntFlgMode = parent.OPMD_UMODE										'⊙: Indicates that current mode is Update mode

    lgQueryOk= True
	lgQueryState = True
	
	Call SetToolbar("1111110100011111")										'⊙: 버튼 툴바 제어 
	Call SetSpreadColor(1, frm1.vspdData.Maxrows)
	 
	frm1.vspdData.Row = 1
	frm1.vspdData.Col = C_RcptType_cd
	varData = frm1.vspdData.text

	Call subVspdSettingChange(C_RcptType_cd,1,frm1.vspdData.Maxrows, varData)
	
	Call ClickTab1()  
	Call DoSum()
	Call txtDocCur_OnChange()
	Call txtDept_OnBlur()
	
	frm1.txtRcptNo.focus
	lgBlnFlgChgValue = False    
	lgQueryOk= False	
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim pAr0061 
    Dim IntRows 
    Dim IntCols 
    Dim vbIntRet 
    Dim lStartRow 
    Dim lEndRow 
    Dim boolCheck 
    Dim lGrpcnt 
	Dim strVal, strDel
	Dim ApAmt, PayAmt
 
    DbSave = False                                                          '⊙: Processing is NG
    
    On Error Resume Next													'☜: Protect system from crashing
 
	Call LayerShowHide(1)
	 
	With frm1
		.txtMode.value = parent.UID_M0002									'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태   
	End With
	 
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data 연결 규칙 
	' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

	lGrpCnt = 1
	    
	strVal = ""
	strDel = ""

	With frm1.vspdData
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
				   
			If .Text  <> ggoSpread.DeleteFlag Then
				strVal = strVal & "C" & parent.gColSep & IntRows & parent.gColSep    

				.Col = C_Rcpttype_Cd						'3
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_NetRcptAmt							'4
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_NetRcptLocAmt						'5
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_NoteNo								'6
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_BankAcct							'7
				strVal = strVal & Trim(.Text) & parent.gColSep          

				.Col = C_hiddDtlSeq							'8 
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_Rcptacct_Cd						'9 
				strVal = strVal & Trim(.Text) & parent.gColSep
				
				.Col = C_Rcptitem_Desc						'10
				strVal = strVal & Trim(.Text) & parent.gRowSep				

				lGrpCnt = lGrpCnt + 1
			End if  
		Next
	End With
	 
	frm1.txtMaxRows.value = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value = strDel & strVal									'☜: Spread Sheet 내용을 저장 

	'권한관리추가 start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'권한관리추가 end
			 
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
	        
	DbSave = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
    Call InitVariables()
    Call ggoOper.ClearField(Document, "2")							'⊙: Clear Contents  Field
    
    Call InitVariables()													'⊙: Initializes local global variables
    Call InitSpreadsheet()													'⊙: Initializes local global variables
    
    Call DbQuery()
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************




'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblTotNetAmt
	Dim dblTotNetLocAmt

	With frm1.vspdData
		dblTotNetAmt = FncSumSheet1(frm1.vspdData,C_NetRcptAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotNetLocAmt = FncSumSheet1(frm1.vspdData,C_NetRcptLocAmt, 1, .MaxRows, False, -1, -1, "V")
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then                     
			frm1.txtTotNetRcptAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotNetAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If 
		frm1.txtTotNetRcptLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotNetLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
	End With 
End Sub    
    
'===================================== CurFormatNumericOCX()  =======================================
' Name : CurFormatNumericOCX()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 입금액 
		ggoOper.FormatFieldByObjectOfCur .txtRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 은행수수료 
		ggoOper.FormatFieldByObjectOfCur .txtBankAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 청산금액 
		ggoOper.FormatFieldByObjectOfCur .txtSttlAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 입금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotNetRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec  
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
' Name : CurFormatNumSprSheet()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' 입금액 
		ggoSpread.SSSetFloatByCellOfCur C_NetRcptAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'==========================================================================================
'   Event Name : subVspdSettingChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub subVspdSettingChange(ByVal Col , ByVal Row,  ByVal Row2, Byval varData) 
	Dim intIndex
	Dim strval
	Dim CurRow
	
	        
	For CurRow = Row To Row2
		frm1.vspdData.Col = C_RcptType_CD
		frm1.vspdData.Row = CurRow
		strval = UCase(TRim(frm1.vspdData.Text))

		If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			Select Case UCase(lgF0)
				Case "CS" & Chr(11)
					ggoSpread.SSSetProtected C_BankAcct,  CurRow, CurRow 
					ggoSpread.SpreadLock C_BankAcctPop,   CurRow, C_BankAcctPop,  CurRow 
					ggoSpread.SSSetProtected C_NoteNo,   CurRow, CurRow  
					ggoSpread.SpreadLock C_NoteNoPop,     CurRow, C_NoteNoPop,    CurRow      
				Case "DP" & Chr(11)   ' 예적금 
					ggoSpread.SpreadUnLock C_BankAcct,    CurRow,   CurRow 
					ggoSpread.SSSetRequired C_BankAcct,    CurRow,   CurRow 
					ggoSpread.SpreadUnLock C_BankAcctPop, CurRow, C_BankAcctPop,  CurRow
					ggoSpread.SpreadLock C_NoteNo,    CurRow, C_NoteNo,       CurRow
					ggoSpread.SpreadLock C_NoteNoPop,     CurRow, C_NoteNoPop,    CurRow  
				Case "NO" & Chr(11)
					ggoSpread.SpreadLock C_BankAcct,   CurRow, C_BankAcct,     CurRow 
					ggoSpread.SpreadLock C_BankAcctPop,   CurRow, C_BankAcctPop,  CurRow 
					ggoSpread.SpreadUnLock C_NoteNo,   CurRow, C_NoteNo,       CurRow
					ggoSpread.SSSetRequired C_NoteNo,   CurRow, CurRow
					ggoSpread.SpreadUnLock C_NoteNoPop,   CurRow, C_NoteNoPop,    CurRow     
				Case Else
					ggoSpread.SSSetProtected C_BankAcct,  CurRow, CurRow 
					ggoSpread.SpreadLock C_BankAcctPop,   CurRow, C_BankAcctPop,  CurRow 
					ggoSpread.SSSetProtected C_NoteNo,   CurRow, CurRow  
					ggoSpread.SpreadLock C_NoteNoPop,     CurRow, C_NoteNoPop,    CurRow      
			End Select
		End If
		If strval = "" Then
			ggoSpread.SSSetProtected C_BankAcct,  CurRow, CurRow 
			ggoSpread.SpreadLock C_BankAcctPop,   CurRow, C_BankAcctPop,  CurRow 
			ggoSpread.SSSetProtected C_NoteNo,   CurRow, CurRow  
			ggoSpread.SpreadLock C_NoteNoPop,     CurRow, C_NoteNoPop,    CurRow      			
		End If			
	Next 
End Sub

'====================================================================================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_NetRcptLocAmt	
			.vspdData.Text = ""    	
			ggoSpread.Source = .vspdData
			ggoSpread.UpdateRow ii	
		Next	
		.txtTotNetRcptLocAmt.text = "0"
		.txtBankLocAmt.text = "0"
	End With
End Sub


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************





'===================================== PopSaveSpreadColumnInf()  ======================================
' Name : PopSaveSpreadColumnInf()
' Description : 이동한 컬럼의 정보를 저장 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : 컬럼의 순서정보를 복원함 
'====================================================================================================
Sub  PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
			.vspdData.Row = NewRow
			.vspdData.Col = 0
			If .vspddata.Text = ggoSpread.DeleteFlag Then
				Exit Sub       
			End if
		End With
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
	
    gMouseClickStatus = "SPC" 'Split 상태코드 
 	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'Descending Sort
			lgSortKey = 1
		End If																
		Exit Sub
	End If		
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick(ByVal Col,ByVal Row)
    Dim iColumnName
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  ------------------------------------------------------------- 
'    LngLastRow = NewTop + 30
'    LngMaxRow = frm1.vspdData.MaxRows
    
'    If LngLastRow = frm1.vspdData.MaxRows Then
'        Call DbQuery()
'    End If    
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
		If Row > 0 then
			Select Case Col
				Case C_Rcpttype_Pb
					.Col = C_Rcpttype_Cd
					.Row = Row
					Call OpenPopup(.value, 6)
				Case C_Rcptacct_Pb
					.Col = C_Rcptacct_Cd
					.Row = Row
					Call OpenPopup(.value, 8)
				Case C_NoteNoPop 
					.Col = C_NoteNo 
					.Row = Row
					Call OpenPopup(.value, 4)
				Case C_BankAcctPop
					.Col = C_BankAcct
					.Row = Row
					Call OpenPopup(.value, 5)
			End Select
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 가수금구분이 예적금,어음일 경우에만 어음번호,계좌번호 Enabled 되게.
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim NetRcptAmt
	
	
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col  
 
	Select Case Col
		Case C_RcptType_cd 
			
			frm1.vspdData.ReDraw = False  
			If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(UCase(frm1.vspddata.Text), "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
				Select case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
					Case "DP" 
						frm1.vspdData.Col  = C_NoteNo
						frm1.vspdData.Row  = Row
						frm1.vspdData.Text = ""   
						frm1.vspdData.Col  = C_RcptType_cd 
					Case "NP" 
						frm1.vspdData.Col  = C_BankAcct
						frm1.vspdData.Row  = Row
						frm1.vspdData.Text = ""   
						frm1.vspdData.Col  = C_RcptType_cd 
					Case "NR"
						frm1.vspdData.Col  = C_BankAcct
						frm1.vspdData.Row  = Row
						frm1.vspdData.Text = ""   
						frm1.vspdData.Col  = C_RcptType_cd 
					Case Else          
						frm1.vspdData.Col  = C_BankAcct
						frm1.vspdData.Row  = Row
						frm1.vspdData.Text = ""   
						frm1.vspdData.Col  = C_NoteNo
						frm1.vspdData.Row  = Row
						frm1.vspdData.Text = ""   
						frm1.vspdData.Col  = C_RcptType_cd      
				End Select
			ENd if
			frm1.vspdData.ReDraw = True 

			Call subVspdSettingChange(Col,Row,Row, frm1.vspddata.Text)

			frm1.vspdData.Col  = C_Rcptacct_Cd
			frm1.vspdData.Text = ""
			frm1.vspdData.Col  = C_Rcptacct_Nm
			frm1.vspdData.Text = ""
		
		Case C_NetRcptLocAmt
			If UNICDbl(frm1.vspdData.text) < 0 Then
				frm1.vspdData.Text  = UNIConvNumPCToCompanyByCurrency(frm1.vspdData.Text * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			End if
			Call DoSum()
		
		Case C_NetRcptAmt
			
			If UNICDbl(frm1.vspdData.text) < 0 Then
				frm1.vspdData.Text  = UNIConvNumPCToCompanyByCurrency(frm1.vspdData.Text * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			End if
			frm1.vspdData.Col  = C_NetRcptLocAmt		
			frm1.vspddata.Text = ""

			Call DoSum()
	End Select 
	

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row     
End Sub





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
' Name : txtDocCur_onblur()
' Description : 
'=======================================================================================================
Function txtDocCur_onblur()
  
End Function

'========================================================================================
' Function Name :txtXchRate_onblur
' Function Desc : 
'========================================================================================
Function txtXchRate_onblur()
	lgBlnFlgChgValue = True
End Function

'========================================================================================
' Function Name :txtBankAmt_onblur
' Function Desc : 
'========================================================================================
Function txtBankAmt_onblur()
	lgBlnFlgChgValue = True
	If UNICDbl(frm1.txtBankAmt.text) < 0 then
		frm1.txtBankAmt.Text = UNIFormatNumber(UNICDbl(frm1.txtBankAmt.Text) * (-1),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)   
	End if
End function
'========================================================================================
' Function Name :txtBankLocAmt_onblur
' Function Desc : 
'======================================================================================== 
Function txtBankLocAmt_onblur() 
	lgBlnFlgChgValue = True
End function


'=======================================================================================================
'   Event Name : txtRcptDt_DblClick(Button)
'   Event Desc : 입금일관련 달력을 호출한다.
'=======================================================================================================
Sub txtRcptDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtRcptDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtRcptDt.Focus 
    End If
    Call txtRcptDt_OnBlur()  
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
		Call DoSum()
	End If    
	
	If lgQueryOk<> True Then
		Call XchLocRate()
	End If	
End Sub

'==========================================================================================
'   Event Name : txtBankAmt_Change
'   Event Desc : 
'==========================================================================================
Sub txtBankAmt_Change()
	lgBlnFlgChgValue = True
	frm1.txtBankLocAmt.text = "0"
End Sub
'==========================================================================================
'   Event Name : txtDept_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtDept_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtRcptDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDept.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDept.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtRcptDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDept.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDept.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub

'==========================================================================================
'   Event Name : txtRcptDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtRcptDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
	
				If LTrim(RTrim(.txtDept.value)) <> "" and Trim(.txtRcptDt.Text <> "") Then
					'----------------------------------------------------------------------------------------
						strSelect	=			 " Distinct org_change_id "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDept.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtRcptDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
							.txtDept.value = ""
							.txtDeptNm.value = ""
							.hOrgChangeId.value = ""
							.txtDept.focus
					End if
				End If
			End With
		'----------------------------------------------------------------------------------------
		End If
	End IF
	
	Call XchLocRate()
End Sub

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.8 HTML Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************

'========================================================================================
' Function Name :txtBpCd_onBlur
' Function Desc : 
'========================================================================================
Function txtBpCd_onBlur()
	If frm1.txtBpCd.value = "" then
	 frm1.txtBpNm.value = ""
	End if
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!-- '#########################################################################################################
'            6. Tag부 
'#########################################################################################################  -->
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>유형별 가수금내역</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>   
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- 본문내용  -->
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
							<TD CLASS="TD5" NOWRAP>가수금번호</TD>
							<TD CLASS="TD6"><INPUT NAME="txtRcptNo" TYPE="Text" MAXLENGTH=18 tag="12XXXU" ALT="가수금번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtRcptNo.value, 0)"></TD>
							<TD CLASS="TDT"></TD>
							<TD CLASS="TD6"></TD>
						</TR>
					</TABLE>
					</FIELDSET>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH="100%">     
				
				<DIV ID="TabDiv" STYLE="FlOAT: left; HEIGHT:100%; OVERFLOW:auto; WIDTH:100%;" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>가수금유형</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRcptType" SIZE=10 MAXLENGTH=20  tag="22XXXU" ALT="가수금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('',7)">&nbsp;<INPUT TYPE=TEXT NAME="txtRcptTypeNm" SIZE=25 tag="24" ALT="가수금유형명"></TD>
							<TD CLASS=TD5 NOWRAP>프로젝트</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="프로젝트" MAXLENGTH=25 SIZE=25 tag="2X"></TD>
<!--							<TD CLASS=TD5 NOWRAP>결재여부</TD>
							<TD CLASS=TD6 NOWRAP><SELECT NAME="cboConfFg" ALT="결재여부" STYLE="WIDTH: 100px" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>-->
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>입금일자</TD>                           
							<TD CLASS="TD6" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtRcptDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="입금일자"> </OBJECT>');</SCRIPT>               
							</TD>
							<TD CLASS=TD5 NOWRAP>수금처</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="수금처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="2XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.value, 2)"> <INPUT NAME="txtBpNm" TYPE="Text" SIZE=25 tag="24"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>부서</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDept" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDept.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ></TD>            							
							<TD CLASS=TD5 NOWRAP>참조번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" STYLE="TEXT-ALIGN: left" tag="24XXXU">&nbsp;</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>거래통화</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 tag="22XXXU" MAXLENGTH="3"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.value, 3)"></TD>
							<TD CLASS=TD5 NOWRAP>환율</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name="txtXchRate" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 70px" title="FPDOUBLESINGLE" ALT="환율" tag="24x5"> </OBJECT>');</SCRIPT>&nbsp;
							</TD>         
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>입금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtRcptAmt CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="입금액" tag="24x2"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>입금액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtRcptLocAmt CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="입금액(자국)" tag="24x2"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
						</TR>
						<TR>                    
							<TD CLASS=TD5 NOWRAP>은행수수료</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtBankAmt CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="은행수수료" tag="21X2"> </OBJECT>');</SCRIPT>&nbsp;
							</TD>
							<TD CLASS=TD5 NOWRAP>은행수수료(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
							 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name=txtBankLocAmt CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="은행수수료(자국)" tag="21X2"> </OBJECT>');</SCRIPT>&nbsp;
							</TD>
						</TR>
						<TR>                    
							<TD CLASS=TD5 NOWRAP>반제금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtClsAmt CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="반제금액" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>반제금액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtClsLocAmt CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="반제금액(자국)" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
						</TR>        
						<TR>                      
							<TD CLASS=TD5 NOWRAP>청산금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtSttlAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="청산금액" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>청산금액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name=txtSttlLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="청산금액(자국)" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>                                 
						</TR>
						<TR>                      
							<TD CLASS=TD5 NOWRAP>잔액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>잔액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24"> </OBJECT>');</SCRIPT> &nbsp;
						    </TD>
						</TR>       
						<TR>
							<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGLNo" SIZE=19 MAXLENGTH=18 tag="24" ALT="전표번호"></TD>
							<TD CLASS="TD5" NOWFRAP>전표번호</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGLNo" SIZE=19 MAXLENGTH=18 tag="24" ALT="전표번호"></TD>
						</TR>        

						<TR>
							<TD CLASS=TD5 NOWRAP>비고</TD>
							<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=80 MAXLENGTH=128 tag="2X" ALT="비고"></TD>        
						</TR>
					</TABLE>
				</DIV>
				
				
				<DIV ID="TabDiv" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%" COLSPAN="4">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=va1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
						<TR>
							<TD COLSPAN=4>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>              
										<TD class=TD5 NOWRAP>입금액</TD>
										<TD class=TD6 NOWRAP>         
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotNetRcptAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="순매출액" tag="24X2" id=OBJECT22> </OBJECT>');</SCRIPT>
										</TD>
										<TD class=TD5 NOWRAP>입금액(자국)</TD>
										<TD class=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotNetRcptLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="순매출액(자국)" tag="24X2" id=OBJECT22> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
								</TABLE>
							</TD>         
						</TR>
					</TABLE>
				</DIV>
			</TD>
		</TR>    
	</TABLE>
  </TD>
 </TR>   
 <TR>
	<TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR>
	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TD>
 </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"  tag="24" TABINDEX="-1"></TEXTAREA><% '업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
	<INPUT TYPE=HIDDEN NAME="txtMode"      tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtMaxRows"   tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode"   tag="24" TABINDEX="-1">
	<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

