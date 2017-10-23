<%@ LANGUAGE="VBSCRIPT" %>
<!--===================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5121ma1
'*  4. Program Name         : 부도어음처리 
'*  5. Program Desc         : 부도어음처리 등록 수정 삭제 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2003/04/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Soo Min, Oh
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************
'==========================================  1.1.1 Style Sheet  ==========================================
'========================================================================================================== -->

<!--========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance
																			'☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Const BIZ_PGM_ID  = "f5121mb1.asp"										'☆: 비지니스 로직 ASP명 
'Const BIZ_PGM_ID2 = "f5121mb2.asp"										'☆: 비지니스 로직 ASP명 
'Const JUMP_PGM_ID_NOTE_CHG = "f5107ma1"									'어음변경등록 

Dim C_SEQ
Dim C_STTL_TYPE
Dim C_STTL_TYPE_NM
Dim C_RCPT_TYPE
Dim C_RCPT_TYPE_BT
Dim C_RCPT_TYPE_NM
Dim C_REF_NOTE_NO
Dim C_REF_NOTE_BT
Dim C_ACCT_CD
Dim C_ACCT_BT
Dim C_ACCT_NM
Dim C_BANK_ACCT
Dim C_BANK_ACCT_BT
Dim C_BANK_CD
Dim C_BANK_BT
Dim C_BANK_NM
Dim C_STTL_AMT
Dim C_NOTE_ITEM_DESC

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE								'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False										'⊙: Indicates that no value changed
    lgIntGrpCount = 0												'⊙: Initializes Group View Size

	lgStrPrevKey = ""
	lgLngCurRows = 0												'initializes Deleted Rows Count
    IsOpenPop = False												'☆: 사용자 변수 초기화 

    lgSortKey = 1
	lgPageNo  = ""

    lgBlnFlgChgValue = False
End Sub

Sub initSpreadPosVariables()
	C_SEQ = 1
	C_STTL_TYPE = 2
	C_STTL_TYPE_NM = 3
	C_RCPT_TYPE = 4
	C_RCPT_TYPE_BT = 5
	C_RCPT_TYPE_NM = 6
	C_REF_NOTE_NO = 7
	C_REF_NOTE_BT = 8	
	C_ACCT_CD = 9
	C_ACCT_BT = 10
	C_ACCT_NM = 11
	C_BANK_ACCT = 12
	C_BANK_ACCT_BT = 13
	C_BANK_CD = 14
	C_BANK_BT = 15
	C_BANK_NM = 16
	C_STTL_AMT = 17
	C_NOTE_ITEM_DESC = 18
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtStsDt.Text	= UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
	frm1.txtNoteNoQry.focus
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()
    Dim sList

    With frm1
		.vspdData.MaxCols = C_NOTE_ITEM_DESC + 1
		.vspdData.Col = .vspdData.MaxCols	:	.vspdData.ColHidden = True				'☜: 공통콘트롤 사용 Hidden Column
		.vspdData.MaxRows = 0
		ggoSpread.Source = .vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread
        Call GetSpreadColumnPos("A")
	
		ggoSpread.SSSetFloat	C_SEQ,			"순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetCombo	C_STTL_TYPE,	"처리유형", 12
		ggoSpread.SSSetCombo	C_STTL_TYPE_NM, "처리유형명", 12
		ggoSpread.SSSetEdit		C_RCPT_TYPE,	"입금유형", 15, , , 20
		ggoSpread.SSSetButton	C_RCPT_TYPE_BT
		ggoSpread.SSSetEdit		C_RCPT_TYPE_NM, "입금유형명", 15, , , 30
		ggoSpread.SSSetEdit		C_REF_NOTE_NO,	"받을어음번호", 15, , , 20
		ggoSpread.SSSetButton	C_REF_NOTE_BT				
		ggoSpread.SSSetEdit		C_ACCT_CD,		"계정코드", 15, , , 20
		ggoSpread.SSSetButton	C_ACCT_BT
		ggoSpread.SSSetEdit		C_ACCT_NM,		"계정코드명", 15, , , 50
		ggoSpread.SSSetEdit		C_BANK_ACCT,	"계좌번호", 20, , , 30
		ggoSpread.SSSetButton	C_BANK_ACCT_BT
		ggoSpread.SSSetEdit		C_BANK_CD,		"은행코드", 15, , , 30
		ggoSpread.SSSetButton	C_BANK_BT
		ggoSpread.SSSetEdit		C_BANK_NM,		"은행명", 15, , , 30
		ggoSpread.SSSetFloat	C_STTL_AMT,		"처리금액", 17, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit		C_NOTE_ITEM_DESC, "비고", 35, , , 128
	
'2003/09/09 수정 필요	 
		Call ggoSpread.MakePairsColumn(C_STTL_TYPE,C_STTL_TYPE_NM,"1")
		Call ggoSpread.MakePairsColumn(C_RCPT_TYPE,C_RCPT_TYPE_NM,"1")
		Call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_NM)
		Call ggoSpread.MakePairsColumn(C_BANK_ACCT,C_BANK_ACCT_BT)
		Call ggoSpread.MakePairsColumn(C_BANK_CD,C_BANK_NM)
		
		Call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
		Call ggoSpread.SSSetColHidden(C_STTL_TYPE,C_STTL_TYPE,True)

		Call SetSpreadLock                                              '바뀐부분 
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
	Dim RowCnt

	ggoSpread.Source = frm1.vspdData	

    With frm1
		.vspdData.ReDraw = False      
'		ggoSpread.SpreadLock		1 ,     -1  

		ggoSpread.SpreadLock		C_SEQ				, -1	, C_SEQ
		ggoSpread.SpreadUnLock		C_STTL_TYPE_NM		, -1    , C_STTL_TYPE_NM
		ggoSpread.SSSetRequired		C_STTL_TYPE_NM		, -1    

		ggoSpread.SpreadLock		C_RCPT_TYPE_NM,			-1		, C_ACCT_NM
		ggoSpread.SpreadLock		C_ACCT_NM,			-1		, C_ACCT_NM
		ggoSpread.SSSetRequired		C_STTL_TYPE_NM,		-1
		ggoSpread.SSSetRequired		C_ACCT_CD,			-1
		ggoSpread.SpreadUnLock		C_STTL_AMT			, -1    , C_STTL_AMT		
		ggoSpread.SSSetRequired		C_STTL_AMT,			-1
		ggoSpread.SpreadUnLock		C_NOTE_ITEM_DESC	, -1    , C_NOTE_ITEM_DESC

		For RowCnt = 1 To .vspdData.MaxRows			
			.vspdData.Col = C_STTL_TYPE
			.vspdData.Row = RowCnt	

			If UCase(Trim(.vspdData.text)) = "RI" Then						'상환 
				ggoSpread.SpreadUnLock		C_RCPT_TYPE,		RowCnt,	C_RCPT_TYPE	,RowCnt			
				ggoSpread.SSSetRequired		C_RCPT_TYPE,		RowCnt,	RowCnt			
				ggoSpread.SpreadUnLock		C_RCPT_TYPE_BT,		RowCnt,	C_RCPT_TYPE_BT	

				ggoSpread.SpreadLock		C_REF_NOTE_NO,		RowCnt,	C_REF_NOTE_NO,RowCnt			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		RowCnt,	RowCnt			

				.vspdData.Col = C_RCPT_TYPE
				.vspdData.Row = RowCnt		

				If Trim(.vspdData.text) = "DP" Then						
					ggoSpread.SpreadUnLock		C_BANK_ACCT,		RowCnt,	C_BANK_ACCT	,RowCnt			
					ggoSpread.SSSetRequired		C_BANK_ACCT,		RowCnt,	RowCnt			
					ggoSpread.SpreadUnLock		C_BANK_ACCT_BT,		RowCnt,	C_BANK_ACCT_BT

					ggoSpread.SpreadUnLock		C_BANK_CD,			RowCnt,	C_BANK_CD	,RowCnt			
					ggoSpread.SSSetRequired		C_BANK_CD,			RowCnt,	RowCnt			
					ggoSpread.SpreadUnLock		C_BANK_BT,			RowCnt,	C_BANK_BT					
				Else
					ggoSpread.SpreadLock		C_BANK_ACCT,		-1		, C_BANK_ACCT_BT
					ggoSpread.SpreadLock		C_BANK_CD,			-1		, C_BANK_NM
				End If
			ElseIf  UCase(Trim(.vspdData.text)) = "NR" Then	
				ggoSpread.SpreadLock		C_RCPT_TYPE,		RowCnt, C_RCPT_TYPE			,RowCnt			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		RowCnt, RowCnt
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		RowCnt,	C_RCPT_TYPE_BT	
				
				ggoSpread.SpreadUnLock		C_REF_NOTE_NO,		RowCnt, C_REF_NOTE_NO			,RowCnt			
				ggoSpread.SSSetRequired		C_REF_NOTE_NO,		RowCnt, RowCnt
				ggoSpread.SpreadUnLock		C_REF_NOTE_BT,		RowCnt,	C_RCPT_TYPE_BT		
				
				ggoSpread.SpreadLock		C_BANK_ACCT,		-1		, C_BANK_ACCT_BT
				ggoSpread.SpreadLock		C_BANK_CD,			-1		, C_BANK_NM
			ElseIf UCase(Trim(.vspdData.text)) <> "AL" Or UCase(Trim(.vspdData.text)) <> "EP" Then
				ggoSpread.SpreadLock		C_RCPT_TYPE,		RowCnt, C_RCPT_TYPE			,RowCnt			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		RowCnt, RowCnt
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		RowCnt,	C_RCPT_TYPE_BT	
				
				ggoSpread.SpreadLock		C_REF_NOTE_NO,		RowCnt,	C_REF_NOTE_NO,RowCnt			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		RowCnt,	RowCnt			
				
				ggoSpread.SpreadLock		C_BANK_ACCT,		-1		, C_BANK_ACCT_BT
				ggoSpread.SpreadLock		C_BANK_CD,			-1		, C_BANK_NM
			End If 
		Next 

		.vspdData.ReDraw = True
   End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal lRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired		C_STTL_TYPE_NM,		lRow,	lRow
		ggoSpread.SSSetProtected	C_RCPT_TYPE,		lRow,	lRow
		ggoSpread.SSSetProtected	C_RCPT_TYPE_BT,		lRow,	lRow
		ggoSpread.SSSetProtected	C_RCPT_TYPE_NM,		lRow,	lRow
		ggoSpread.SSSetProtected	C_REF_NOTE_NO,		lRow,	lRow
		ggoSpread.SSSetProtected	C_REF_NOTE_BT,		lRow,	lRow
		ggoSpread.SSSetRequired		C_ACCT_CD,			lRow,	lRow
		ggoSpread.SSSetProtected	C_ACCT_NM,			lRow,	lRow
		ggoSpread.SSSetProtected	C_BANK_ACCT,		lRow,	lRow
		ggoSpread.SSSetProtected	C_BANK_ACCT_BT,		lRow,	lRow		
		ggoSpread.SpreadLock		C_BANK_CD,			lRow,	lRow
		ggoSpread.SSSetRequired		C_STTL_AMT,			lRow,	lRow
		
		.vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Function InitCombo()
	ggoSpread.Source = frm1.vspdData
	                   'Select                 From        Where                Return value list  
    Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("F1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_STTL_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_STTL_TYPE_NM       

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))    
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))    
End Function

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenNoteInfo()
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("f5121ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f5121ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		frm1.txtNoteNoQry.focus
		Exit Function
	Else
		frm1.txtNoteNoQry.value  = arrRet(0)
		frm1.txtNoteNoQry.focus
	End If	
End Function

'------------------------------------------  OpenPopUpNoteNo()  ---------------------------------------------
'	Name : OpenPopUpNoteNo()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function  OpenPopUpNoteNo()
	Dim strNoteFg
	Dim IntRetCd
	
	strNoteFg = frm1.cboNoteFg.Value
	
	If strNoteFg = "" Then
	    IntRetCD = DisplayMsgBox("141327","x","x","x")	'어음구분을 먼저 입력하십시오.
		Exit Function		      	
	ElseIf strNoteFg = "D3" Then 
		Call OpenPopUp(frm1.txtNoteNo.Value, 1) '지급어음 
	Else  
	    IntRetCD = DisplayMsgBox("141220","x","x","x")	'어음번호를 직접 입력해주십시오.
		Exit Function		
    End If	
End Function  

'==================================================================================
'	Name : OpenPopUp()
'	Description : 공통팝업 정의 
'==================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 2		'입금유형 
 			arrParam(0) = "입금유형팝업"
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " "
			arrParam(4) = arrParam(4) & "AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD "
			arrParam(4) = arrParam(4) & "AND B.SEQ_NO = 1 AND B.REFERENCE = " & FilterVar("RP", "''", "S") & "  "
			arrParam(5) = strCode

			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"

			arrHeader(0) = "입금유형"
			arrHeader(1) = "입금유형명"
		Case 3		' 은행 
			arrParam(0) = "은행 팝업"	' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "									' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(5) = strCode													' 조건필드의 라벨 명칭 
			
			arrField(0) = "A.BANK_CD"						' Field명(0)
			arrField(1) = "A.BANK_NM"						' Field명(1)
			arrField(2) = "B.BANK_ACCT_NO"					' Field명(2)
			
			arrHeader(0) = "은행코드"					' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)
			arrHeader(2) = "계좌번호"					' Header명(2)				
		Case 4		' 계좌번호 
'			If frm1.txtBankAcct1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "계좌번호 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"				' TABLE 명칭 
			arrParam(2) = strCode							' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "												' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "	
			arrParam(5) = "계좌번호"					' 조건필드의 라벨 명칭 
			arrField(0) = "B.BANK_ACCT_NO"					' Field명(0)
			arrField(1) = "A.BANK_CD"						' Field명(0)
			arrField(2) = "A.BANK_NM"						' Field명(0)
			arrHeader(0) = "계좌번호"					' Header명(0)
			arrHeader(1) = "은행코드"					' Header명(0)
			arrHeader(2) = "은행명"						' Header명(0)	
		Case 5		' 입금유형계정코드 
			arrParam(0) = "입금계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & " " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  "
			arrParam(4) = arrParam(4) & " AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ "
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  and D.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "			
			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = C_STTL_TYPE
			
			If Trim(frm1.vspdData.Text) <> "" Then
				arrParam(4) = arrParam(4) & " AND D.JNL_CD = " & FilterVar(frm1.vspdData.Text, "''", "S") 					 				 		
			End If 
			
			frm1.vspdData.Col = C_RCPT_TYPE
			
			If Trim(frm1.vspdData.Text) <> "" Then
			arrParam(4) = arrParam(4) & " AND D.EVENT_CD = " & FilterVar(frm1.vspdData.Text, "''", "S")
			End If

			arrParam(5) = strCode											' 조건필드의 라벨 명칭 

			arrField(0) = "A.ACCT_CD"										' Field명(0)
			arrField(1) = "A.ACCT_NM"										' Field명(1)
			arrField(2) = "B.GP_CD"											' Field명(2)
			arrField(3) = "B.GP_NM"					 						' Field명(3)

			arrHeader(0) = "입금유형계정코드"							' Header명(0)
			arrHeader(1) = "입금유형계정명"								' Header명(1)
			arrHeader(2) = "그룹코드"									' Hea der명(2)
			arrHeader(3) = "그룹명"										' Header명(3)	
		Case 6																'어음번호 POPUP
 			arrParam(0) = "어음번호팝업"
			arrParam(1) = "F_NOTE A, B_BANK	B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.NOTE_STS = " & FilterVar("BG", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND A.BANK_CD = B.BANK_CD "
			arrParam(5) = strCode

			arrField(0) = "A.NOTE_NO"			
			arrField(1) = "A.NOTE_AMT"
			arrField(2) = "B.BANK_NM"

			arrHeader(0) = "어음번호"
			arrHeader(1) = "어음금액"
			arrHeader(2) = "발행은행"
		Case 7		' 이자수익계정코드 
			arrParam(0) = "이자수익계정팝업"							' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & " " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  "
			arrParam(4) = arrParam(4) & " AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ "
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  and D.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  "						
			arrParam(4) = arrParam(4) & " AND D.JNL_CD = " & FilterVar("IR", "''", "S") & "  " 					 				 					
			arrParam(5) = strCode											' 조건필드의 라벨 명칭 

			arrField(0) = "A.ACCT_CD"										' Field명(0)
			arrField(1) = "A.ACCT_NM"										' Field명(1)
			arrField(2) = "B.GP_CD"											' Field명(2)
			arrField(3) = "B.GP_NM"					 						' Field명(3)

			arrHeader(0) = "이자수익계정코드"							' Header명(0)
			arrHeader(1) = "이자수익계정명"								' Header명(1)
			arrHeader(2) = "그룹코드"									' Hea der명(2)
			arrHeader(3) = "그룹명"										' Header명(3)
	End Select
  
	IsOpenPop = True
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  EscPopUp()  --------------------------------------------------
'	Name : EscPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 2		' 입금유형		
				Call SetActiveCell(.vspdData,C_RCPT_TYPE,.vspdData.ActiveRow ,"M","X","X")
				Call SetActiveCell(.vspdData,C_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 3		' 은행 
				.txtBankCD.focus
			Case 4		' 거래처 
				.txtBpCd.focus
			Case 5		' 입금유형계정코드 
				Call SetActiveCell(.vspdData,C_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 6		' 받을어음 
				Call SetActiveCell(.vspdData,C_REF_NOTE_NO,.vspdData.ActiveRow ,"M","X","X")
			Case 7		' 이자수익 
				.txtIntAcctCd.focus
		End Select
	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 2		' 입금유형										
				.vspdData.Col = C_RCPT_TYPE_NM
				.vspdData.Text = arrRet(1)					
				.vspdData.Col = C_RCPT_TYPE
				.vspdData.Text = arrRet(0)

				If UCase(Trim(.vspdData.Text)) = "DP" Then												
					ggoSpread.SpreadUnLock		C_BANK_ACCT,		.vspdData.ActiveRow,	C_BANK_ACCT	,.vspdData.ActiveRow			
					ggoSpread.SSSetRequired		C_BANK_ACCT,		.vspdData.ActiveRow,	.vspdData.ActiveRow			
					ggoSpread.SpreadUnLock		C_BANK_ACCT_BT,		.vspdData.ActiveRow,	C_BANK_ACCT_BT										
				
					ggoSpread.SpreadUnLock		C_BANK_CD,			.vspdData.ActiveRow,	C_BANK_CD	,.vspdData.ActiveRow			
					ggoSpread.SSSetRequired		C_BANK_CD,			.vspdData.ActiveRow,	.vspdData.ActiveRow			
					ggoSpread.SpreadUnLock		C_BANK_BT,			.vspdData.ActiveRow,	C_BANK_BT
					ggoSpread.SpreadLock		C_BANK_NM,			.vspdData.ActiveRow,	C_BANK_NM					
				Else
					ggoSpread.SpreadLock		C_BANK_ACCT,		.vspdData.ActiveRow,	C_BANK_ACCT			,.vspdData.ActiveRow			
					ggoSpread.SSSetProtected	C_BANK_ACCT,		.vspdData.ActiveRow,	.vspdData.ActiveRow
					ggoSpread.SpreadLock		C_BANK_CD,			.vspdData.ActiveRow,	C_BANK_CD			,.vspdData.ActiveRow			
					ggoSpread.SSSetProtected	C_BANK_CD,			.vspdData.ActiveRow,	.vspdData.ActiveRow				
					ggoSpread.SpreadLock		C_BANK_NM,			.vspdData.ActiveRow,	C_BANK_NM
				End If
				
				.vspdData.Col = C_ACCT_CD
				.vspdData.Text = ""
				.vspdData.Col = C_ACCT_NM
				.vspdData.Text = ""
				.vspdData.Col = C_BANK_ACCT
				.vspdData.Text = ""
				.vspdData.Col = C_BANK_CD
				.vspdData.Text = ""
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = ""				
				
				ggoSpread.SpreadUnLock		C_ACCT_CD,			.vspdData.ActiveRow,	C_ACCT_CD	,.vspdData.ActiveRow			
				ggoSpread.SSSetRequired		C_ACCT_CD,			.vspdData.ActiveRow,	.vspdData.ActiveRow			
				ggoSpread.SpreadUnLock		C_ACCT_BT,			.vspdData.ActiveRow,	C_ACCT_BT
				ggoSpread.SpreadLock		C_ACCT_NM,			.vspdData.ActiveRow,	C_ACCT_NM									

				Call SetActiveCell(.vspdData,C_RCPT_TYPE,.vspdData.ActiveRow ,"M","X","X")
			Case 3		' 은행 
				.vspdData.Col = C_BANK_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_BANK_ACCT
				.vspdData.Text = arrRet(2)
			Case 4		' 계좌번호 
				.vspdData.Col = C_BANK_ACCT
				.vspdData.Text = arrRet(0)				
				.vspdData.Col = C_BANK_CD
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = arrRet(2)
			Case 5		' 입금유형계정코드 
				.vspdData.Col = C_ACCT_CD
				.vspdData.Text = arrRet(0)				
				.vspdData.Col = C_ACCT_NM
				.vspdData.Text = arrRet(1)				
				Call SetActiveCell(.vspdData,C_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 6		' 받을어음 
				.vspdData.Col = C_REF_NOTE_NO
				.vspdData.Text = arrRet(0)				
				.vspdData.Col = C_STTL_AMT
				.vspdData.Text = arrRet(1)	
				Call SetActiveCell(.vspdData,C_REF_NOTE_NO,.vspdData.ActiveRow ,"M","X","X")
			Case 7		' 이자수익 
				.txtIntAcctCd.value = arrRet(0)
				.txtIntAcctNM.value = arrRet(1)
				.txtIntAcctCd.focus
		End Select

		lgBlnFlgChgValue = True
	End With
End Function

'------------------------------------------  OpenPopupDept()  --------------------------------------------
'	Name : OpenPopupDept()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'부서코드 
	arrParam(1) = frm1.txtStsDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"							'부서권한(lgUsrIntCd)
	arrParam(3) = "F"							'부서권한(lgUsrIntCd)

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCD.focus
		Exit Function
	End If

	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	frm1.txtStsDt.text = arrRet(3)
	Call txtDeptCD_Change()
	frm1.txtDeptCD.focus

	lgBlnFlgChgValue = True
End Function

'------------------------------------------  OpenPopupTempGL()  --------------------------------------------
'	Name : OpenPopupTempGL()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
    Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	With frm1		
		arrParam(0) = Trim(.hTempGlNo.value)	'전표번호 
		arrParam(1) = ""						'Reference번호		
	End With
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'------------------------------------------  OpenPopupGL()  --------------------------------------------
'	Name : OpenPopupGL()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
		
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	With frm1		
		arrParam(0) = Trim(.hGlNo.value)	'전표번호 
		arrParam(1) = ""			'Reference번호		
	End With
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function CookiePage(ByVal Kubun)
	Dim strTemp

	Select Case Kubun		
		Case "FORM_LOAD"		
			strTemp = ReadCookie("NOTE_NO")
			
			Call WriteCookie("NOTE_NO", "")
			
			If strTemp = "" Then Exit Function

			frm1.txtNoteNoQry.value = strTemp
	
			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("NOTE_NO", "")
				Exit Function 
			End If
					
			Call MainQuery()
		Case Else
			Exit Function
	End Select
End Function	

'========================================================================================================
'	Desc : 화면이동 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD

	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 계속하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call CookiePage(strPgmId)
    Call PgmJump(strPgmId)
End Function

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029															'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.ClearField(Document, "1")										'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field

	Call InitSpreadSheet                                                        'Setup the Spread sheet			
	Call InitCombo    	
    Call SetDefaultVal    
    Call InitVariables															'⊙: Initializes local global variables

	Call SetToolbar("1110110000001111")
'    Call CookiePage("FORM_LOAD")
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SEQ           =   iCurColumnPos(1)
			C_STTL_TYPE		=	iCurColumnPos(2)
			C_STTL_TYPE_NM	= 	iCurColumnPos(3)
			C_RCPT_TYPE		=	iCurColumnPos(4)
			C_RCPT_TYPE_BT	=	iCurColumnPos(5)
			C_RCPT_TYPE_NM	=	iCurColumnPos(6)
			C_REF_NOTE_NO	=	iCurColumnPos(7)
			C_REF_NOTE_BT	=	iCurColumnPos(8)			
			C_ACCT_CD		=	iCurColumnPos(9)
			C_ACCT_BT		=	iCurColumnPos(10)
			C_ACCT_NM		=	iCurColumnPos(11)
			C_BANK_ACCT		=	iCurColumnPos(12)
			C_BANK_ACCT_BT	=	iCurColumnPos(13)
			C_BANK_CD		=	iCurColumnPos(14)
			C_BANK_BT		=	iCurColumnPos(15)
			C_BANK_NM		=	iCurColumnPos(16)
			C_STTL_AMT		=	iCurColumnPos(17)
			C_NOTE_ITEM_DESC=	iCurColumnPos(18)
    End Select    
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow
		frm1.vspdData.Col = C_STTL_TYPE
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_STTL_TYPE_NM
		frm1.vspdData.value = intindex
	Next
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStsDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStsDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStsDt.Focus 
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStsDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStsDt_Change()
    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtStsDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStsDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
								
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
	
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDeptCD_Change()
'   Event Desc : Vlidation Check of Department Code
'=======================================================================================================
Sub txtDeptCD_Change()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii

	If Trim(frm1.txtDeptCd.value) = "" And Trim(frm1.txtStsDt.Text = "") Then Exit Sub

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStsDt.Text, gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			

		For ii = 0 to Ubound(arrVal1,1) - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next	
	End If

     lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtStsDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCashRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtNoteAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtSttlAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtIntRevAmt_Change()
    lgBlnFlgChgValue = True
    If unicdbl(frm1.txtIntRevAmt.Text) > 0 Then   		
		Call ggoOper.SetReqAttr(frm1.txtIntAcctCd, "N")		
	Else 
		frm1.txtIntAcctCd.value = ""
		frm1.txtIntAcctNm.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtIntAcctCd, "Q")			
	End If		
End Sub

 '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
Sub cboNoteFg_OnChange1()							'dbqueryok 시의 event (field not clear)
	with frm1
		Select Case frm1.cboNoteFg.value
			Case "D1"	'받을어음				
				Call ggoOper.SetReqAttr(.txtCashRate, "N")	'N:Required, Q:Protected, D:Default				
			Case "D3"	'지급어음			
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default				
			Case Else				
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default
		End Select
	End with
End Sub

Sub cboPlace_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboRcptFg_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtNoteNo_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtDeptCD_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtPublisher_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBpCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtNoteDesc_OnChange()
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'   Event Name : vspd	
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    lgBlnFlgChgValue = True
End Sub 
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			.Col = Col
			.Row = Row
			Select Case Col
			Case C_RCPT_TYPE_BT
				Call OpenPopup(.Text, 2)
			Case C_REF_NOTE_BT			
				Call OpenPopup(.Text, 6)
			Case C_ACCT_BT
				Call OpenPopup(.Text, 5)
			Case C_BANK_ACCT_BT
				Call OpenPopup(.Text, 4)
			Case C_BANK_BT
				Call OpenPopup(.Text, 3)
			Case Else
			End Select
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data clicked
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData
    
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
	
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
    End If     
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgPageNo <> "" Then                         
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData
	
	With frm1.vspdData
		.ReDraw = False
		.Row = Row
    
		Select Case Col
			Case  C_STTL_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_STTL_TYPE
				.Value = intIndex
				varData = .text							
			Case C_STTL_TYPE
				.Col = Col
				intIndex = .Value
				.Col = C_STTL_TYPE_NM
				.Value = intIndex
				varData = .text				
		End Select
		
		ggoSpread.Source = frm1.vspdData												
		
		Select Case UCase(Trim(.text))		
			Case "RI"						'상환(ReImbursement)												
				ggoSpread.SpreadUnLock		C_RCPT_TYPE,		Row, C_RCPT_TYPE	,Row			
				ggoSpread.SSSetRequired		C_RCPT_TYPE,		Row, Row			
				ggoSpread.SpreadUnLock		C_RCPT_TYPE_BT,		Row, C_RCPT_TYPE_BT	,Row
						
				ggoSpread.SpreadLock		C_REF_NOTE_NO,		Row, C_REF_NOTE_NO	,Row			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		Row, Row				
			Case "NR"						'신규받을어음(Note Receivable)			
				ggoSpread.SpreadLock		C_RCPT_TYPE,		Row, C_RCPT_TYPE	,Row			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		Row, Row
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		Row, C_RCPT_TYPE_BT	,Row
						
				ggoSpread.SpreadUnLock		C_REF_NOTE_NO,		Row, C_REF_NOTE_NO	,Row			
				ggoSpread.SSSetRequired		C_REF_NOTE_NO,		Row, Row
				ggoSpread.SpreadUnLock		C_REF_NOTE_BT,		Row, C_RCPT_TYPE_BT	,Row	
						
				ggoSpread.SpreadLock		C_BANK_ACCT,		Row, C_BANK_ACCT	,Row			
				ggoSpread.SSSetProtected	C_BANK_ACCT,		Row, Row	
				ggoSpread.SpreadLock		C_BANK_CD,			Row, C_BANK_CD		,Row			
				ggoSpread.SSSetProtected	C_BANK_CD,			Row, Row		
			Case Else
				ggoSpread.SpreadLock		C_RCPT_TYPE,		Row, C_RCPT_TYPE	,Row			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		Row, Row
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		Row, C_RCPT_TYPE_BT	,Row			
			
				ggoSpread.SpreadLock		C_REF_NOTE_NO,		Row, C_REF_NOTE_NO	,Row			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		Row, Row							
			
				ggoSpread.SpreadLock		C_BANK_ACCT,		Row, C_BANK_ACCT	,Row			
				ggoSpread.SSSetProtected	C_BANK_ACCT,		Row, Row	
				ggoSpread.SpreadLock		C_BANK_CD,			Row, C_BANK_CD		,Row			
				ggoSpread.SSSetProtected	C_BANK_CD,			Row, Row					
		End Select
		
		ggoSpread.SpreadLock		C_BANK_ACCT,		Row, C_BANK_ACCT			,Row			
		ggoSpread.SSSetProtected	C_BANK_ACCT,		Row, Row	
		ggoSpread.SpreadLock		C_BANK_CD,			Row, C_BANK_CD			,Row			
		ggoSpread.SSSetProtected	C_BANK_CD,			Row, Row
		ggoSpread.SpreadLock		C_BANK_NM,			Row, C_BANK_NM			,Row			
		ggoSpread.SSSetProtected	C_BANK_NM,			Row, Row				
		
'		ggoSpread.SpreadLock		C_ACCT_CD,			Row, C_ACCT_CD			,Row			
'		ggoSpread.SSSetProtected	C_ACCT_CD,			Row, Row	
		
		.Col = C_RCPT_TYPE			
		.Text = ""
		.Col = C_RCPT_TYPE_NM			
		.Text = ""
		.Col = C_REF_NOTE_NO			
		.Text = ""
		.Col = C_ACCT_CD			
		.Text = ""
		.Col = C_ACCT_NM			
		.Text = ""
		.Col = C_BANK_ACCT			
		.Text = ""
		.Col = C_BANK_CD			
		.Text = ""
		.Col = C_BANK_NM			
		.Text = ""
			
		.ReDraw = True	
	End With
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    	
    '-----------------------
    'Check previous data area
    '----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")		'☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables														'⊙: Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then										'⊙: This function check indispensable field
		Exit Function
    End If
    
    Call ggoOper.LockField(Document, "N")									'⊙: This function lock the suitable field

	'-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery															'☜: Query db data
       
    FncQuery = True															'⊙: Processing is OK
	Set gActiveElement = document.activeElement          
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False														'⊙: Processing is NG
    
	'-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
	    IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")	'☜ 바뀐부분 
	     If IntRetCD = vbNo Then
	         Exit Function
	     End If
	End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")								'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")								'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")								'⊙: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables													'⊙: Initializes local global variables

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call SetToolbar("1110100000000011")									'⊙: 버튼 툴바 제어 

    FncNew = True														'⊙: Processing is OK

	frm1.txtNoteNoQry.focus 
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False														'⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        intRetCD = DisplayMsgBox("900002","x","x","x")						'☜ 바뀐부분 
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If    
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")			'☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If    
    
	'-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001","x","x","x")  '☜ 바뀐부분 
		Exit Function
	End If
	
	'-----------------------
	  'Check content area
	'-----------------------
    If Not chkField(Document, "2") Then										'⊙: Check contents area
		Exit Function
    End If
   
	'-----------------------
	'sum(single amt) = sum(multi amt) check
	'----------------------- 
	Call DoSum()
	 
	If chkSttlAmt() = False Then
		DisplayMsgBox "113119","X","X","X"
		Exit Function 
	End If 		    
  
	'-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                '☜: Save db data
       
    FncSave = True                                                          '⊙: Processing is OK
    Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	With frm1
   		.vspdData.ReDraw = False

		If .vspdData.MaxRows < 1 Then Exit Function

		ggoSpread.Source = .vspdData
		ggoSpread.CopyRow

		MaxSpreadVal .vspdData, C_SEQ , .vspdData.ActiveRow

		Call SetSpreadColor(.vspdData.ActiveRow)

		.vspdData.ReDraw = True
	End With

	Set gActiveElement = document.activeElement    
End Function

'==========================================================================================
'   Event Desc : Grid의 Max Count 를 찾는다.
'==========================================================================================
Function MaxSpreadVal(ByVal objSpread, ByVal intCol, byval Row)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue   Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1

	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.Text	= MaxValue

end Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo

	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim imRow
    Dim ii,iCurRowPos
    
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error status
    
    FncInsertRow = False															'☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1
		iCurRowPos = .vspdData.ActiveRow
        .vspdData.Redraw = False
        ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		
		For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
			Call MaxSpreadVal(.vspdData, C_SEQ, ii)
		Next
		
		.Col = 2																	' 컬럼의 절대 위치로 이동      
		.Row = 	ii - 1
		.Action = 0
		
        Call SetSpreadColor(iCurRowPos + 1)
        .ReDraw = True
	End With        

    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    If Frm1.vspdData.MaxRows < 1 Then
       Exit function
	End if	

	lgBlnFlgChgValue = True

    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With

    Set gActiveElement = document.ActiveElement  
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'☜: 화면 유형 
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                     '☜:화면 유형, Tab 유무 
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
	Call InitCombo()
    Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadLock()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003				'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNoQry.value)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&hGlNo=" & Trim(frm1.hGlNo.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&hTempGlNo=" & Trim(frm1.hTempGlNo.value)		'☜: 삭제 조건 데이타 
       
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	Call FncQuery()
End Function

'========================================================================================
' Function Name : DoSum() 
' Function Desc : 스프레드의 합을 구한다.
'========================================================================================
Sub DoSum()
	Dim tmpSttlSum		
	DIm Row	
	
	tmpSttlSum = 0 	
	
	With frm1
		For row = 1 To .vspdData.maxRows
			.vspdData.Col = 0
			.vspdData.Row = row
				
			If .vspdData.Text <> ggoSpread.DeleteFlag Then
				.vspdData.Col = C_STTL_AMT
				tmpSttlSum = CDbl(tmpSttlSum) + unicdbl(.vspdData.text) 
				
				'UNIConvNumPCToCompanyByCurrency(tmpSttlSum,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				.htxtSumSttlAmt.text = tmpSttlSum				
			End If	
		Next	
	End With 
End Sub

Function chkSttlAmt()
	chkSttlAmt = True  
	
	With frm1		
		If uniCdbl(.htxtSumSttlAmt.text) <> uniCdbl(.txtNoteAmt.text) +  uniCdbl(.txtIntRevAmt.text) Then 
			chkSttlAmt = False 
			Exit Function 
		End If 		
	End With
End Function 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbQuery = False                                                         '⊙: Processing is NG
    
	Call LayerShowHide(1)

	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001			'☜: 비지니스 처리 ASP의 상태 
			strVal = strVal & "&txtNoteNoQry=" & Trim(.hNoteNo.value)		'☆: 조회 조건 데이타 
		Else		
			strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001			'☜: 비지니스 처리 ASP의 상태 
			strVal = strVal & "&txtNoteNoQry=" & Trim(.txtNoteNoQry.value)	'☆: 조회 조건 데이타 
		End If
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo	=" & lgPageNo         
			strVal = strVal & "&txtMaxRows	=" & .vspdData.MaxRows
	End With

    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
 	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call InitData
	Call SetToolbar("1111111100111111")
	
	Call SetSpreadLock()		
	
	If frm1.vspdData.MaxRows > 0 Then  
		lgIntFlgMode = Parent.OPMD_UMODE									'⊙: Indicates that current mode is Update mode
	Else 
		lgIntFlgMode = Parent.OPMD_CMODE									'⊙: Indicates that current mode is Update mode
	End If
	
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Dim strVal
	Dim lRow
	Dim lGrpCnt	
	Dim	intRetCd			

    Err.Clear																'☜: Protect system from crashing
	DbSave = False															'⊙: Processing is NG	
	
	With frm1
		.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode	
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
		    .vspdData.Col = 0		    			
			If  .vspdData.Text <> ggoSpread.DeleteFlag Then 			
				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep	'☜: C=Create, 순번		0,1 
									
				.vspdData.Col = C_SEQ
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 순번	2
				.vspdData.Col = C_STTL_TYPE
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 처리유형	3
				.vspdData.Col = C_RCPT_TYPE
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 입금유형	4
				.vspdData.Col = C_ACCT_CD 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 입금계정코드	5
				.vspdData.Col = C_REF_NOTE_NO 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 받을어음번호	6
				.vspdData.Col = C_BANK_ACCT 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 계좌번호		7
				.vspdData.Col = C_BANK_CD
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 은행코드		8
				.vspdData.Col = C_STTL_AMT
				strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & Parent.gColSep	' 처리금액		9
				.vspdData.Col = C_NOTE_ITEM_DESC 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep		' 비고			10						
				
				lGrpCnt = lGrpCnt + 1
			End If 				
		Next			
		
		frm1.txtSpread.Value = strVal						
		
		If frm1.txtSpread.Value = "" Then
		'☜ spread전체 delete시[부도어음결제를 취소하시겠습니까?]
			intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")   
			If intRetCd = VBNO Then
				Exit Function
			End If
		
			If  DbDelete = False Then
				Exit Function
			End If
			Exit Function
		End If 	
		
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal				
		
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)											
	End With		
		
    DbSave = True																'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk(Byval ptxtNoteNo)												'☆: 저장 성공후 실행 로직 
    Select Case lgIntFlgMode
		Case Parent.OPMD_CMODE
			frm1.txtNoteNoQry.value = ptxtNoteNo
    End Select

    Call InitVariables
    Call MainQuery()
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>부도어음처리</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopuptempGL()">결의전표</A>|
											<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
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
									<TD CLASS=TD5 NOWRAP>어음번호</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtNoteNoQry" NAME="txtNoteNoQry" SIZE=30 MAXLENGTH=30 tag="12XXXU"ALT="어음번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteQry" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenNoteInfo"></TD>
								<TR>		
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
								<TD CLASS=TD5 NOWRAP>처리일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDateTime1_txtStsDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptCD" NAME="txtDeptCD" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON"  ONCLICK="vbscript:Call OpenPopUpDept(frm1.txtDeptCD.Value, 1)">&nbsp;
													<INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 MAXLENGTH=40 STYLE="TEXT-ALIGN: left" tag="24X" ALT="부서"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>어음구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="어음구분" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
								<TD CLASS=TD5 NOWRAP>어음상태</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="어음상태" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발행일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDateTime1_txtIssueDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>만기일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDateTime2_txtDueDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10   tag="24XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBpCd.Value, 4)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="24X" ALT="거래처"> </TD>
								<TD CLASS=TD5 NOWRAP>지급은행</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10   tag="24XXXU" ALT="은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 5)">&nbsp;
													 <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="은행"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>어음금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDoubleSingle1_txtNoteAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>결제금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDoubleSingle2_txtSttlAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>이자수익</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDoubleSingle1_txtIntRevAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>이자계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIntAcctCd" ALT="이자수익계정" SIZE="10" MAXLENGTH="20"  tag="22X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIntAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtIntAcctCd.value, 7)">
													 <INPUT NAME="txtIntAcctNm" ALT="이자수익계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=4><INPUT CLASS="clstxt" TYPE=TEXT ID="txtNoteDesc" NAME="txtNoteDesc" SIZE=70 MAXLENGTH=128  tag="2XX" ALT="비고"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<script language =javascript src='./js/f5121ma1_OBJECT1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hGlNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTempGlNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="horgchangeid" tag="2" TABINDEX="-1">
<!-- 스프래드의 처리금액 sum -->
<script language =javascript src='./js/f5121ma1_hOBJECT1_htxtSumSttlAmt.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

