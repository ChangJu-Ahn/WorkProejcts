
<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5104ma1
'*  4. Program Name         : 만기어음일괄처리 
'*  5. Program Desc         : 만기어음일괄처리 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/10/16
'*  8. Modified date(Last)  : 2002/02/15
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Soo Min, Oh
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   ***************************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->			<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Const BIZ_PGM_ID = "f5104mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "f5104mb2.asp"											 '☆: 비지니스 로직 ASP명 : Tab1의 ADO 조회용  
Const BIZ_PGM_ID3 = "f5104mb3.asp"											 '☆: 비지니스 로직 ASP명 : Tab2의 ADO 조회용 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

'TAB1, vspddata
Dim C_PROC_CHK
Dim C_NOTE_NO
Dim C_NOTE_AMT
Dim C_DUE_DT
Dim C_NOTE_STS
Dim C_BANK_CD
Dim C_BANK_NM
Dim C_BP_CD	
Dim C_BP_NM	
Dim C_DEPT_CD
Dim C_DEPT_NM
Dim C_NOTE_ITEM_DESC
Dim C_GL_NO
Dim C_TEMP_GL_NO		
Dim C_COL_END

'TAB2, vspddata2
Dim C_CNCL_CHK	
Dim C_CNCL_NOTE_NO	
Dim C_CNCL_TEMP_GL_NO	
Dim C_CNCL_TEMP_GL_DT	
Dim C_CNCL_GL_NO	
Dim C_CNCL_GL_DT	
Dim C_CNCL_NOTE_AMT	
Dim C_CNCL_BP_CD	
Dim C_CNCL_BP_NM	
Dim C_CNCL_DEPT_CD	
Dim C_CNCL_DEPT_NM	
Dim C_CNCL_RCPT_TYPE		'☜: hidden field(11~16, 취소시 필요)	
Dim C_CNCL_ORG_CHANGE_ID
Dim C_CNCL_GL_DEPT_CD	
Dim C_CNCL_INTERNAL_CD	
Dim C_CNCL_NOTE_ITEM_DESC
Dim C_CNCL_COL_END		


'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================

Dim lgBlnFlgConChg				'☜: Condition 변경 Flag
Dim  gSelframeFlg

Dim lgStrPrevKeyNoteNo	' 이전 값 (CG, DG)
Dim lgStrPrevKeyTempGlNo		'이전 TEmp Gl 값(DG)
Dim lgStrPrevKeyGlNo    ' 이전 GL 값 (DG)

Dim lgStrPrevKeyNoteNo1	' 이전 값 (CG, DG)
Dim lgStrPrevKeyTempGlNo1		'이전 TEmp Gl 값(DG)
Dim lgStrPrevKeyGlNo1    ' 이전 GL 값 (DG)

Dim  IsOpenPop          

Dim lgPageNo1
Dim lstxtPlanAmtSum

'++++++++변수 선언 2002.01.10 추가 사항 ++++++++++++++++
<%
Dim dtToday 
dtToday = GetSvrDate
%>

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

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

    lgIntFlgMode = Parent.OPMD_CMODE   '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '⊙: Indicates that no value changed
    lgIntGrpCount = 0           '⊙: Initializes Group View Size
    lgPageNo         = 0
	lgPageNo1        = 0
	lgStrPrevKeyNoteNo	= ""
	lgStrPrevKeyGlNo	= ""
	lgStrPrevKeyNoteNo1 = ""
	lgStrPrevKeyGlNo1	= ""
	
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'☆: 사용자 변수 초기화 
    lgSortKey = 1
    
End Sub

Sub initSpreadPosVariables(ByVal spdsep2)
      
      select case spdsep2
      case "A"
      C_PROC_CHK	= 1
      C_NOTE_NO		= 2
      C_NOTE_AMT	= 3
      C_DUE_DT		= 4
      C_NOTE_STS	= 5 
      C_BANK_CD		= 6
      C_BANK_NM		= 7
      C_BP_CD		= 8
      C_BP_NM		= 9
      C_DEPT_CD		= 10
      C_DEPT_NM		= 11
      C_NOTE_ITEM_DESC		= 12
      C_GL_NO		= 13
      C_TEMP_GL_NO	= 14
      C_COL_END		= 15
      
      Case "B"
      C_CNCL_CHK			= 1
      C_CNCL_NOTE_NO		= 2
      C_CNCL_TEMP_GL_NO		= 3
      C_CNCL_TEMP_GL_DT		= 4
      C_CNCL_GL_NO			= 5
      C_CNCL_GL_DT			= 6	
      C_CNCL_NOTE_AMT		= 7
      C_CNCL_BP_CD			= 8
      C_CNCL_BP_NM			= 9
      C_CNCL_DEPT_CD		= 10
	  C_CNCL_DEPT_NM		= 11
      C_CNCL_RCPT_TYPE		= 12
      C_CNCL_ORG_CHANGE_ID	= 13
      C_CNCL_GL_DEPT_CD		= 14		
      C_CNCL_INTERNAL_CD	= 15
      C_CNCL_NOTE_ITEM_DESC	= 16
      C_CNCL_COL_END		= 17
      End Select 
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("A","*","NOCOOKIE","BA") %>

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
	Dim strSvrDate
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	
	frDt = UNIDateAdd("M", -1, strSvrDate,Parent.gServerDateFormat)		
	frm1.txtDueDtStart.Text = UNIConvDateAToB(frDt,parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDueDtEnd.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 	
	frm1.txtStsDtStart.Text = UniConvDateAToB(frDt,Parent.gServerDateFormat,Parent.gDateFormat) 
	frm1.txtStsDtEnd.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)	
	frm1.txtGLDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
	
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(ByVal spdsep)        
    
    select case spdsep
    
	Case "A"
    Call initSpreadPosVariables("A")
	     
	  With frm1
    
		.vspdData.MaxCols = C_COL_END
		.vspdData.Col = .vspdData.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.vspdData.ColHidden = True
		.vspdData.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
        Call GetSpreadColumnPos("A")
     
		ggoSpread.SSSetCheck	C_PROC_CHK,		"선택"	  , 10, , "", True, -1
		ggoSpread.SSSetEdit		C_NOTE_NO,		"어음번호", 15, , , 30
		ggoSpread.SSSetFloat	C_NOTE_AMT,		"어음금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetDate		C_DUE_DT,		"만기일자", 10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit		C_NOTE_STS,		"어음상태", 10, , , 15   '''2004.04.08 khj max length 수정 
		ggoSpread.SSSetEdit		C_BANK_CD,		"은행", 10, , , 10
		ggoSpread.SSSetEdit		C_BANK_NM,		"은행명", 20, , , 30
		ggoSpread.SSSetEdit		C_BP_CD,		"거래처", 10, , , 10
		ggoSpread.SSSetEdit		C_BP_NM,		"거래처명", 20, , , 50
		ggoSpread.SSSetEdit		C_DEPT_CD,		"부서", 10, , , 10
		ggoSpread.SSSetEdit		C_DEPT_NM,		"부서명", 20, , , 40
		ggoSpread.SSSetEdit     C_NOTE_ITEM_DESC,	"비고", 30, , , 128        
		ggoSpread.SSSetEdit		C_GL_NO,			"전표번호", 15, , , 18
		ggoSpread.SSSetEdit		C_TEMP_GL_NO,	"결의전표번호", 15, , , 18
		
		
        Call ggoSpread.SSSetColHidden(C_GL_NO,C_GL_NO,True)
        Call ggoSpread.SSSetColHidden(C_TEMP_GL_NO,C_TEMP_GL_NO,True)
         
    End With
    	Call SetSpreadLock("A")                                              '바뀐부분 

	
	case "B"	
    Call initSpreadPosVariables("B")

    With frm1
    
		.vspdData2.MaxCols = C_CNCL_COL_END
		.vspdData2.Col = .vspdData2.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.vspdData2.ColHidden = True
		.vspdData2.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
        Call GetSpreadColumnPos("B")

		ggoSpread.SSSetCheck	C_CNCL_CHK,				"선택"	  , 10, , "", True, -1
		ggoSpread.SSSetEdit		C_CNCL_NOTE_NO,			"어음번호", 15, , , 30				
		ggoSpread.SSSetEdit		C_CNCL_TEMP_GL_NO,		"결의전표번호", 15, , , 18		
		ggoSpread.SSSetDate		C_CNCL_TEMP_GL_DT,		"결의전표일자", 10, 2, Parent.gDateFormat		
		ggoSpread.SSSetEdit		C_CNCL_GL_NO,			"회계전표번호", 15, , , 18		
		ggoSpread.SSSetDate		C_CNCL_GL_DT,			"회계전표일자", 10, 2, Parent.gDateFormat
		ggoSpread.SSSetFloat	C_CNCL_NOTE_AMT,		"어음금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit		C_CNCL_BP_CD,			"거래처", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_BP_NM,			"거래처명", 20, , , 50
		ggoSpread.SSSetEdit		C_CNCL_DEPT_CD,			"부서", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_DEPT_NM,			"부서명", 20, , , 40
		ggoSpread.SSSetEdit		C_CNCL_RCPT_TYPE,		"입금유형", 10, , , 10		
		ggoSpread.SSSetEdit		C_CNCL_ORG_CHANGE_ID,	"ORG CHANGE ID", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_GL_DEPT_CD,		"GL DEPT CODE", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_INTERNAL_CD,		"INTERNAL CODE", 10, , , 10		
		ggoSpread.SSSetEdit     C_CNCL_NOTE_ITEM_DESC,	"비고", 30, , , 128    
  
        Call ggoSpread.SSSetColHidden(C_CNCL_RCPT_TYPE,C_CNCL_RCPT_TYPE,True)
        Call ggoSpread.SSSetColHidden(C_CNCL_ORG_CHANGE_ID,C_CNCL_ORG_CHANGE_ID,True)
        Call ggoSpread.SSSetColHidden(C_CNCL_GL_DEPT_CD,C_CNCL_GL_DEPT_CD,True) 
        Call ggoSpread.SSSetColHidden(C_CNCL_INTERNAL_CD,C_CNCL_INTERNAL_CD,True)

    End With

	Call SetSpreadLock("B")                                              '바뀐부분 
	End Select 


End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(ByVal spdsep1)
	Dim RowCnt
	Dim strTempGlNo
	Dim strGlNo
	
	select case spdsep1
	
	case "A"
	ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
		.ReDraw = False
		ggoSpread.SpreadLock C_NOTE_NO,	-1, C_NOTE_NO			' 어음번호 
		ggoSpread.SpreadLock C_NOTE_AMT,-1, C_NOTE_AMT			' 어음금액 
		ggoSpread.SpreadLock C_DUE_DT,	-1, C_DUE_DT			' 만기일 
		ggoSpread.SpreadLock C_NOTE_STS,-1, C_NOTE_STS			' 어음상태 
		ggoSpread.SpreadLock C_BANK_CD,	-1, C_BANK_CD			' 발행은행 
		ggoSpread.SpreadLock C_BANK_NM,	-1, C_BANK_NM			' 발행은행 
		ggoSpread.SpreadLock C_BP_CD,	-1, C_BP_CD				' 거래처코드 
		ggoSpread.SpreadLock C_BP_NM,	-1, C_BP_NM				' 거래처명 
		ggoSpread.SpreadLock C_DEPT_CD,	-1, C_DEPT_CD			' 부서코드 
		ggoSpread.SpreadLock C_DEPT_NM,	-1, C_DEPT_NM			' 부서명 
		ggoSpread.SpreadUnLock C_NOTE_ITEM_DESC, -1, C_NOTE_ITEM_DESC ' 비고 
		.ReDraw = True

    End With

    Case "B"
    ggoSpread.Source = frm1.vspdData2
    With frm1.vspdData2
		.ReDraw = False			    		
		ggoSpread.SpreadLock C_CNCL_NOTE_NO,		-1, C_CNCL_NOTE_NO			' 어음번호		
		ggoSpread.SpreadLock C_CNCL_TEMP_GL_NO,		-1, C_CNCL_TEMP_GL_NO		' 결의전표번호		
		ggoSpread.SpreadLock C_CNCL_TEMP_GL_DT,		-1, C_CNCL_TEMP_GL_DT		' 결의전표일자 
		ggoSpread.SpreadLock C_CNCL_GL_NO,			-1, C_CNCL_GL_NO			' 회계전표번호 
		ggoSpread.SpreadLock C_CNCL_GL_DT,			-1, C_CNCL_GL_DT			' 전표일자 
		ggoSpread.SpreadLock C_CNCL_NOTE_AMT,		-1, C_CNCL_NOTE_AMT			' 전표금액 
		ggoSpread.SpreadLock C_CNCL_BP_CD,			-1, C_CNCL_BP_CD			' 거래처코드 
		ggoSpread.SpreadLock C_CNCL_BP_NM,			-1, C_CNCL_BP_NM			' 거래처명 
		ggoSpread.SpreadLock C_CNCL_DEPT_CD,		-1, C_CNCL_DEPT_CD			' 부서코드 
		ggoSpread.SpreadLock C_CNCL_DEPT_NM,		-1, C_CNCL_DEPT_NM			' 부서명 
		ggoSpread.SpreadLock C_CNCL_NOTE_ITEM_DESC,	-1,	C_CNCL_NOTE_ITEM_DESC	' 비고 
    	.ReDraw = True

    End With
    
    Case "C"
    ggoSpread.Source = frm1.vspdData2
    With frm1.vspdData2
		.ReDraw = False			    
		For RowCnt = 1 To .MaxRows
			
			.Row = RowCnt
			.Col = C_CNCL_TEMP_GL_NO
			strTempGlNo = .Text
			.Col = C_CNCL_GL_NO
			strGlNo = .Text
			If strTempGlNo <> "" and strGlNo <> ""Then				
				ggoSpread.SpreadLock		C_CNCL_CHK	, RowCnt	, C_CNCL_CHK	, RowCnt				
			Else 				
				ggoSpread.SpreadUnLock	C_CNCL_CHK	, RowCnt	, C_CNCL_CHK	, RowCnt
			End If
			
		Next		
    End With
    End Select

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1

		.vspdData.ReDraw = False
		.vspdData.ReDraw = True
    
    End With

End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
  	Dim arrData

		'어음구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteFg1 ,lgF0  ,lgF1  ,Chr(11))
	
	'어음구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteFg2 ,lgF0  ,lgF1  ,Chr(11))
	
	'어음상태 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1011", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))

	
	
End Sub


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

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0			'입금/출금유형 
			arrParam(0) = "입금/출금유형 팝업"
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 1 AND B.REFERENCE = " & FilterVar("RP", "''", "S") & "  "
			arrParam(5) = "입금/출금유형"
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtRcptType.Alt
			arrHeader(1) = frm1.txtRcptTypeNm.Alt

		Case 1			'입출금은행 
		    arrParam(0) = "은행 팝업"	' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE 명칭 
			arrParam(2) = strCode																	' Code Condition
			arrParam(3) = ""																			' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "											' Where Condition			
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD " 
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO " 
			arrParam(4) = arrParam(4) & "AND (C.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR C.DPST_FG = " & FilterVar("ET", "''", "S") & " ) " 
			arrParam(5) = "은행코드"															' 조건필드의 라벨 명칭 

			arrField(0) = "A.BANK_CD"							' Field명(0)
			arrField(1) = "A.BANK_NM"							' Field명(1)	
			arrField(2) = "B.BANK_ACCT_NO" 				' Field명(2) 		
    
			arrHeader(0) = "은행코드"					' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)			
			arrHeader(2) = "계좌번호" 					' Header명(2)

		Case 3			'부서 
			arrParam(0) = "부서 팝업"	' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT"		 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(parent.gChangeOrgId, "''", "S") & ""	' Where Condition
			arrParam(5) = "부서"					' 조건필드의 라벨 명칭 

			arrField(0) = "DEPT_CD"						' Field명(0)
			arrField(1) = "DEPT_NM"						' Field명(1)
    
			arrHeader(0) = "부서코드"					' Header명(0)
			arrHeader(1) = "부서명"						' Header명(1)

		Case 4			'계좌번호 
		
			arrParam(0) = "계좌번호 팝업" 							' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, B_MINOR C, B_MINOR D, F_DPST E " 		' TABLE 명칭 
			arrParam(2) = strCode 								' Code Condition 
			arrParam(3) = "" 									' Name Cindition 
			arrParam(4) = "A.BANK_CD = B.BANK_CD " 						' Where Condition 
			arrParam(4) = arrParam(4) & "AND C.MAJOR_CD = " & FilterVar("F3011", "''", "S") & "  AND C.MINOR_CD = B.BANK_ACCT_TYPE " 
			arrParam(4) = arrParam(4) & "AND D.MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  AND D.MINOR_CD = B.DPST_TYPE " 
			arrParam(4) = arrParam(4) & "AND (E.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR E.DPST_FG = " & FilterVar("ET", "''", "S") & " ) " 
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = E.BANK_ACCT_NO " 
			arrParam(4) = arrParam(4) & "AND B.BANK_CD = E.BANK_CD " 
			arrParam(5) = "계좌번호" 							' 조건필드의 라벨 명칭 
				
			arrField(0) = "B.BANK_ACCT_NO" 							' Field명(0) 
			arrField(1) = "A.BANK_CD" 								' Field명(1) 
			arrField(2) = "A.BANK_NM" 								' Field명(2) 
			arrField(3) = "C.MINOR_NM" 							' Field명(3) 
			arrField(4) = "D.MINOR_NM" 							' Field명(4) 
			arrField(5) = "HH" & parent.gColSep & "C.MINOR_CD" 					' Field명(5) - Hidden 
			arrField(6) = "HH" & parent.gColSep & "D.MINOR_CD" 					' Field명(6) - Hidden 

			arrHeader(0) = "계좌번호" 							' Header명(0) 
			arrHeader(1) = "은행코드" 							' Header명(1) 
			arrHeader(2) = "은행명" 							' Header명(2)
			arrHeader(3) = "예적금구분" 							' Header명(3) 
			arrHeader(4) = "예적금유형" 							' Header명(4)
			
		Case 5, 6			'발행은행 
	'	 If frm1.txtBankCd1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "은행 팝업"	' 팝업 명칭 
			arrParam(1) = "B_BANK "						' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = " "									' Where Condition						
			arrParam(5) = "은행코드"											' 조건필드의 라벨 명칭 

			arrField(0) = "BANK_CD"						' Field명(0)
			arrField(1) = "BANK_NM"						' Field명(1)			
    
			arrHeader(0) = "은행코드"					' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)			
			
		Case 7
			If frm1.txtNoteAcctCd.className = "protected" Then Exit Function    

			arrParam(0) = "입금/출금계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN001", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN001", "''", "S") & " " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ"
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND C.JNL_CD =  " & FilterVar(frm1.cboNoteFg1.Value, "''", "S") 	 	
			If frm1.txtRcptType.Value<>"" then
				arrParam(4) = arrParam(4) & " AND D.EVENT_CD =  " & FilterVar(UCase(frm1.txtRcptType.Value), "''", "S")
			End if
			arrParam(5) = frm1.txtNoteAcctCd.Alt							' 조건필드의 라벨 명칭 
			
			arrField(0) = "A.ACCT_CD"									' Field명(0)
			arrField(1) = "A.ACCT_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"					 					' Field명(3)
			
			arrHeader(0) = frm1.txtNoteAcctCd.Alt									' Header명(0)
			arrHeader(1) = frm1.txtNoteAcctNm.Alt								' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)	
		
		Case 8, 9			'어음번호 
	'	 If frm1.txtBankCd1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "어음번호 팝업"						' 팝업 명칭 
			arrParam(1) = "F_NOTE	A, B_BANK	B, B_MINOR C "		' TABLE 명칭 
			arrParam(2) = strCode									' Code Condition
			arrParam(3) = ""										' Name Condition
			arrParam(4) = " A.NOTE_FG = " & FilterVar(UCase(frm1.cboNoteFg1.Value), "''", "S") 	' Where Condition
			arrParam(4) = arrParam(4) & " AND C.MAJOR_CD = " & FilterVar("F1011", "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND A.NOTE_STS = C.MINOR_CD "
			arrParam(4) = arrParam(4) & " AND A.BANK_CD = B.BANK_CD "
			arrParam(5) = "어음번호"											' 조건필드의 라벨 명칭 

			arrField(0) = "A.NOTE_NO"						' Field명(0)
			arrField(1) = "A.BANK_CD"						' Field명(1)			
			arrField(2) = "B.BANK_NM"						' Field명(0)			
    
			arrHeader(0) = "어음번호"					' Header명(0)
			arrHeader(1) = "은행코드"					' Header명(0)
			arrHeader(2) = "은행명"						' Header명(1)	
			
			
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	If  (iWhere =1 or iWhere = 4) Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
			Select Case iWhere
				Case 0		' 입금/출금유형 
					.txtRcptType.focus
				Case 1		' 은행 
					.txtBankCd1.focus												
				Case 2		' 발행은행(tab1)
					.txtBankCd.focus			
				Case 3		' 부서 
					.txtDeptCd.focus
				Case 4		' 계좌번호 
					.txtBankAcctNo.focus
				Case 5		' 발행은행(tab2)
					.txtBankCd2.focus
				Case 6 
					.txtBankCd.focus
				Case 7
					.txtNoteAcctCd.focus
				Case 8
					.txtNoteNo.focus
				Case 9
					.txtNoteNo1.focus
			End Select
		End With
		Exit Function
	End If	

	With frm1
		Select Case iWhere
			Case 0		' 입금/출금유형 
				.txtRcptType.value	= arrRet(0)
				.txtRcptTypeNm.value= arrRet(1)		
				.txtRcptType.focus
				Call txtRcptType_OnChange()
			Case 1		' 은행 
				.txtBankCd1.value	= arrRet(0)
				.txtBankNm1.value	= arrRet(1)
				.txtBankAcctNo.value =  arrRet(2)
				.txtBankCd1.focus												
			Case 2		' 발행은행(tab1)
				.txtBankCd.value	= arrRet(0)
				.txtBankNM.value	= arrRet(1)	
				.txtBankCd.focus			
			Case 3		' 부서 
				.txtDeptCd.value	= arrRet(0)
				.txtDeptNm.value	= arrRet(1)
				.txtDeptCd.focus
			Case 4		' 계좌번호 
				.txtBankAcctNo.value =  arrRet(0)					
				.txtBankCd1.value	= arrRet(1)
				.txtBankNm1.value	= arrRet(2)					
				.txtBankAcctNo.focus
			Case 5		' 발행은행(tab2)
				.txtBankCd2.value	= arrRet(0)
				.txtBankNM2.value	= arrRet(1)	
				.txtBankCd2.focus
			Case 6 
				.txtBankCd.value	= arrRet(0)
				.txtBankNM.value	= arrRet(1)
				.txtBankCd.focus
			Case 7
				.txtNoteAcctCd.value	= arrRet(0)
				.txtNoteAcctNm.value	= arrRet(1)
				.txtNoteAcctCd.focus
			Case 8
				.txtNoteNo.value	= arrRet(0)
			Case 9
				.txtNoteNo1.value	= arrRet(0)
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'============================================================
'회계전표 팝업 
'============================================================
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

	With frm1.vspdData2
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_CNCL_GL_NO
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		End If
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'============================================================
'결의전표 팝업 
'============================================================
Function OpenPopupTempGL()

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

	With frm1.vspdData2
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_CNCL_TEMP_GL_NO
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		End If
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function


Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'부서코드 
	arrParam(1) = frm1.txtGLDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"							'부서권한(lgUsrIntCd)
	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If

	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	Call txtDeptCD_OnChange()
	frm1.txtDeptCd.focus

	lgBlnFlgChgValue = True
End Function

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function


'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtfromBizAreaCd.Value	= arrRet(0)
			frm1.txtfromBizAreaNm.Value	= arrRet(1)
			frm1.txtfromBizAreaCd.focus
		case 1
			frm1.txttoBizAreaCd.Value = arrRet(0)
			frm1.txttoBizAreaNm.Value = arrRet(1)
			frm1.txttoBizAreaCd.focus
		case 2
			frm1.txtfromBizAreaCd1.Value = arrRet(0)
			frm1.txtfromBizAreaNm1.Value = arrRet(1)
			frm1.txtfromBizAreaCd1.focus
		case 3
			frm1.txttoBizAreaCd1.Value	= arrRet(0)
			frm1.txttoBizAreaNm1.Value	= arrRet(1)
			frm1.txttoBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function
'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

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
	
	Call InitVariables							'⊙: Initializes local global variables
    Call LoadInfTB19029							'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.ClearField(Document, "1")      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")		'⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggospread.ClearSpreadData

    Call InitSpreadSheet("A")                                                        'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                        'Setup the Spread sheet
	Call InitComboBox
	Call txtRcptType_OnChange()
    Call cboNoteFg1_OnChange()
    Call SetDefaultVal
    Call ClickTab1

    gIsTab     = "Y" 
	gTabMaxCnt = 2  	

	' [Main Menu ToolBar]의 각 버튼을 [Enable/Disable] 처리하는 부분 
	'1메뉴탐색기/2조회/3신규/4삭제/5저장/6행추가/7행삭제/8취소/9이전/10다음/11레코드복사/12EXPORT/13인쇄/14찾기/15도움말 
    Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어 
	
	Set gActiveElement = document.activeElement
	
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
            
            
             C_PROC_CHK = iCurColumnPos(1)
             C_NOTE_NO = iCurColumnPos(2)
             C_NOTE_AMT = iCurColumnPos(3)
             C_DUE_DT = iCurColumnPos(4)
             C_NOTE_STS = iCurColumnPos(5)
             C_BANK_CD = iCurColumnPos(6)
             C_BANK_NM = iCurColumnPos(7)
             C_BP_CD	= iCurColumnPos(8)
             C_BP_NM	= iCurColumnPos(9)
             C_DEPT_CD= iCurColumnPos(10)
             C_DEPT_NM= iCurColumnPos(11)
             C_NOTE_ITEM_DESC   = iCurColumnPos(12)
             C_GL_NO	= iCurColumnPos(13)
             C_TEMP_GL_NO = iCurColumnPos(14)
             C_COL_END= iCurColumnPos(15)
            
      Case "B"
      
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
             C_CNCL_CHK	= iCurColumnPos(1)
             C_CNCL_NOTE_NO	= iCurColumnPos(2) 
			 C_CNCL_TEMP_GL_NO	= iCurColumnPos(3)
			 C_CNCL_TEMP_GL_DT =  iCurColumnPos(4)
             C_CNCL_GL_NO	= iCurColumnPos(5)
             C_CNCL_GL_DT	= iCurColumnPos(6)
             C_CNCL_NOTE_AMT	= iCurColumnPos(7)
             C_CNCL_BP_CD	= iCurColumnPos(8)
             C_CNCL_BP_NM	= iCurColumnPos(9)
             C_CNCL_DEPT_CD	= iCurColumnPos(10)
             C_CNCL_DEPT_NM	= iCurColumnPos(11)
             C_CNCL_RCPT_TYPE		= iCurColumnPos(12)
             C_CNCL_ORG_CHANGE_ID= iCurColumnPos(13)
             C_CNCL_GL_DEPT_CD	= iCurColumnPos(14)
             C_CNCL_INTERNAL_CD	= iCurColumnPos(15)              
             C_CNCL_NOTE_ITEM_DESC = iCurColumnPos(16)
             C_CNCL_COL_END		= iCurColumnPos(17)
 
    End Select    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Desc : 입금유형별 Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptType_OnChange()
	'은행코드, 계좌번호 Protected Setting
	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	strval = UCase(frm1.txtRcptType.value)
	
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
	
			Select Case UCase(lgF0)
				Case "DP" & Chr(11)			' 예적금			
					Call ggoOper.SetReqAttr(frm1.txtBankCd1, "N")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "N")
				Case Else
					frm1.txtBankCd1.value = ""
					frm1.txtBankNm1.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")									
			End Select
	Else
			frm1.txtBankCd1.value = ""
			frm1.txtBankNm1.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")											
	End If 
	frm1.txtNoteAcctCd.value = ""
	frm1.txtNoteAcctNm.value = ""
	
End Sub

Sub cboNoteFg1_OnChange()		
	If frm1.cboNoteFg1.value = "D1" Then			
		Call ggoOper.SetReqAttr(frm1.cboNoteSts, "D")
	Else 
		frm1.cboNoteSts.value = ""		
		Call ggoOper.SetReqAttr(frm1.cboNoteSts, "Q")
	End If
	
End Sub
'2005/05/24 배서어음,할인어음일때는 입/출금유형,계정 필수입력에서 제외 
Sub cboNoteSts_OnChange()		
	If frm1.cboNoteSts.value = "DC" or frm1.cboNoteSts.value = "SE"  Then			
		Call ggoOper.SetReqAttr(frm1.txtRcptType, "D")
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "D")
	Else 
		Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "N") 
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDtStart_DblClick(Button)
	if Button = 1 then
		frm1.txtDueDtStart.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDueDtStart.Focus
	End if
End Sub

Sub txtDueDtEnd_DblClick(Button)
	if Button = 1 then
		frm1.txtDueDtEnd.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDueDtEnd.Focus
	End if
End Sub

Sub txtGLDt_DblClick(Button)
	if Button = 1 then
		frm1.txtGLDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtGLDt.Focus
	End if
End Sub

Sub txtStsDtStart_DblClick(Button)
	if Button = 1 then
		frm1.txtStsDtStart.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtStsDtStart.Focus
	End if
End Sub

Sub txtStsDtEnd_DblClick(Button)
	if Button = 1 then
		frm1.txtStsDtEnd.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtStsDtEnd.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name :txtDueDt_keypress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDtStart_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtDueDtEnd.focus
		Call MainQuery
	End If   
End Sub
Sub txtDueDtEnd_KeyPress(KeyAscii)
	If KeyAscii = 13 Then  
		frm1.txtDueDtStart.focus
		Call MainQuery
	End If   
End Sub

Sub txtStsDtStart_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtStsDtEnd.focus
	   Call MainQuery
	End If   
End Sub

Sub txtStsDtEnd_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtStsDtStart.focus
	   Call MainQuery
	End If   
End Sub

Sub txtDueDtEnd_Change()
End Sub

Sub txtGLDt_Change()
    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtGLDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

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
End Sub

'=======================================================================================================
'   Event Name : txtBankAcctNo_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDeptCD_OnChange()
	If frm1.txtDeptCD.value = "" then
		frm1.txtDeptNm.value = ""
		Exit Sub
	End If
	
	Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii
	If Trim(frm1.txtDeptCd.value) = "" and Trim(frm1.txtGLDt.Text <> "") Then		Exit Sub

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"			
		
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
		'----------------------------------------------------------------------------------------

     lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_onBlur()
	if frm1.txtBankCd.value = "" then
		frm1.txtBankNm.value = ""
	end if
End Sub	

Sub txtBankCd1_onBlur()
	if frm1.txtBankCd1.value = "" then
		frm1.txtBankNm1.value = ""
	end if
End Sub	

Sub txtRcptType_onBlur()
	if frm1.txtRcptType.value = "" then
		frm1.txtRcptTypeNm.value = ""
	end if
End Sub	

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()	
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
	    Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어	    
	Else                 
	    Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어 
	End If
	
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)														'첫번째 Tab 	
	
	gSelframeFlg = TAB1
	frm1.hProcFg.value = "CG"
	frm1.cboNoteFg1.focus
						 
End Function

Function ClickTab2()
      
   If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetToolBar("1100000000001111")
	ELSE                 
		Call SetToolBar("1100000000001111")
	END IF	
	
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)														'두번째 Tab 
	
	gSelframeFlg = TAB2
	frm1.hProcFg.value = "DG"
	frm1.cboNoteFg2.focus
	
	
End Function


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


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This function is data query with spread sheet scrolling
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


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")
   	gMouseClickStatus = "SPC"	'Split 상태코드 
	
  	Set gActiveSpdSheet = frm1.vspdData2

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData2
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
    
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
   Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
    
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
    
    

End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    
    
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
    
    

End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)

    
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)

    
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
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
         
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If lgStrPrevKeyNoteNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If    
    
End Sub



'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    
    ggoSpread.Source = Frm1.vspdData
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData2_GotFocus()
    
    ggoSpread.Source = Frm1.vspdData2
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    If Col <= C_PROC_CHK Or NewCol <= C_PROC_CHK Then
        Cancel = True
        Exit Sub
    End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
    
  
End Sub

Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
     If Col <= C_CNCL_CHK Or NewCol <= C_CNCL_CHK Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
    
  
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If 
    
   	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	    
		If lgPageNo <> "" Then								
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If
    
End Sub


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData
		.Row = Row
		.Col = C_PROC_CHK
		
		ggoSpread.Source = frm1.vspdData
		
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
		
		.Row = Row
		.Col = C_PROC_CHK
				
		ggoSpread.Source = frm1.vspdData
				
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
				.col = C_NOTE_AMT
				lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumNoteAmt.Text) + UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				frm1.txtSumNoteAmt.Text = lstxtPlanAmtSum
			Else
				ggoSpread.SSDeleteFlag Row,Row				
				.col = C_NOTE_AMT
				lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumNoteAmt.Text) - UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				frm1.txtSumNoteAmt.Text = lstxtPlanAmtSum
			End If		
		End If
			
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData2
		.Row = Row
		.Col = C_PROC_CHK
		
		ggoSpread.Source = frm1.vspdData2
		
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
	End With
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    'FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
   '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.ClearField(Document, "3")			'⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData			'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData2
    ggospread.ClearSpreadData			'⊙: Clear Contents  Field
	
    If gSelframeFlg = TAB1 Then 
       Call InitSpreadSheet("A")
     else
       Call InitSpreadSheet("B")
     End if 
       
    Call InitVariables() 
    Call cboNoteFg1_OnChange()

    frm1.vspdData.MaxRows = 0
	
    '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field     						
	   Exit Function	
    end if 
    
    If frm1.txtfromBizAreaCd.value = "" Then
		frm1.txtfromBizAreaNm.value = ""
	End If
	
	If frm1.txttoBizAreaCd.value = "" Then
		frm1.txttoBizAreaNm.value = ""
	End If
	
	If frm1.txtfromBizAreaCd1.value = "" Then
		frm1.txtfromBizAreaNm1.value = ""
	End If
	
	If frm1.txttoBizAreaCd1.value = "" Then
		frm1.txttoBizAreaNm1.value = ""
	End If
	
    
    If gSelframeFlg = "1" Then
		If (frm1.txtDueDtStart.Text <> "") And (frm1.txtDueDtEnd.Text <> "") Then
			If CompareDateByFormat(frm1.txtDueDtStart.Text, frm1.txtDueDtEnd.Text, frm1.txtDueDtStart.Alt, frm1.txtDueDtEnd.Alt, _
						"970025", frm1.txtDueDtStart.UserDefinedFormat, Parent.gComDateType, true) = False Then
				frm1.txtDueDtStart.focus											
				Exit Function
			End if	
		End If
		
		If Trim(frm1.txtfromBizAreaCd.value) <> "" and   Trim(frm1.txttoBizAreaCd.value) <> "" Then				
		  If UCase(Trim(frm1.txtfromBizAreaCd.value)) > UCase(Trim(frm1.txttoBizAreaCd.value)) Then
		  		
		  	IntRetCd = DisplayMsgBox("970025", "X", frm1.txtfromBizAreaCd.Alt, frm1.txttoBizAreaCd.Alt)
		  	frm1.txtfromBizAreaCd.focus
		  	Exit Function
		  End If
		End If
	  
	End IF
	
    If gSelframeFlg = "2" Then
		If (frm1.txtStsDtStart.Text <> "") And (frm1.txtStsDtEnd.Text <> "") Then
			If CompareDateByFormat(frm1.txtStsDtStart.Text, frm1.txtStsDtEnd.Text, frm1.txtStsDtStart.Alt, frm1.txtStsDtEnd.Alt, _
						"970025", frm1.txtStsDtStart.UserDefinedFormat, Parent.gComDateType, true) = False Then
				frm1.txtStsDtStart.focus											
				Exit Function
			End if	
		End If
		
		If Trim(frm1.txtfromBizAreaCd1.value) <> "" and   Trim(frm1.txttoBizAreaCd1.value) <> "" Then				
		If UCase(Trim(frm1.txtfromBizAreaCd1.value)) > UCase(Trim(frm1.txttoBizAreaCd1.value)) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtfromBizAreaCd1.Alt, frm1.txttoBizAreaCd1.Alt)
			frm1.txtfromBizAreaCd1.focus
			Exit Function
		End If
	  End If
	  
	End IF
	
    If frm1.txtBankCd.value = "" Then
		frm1.txtBankNm.value = ""
	End If
	
    Call ggoOper.LockField(Document, "N")			'⊙: This function lock the suitable field

    '-----------------------
    'Query function call area
    '----------------------- 
	
    IF  DbQuery	= False Then						'☜: Query db data
       Exit Function	
    End If
    
    FncQuery = True		
	Set gActiveElement = document.activeElement    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
	Set gActiveElement = document.activeElement    
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                     '☜:화면 유형, Tab 유무 
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
	
	if gMouseClickStatus = "SPCRP" then
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
	
	End If
	
	If gMouseClickStatus = "SP2CRP" Then
	
	iColumnLimit = frm1.vspdData2.MaxCols
	
	ACol = frm1.vspdData2.ActiveCol
	ARow = frm1.vspdData2.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData2.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData2.Col = ACol
	frm1.vspdData2.Row = ARow
	frm1.vspdData2.Action = Parent.SS_ACTION_ACTIVE_CELL
	frm1.vspdData2.ScrollBars = Parent.SS_SCROLLBAR_BOTH
	
	
	end if
End Function

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
    if gSelframeFlg = TAB1 Then
		Call InitSpreadSheet("A") 
    else
		Call InitSpreadSheet("B")      
    end if
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
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
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
Dim strVal

    DbQuery = False
    Err.Clear                '☜: Protect system from crashing
    
	Call LayerShowHide(1)	
	Call txtRcptType_OnChange()		
	Call cboNoteFg1_OnChange()	
	
		
	With frm1
		If gSelframeFlg = "1" Then 														'☜: 일괄처리(tab1) 조회 
		    If lgIntFlgMode = Parent.OPMD_UMODE Then		    
				strVal = BIZ_PGM_ID2 & "?txtMode = " & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
				strVal = strVal & "&cboProcFg=" & Trim(frm1.hProcFg.value)				'☆: 조회 조건 데이타				
				strVal = strVal & "&cboNoteFg=" & Trim(frm1.hNoteFg1.value)
				strVal = strVal & "&cboNoteSts=" & Trim(frm1.hNoteSts.value)				
				strVal = strVal & "&txtDueDtStart=" & Trim(frm1.hDueDtStart.value)
				strVal = strVal & "&txtDueDtEnd=" & Trim(frm1.hDueDtEnd.value)
				strVal = strVal & "&txtBankCd=" & Trim(frm1.hBankCd.value)
				strVal = strVal & "&txtBankCd_Alt=" & Trim(frm1.txtBankCd.Alt)
				'2003/12/12 Oh Soo Min 추가 
				strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.hfromtxtBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(frm1.txtfromBizAreaCd.alt)				
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.htotxtBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd1_Alt=" & Trim(frm1.txttoBizAreaCd.alt)
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo=" & lgStrPrevKeyGlNo
				strVal = strVal & "&lgPageNo=" & lgPageNo
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			
			Else			
				strVal = BIZ_PGM_ID2 & "?txtMode = " & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
				strVal = strVal & "&cboProcFg=" & Trim("CG")							'☆: 조회 조건 데이타				
				strVal = strVal & "&cboNoteFg=" & Trim(frm1.cboNoteFg1.value)				
				strVal = strVal & "&cboNoteSts=" & Trim(frm1.cboNoteSts.value)				
				strVal = strVal & "&txtDueDtStart=" & Trim(frm1.txtDueDtStart.Text)
				strVal = strVal & "&txtDueDtEnd=" & Trim(frm1.txtDueDtEnd.Text)
				strVal = strVal & "&txtBankCd=" & Trim(frm1.txtBankCd.value)
				strVal = strVal & "&txtBankCd_Alt=" & Trim(frm1.txtBankCd.Alt)
				'2003/12/12 Oh Soo Min 추가 
				strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.txtfromBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(frm1.txtfromBizAreaCd.alt)
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.txttoBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd1_Alt=" & Trim(frm1.txttoBizAreaCd.alt)
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo=" & lgStrPrevKeyGlNo
				strVal = strVal & "&lgPageNo=" & lgPageNo
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			End If   						
		Else 
		    If lgIntFlgMode = Parent.OPMD_UMODE Then
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 				
				strVal = strVal & "&cboProcFg=" &  Trim(frm1.hProcFg.value)						'☆: 조회 조건 데이타 				
				strVal = strVal & "&cboNoteFg=" & Trim(.hNoteFg2.value)				
				strVal = strVal & "&txtStsDtStart=" & Trim(.hFrStsDT1.value)
				strVal = strVal & "&txtStsDtEnd=" & Trim(.hToStsDT1.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.hfrBizAreaCd1.value)
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.htoBizAreaCd1.value)
				strVal = strVal & "&txtNoteNo1=" & Trim(.htxtNoteNo1.value)
				strVal = strVal & "&lgPageNo=" & lgPageNo
				strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			Else
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 				
				strVal = strVal & "&cboProcFg=" & Trim("DG")									'☆: 조회 조건 데이타 				
				strVal = strVal & "&cboNoteFg=" & Trim(.cboNoteFg2.value)
				strVal = strVal & "&txtStsDtStart=" & Trim(.txtStsDtStart.Text)
				strVal = strVal & "&txtStsDtEnd=" & Trim(.txtStsDtEnd.Text)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.txtfromBizAreaCd1.value)
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.txttoBizAreaCd1.value)
				strVal = strVal & "&txtNoteNo1=" & Trim(.txtNoteNo1.value)			
				strVal = strVal & "&lgPageNo=" & lgPageNo				
				strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			End If	
		End If		
	End With 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인			
	
	Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()							'☆: 조회 성공후 실행로직 
	
	If gSelframeFlg = "2" Then 					
		Call SetSpreadLock("C")
	End If 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE	'⊙: Indicates that current mode is Update mode    
    
    
     Call InitData()
    
	lgBlnFlgChgValue = False
	
	' 현재 Page의 From Element들을 사용자가 입력을 받지 못하게 하거나 필수입력사항을 표시한다.
	' LockField(pDoc, pACode)
	
'   Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
    Call txtRcptType_OnChange()
    Call cboNoteFg1_OnChange()
        
    frm1.txtGLDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)	

End Function


'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row)

	Dim strVal	
	Dim lngRows
		
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
	
	Dim strTableid
	Dim strColid
	Dim strColNm	
	Dim strMajorCd	
	Dim strNmwhere
	Dim i
	Dim arrVal
	
	With frm1
		.htxtGlNo.value = frm1.txtGlNo.value
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
	    frm1.vspdData2.ReDraw = false	
	    
        If CopyFromData(.hItemSeq.Value) = True Then
			SetSpread2Color 	
			frm1.vspdData2.ReDraw = True
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =	" a.note_no,  b.note_amt, b.due_dt, e.minor_nm,b.bank_Cd, c.bank_nm,  b.bp_Cd, d.bp_nm "
		
    	strFrom = " f_note_item a, f_note b, b_bank c, b_biz_partner d, b_minor e (NOLOCK) "
		
		strWhere =			  " a.note_no = b.note_no & "' "
		strWhere = strWhere & " AND a.gl_no = " & .htxtGlNo.Value & " "
		strWhere = strWhere & " AND a.note_sts= " & FilterVar("SM", "''", "S") & "   "
		strWhere = strWhere & " AND b.bank_cd = c.bank_cd "
		strWhere = strWhere & "	AND b.bp_cd = d.bp_cd "
		strWhere = strWhere & " AND e.major_cd = " & FilterVar("f1008", "''", "S") & "  "		
		strWhere = strWhere & " AND e.minor_cd = a.note_sts "
		strWhere = strWhere & " ORDER BY a.note_no "	
				
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SSShowData lgF2By2							
			
			For lngRows = 1 To frm1.vspdData3.Maxrows
				frm1.vspddata3.row = lngRows	
				frm1.vspddata3.col = C_Tableid 
				IF Trim(frm1.vspdData3.text) <> "" Then
					frm1.vspddata3.col = C_Tableid
					strTableid = frm1.vspdData3.text
					frm1.vspddata3.col = C_Colid
					strColid = frm1.vspdData3.text
					frm1.vspddata3.col = C_ColNm
					strColNm = frm1.vspdData3.text	
					frm1.vspddata3.col = C_MajorCd					
					strMajorCd = frm1.vspdData3.text	
					
					frm1.vspddata3.col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspdData3.text , "''", "S") & " " 
					
					IF Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") & " "
					End IF				 
					
					IF CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspdData3.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspdData3.text = arrVal(0)
					End IF
				End IF								
				
				strVal = strVal & Chr(11) & .hItemSeq.Value 
				For i = 1 To  C_MajorCd					
				
					frm1.vspdData3.col = i
					strVal = strVal & Chr(11) & frm1.vspdData3.text
				Next	
								
				strVal = strVal & Chr(11) & Chr(12)									
			NEXT					

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
			
		END IF 		
		
		
		intItemCnt = .vspddata.MaxRows
		SetSpread2Color 	
		
	End With
	
	frm1.vspdData3.ReDraw = True
	
	Call LayerShowHide(0)
	
	DbQuery2 = True
	
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim intIndex2 
	Dim strval
	
	strval = 0 
	
	With frm1.vspdData		
		For intRow = 1 To .MaxRows			
			.Row = intRow	
			.Col = C_PROC_CHK
			.text= "1"
		Next
		
		For intRow = 1 To .MaxRows			
			.Row = intRow										
			.Col = C_NOTE_AMT		
			strval = UniCdbl(strval) + UniCdbl(.text)
		Next				
	End With			
		
		frm1.txtSumNoteAmt.text = UNIFormatNumber(UniCdbl(strval), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)				
	
End Sub

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	Dim NoteAmtSum
	Dim ChkCnt
	Dim strGLNo
	Dim ChkFlag
	Dim BatchChk
	Dim intRetCD
	
	DbSave = False				'⊙: Processing is NG

	'2001.03.01 added
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		IntRetCD = DisplayMsgBox("900002","x","x","x")  '조회를 먼저 하십시오.
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"x","x")	'작업을 수행하시겠습니까?

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	If frm1.hProcFg.value = "CG" Then
		If Not chkField(Document, "2") Then                                   '⊙: Check contents area
			Exit Function
		End If
	End If
    
	IF Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
	
	With frm1

		.txtMode.value = Parent.UID_M0002			'☜: 비지니스 처리 ASP 의 상태 
		.txtInsrtUserId.value = Parent.gUsrID

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		If .hProcFg.value = "CG" Then										'☜:일괄처리 
			For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				
				.vspdData.Col = C_PROC_CHK
				
				If .vspdData.Text = "1" Then
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData.Col = C_NOTE_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 어음번호 
					.vspdData.Col = C_TEMP_GL_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 전표번호 
					.vspdData.Col = C_GL_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 결의전표번호 
					.vspdData.Col = C_NOTE_ITEM_DESC
					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep		' 비고 

					lGrpCnt = lGrpCnt + 1
				End If
			Next	
		ElseIf .hProcFg.value = "DG" Then									 '☜:일괄취소 
			For lRow = 1 To .vspdData2.MaxRows
				.vspdData2.Row = lRow
				
				.vspdData2.Col = C_CNCL_CHK
				
				If .vspdData2.Text = "1" Then
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData2.Col = C_CNCL_NOTE_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' 어음번호				
					.vspdData2.Col = C_CNCL_TEMP_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' 회계전표번호 
					.vspdData2.Col = C_CNCL_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gRowSep		' 결의전표번호				
					
					lGrpCnt = lGrpCnt + 1
				End If
			Next	
		End If
		
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
	
		If .txtMaxRows.value <= 0 Then
			Call DisplayMsgBox("900025","x","x","x")  '선택된 항목이 없습니다.
			Exit Function
		End If

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
		
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()			'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0	
	Call MainQuery
	
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

'=======================================================================================================
'   Event Name : Rowcancel() / Rowselect()
'   Event Desc :
'=======================================================================================================    
Function Rowcancel()
	Dim lRow

	If gSelframeFlg = "1" Then 						'☜: 일괄처리(tab1) 조회 
		With Frm1.vspdData
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) = ggoSpread.UPDATEFlag OR Trim(.TEXT) = ggoSpread.INSERTFlag THEN
					.Col = C_PROC_CHK
					.Text = "0"
					IF Trim(.TEXT) = ggoSpread.UPDATEFlag THEN
						ggoSpread.SSDeleteFlag lRow,lRow
					END IF
				END IF
			Next
		End With
	Else
		With Frm1.vspdData2
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) = ggoSpread.UPDATEFlag OR Trim(.TEXT) = ggoSpread.INSERTFlag THEN
					.Col = C_PROC_CHK
					.Text = "0"
					IF Trim(.TEXT) = ggoSpread.UPDATEFlag THEN
						ggoSpread.SSDeleteFlag lRow,lRow
					END IF
				END IF
			Next
		End With
	End If 
End Function

Function Rowselect()
	Dim lRow
	
	If gSelframeFlg = "1" Then 						'☜: 일괄처리(tab1) 조회 
		With Frm1.vspdData
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) <> ggoSpread.DELETEFlag THEN
					.Col = C_PROC_CHK
					If .Lock = False Then
						.Col = C_PROC_CHK
						.Text = "1"
						ggoSpread.UpdateRow lRow
					End If
				END IF
			Next
		End With
	Else
		With Frm1.vspdData2
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) <> ggoSpread.DELETEFlag THEN
					.Col = C_PROC_CHK
					If .Lock = False Then
						.Col = C_PROC_CHK
						.Text = "1"
						ggoSpread.UpdateRow lRow
					End If
				END IF
			Next
		End With
	End If 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
							</TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>만기어음일괄취소</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>
												<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A> &nbsp;|
												<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=10>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
	
			<DIV ID="TabDiv" SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>어음구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg1" NAME="cboNoteFg1" ALT="어음구분" STYLE="WIDTH: 132px" tag="12X"></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>어음상태</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="어음상태" STYLE="WIDTH: 132px" tag="12X"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>만기일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtDueDtStart CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작만기일"></OBJECT>');</SCRIPT>&nbsp; ~ &nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtDueDtEnd CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료만기일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtfromBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtfromBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtfromBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="1XX" ALT="발행은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 6)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNm" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14X" ALT="발행은행명"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txttoBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txttoBizAreaCd.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txttoBizAreaNm" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>어음번호</TD>								
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtNoteNo" NAME="txtNoteNo" SIZE=30 MAXLENGTH=30  tag="1XX" ALT="기준어음번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNo.Value, 8)"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>									
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>회계일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="회계일자" tag="22X1" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" Size= "10" MAXLENGTH="10" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUpDept(frm1.txtDeptCd.value, 3)">&nbsp;<INPUT NAME="txtDeptNm" ALT="부서명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>입금/출금유형</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRcptType" ALT="입금/출금유형코드" SIZE="10" MAXLENGTH="2" tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptType.value, 0)">&nbsp;<INPUT NAME="txtRcptTypeNm" ALT="입금/출금유형명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>입금/출금계정</TD>												
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtNoteAcctCd" ALT="입금/출금계정" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteAcctCd.value, 7)">
																	  <INPUT NAME="txtNoteAcctNm" ALT="입금/출금계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>																													  
								<TD CLASS=TD5 NOWRAP>은행</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd1" NAME="txtBankCd1" SIZE=10 MAXLENGTH=10  tag="21XXXU" ALT="은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd1.Value, 1)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNm1" NAME="txtBankNm1" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="은행명"></TD>
								<TD CLASS=TD5 NOWRAP>계좌번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankAcctNo" NAME="txtBankAcctNo" SIZE=20 MAXLENGTH=30 tag="21XXXU" ALT="계좌번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.Value, 4)"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="33" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TDT">
									<TD CLASS="TD6">
									<TD CLASS="TD5" NOWRAP>어음결제총액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtSumNoteAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 160px" title="FPDOUBLESINGLE" ALT="결제총액" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
				                    </TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
			</DIV>
		 

			<DIV ID="TabDiv"  SCROLL=no>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>어음구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg2" NAME="cboNoteFg2" ALT="어음구분" STYLE="WIDTH: 132px" tag="12X"></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtfromBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtfromBizAreaCd1.value, 2)">&nbsp;<INPUT TYPE=TEXT NAME="txtfromBizAreaNm1" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>회계일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtStsDtStart CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작회계일" tag="12X1"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtStsDtEnd CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료회계일" tag="12X1" ></OBJECT>');</SCRIPT></TD>																																
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txttoBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txttoBizAreaCd1.value, 3)">&nbsp;<INPUT TYPE=TEXT NAME="txttoBizAreaNm1" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>어음번호</TD>								
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtNoteNo1" NAME="txtNoteNo1" SIZE=30 MAXLENGTH=30  tag="1XX" ALT="취소대상어음번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNo1.Value, 9)"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>									
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>						
								</TR>								
							</TABLE>
							    <TR>
									<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="23" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD WIDTH=100% HEIGHT="50%" colspan=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="33" ID=vspdData3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
								</TR>						
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
  		<TD WIDTH="100%">
  			<TABLE <%=LR_SPACE_TYPE_30%>>
   				<TR>
   					<TD WIDTH=10>&nbsp;</TD>
   					<TD><BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:DBSave()" Flag=1>실행</BUTTON>&nbsp;
   						<BUTTON NAME="btncancel" CLASS="CLSSBTN" ONCLICK="vbscript:Rowselect()">전체선택</BUTTON>&nbsp;
						<BUTTON NAME="btnselect" CLASS="CLSSBTN" ONCLICK="vbscript:Rowcancel()">전체취소</BUTTON>
					</TD>
   				</TR>
   			</TABLE> 
  		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hProcFg"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg1"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg2"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteSts"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtStart"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtEnd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtStart"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtEnd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBankCd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"		tag="1" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtGlNo"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hfromtxtBizAreaCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hfrBizAreaCd1"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htotxtBizAreaCd"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htoBizAreaCd1"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtNoteNo1"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFrStsDT1"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hToStsDT1"			tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="CtxtGLDt"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtCRAmt"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtCRLocAmt"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDRAmt"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDRLocAmt"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDocCur"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtXchRate"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtOrgChangeId"	tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDeptCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtAcctCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="GtxtBankCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtBankAcctNo"	tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="DtxtNoteNo"		tag="2" TABINDEX="-1">
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

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
