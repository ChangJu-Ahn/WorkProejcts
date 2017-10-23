<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5122ma1
'*  4. Program Name         : 받을어음이동처리 
'*  5. Program Desc         : 받을어음이동처리 
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
Const BIZ_PGM_ID  = "f5122mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "f5122mb2.asp"											 '☆: 비지니스 로직 ASP명 : Tab1의 ADO 조회용  
Const BIZ_PGM_ID3 = "f5122mb3.asp"											 '☆: 비지니스 로직 ASP명 : Tab2의 ADO 조회용 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

'TAB1, vspddata
Dim C_PROC_CHK
Dim C_FR_DEPT_CD
Dim C_FR_DEPT_NM
Dim C_NOTE_NO	
Dim C_NOTE_AMT
Dim C_NOTE_STS
Dim C_TO_DEPT_CD
Dim C_TO_DEPT_POP
Dim C_TO_DEPT_NM
Dim C_MOVE_DESC
Dim C_BP_CD	
Dim C_BP_NM	
Dim C_ISSUED_DT
Dim C_DUE_DT	

'TAB2, vspddata2
Dim C_CNCL_CHK	
Dim C_CNCL_TO_DEPT_CD	
Dim C_CNCL_TO_DEPT_NM	
Dim C_CNCL_NOTE_NO	
Dim C_CNCL_NOTE_AMT	
DIm C_CNCL_FR_DEPT_CD
Dim C_CNCL_FR_DEPT_NM
Dim C_CNCL_TEMP_GL_NO
Dim C_CNCL_TEMP_GL_DT
Dim C_CNCL_GL_NO	
Dim C_CNCL_GL_DT	
Dim C_CNCL_BP_CD	
Dim C_CNCL_BP_NM	
Dim C_CNCL_ISSUED_DT
Dim C_CNCL_DUE_DT	

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

Dim IsOpenPop          

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
	Select case spdsep2
		Case "A"
			C_PROC_CHK		= 1
			C_FR_DEPT_CD	= 2
			C_FR_DEPT_NM	= 3        
			C_NOTE_NO		= 4
			C_NOTE_AMT		= 5  
			C_NOTE_STS      = 6  
			C_TO_DEPT_CD	= 7
			C_TO_DEPT_POP	= 8
			C_TO_DEPT_NM	= 9
			C_MOVE_DESC     = 10			
			C_BP_CD			= 11	  
			C_BP_NM			= 12 
			C_ISSUED_DT		= 13              
			C_DUE_DT		= 14     
		Case "B"
			C_CNCL_CHK			= 1
			C_CNCL_TO_DEPT_CD	= 2
			C_CNCL_TO_DEPT_NM	= 3              
			C_CNCL_GL_NO		= 4
			C_CNCL_GL_DT		= 5
			C_CNCL_TEMP_GL_NO	= 6      
			C_CNCL_TEMP_GL_DT	= 7
			C_CNCL_NOTE_NO		= 8
			C_CNCL_NOTE_AMT		= 9
			C_CNCL_FR_DEPT_CD	= 10
			C_CNCL_FR_DEPT_NM	= 11        
			C_CNCL_BP_CD		= 12
			C_CNCL_BP_NM		= 13
			C_CNCL_ISSUED_DT	= 14              
			C_CNCL_DUE_DT		= 15     
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
	frm1.txtFromDt.Text = UNIConvDateAToB(frDt,parent.gServerDateFormat,parent.gDateFormat)  '발행일자 Fr
	frm1.txtToDt.Text   = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) '~ To	
	frm1.txtFrGlDt.Text = UniConvDateAToB(frDt,Parent.gServerDateFormat,Parent.gDateFormat)               '두번째 Tab 회계일자 Fr
	frm1.txtToGlDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)                      '~ To	
	frm1.txtGLDt.text   = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)  '이동일 
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
    frm1.hProcFg.value = "CG"
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(ByVal spdsep)        
    Select case spdsep
		Case "A"
			Call initSpreadPosVariables("A")
			     
			With frm1
				.vspdData.MaxCols = C_DUE_DT
				.vspdData.Col = .vspdData.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
				.vspdData.ColHidden = True
				.vspdData.MaxRows = 0
				
				ggoSpread.Source = frm1.vspdData
  				
			    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
			 
			    Call GetSpreadColumnPos("A")

				ggoSpread.SSSetCheck	C_PROC_CHK,		"선택",       5, , "", True, -1
				ggoSpread.SSSetEdit		C_FR_DEPT_CD,	"현재 부서",   8, , , 10
				ggoSpread.SSSetEdit		C_FR_DEPT_NM,	"현재 부서명", 10, , , 40				
				ggoSpread.SSSetEdit		C_NOTE_NO,		"어음번호",   15, , , 30
				ggoSpread.SSSetFloat	C_NOTE_AMT,		"어음금액",   12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetEdit		C_NOTE_STS,		"어음상태",   8, , , 30
				ggoSpread.SSSetEdit		C_TO_DEPT_CD,	"이동부서",   8, , , 10
				ggoSpread.SSSetButton   C_TO_DEPT_POP
				ggoSpread.SSSetEdit		C_TO_DEPT_NM,	"이동부서명", 10, , , 40
				ggoSpread.SSSetEdit		C_MOVE_DESC,	"비고"		, 15, , , 100		
				ggoSpread.SSSetEdit		C_BP_CD,		"거래처",     10, , , 10
				ggoSpread.SSSetEdit		C_BP_NM,		"거래처명",   15, , , 50
				ggoSpread.SSSetDate		C_ISSUED_DT,	"발행일",     10, 2, Parent.gDateFormat
				ggoSpread.SSSetDate		C_DUE_DT,		"만기일",     10, 2, Parent.gDateFormat

			    'Call ggoSpread.SSSetColHidden(C_GL_NO,C_GL_NO,True)
			    'Call ggoSpread.SSSetColHidden(C_TEMP_GL_NO,C_TEMP_GL_NO,True)
			End With
    
			Call SetSpreadLock("A")                                              '바뀐부분 
		Case "B"	
			Call initSpreadPosVariables("B")

			With frm1
				.vspdData2.MaxCols = C_CNCL_DUE_DT
				.vspdData2.Col = .vspdData2.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
				.vspdData2.ColHidden = True
				.vspdData2.MaxRows = 0
				
				ggoSpread.Source = frm1.vspdData2
			    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
			    
			    Call GetSpreadColumnPos("B")

				ggoSpread.SSSetCheck	C_CNCL_CHK,				"선택"	  ,      5, , "", True, -1
				ggoSpread.SSSetEdit		C_CNCL_TO_DEPT_CD,	    "현재 부서",      8, , , 10
				ggoSpread.SSSetEdit		C_CNCL_TO_DEPT_NM,	    "현재 부서명",   10, , , 40	
				ggoSpread.SSSetEdit		C_CNCL_NOTE_NO,			"어음번호",     15, , , 30				
				ggoSpread.SSSetFloat	C_CNCL_NOTE_AMT,		"어음금액",     12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec		
				ggoSpread.SSSetEdit		C_CNCL_FR_DEPT_CD,		"전부서",    8, , , 10
				ggoSpread.SSSetEdit		C_CNCL_FR_DEPT_NM,		"전부서명", 10, , , 40		
				ggoSpread.SSSetEdit		C_CNCL_TEMP_GL_NO,		"결의전표번호", 12, , , 18		
				ggoSpread.SSSetDate		C_CNCL_TEMP_GL_DT,		"결의전표일",	10, 2, Parent.gDateFormat		
				ggoSpread.SSSetEdit		C_CNCL_GL_NO,			"회계전표번호", 12, , , 18		
				ggoSpread.SSSetDate		C_CNCL_GL_DT,			"회계전표일",   10, 2, Parent.gDateFormat
				ggoSpread.SSSetEdit		C_CNCL_BP_CD,			"거래처",		10, , , 10
				ggoSpread.SSSetEdit		C_CNCL_BP_NM,			"거래처명",		15, , , 50
				ggoSpread.SSSetDate		C_CNCL_ISSUED_DT,		"발행일",		10, 2, Parent.gDateFormat
				ggoSpread.SSSetDate		C_CNCL_DUE_DT,			"만기일",		10, 2, Parent.gDateFormat  
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
	
	Select case spdsep1
		Case "A"
			ggoSpread.Source = frm1.vspdData
			With frm1.vspdData
				.ReDraw = False
				ggoSpread.SpreadLock	C_FR_DEPT_CD,	-1, C_FR_DEPT_CD		' 변경전부서 
				ggoSpread.SpreadLock	C_FR_DEPT_NM,	-1, C_FR_DEPT_NM		' 변경전부서명 
				ggoSpread.SpreadLock	C_NOTE_NO,		-1, C_NOTE_NO			' 어음번호 
				ggoSpread.SpreadLock	C_NOTE_AMT,		-1, C_NOTE_AMT			' 어음금액 
				ggoSpread.SpreadLock	C_NOTE_STS,		-1, C_NOTE_STS			' 어음금액				
				ggoSpread.SSSetRequired C_TO_DEPT_CD,	-1, C_TO_DEPT_CD		' 변경후부서 
				ggoSpread.SpreadUnLock	C_TO_DEPT_POP,	-1, C_TO_DEPT_POP		' 변경후부서팝업 
				ggoSpread.SpreadLock	C_TO_DEPT_NM,	-1, C_TO_DEPT_NM		' 변경후부서명 
				ggoSpread.SpreadUnLock	C_MOVE_DESC,	-1, C_TO_DEPT_NM		' 이동시비고 
				ggoSpread.SpreadLock	C_BP_CD,		-1, C_BP_CD				' 거래처코드 
				ggoSpread.SpreadLock	C_BP_NM,		-1, C_BP_NM				' 거래처명 
				ggoSpread.SpreadLock	C_ISSUED_DT,	-1, C_ISSUED_DT			' 어음발행일 
				ggoSpread.SpreadLock	C_DUE_DT,		-1, C_DUE_DT			' 어음만기일 
				
				.ReDraw = True
			End With
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			With frm1.vspdData2
				.ReDraw = False			    		
				ggoSpread.SpreadLock C_CNCL_TO_DEPT_CD,		-1, C_CNCL_TO_DEPT_CD			' 어음번호		
				ggoSpread.SpreadLock C_CNCL_TO_DEPT_NM,		-1, C_CNCL_TO_DEPT_NM			' 어음번호		
				ggoSpread.SpreadLock C_CNCL_NOTE_NO,		-1, C_CNCL_NOTE_NO			' 어음번호		
				ggoSpread.SpreadLock C_CNCL_NOTE_AMT,		-1, C_CNCL_NOTE_AMT			' 전표금액		
				
				ggoSpread.SpreadLock C_CNCL_FR_DEPT_CD,		-1, C_CNCL_FR_DEPT_CD			' 어음번호		
				ggoSpread.SpreadLock C_CNCL_FR_DEPT_NM,		-1, C_CNCL_FR_DEPT_NM			' 어음번호	
						
				ggoSpread.SpreadLock C_CNCL_TEMP_GL_NO,		-1, C_CNCL_TEMP_GL_NO		' 결의전표번호		
				ggoSpread.SpreadLock C_CNCL_TEMP_GL_DT,		-1, C_CNCL_TEMP_GL_DT		' 결의전표일자 
				ggoSpread.SpreadLock C_CNCL_GL_NO,			-1, C_CNCL_GL_NO			' 회계전표번호 
				ggoSpread.SpreadLock C_CNCL_GL_DT,			-1, C_CNCL_GL_DT			' 전표일자 

				ggoSpread.SpreadLock C_CNCL_BP_CD,			-1, C_CNCL_BP_CD			' 거래처코드 
				ggoSpread.SpreadLock C_CNCL_BP_NM,			-1, C_CNCL_BP_NM			' 거래처명 
				ggoSpread.SpreadLock C_CNCL_ISSUED_DT,		-1, C_CNCL_ISSUED_DT			' 부서코드 
				ggoSpread.SpreadLock C_CNCL_DUE_DT,			-1, C_CNCL_DUE_DT			' 부서명 
				
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
	Dim arrParam(8), arrField(6), arrHeader(6)
	Dim strBizAreaCd
	Dim iCalledAspName	

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	    Case 1,2
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			If iWhere = "1" Then
				' 권한관리 추가 
				If lgAuthBizAreaCd <> "" Then
					arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
				Else
					arrParam(4) = ""
				End If
			Else
				strBizAreaCd = Trim(frm1.txtFrBizCd.value)
				
				If strBizAreaCd = "" Then
					strBizAreaCd = "%"
					arrParam(4) = ""						' Where Condition
				Else
					arrParam(4) = "BIZ_AREA_CD NOT LIKE  " & FilterVar(strBizAreaCd, "''", "S") & ""  	' Where Condition						
				End If	
			End If
			
			arrParam(5) = "사업장코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"						' Field명(0)
			arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
			arrHeader(0) = "사업장코드"			' Header명(0)
			arrHeader(1) = "사업장명"
		Case 3		'권한에 의한 부서코드만 Popup
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
		Case 4,5			'부서 
			' 선택한 사업장에 속한 부서만 PopUp
			If iWhere = "3" Then
				strBizAreaCd = Trim(frm1.txtFrBizCd.value)
			Else
				strBizAreaCd = Trim(frm1.txtToBizCd.value)
			End If

			If strBizAreaCd = "" Then
				strBizAreaCd = "%"
			End If

			arrParam(0) = "부서코드팝업"			' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, A_ACCT D "    				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id" & _
			              " from b_acct_dept where org_change_dt = ( select max(org_change_dt)" & _
			              " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))" & _
			              " and C.biz_area_cd LIKE " & FilterVar(strBizAreaCd, "''", "S")  & _
			              " AND B.cost_cd = A.cost_cd " & _
			              " AND C.biz_area_cd = B.biz_area_cd AND D.REL_BIZ_AREA_CD = C.BIZ_AREA_CD"
			              
			arrParam(5) = "부서코드"				' 조건필드의 라벨 명칭 
			
			arrField(0) = "A.DEPT_CD"	     				' Field명(0)
			arrField(1) = "A.DEPT_NM"			    		' Field명(1)
			arrField(2) = "C.BIZ_AREA_CD"			    		' Field명(2)
			arrField(3) = "C.BIZ_AREA_NM"			    		' Field명(3)
			arrField(4) = "A.INTERNAL_CD"
    
			arrHeader(0) = "부서코드"				' Header명(0)
			arrHeader(1) = "부서명"				    ' Header명(1)						
			arrHeader(2) = "사업장코드"				' Header명(2)		
			arrHeader(3) = "사업장명"				' Header명(3)	
			arrHeader(4) = "내부부서코드"			
			
		Case 8			'어음번호 
	'	 If frm1.txtBankCd1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "어음번호 팝업"						' 팝업 명칭 
			arrParam(1) = "F_NOTE	A"		' TABLE 명칭 
			arrParam(2) = strCode									' Code Condition
			arrParam(3) = ""										' Name Condition
			arrParam(4) = " A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.NOTE_STS = " & FilterVar("OC", "''", "S") & " "	  ' Where Condition
			arrParam(5) = "어음번호"											' 조건필드의 라벨 명칭 

			arrField(0) = "A.NOTE_NO"						' Field명(0)
			arrField(1) = "A.ISSUE_DT"						' Field명(1)			
			arrField(2) = "A.NOTE_AMT"						' Field명(0)			
			arrField(3) = "A.DEPT_CD"						' Field명(0)			
			
			arrHeader(0) = "어음번호"					' Header명(0)
			arrHeader(1) = "발행일"					' Header명(0)
			arrHeader(2) = "어음금액"						' Header명(1)	
 			arrHeader(3) = "발생부서"		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	Select Case iWhere				
		Case 1, 2
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		
	    Case 3
			arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
					"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	    Case 4, 5
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
					 "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	    					 	    
        Case Else 
		     arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			     	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")					 
	End Select			
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
			Select Case iWhere				
				Case 1		'최초사업장 
					.txtFrBizCd.Focus
				Case 2		'이동사업장 
					.txtToBizCd.Focus
				Case 3		' 부서 
					.txtFrDeptCd.focus
				Case 4		' 계좌번호 
					.txtToDeptCd.focus
				Case 8  '어음번호 
					.txtNoteNo.focus 			
			End Select
		End With
		Exit Function
	End If	

	With frm1
		Select Case iWhere
			Case 1		'최초사업장 
				.txtFrBizCd.value = arrRet(0)
				.txtFrBizNm.value = arrRet(1)
				.txtFrDeptCd.focus
			Case 2		'이동사업장 
				.txtToBizCd.value = arrRet(0)
				.txtToBizNm.value = arrRet(1)
				.txtToDeptCd.focus
			Case 3		' 최초부서 
				.txtFrDeptCd.value	= Trim(arrRet(0))
				.txtFrDeptNm.value	= Trim(arrRet(1))
			Case 4		'이동부서 
				.txtToDeptCd.value = Trim(arrRet(0))
				.txtToDeptNm.value = Trim(arrRet(1))

				Call fncToDeptIntoSheet(Trim(arrRet(0)),Trim(arrRet(1)))
			Case 5
       			ggoSpread.Source = frm1.vspdData

				frm1.vspdData.Row = frm1.vspdData.ActiveRow
				frm1.vspdData.Col = C_TO_DEPT_CD

				frm1.vspdData.Text = arrRet(0)

				frm1.vspdData.Col = C_TO_DEPT_NM
				frm1.vspdData.Text = arrRet(1)

				ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			Case 8
				.txtNoteNo.value  = arrRet(0)
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
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode					'부서코드 
	arrParam(1) = frm1.txtGLDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"						'부서권한(lgUsrIntCd)
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If
	
	if iWhere = "1" then
		frm1.txtFrDeptCd.value = arrRet(0)
		frm1.txtFrDeptNm.value = arrRet(1)
		'Call txtFrDeptCD_OnChange()
		frm1.txtFrDeptCd.focus
	else	
		frm1.txtFrDeptCd.value = arrRet(0)
		frm1.txtFrDeptNm.value = arrRet(1)
		'Call txtFrDeptCD_OnChange()
		frm1.txtFrDeptCd.focus
	end if
			
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

    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggospread.ClearSpreadData

    Call InitSpreadSheet("A")                                                        'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                        'Setup the Spread sheet
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
            
            C_PROC_CHK    = iCurColumnPos(1)
            C_FR_DEPT_CD  = iCurColumnPos(2)
            C_FR_DEPT_NM  = iCurColumnPos(3)             
            C_NOTE_NO	  = iCurColumnPos(4)
            C_NOTE_AMT    = iCurColumnPos(5)
            C_NOTE_STS    = iCurColumnPos(6)
            C_TO_DEPT_CD  = iCurColumnPos(7)
            C_TO_DEPT_POP = iCurColumnPos(8)
            C_TO_DEPT_NM  = iCurColumnPos(9)
            C_MOVE_DESC   = iCurColumnPos(10)                          
            C_BP_CD		  = iCurColumnPos(11)
            C_BP_NM	      = iCurColumnPos(12)             
            C_ISSUED_DT   = iCurColumnPos(13)             
            C_DUE_DT      = iCurColumnPos(14)
		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
             
            C_CNCL_CHK			= iCurColumnPos(1)
            C_CNCL_TO_DEPT_CD	= iCurColumnPos(2)
            C_CNCL_TO_DEPT_NM	= iCurColumnPos(3)                                       
            C_CNCL_GL_NO		= iCurColumnPos(4)
            C_CNCL_GL_DT		= iCurColumnPos(5)
			C_CNCL_TEMP_GL_NO	= iCurColumnPos(6)
			C_CNCL_TEMP_GL_DT	= iCurColumnPos(7)
            C_CNCL_NOTE_NO		= iCurColumnPos(8)              
            C_CNCL_NOTE_AMT		= iCurColumnPos(9)              
            C_CNCL_FR_DEPT_CD	= iCurColumnPos(10)
            C_CNCL_FR_DEPT_NM	= iCurColumnPos(11)                                       
			C_CNCL_BP_CD		= iCurColumnPos(12)
            C_CNCL_BP_NM		= iCurColumnPos(13)
            C_CNCL_ISSUED_DT	= iCurColumnPos(14)             
            C_CNCL_DUE_DT		= iCurColumnPos(15)
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
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End if
End Sub

Sub txtTodt_DblClick(Button)
	if Button = 1 then
		frm1.txtTodt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTodt.Focus
	End if
End Sub

Sub txtGLDt_DblClick(Button)
	if Button = 1 then
		frm1.txtGLDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtGLDt.Focus
	End if
End Sub

Sub txtFrGlDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrGlDt.Focus
	End if
End Sub

Sub txtToGlDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToGlDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name :txtDueDt_keypress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtFromDt.focus
		Call MainQuery
	End If   
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then  
		frm1.txtToDt.focus
		Call MainQuery
	End If   
End Sub

Sub txtFrGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToGlDt.focus
	   Call MainQuery
	End If   
End Sub

Sub txtToGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFrGlDt.focus
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

	If Trim(frm1.txtToDeptCd.value) <> "" and Trim(frm1.txtGLDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtToDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtToDeptCd.value = ""
			frm1.txtToDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtToDeptCd.value = ""
					frm1.txtToDeptNm.value = ""
				    frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFrDeptCd_onBlur()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrBizCd_onBlur()
	If Trim(frm1.txtFrBizCd.value) = "" Then
		frm1.txtFrBizNm.value = ""
	End If
End Sub	

Sub txtFrDeptCd_onBlur()
	If Trim(frm1.txtFrDeptCd.value) = "" Then
		frm1.txtFrDeptNm.value = ""
	End If
End Sub	

Sub txtToBizCd_onBlur()
	If Trim(frm1.txtToBizCd.value) = "" Then
		frm1.txtToBizNm.value = ""
	End If
End Sub	

Sub txtToDeptCd_onBlur()	
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	
	If Trim(frm1.txtToDeptCd.value) = "" Then
		frm1.txtToDeptNm.value = ""		
	Else
		strSelect	= " dept_nm"    		
		strFrom		= " b_acct_dept(NOLOCK) "		
		strWhere	= " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtToDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtToDeptCd.value = ""
			frm1.txtToDeptNm.value = ""
			'frm1.hOrgChangeId.value = ""
		Else
			frm1.txtToDeptNm.value = mid(Trim(lgF2By2),2,len(lgF2By2) - 3)
			Call fncToDeptIntoSheet(UCase(Trim(frm1.txtToDeptCd.value)), Trim(frm1.txtToDeptNm.value))
		End If
	End If
End Sub	

Function fncToDeptIntoSheet(ByVal pDeptCd, pDeptNm)
	Dim IRow
	
	If frm1.vspdData.MaxRows < 1 Then
		Exit Function
	End If 

	ggoSpread.Source = frm1.vspdData
	
	For IRow = 1 To frm1.vspdData.MaxRows 
		frm1.vspdData.Row  = IRow
		frm1.vspdData.Col  = C_TO_DEPT_CD
		frm1.vspdData.Text = pDeptCd
		
		frm1.vspdData.Col  = C_TO_DEPT_NM
		frm1.vspdData.Text = pDeptNm
	Next

	lgBlnFlgChgValue = True
End Function

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
End Function

Function ClickTab2()
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetToolBar("1100000000001111")
	Else                 
		Call SetToolBar("1100000000001111")
	End If	

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)														'두번째 Tab 
	
	gSelframeFlg = TAB2
	frm1.hProcFg.value = "DG"
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
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
    End If     
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
       Exit Sub
    End If     
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
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
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
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If 
    
   	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	    
		If (lgStrPrevKeyNoteNo <> "" or  lgStrPrevKeyGlNo <> "" or lgStrPrevKeyTempGlNo <> "" ) Then								
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
	'2004.05.27
	If Col = C_PROC_CHK Then
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
					ggoSpread.UpdateRow Row  ''2004.03.19 comment 처리				
				Else
					ggoSpread.SSDeleteFlag Row,Row
					ggoSpread.SSDeleteFlag Row,Row
				End If			
			End If
		End With
	Elseif	Col = C_TO_DEPT_POP then
		With frm1.vspdData
			.Row = Row
			.Col = C_TO_DEPT_CD
			
			Call OpenPopUp(frm1.vspdData.text, 5)			
		End With
	End if
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
    Else
		Call InitSpreadSheet("B")
    End If 
       
    Call InitVariables() 
    
    frm1.vspdData.MaxRows = 0
	
    '-----------------------
    'Check condition area
    '----------------------- 
    If gSelframeFlg = TAB1 Then 
		If Not chkField(Document, "1") Then									'⊙: This function check indispensable field     						
			Exit Function	
		End If 
	Else
		If (frm1.txtFrGlDt.Text = "") or (frm1.txtToGlDt.Text = "") Then
			Call DisplayMsgBox("17A002", parent.VB_INFORMATION, "X", "X")
			Exit Function
		End if
		''KO      17A002 A        2        %1을 입력하세요.
	End if
	    
    If gSelframeFlg = "1" Then
		If (frm1.txtFromDt.Text <> "") And (frm1.txtTodt.Text <> "") Then
			If CompareDateByFormat(frm1.txtFromDt.Text, frm1.txtTodt.Text, frm1.txtFromDt.Alt, frm1.txtToDt.Alt, _
						"970025", frm1.txtFromDt.UserDefinedFormat, Parent.gComDateType, true) = False Then

				frm1.txtFromDt.focus											
				Exit Function
			End if	
		End If
	End If
	
    If gSelframeFlg = "2" Then
		If (frm1.txtFrGlDt.Text <> "") And (frm1.txtToGlDt.Text <> "") Then
			If CompareDateByFormat(frm1.txtFrGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFrGlDt.Alt, frm1.txtToGlDt.Alt, _
						"970025", frm1.txtFrGlDt.UserDefinedFormat, Parent.gComDateType, true) = False Then
				frm1.txtFrGlDt.focus											
				Exit Function
			End if	
		End If
	End If
	
    If frm1.txtToBizCd.value  = "" Then
		frm1.txtToBizNm.value = ""
	End If
    
    If frm1.txtToDeptCd.value  = "" Then
		frm1.txtToDeptNm.value = ""
	End If
		
    Call ggoOper.LockField(Document, "N")			'⊙: This function lock the suitable field

    '-------------------------
    'Query function call area
    '-------------------------	   
  
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
	dbsave()
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
    
    If gSelframeFlg = TAB1 Then
		Call InitSpreadSheet("A") 
    Else
		Call InitSpreadSheet("B")      
    End If
    
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

	With frm1
		If gSelframeFlg = "1" Then 														'☜: 일괄처리(tab1) 조회 
		    If lgIntFlgMode = Parent.OPMD_UMODE Then		    
				strVal = BIZ_PGM_ID2 & "?txtMode = " & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
					
				strVal = strVal & "&txtFromDt="	   & Trim(frm1.txtFromDt.Text)
				strVal = strVal & "&txtToDt="      & Trim(frm1.txtToDt.Text)
				strVal = strVal & "&txtFrBizCd="   & Trim(frm1.txtFrBizCd.value)
				
				strVal = strVal & "&txtBpCd="      & Trim(frm1.txtBpCd.value)		
				strVal = strVal & "&txtFrDeptCd=" & Trim(frm1.txtFrDeptCd.value)
				strVal = strVal & "&gChangeOrgId=" & Trim(.hOrgChangeId.Value)
				strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)
												
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="   & lgStrPrevKeyGlNo
				strVal = strVal & "&lgPageNo="           & lgPageNo
				strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
			Else			
				strVal = BIZ_PGM_ID2 & "?txtMode = " & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
				
				strVal = strVal & "&txtFromDt="	   & Trim(frm1.txtFromDt.Text)
				strVal = strVal & "&txtToDt="      & Trim(frm1.txtToDt.Text)
				strVal = strVal & "&txtFrBizCd="   & Trim(frm1.txtFrBizCd.value)
				
				strVal = strVal & "&txtBpCd="      & Trim(frm1.txtBpCd.value)		
				strVal = strVal & "&txtFrDeptCd=" & Trim(frm1.txtFrDeptCd.value)
				strVal = strVal & "&gChangeOrgId=" & Trim(.hOrgChangeId.Value)
				strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)
				
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="   & lgStrPrevKeyGlNo				
				strVal = strVal & "&lgPageNo="           & lgPageNo
				strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
			End If   						
		Else 																			'☜: 일괄취소(tab2) 조회																				
		    If lgIntFlgMode = Parent.OPMD_UMODE Then
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 				
				
				strVal = strVal & "&txtFrGlDt="	  & Trim(frm1.txtFrGlDt.Text)
				strVal = strVal & "&txtToGlDt="   & Trim(frm1.txtToGlDt.Text)				
				strVal = strVal & "&lgStrPrevKeyNoteNo="	& lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="		& lgStrPrevKeyGlNo
				strVal = strVal & "&lgStrPrevKeyTempGlNo="	& lgStrPrevKeyTempGlNo				
				strVal = strVal & "&lgPageNo="				& lgPageNo
				strVal = strVal & "&txtMaxRows="			& .vspdData2.MaxRows
			Else
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 				

				strVal = strVal & "&txtFrGlDt=" & Trim(frm1.txtFrGlDt.Text)
				strVal = strVal & "&txtToGlDt=" & Trim(frm1.txtToGlDt.Text)
				strVal = strVal & "&lgStrPrevKeyNoteNo="	& lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="		& lgStrPrevKeyGlNo
				strVal = strVal & "&lgStrPrevKeyTempGlNo="	& lgStrPrevKeyTempGlNo				
				strVal = strVal & "&lgPageNo="				& lgPageNo
				strVal = strVal & "&txtMaxRows="			& .vspdData2.MaxRows
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
    
	lgBlnFlgChgValue = False
	
	' 현재 Page의 From Element들을 사용자가 입력을 받지 못하게 하거나 필수입력사항을 표시한다.
	' LockField(pDoc, pACode)
	
'   Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
    frm1.txtGLDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)	
    Call SetToolBar("1100100000001111")
End Function

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
	
	'If frm1.hProcFg.value = "CG" Then
	If gSelframeFlg = TAB1 then  ''''이동처리 Tab이면  
		If Not chkField(Document, "2") Then                                   '⊙: Check contents area
			Exit Function
		End If
	End If
    
	IF Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
	
	If gSelframeFlg = TAB1 then  ''''이동처리 Tab이면 
		If UCase(Trim(frm1.txtFrBizCd.value)) = UCase(Trim(frm1.txtToBizCd.value)) then
			IntRetCD = DisplayMsgBox("141445","x","x","x")
			Exit Function
		End If	
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
		'If .hProcFg.value = "CG" Then										'☜:일괄처리 
		If gSelframeFlg = TAB1 then  ''''이동처리 Tab이면 
			For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				.vspdData.Col = C_PROC_CHK
				
				If .vspdData.Text = "1" Then				
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData.Col = C_NOTE_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 어음번호 
					.vspdData.Col = C_TO_DEPT_CD
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' 이동부서코드					
					.vspdData.Col = C_MOVE_DESC
					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep		' 비고 
					lGrpCnt = lGrpCnt + 1
				End If
			Next

			.hProcFg.value = "CG"	
		ElseIf gSelframeFlg = TAB2 then  ''''이동취소  Tab이면  Then									 '☜:일괄취소 
			For lRow = 1 To .vspdData2.MaxRows
				.vspdData2.Row = lRow
				.vspdData2.Col = C_CNCL_CHK
				
				If .vspdData2.Text = "1" Then
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData2.Col = C_CNCL_NOTE_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' 어음번호				
					.vspdData2.Col = C_CNCL_TEMP_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' 결의전표번호 
					.vspdData2.Col = C_CNCL_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gRowSep		' 회계전표번호				
					
					lGrpCnt = lGrpCnt + 1
				End If
			Next	
			.hProcFg.value = "DG"
		End If
			
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		If .txtMaxRows.value <= 0 Then
			Call DisplayMsgBox("900025","x","x","x")	'선택된 항목이 없습니다.
			Exit Function
		End If

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
	
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)				'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True										'⊙: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function DbSaveOk()										'☆: 저장 성공후 실행 로직 
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

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode						'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""							'FrDt
	arrParam(3) = ""							'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		'Call SetPopUp(arrRet, iWhere)
		frm1.txtBpCd.value  = arrRet(0)
		frm1.txtBpNm.value  = arrRet(1)
	End If	
End Function

Sub txtFrDeptCd_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If frm1.txtFrDeptCd.value = "" Then
		frm1.txtFrDeptNm.value = ""
	End If
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtFrDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect = "dept_cd, ORG_CHANGE_ID"
		strFrom =  " B_ACCT_DEPT "
		strWhere = " ORG_CHANGE_DT >= "
		strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
		strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
		strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtTodt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
		strWhere = strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtFrDeptCd.value)), "''", "S")
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtFrDeptCd.value = ""
			frm1.txtFrDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtFrDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
		End If
	End If
End Sub

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
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>받을어음이동취소</font></td><td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A> &nbsp;|<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
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
									<TD CLASS=TD5 NOWRAP>발행일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작발행일"></OBJECT>');</SCRIPT>&nbsp; ~ &nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtTodt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료발행일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10  tag="1XX" ALT="거래처코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value)">
									                     <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNm" NAME="txtBpNm" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14X" ALT="거래처명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>현재 사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFrBizCd" ALT="현재 사업장코드" Size= "12" MAXLENGTH="10" tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtFrBizCd.value, 1)">
														 <INPUT NAME="txtFrBizNm" ALT="현재 사업장명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
									<TD CLASS=TD5 NOWRAP>현재 부서</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFrDeptCd" ALT="현재 부서코드" Size= "10" MAXLENGTH="10" tag="11XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtFrDeptCd.value,3)">
														 <INPUT NAME="txtFrDeptNm" ALT="현재 부서명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
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
								<TD CLASS=TD5 NOWRAP>이동일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="이동일" tag="22X1" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>					
							<TR>
								<TD CLASS=TD5 NOWRAP>이동사업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtToBizCd" ALT="이동사업장코드" Size= "12" MAXLENGTH="10" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtToBizCd.value, 2)">
													 <INPUT NAME="txtToBizNm" ALT="이동사업장명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
								<TD CLASS=TD5 NOWRAP>이동부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtToDeptCd" ALT="이동부서코드" Size= "12" MAXLENGTH="10" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtToDeptCd.value, 4)">
													 <INPUT NAME="txtToDeptNm" ALT="이동사업장명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
							</TR>									
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtNoteDesc" ALT="비고" SIZE = "90" STYLE="TEXT-ALIGN: left" tag="21X"></TD></TD>
							</TR>	
							<TR>
								<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="33" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
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
									<TD CLASS=TD5 NOWRAP>회계일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtFrGlDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작회계일" tag="12X1"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtToGlDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료회계일" tag="12X1" ></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>														 
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
   					<TD><BUTTON NAME="btncancel" CLASS="CLSSBTN" ONCLICK="vbscript:Rowselect()">전체선택</BUTTON>&nbsp;
						<BUTTON NAME="btnselect" CLASS="CLSSBTN" ONCLICK="vbscript:Rowcancel()">전체취소</BUTTON>
					</TD>
   				</TR>
   			</TABLE> 
  		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"		tag="1" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hToBizAreaCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hProcFg"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg1"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg2"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteSts"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtStart"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtEnd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtStart"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtEnd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtGlNo"			tag="2" TABINDEX="-1">
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
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 

src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
