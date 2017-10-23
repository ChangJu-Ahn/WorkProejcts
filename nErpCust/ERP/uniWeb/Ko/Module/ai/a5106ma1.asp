<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!--'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : 기초계정등록 
'*  3. Program ID        : a5104ma
'*  4. Program 이름      : 기초계정등록 
'*  5. Program 설명      : 회계전표 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : a5104ma
'*  7. 최초 작성년월일   : 2003/01/02
'*  8. 최종 수정년월일   : 2003/10/24
'*  9. 최초 작성자       : 김호영 
'* 10. 최종 작성자       : Jeong Yong Kyun
'* 11. 전체 comment      :
'*
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../AG/Acctctrl.vbs"></SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID      = "a5106mb1.asp"			'☆: 비지니스 로직 ASP명 
Const JUMP_PGM_ID_TAX_REP = "a6114ma1"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns

'@Grid_Column       
Dim  C_ItemSeq		
Dim  C_deptcd		
Dim  C_deptPopup	
Dim  C_deptnm	   	
Dim  C_AcctCd		
Dim  C_AcctPopup	
Dim  C_AcctNm		
Dim  C_DrCrFg		
Dim  C_DrCrNm		
Dim  C_DocCur		
Dim  C_DocCurPopup	
Dim  C_ExchRate	
Dim  C_ItemAmt		
Dim  C_ItemLocAmt	
Dim  C_IsLAmtChange
Dim  C_ItemDesc	
Dim  C_VatType		
Dim  C_VatNm		
Dim  C_AcctCd2		

Const C_SHEETMAXROWS = 30 ' : 한 화면에 보여지는 최대갯수*1.5<BR>
Const C_GLINPUTTYPE = "TR"

Const MENU_NEW	=	"1110010000011111"
Const MENU_CRT	=	"1110111100111111"
Const MENU_UPD	=	"1111111100111111"
Const MENU_PRT	=	"1110000000011111"	

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgCurrRow
Dim lgStrPrevKeyDtl
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgFormLoad
Dim lgQueryOk
Dim lgstartfnc
Dim intItemCnt
Dim lgBlnExecDelete
'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim lgArrAcctForVat
Dim lgBlnGetAcctForVat

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
Sub initSpreadPosVariables()
	C_ItemSeq		= 1 
	C_deptcd		= 2 
	C_deptPopup		= 3 
	C_deptnm		= 4	
	C_AcctCd		= 5 
	C_AcctPopup		= 6 
	C_AcctNm		= 7 
	C_DrCrFg		= 8 
	C_DrCrNm		= 9 
	C_DocCur		= 10
	C_DocCurPopup	= 11
	C_ExchRate		= 12
	C_ItemAmt		= 13
	C_ItemLocAmt	= 14
	C_IsLAmtChange	= 15
	C_ItemDesc		= 16
	C_VatType		= 17
	C_VatNm			= 18
	C_AcctCd2		= 19
End Sub

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE					'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
    lgStrPrevKey = ""									'initializes Previous Key
    lgLngCurRows = 0									'initializes Deleted Rows Count
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
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear

	With frm1
		.txtGLDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
		.txtCommandMode.Value = "CREATE"
		.cboGlInputType.Value = C_GLINPUTTYPE
		.cboGlType.Value      = "04"
		.txtDeptCd.Value	  = parent.gDepart
		.vspdData3.MaxRows    = 0
		.vspdData3.MaxCols    = 16    
		.hOrgChangeId.Value   = parent.gChangeOrgId 
		'현금계정을 가지고온다.
		Call GetCheckAcct	
		    
		.txtGLNo.focus
		lgBlnFlgChgValue = False
	End With		
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================== 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	With frm1.vspdData
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20030218",,parent.gAllowDragDropSpread    
	
		.MaxCols = C_AcctCd2 + 1
		.Col = .MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
		.MaxRows = 0
		.ReDraw = False

		Call AppendNumberPlace("6","3","0")
        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetFloat  C_ItemSeq,    " ", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd,     "부서코드",   10, , , 10, 2
        ggoSpread.SSSetButton C_deptpopup
        ggoSpread.SSSetEdit   C_deptnm,     "부서명",     17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd,     "계정코드",   15, , , 18
		ggoSpread.SSSetButton C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,     "계정코드명", 20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg,     "", 8
	    ggoSpread.SSSetCombo  C_DrCrNm,     "차대구분",   11
		ggoSpread.SSSetEdit   C_DocCur,     "거래통화",   10, , , 10, 2
        ggoSpread.SSSetButton C_DocCurPopup
		ggoSpread.SSSetFloat  C_ExchRate,   "환율", 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ItemAmt,    "금액",       15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)", 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_IsLAmtChange,   "",     30, , , 128
		ggoSpread.SSSetEdit   C_ItemDesc,   "비  고",     30, , , 128
		ggoSpread.SSSetCombo  C_VATTYPE,     "", 8
	    ggoSpread.SSSetCombo  C_VATNM,     "계산서유형",   20	    		
		ggoSpread.SSSetEdit   C_AcctCd2,   "",     30, , , 128

		Call ggoSpread.MakePairsColumn(C_deptcd,C_deptpopup)
		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)
		Call ggoSpread.MakePairsColumn(C_DrCrFg,C_DrCrNm,"1")
		Call ggoSpread.MakePairsColumn(C_VATTYPE,C_VATNM,"1")

		Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)							'공통콘트롤 사용 Hidden Column
		Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)
		Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
		Call ggoSpread.SSSetColHidden(C_VatNm,C_VatNm,True)
		Call ggoSpread.SSSetColHidden(C_IsLAmtChange,C_IsLAmtChange,True)
		Call ggoSpread.SSSetColHidden(C_AcctCd2,C_AcctCd2,True)

		.ReDraw = True
	End With

    SetSpreadLock "I", 0, 1, ""
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False

		Select Case Index
			Case 0			
				ggoSpread.SpreadUnLock		C_deptcd		, -1    , C_deptcd
				ggoSpread.SSSetRequired		C_deptcd		, -1    , C_deptcd
				ggoSpread.SpreadUnLock		C_deptpopup		, -1    , C_deptpopup
				ggoSpread.SpreadLock		C_deptnm		, -1    , C_deptnm
				ggoSpread.SpreadLock		C_AcctCd		, -1    , C_AcctCd
				ggoSpread.SpreadLock		C_AcctPopup		, -1    , C_AcctPopup
				ggoSpread.SpreadLock		C_AcctNm		, -1    , C_AcctNm
				ggoSpread.SpreadUnLock		C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SSSetRequired		C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SpreadUnLock		C_DocCur		, -1    , C_DocCur
				ggoSpread.SSSetRequired		C_DocCur		, -1    , C_DocCur
				ggoSpread.SpreadUnLock		C_DocCurPopup	, -1    , C_DocCurPopup
				ggoSpread.SpreadUnLock		C_ExchRate		, -1    , C_ExchRate
				ggoSpread.SpreadUnLock		C_ItemAmt		, -1    , C_ItemAmt
				ggoSpread.SSSetRequired		C_ItemAmt		, -1    , C_ItemAmt
				ggoSpread.SpreadUnLock		C_ItemLocAmt	, -1    , C_ItemLocAmt
				ggoSpread.SpreadUnLock		C_ItemDesc		, -1    , C_ItemDesc
				ggoSpread.SpreadUnLock		C_VATNM			, -1    , C_VATNM
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			Case 1
				ggoSpread.SpreadLock C_deptcd		, -1    , C_deptcd
				ggoSpread.SpreadLock C_ItemSeq		, -1	, C_ItemSeq 'Item Grid 전체 Lock설정 
				ggoSpread.SpreadLock C_deptpopup	, -1	, C_deptpopup    ', lRow2
				ggoSpread.SpreadLock C_ItemLocAmt	, -1	, C_ItemLocAmt ', lRow2
				ggoSpread.SpreadLock C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SpreadLock C_ItemDesc		, -1	, C_ItemDesc ', lRow2
				ggoSpread.SpreadLock C_AcctPopup	, -1	, C_AcctPopup ', lRow2
				ggoSpread.SpreadLock C_DrCrNm		, -1	, C_DrCrNm    ', lRow2
				ggoSpread.SpreadLock C_DocCur		, -1	, C_DocCur    ', lRow2
				ggoSpread.SpreadLock C_DocCurPopup	, -1	, C_DocCurPopup    ', lRow2
				ggoSpread.SpreadLock C_ExchRate		, -1	, C_ExchRate    ', lRow2				
				ggoSpread.SpreadLock C_ItemAmt		, -1	, C_ItemAmt    ', lRow2
				ggoSpread.SpreadLock C_VATNM		, -1	, C_VATNM    ', lRow2               
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		End Select
    
		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

Sub SetSpread2Lock(Byval stsFg,Byval Index,ByVal lRow  ,ByVal lRow2 )
    With frm1
		ggoSpread.Source = .vspdData2
		If lRow = "" Then
			lRow = 1
		End If

		If lRow2 = "" Then
			lRow2 = .vspdData2.MaxRows
		End If

		.vspdData2.Redraw = False

		Select Case Index
			Case 0			
			Case 1
				ggoSpread.SpreadLock 1, lRow, .vspdData2.MaxCols, lRow2	
		End Select

		.vspdData2.Redraw = True
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(Byval stsFg, Byval Index, ByVal lRow, ByVal lRow2)
    With frm1
		If  lRow2 = "" Then	lRow2 = lRow
		
		.vspdData.ReDraw = False
		' SSSetRequired(ByVal Col, ByVal Row, Optional Row2)
		ggoSpread.SSSetProtected C_ItemSeq, lRow, lRow2			
		ggoSpread.SSSetProtected C_deptNm,    lRow, lRow2
		ggoSpread.SSSetProtected C_AcctNm, lRow, lRow2			' 계정코드명		
		ggoSpread.SSSetRequired  C_deptcd,    lRow, lRow2		' 부서코드 
						
		Select Case stsFg
			Case "I"							
				ggoSpread.SSSetRequired C_AcctCd, lRow, lRow2	' 계정코드 
			Case "Q"						
				ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2	' 계정코드				
		End Select	
		
		ggoSpread.SSSetRequired C_DrCrNm, lRow, lRow2			' 차대구분 
		ggoSpread.SSSetRequired C_DocCur, lRow, lRow2			' 부서코드 
		ggoSpread.SSSetRequired C_ItemAmt, lRow, lRow2			' 금액 

		.vspdData.ReDraw = True
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Err.clear
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlType ,lgF0  ,lgF1  ,Chr(11))
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
End Sub

Function InitComboBoxGrid()
    ggoSpread.Source = frm1.vspdData
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm
    
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("B9001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_VatType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_VatNm
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
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'권한관리 추가   							  
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
		
		Case 1
			If frm1.txtDeptCd.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
			
			arrStrRet =  AutorityMakeSql("DEPT",frm1.hORGCHANGEID.Value, "","","","")	' 권한관리 추가   							  
			
			arrParam(0) = "부서 팝업"												' 팝업 명칭 
			arrParam(1) = arrstrRet(0)													' 권한관리 추가   							  
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = arrstrRet(1)													' 권한관리 추가   							  
			arrParam(5) = "부서코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "DEPT_CD"	     												' Field명(0)
			arrField(1) = "DEPT_NM"			    										' Field명(1)
    
			arrHeader(0) = "부서코드"												' Header명(0)
			arrHeader(1) = "부서명"													' Header명(1)
		Case 2
			arrParam(0) = "통화코드 팝업"											' 팝업 명칭 
			arrParam(1) = "B_Currency"	    											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = ""															' Where Condition
			arrParam(5) = "통화코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "Currency"	    											' Field명(0)
			arrField(1) = "Currency_desc"	    										' Field명(1)
    
			arrHeader(0) = "통화코드"												' Header명(0)
			arrHeader(1) = "통화코드명"												' Header명(1)
		Case 3
			arrParam(0) = "계정코드팝업"											' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "		' Where Condition
			arrParam(5) = "계정코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"												' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"												' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"												' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"												' Field명(3)
			
			arrHeader(0) = "계정코드"												' Header명(0)
			arrHeader(1) = "계정코드명"												' Header명(1)
			arrHeader(2) = "그룹코드"												' Header명(2)
			arrHeader(3) = "그룹명"													' Header명(3)
		Case 4
			arrStrRet =  AutorityMakeSql("DEPT_ITEM",frm1.hORGCHANGEID.Value, frm1.txtDeptCd.Value,"","","")'권한관리 추가 
			
			arrParam(0) = "부서 팝업"												' 팝업 명칭 
			arrParam(1) = arrstrRet(0)													' 권한관리 추가 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = arrstrRet(1)													' 권한관리 추가									   
																						' Where Condition
			arrParam(5) = "부서코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "A.DEPT_CD"	     											' Field명(0)
			arrField(1) = "A.DEPT_NM"			    									' Field명(1)
    
			arrHeader(0) = "부서코드"												' Header명(0)
			arrHeader(1) = "부서명"													' Header명(1)
	End Select

	IsOpenPop = True    
   	If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetPopUp()  -------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtGlNo.Value = UCase(Trim(arrRet(0)))
			Case 1
				.txtDeptCd.Value = UCase(Trim(arrRet(0)))
				.txtDeptNm.Value = arrRet(1)
				Call txtDeptCd_OnChange()
			Case 2
				.vspdData.Row = .vspdData.ActiveRow 
				
				ggoSpread.Source = .vspdData
				ggoSpread.UpdateRow .vspdData.ActiveRow 
				.vspdData.Col  = C_ItemLocAmt
				.vspdData.Text = ""
				.vspdData.Col  = C_DocCur 
				.vspdData.Text = UCase(Trim(arrRet(0)))

				If Trim(.vspdData.Text) = parent.gCurrency Then
					.vspdData.Col  = C_ExchRate
					.vspdData.Text = 1
				Else
					Call FindExchRate(UniConvDateToYYYYMMDD(.txtGLDt.text,parent.gDateFormat,""), UCase(Trim(arrRet(0))),.vspdData.ActiveRow)
				End If

				Call DocCur_OnChange(.vspdData.ActiveRow,.vspdData.ActiveRow)
			Case 3
				.vspdData.Row  = .vspdData.ActiveRow 
				.vspdData.Col  = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_AcctNm
				.vspdData.Text = arrRet(1)
                Call vspdData_Change(C_AcctCd, .vspddata.activerow)
			Case 4
				.vspdData.Row = .vspdData.ActiveRow 
				
				ggoSpread.Source = .vspdData
				ggoSpread.UpdateRow .vspdData.ActiveRow 
				
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)				
		End Select
	End With	
End Function


Function OpenRefGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(4)	                           '권한관리 추가 (3 -> 4)
	
	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5106ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5106ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(4)	= lgAuthorityFlag 
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		Exit Function
	Else		
		Call SetRefGL(arrRet)
	End If
End Function

Function SetRefGL(Byval arrRet)
	frm1.txtGlNo.Value = UCase(Trim(arrRet(0)))
	frm1.txtGLNo.focus 
End Function

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a5121ma1"
	
	With frm1
		VarDateFr = UniConvDateToYYYYMMDD(.txtGlDt.Text, parent.gDateFormat,"")	
		VarDateTo = UniConvDateToYYYYMMDD(.txtGlDt.Text, parent.gDateFormat,"")	
		' 회계전표의 key는 temp_GL_NO이기 때문에 temp_GL_NO만 넘긴다.	
		VarDeptCd      = "%"
		VarBizAreaCd   = "%"
		varGlNoFr      = Trim(.txtGlNo.Value)
		varGlNoTo	   = Trim(.txtGlNo.Value)
		varOrgChangeId = Trim(.hOrgChangeId.Value)	
	End With
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName
	
    If Not chkField(Document, "1") Then										'⊙: This function check indispensable field
		Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId)
	
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	On Error Resume Next                                                    '☜: Protect system from crashing

    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName

    If Not chkField(Document, "1") Then										'⊙: This function check indispensable field
		Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId)

    StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function

'========================================================================================
' Function Name : FncBtnCalc
' Function Desc : This function calculate local amt from amt of multi
'========================================================================================
Function FncBtnCalc() 
	Dim ii
	Dim tempAmt, tempLocAmt, tempExch, TempSep, tempDoc
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strDate
	Dim strExchFg
	Dim IntRetCD
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6		

	With frm1
		strSelect	= "b.minor_cd"
		strFrom		= "b_company a, b_minor b"
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If	

		strDate = UniConvDateToYYYYMMDD(.txtGLDt.text,parent.gDateFormat,"")
		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DocCur			
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_ItemAmt
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_ExchRate
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt = tempAmt * TempExch
					Else
						tempLocAmt = tempAmt / TempExch
					End If
					
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				End If
			Next		
		End If
	End With

	Call SetSumItem	
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : ExchRateCheck
' Function Desc : 
'========================================================================================
Function ExchRateCheck()
	Call FncBtnCalc()
End Function 

'***++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'========================================================================================
' Function Name : gfRealRound
' Function Desc : Arithmetic Rounding Function
'========================================================================================
Function gfRealRound(ByVal x, ByVal Factor )
    Dim lcSwitch, iCurResult

    If x < 0 Then lcSwitch = -1 Else lcSwitch = 1
    x = x * lcSwitch
    iCurResult = Int(x * 10 ^ Factor + 0.5) / 10 ^ Factor
    gfRealRound = iCurResult * lcSwitch
End Function

'==========================================  2.4.3 OpenDept()  =============================================
'	Name : OpenDept()
'	Description : 
'========================================================================================================= 

'***++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(3)
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtDeptCd.readOnly = true then
		IsOpenPop = False
		Exit Function
	End If

	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode									'  Code Condition
   	arrParam(1) = frm1.txtGLDt.Text
	arrParam(2) = lgUsrIntCd								' 자료권한 Condition  

	If lgIntFlgMode = parent.OPMD_UMODE then
		arrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		arrParam(3) = "F"									' 결의일자 상태 Condition  
	End If
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else		
		Call SetDept(arrRet, iWhere)
	End If	
End Function

Function OpenUnderDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim field_fg   	

	If RTrim(LTrim(frm1.txtDeptCd.Value)) <> "" 	Then
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"				
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.Value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD = ( SELECT B.BIZ_AREA_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.Value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.Value , "''", "S") & ")"
		arrParam(5) = "부서코드"			
		arrField(0) = "A.DEPT_CD"	
		arrField(1) = "A.DEPT_Nm"
		arrField(2) = "B.BIZ_AREA_CD"
		arrHeader(0) = "부서코드"		
		arrHeader(1) = "부서코드명"
		arrHeader(2) = "사업장코드"				
	Else
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A"				
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		arrParam(5) = "부서코드"			
		arrField(0) = "A.DEPT_CD"	
		arrField(1) = "A.DEPT_Nm"
		arrHeader(0) = "부서코드"		
		arrHeader(1) = "부서코드명"
	End If

	IsOpenPop = True
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
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
               .txtDeptCd.Value = arrRet(0)
               .txtDeptNm.Value = arrRet(1)
               .txtInternalCd.Value = arrRet(2)
  				If lgQueryOk <> True Then
					.txtGLDt.text = arrRet(3)
				End If           
				Call txtDeptCd_OnChange()  
            Case "1"  
				.vspdData.Row = .vspdData.ActiveRow 
				
				ggoSpread.Source = .vspdData
				ggoSpread.UpdateRow .vspdData.ActiveRow 
				
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)
				
				Call deptCd_underChange(arrRet(0))
            Case Else
        End Select
	End With
End Function       

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'=======================================================================================================
'   Function Name : CheckVATType
'   Function Desc : 
'=======================================================================================================
Function CheckVATType(ByVal strAcctCd)
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim arrTemp
	Dim strAcctType

	CheckVATType = False

	strSelect	= "acct_type"
	strFrom		= "a_acct"
	strWhere	= "acct_cd =  " & FilterVar(strAcctCd , "''", "S") & ""

	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrTemp		= Split(lgF0, chr(11))
		strAcctType =  arrTemp(0)
	End If

	If UCase(Trim(strAcctType)) = "VP" Or UCase(Trim(strAcctType)) = "VR" Then
		CheckVATType = True
	End If
End Function

'=======================================================================================================
'   Function Name : FindNumber
'   Function Desc : 
'=======================================================================================================
Function FindNumber(ByVal objSpread, ByVal intCol)
	Dim lngRows
	Dim lngPrevNum
	Dim lngNextNum

    FindNumber = 0

    lngPrevNum = 0
    lngNextNum = 0

    If objSpread.MaxRows = 0 Then
        Exit Function
    End If
        
    For lngRows = 1 To objSpread.MaxRows
        objSpread.Row = lngRows
        objSpread.Col = intCol
        lngNextNum = Clng(objSpread.Text)
            
        If lngNextNum > lngPrevNum Then
            lngPrevNum = lngNextNum
        End If
    Next
    
    FindNumber = lngPrevNum
End Function

'======================================================================================================
' Function Name : SetSpreadFG
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Sub SetSpreadFG(ByVal pobjSpread , ByVal pMaxRows )
    Dim lngRows 

    For lngRows = 1 To pMaxRows
        pobjSpread.Col  = 0
        pobjSpread.Row  = lngRows
        pobjSpread.Text = ""
    Next
End Sub

'======================================================================================================
' Function Name : SetSumItem
' Function Desc :
'=======================================================================================================
Function SetSumItem()
    Dim DblTotDrAmt 
    Dim DblTotLocDrAmt 
    Dim DblTotCrAmt 
    Dim DblTotLocCrAmt 
    Dim lngRows 

	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData 
		If .MaxRows > 0 Then
		    For lngRows = 1 To .MaxRows
		        .Row = lngRows
	            .Col = 0
	            If .text <> ggoSpread.DeleteFlag Then
		 			.col = C_DrCrFg
				    
		 			If .text = "DR" Then
		 			    .Col = C_ItemAmt	'6
		 			    If .Text = "" Then
		 			        DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
		 			    Else
		 			        DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
		 			    End If

		 			    .Col = C_ItemLocAmt	'7
		 			    If .Text = "" Then
		 			        DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
		 			    Else
		 			        DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
		 			    End If
		 			Elseif .text = "CR" then
		 			    .Col = C_ItemAmt	'6
		 			    If .Text = "" Then
		 			        DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
		 			    Else
		 			        DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
		 			    End If

		 			    .Col = C_ItemLocAmt	'7
		 			    If .Text = "" Then
		 			        DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
		 			    Else
		 			        DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
		 			    End If
		 			End If
		 		End If
		    Next
	    End If

'       IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(UCase(Trim(frm1.txtDocCur.Value)),"''","S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
'			frm1.txtDrAmt.Text	= UNIConvNumPCToCompanyByCurrency(DblTotDrAmt,frm1.txtDocCur.Value,parent.ggAmtOfMoneyNo, "X", "X")
'			frm1.txtCrAmt.Text	= UNIConvNumPCToCompanyByCurrency(DblTotCrAmt,frm1.txtDocCur.Value,parent.ggAmtOfMoneyNo, "X", "X")
'		END IF
		With frm1	
	        .txtDrLocAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	        .txtCrLocAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")

	        If .cboGlType.value = "01" Then
				.txtDrLocAmt.text = .txtCrLocAmt.text
			ElseIf .cboGlType.value = "02" Then
				.txtCrLocAmt.text = .txtDrLocAmt.text
			End If
		End With	
	End With
End Function

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)
'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp
	Dim strNmwhere
	Dim arrVal

	With frm1	
		Select Case Kubun		
			Case "FORM_LOAD"	
				strTemp = ReadCookie("GL_NO")

				Call WriteCookie("GL_NO", "")
				
				If strTemp = "" then Exit Function
							
				.txtGlNo.Value = strTemp
						
				If Err.number <> 0 Then
					Err.Clear
					Call WriteCookie("GL_NO", "")
					Exit Function 
				End If
						
				Call FncQuery()
			Case JUMP_PGM_ID_TAX_REP
				ggoSpread.Source = .vspdData

				If .vspddata.MaxRows < 1  Then
					Exit Function
				End If

				.vspddata.row = .vspddata.ActiveRow	
				.vspddata.Col = C_VatType

				If .vspddata.Value = "" Then
					Exit Function
				End If

				.vspddata.Col = C_ItemSeq

				strNmwhere = " GL_NO  = " & FilterVar(.txtGlNo.Value, "''", "S")
				strNmwhere = strNmwhere & " AND ITEM_SEQ = " & .vspddata.text & " "

				If CommonQueryRs( "VAT_NO" , "A_VAT" ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
					arrVal = Split(lgF0, Chr(11))
					strTemp = arrVal(0)
				End If
				
				Call WriteCookie("VAT_NO", strTemp)	
			Case "GL_POPUP"				
				Call WriteCookie("PGMID", "A5104MA1")
			Case Else
				Exit Function
		End Select
	End With		
End Function

'========================================================================================================
'	Desc : 화면이동 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'------------------------ 	   
	With frm1
		ggoSpread.Source = .vspdData
		If (lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True ) And C_GLINPUTTYPE = .cboGlInputType.Value Then
			IntRetCD = DisplayMsgBox("990027", "X", "X", "X")                          'No data changed!!
		    Exit Function
		End If

		Select Case strPgmId
			Case JUMP_PGM_ID_TAX_REP
				ggoSpread.Source = .vspdData

				If .vspddata.MaxRows < 1  Then
					IntRetCD = DisplayMsgBox("900002", "X","X","X")
					Exit Function
				End If

				.vspddata.row = .vspddata.ActiveRow
				.vspddata.Col = C_VatType

				If .vspddata.Value	=	"" Then
					IntRetCD = DisplayMsgBox("205600", "X","X","X")
					Exit Function
				End If
		End Select
	End With

	Call CookiePage(strPgmId)
	Call PgmJump(strPgmId)
End Function

'========================================================================================================
'	Desc : 입출금 화면에 따른 Grid의 Protect변환 
'========================================================================================================
Sub CboGLType_ProtectGrid(Byval GlType)
	With frm1
		ggoSpread.Source = .vspdData
		Select Case GlType		
			Case "01"			
				ggoSpread.SSSetProtected C_DocCur     , 1, .vspddata.maxrows					' 거래통화 
				ggoSpread.SSSetProtected C_DocCurPopup, 1, .vspddata.maxrows					' 거래통화팝업 
				ggoSpread.SSSetProtected C_DrCrfg     , 1, .vspddata.maxrows					' 차대구분 
				ggoSpread.SSSetProtected C_DrCrNm     , 1, .vspddata.maxrows					' 차대변 
			Case "02"			
				ggoSpread.SSSetProtected C_DocCur     , 1, .vspddata.maxrows					' 거래통화 
				ggoSpread.SSSetProtected C_DocCurPopup, 1, .vspddata.maxrows					' 거래통화팝업 
				ggoSpread.SSSetProtected C_DrCrfg     , 1, .vspddata.maxrows					' 차대구분 
				ggoSpread.SSSetProtected C_DrCrNm     , 1, .vspddata.maxrows					' 차대변 
			Case "03"			
				ggoSpread.SSSetRequired  C_DocCur     , 1, .vspddata.maxrows					' 차대구분 
				ggoSpread.SpreadUnLock   C_DocCurPopup, 1, C_DocCurPopup , .vspddata.maxrows	' 거래통화팝업 
				ggoSpread.SpreadUnLock   C_DrCrfg     , 1, C_DrCrNm, .vspddata.maxrows
				ggoSpread.SSSetRequired  C_DrCrfg     , 1, .vspddata.maxrows					' 차대구분 
				ggoSpread.SSSetRequired  C_DrCrNm     , 1, .vspddata.maxrows					' 차대변 
		End Select 				
	End With
End Sub

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
			C_ItemSeq	   = iCurColumnPos(1)
			C_deptcd	   = iCurColumnPos(2)
			C_deptPopup	   = iCurColumnPos(3)
			C_deptnm	   = iCurColumnPos(4)
			C_AcctCd	   = iCurColumnPos(5)
			C_AcctPopup	   = iCurColumnPos(6)
			C_AcctNm	   = iCurColumnPos(7)
			C_DrCrFg	   = iCurColumnPos(8)
			C_DrCrNm	   = iCurColumnPos(9)
			C_DocCur	   = iCurColumnPos(10)
			C_DocCurPopup  = iCurColumnPos(11)
			C_ExchRate	   = iCurColumnPos(12)
			C_ItemAmt	   = iCurColumnPos(13)
			C_ItemLocAmt   = iCurColumnPos(14)
			C_IsLAmtChange = iCurColumnPos(15)
			C_ItemDesc	   = iCurColumnPos(16)
			C_VatType	   = iCurColumnPos(17)
			C_VatNm		   = iCurColumnPos(18)
			C_AcctCd2	   = iCurColumnPos(19)
    End Select    
End Sub

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################

'*****************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitCtrlSpread()
    Call InitCtrlHSpread()
    Call InitComboBox
    Call InitComboBoxGrid           
    Call SetAuthorityFlag													'권한관리 추가    
    Call SetToolbar(MENU_NEW)												'⊙: 버튼 툴바 제어    
    Call SetDefaultVal
	Call InitVariables                                                      '⊙: Initializes local global variables

	Call CookiePage("FORM_LOAD")	
	Call GetAcctForVat  
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
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData_onfocus()
	lgCurrRow = frm1.vspdData.ActiveRow
	If lgIntFlgMode <> parent.OPMD_UMODE Then    
		Call SetToolbar(MENU_CRT)                                     '버튼 툴바 제어 
		
        If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetToolbar(MENU_PRT) 
		Else
			Call SetToolbar(MENU_CRT)
		End if
    Else        
        If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetToolbar(MENU_PRT) 
		Else      
			Call SetToolbar(MENU_UPD)                                     '버튼 툴바 제어 
		End If
    End If  
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtGLDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGLDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim i
    Dim tmpDrCrFG
    
    Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 

	With frm1
		Set gActiveSpdSheet = .vspdData
    
		If .vspdData.MaxRows <= 0 Then                                                    'If there is no data.
			Exit Sub
   		End If

		If Row <= 0 Then
		    ggoSpread.Source = .vspdData
		    If lgSortKey = 1 Then
		        ggoSpread.SSSort Col
		        lgSortKey = 2
		    Else
		        ggoSpread.SSSort Col,lgSortKey
		        lgSortKey = 1
		    End If    
		    Exit Sub
		End If
	
		ggoSpread.Source = .vspdData
		.vspddata.row    = .vspddata.ActiveRow	

 		.vspdData.Col = C_AcctCd
	
		If Len(.vspdData.Text) < 1 Then
		    ggoSpread.Source = .vspdData2
		    ggoSpread.ClearSpreadData
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				

End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================== 
' Event Name : vspdData_LeaveCell 
' Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq

            .hItemSeq.Value = .vspdData.Text
            ggoSpread.Source = frm1.vspdData2
            ggoSpread.ClearSpreadData

			.vspddata.Col = 0
			If .vspddata.Text = ggoSpread.DeleteFlag Then
				Exit Sub
			End if
        End With

		lgCurrRow = NewRow
		Call DbQuery2(lgCurrRow)
    End If
End Sub

'==========================================================================================
' Event Name : vspdData_ButtonClicked
' Event Desc : 버튼 컬럼을 클릭할 경우 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim iFld1 
	Dim iFld2
	Dim iTable
	Dim istrCode

	With frm1.vspdData
		If Row > 0 And Col = C_AcctPopUp Then
			.Col = Col - 1
			.Row = Row
									
			Call OpenPopUp(.Text, 3)
		End If
		
		If Row > 0 And Col = C_deptPopup Then
			.Col = Col - 1
			.Row = Row							
			Call OpenUnderDept(.Text, 1)
			'//Call OpenPopUp(.Text, 4 )
    	End If    	

		If Row > 0 And Col = C_DocCurPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenPopUp(.Text, 2)
		End If
	End With
End Sub

'=======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	Dim tmpDrCrFG
	Dim IntRetCD
	Dim TempExchRate
	Dim TempAmt
	
	With frm1
	    ggoSpread.Source = .vspdData
	    ggoSpread.UpdateRow Row
	    
	    .vspdData.Row = Row
	    
	    Select Case Col
			Case   C_DeptCd
				.vspdData.Col = C_DeptCd
				Call DeptCd_underChange(.vspdData.text)
		    Case   C_AcctCd
			    .vspdData.Col = 0
				If .vspdData.Text = ggoSpread.InsertFlag Then
					.vspdData.Col   = C_ItemSeq
					.hItemSeq.Value = .vspdData.Text
					.vspdData.Col   = C_AcctCd			

					If Len(.vspdData.Text) > 0 Then
						.vspdData.Row = Row
						.vspdData.Col = C_ItemSeq	 
						DeleteHsheet .vspdData.Text
						
						.vspdData.Col = C_DrCrFg
						tmpDrCrFG = .vspdData.text
						.vspdData.Col = C_AcctCd

						If (.cboGlType.value = "01" Or .cboGlType.value = "02") And .vspdData.text = lgCashAcct Then
							IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
							.vspdData.Text = ""
							.vspdData.Col = C_AcctNm
							.vspdData.Text = ""
							Exit Sub
						Else
	'						If  CheckVATType(.vspdData.Text ) Then
	'							ggoSpread.SSSetRequired C_VatNm    , -1    , C_VatNm
	'						else
	'							ggoSpread.SpreadUnLock C_VatNm		, -1    , C_VatNm
	'						End If
							Call Dbquery3(Row)
							Call InputCtrlVal(Row)
						End If

						'.vspdData.Col = C_AcctCd
						'If .cboGlType.Value <> "03" And .vspdData.text = lgCashAcct Then
						'	IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
						'	.vspdData.Text = ""
						'	.vspdData.Col = C_AcctNm
						'	.vspdData.Text = ""
						'	Exit Sub
						'Else
						'	Call Dbquery3(Row)				
						'	Call InputCtrlVal(Row)						
						'End If	
					Else
						.vspdData.Col = C_AcctNm
						.vspdData.Text = ""
					End If   
				End If
	    	Case C_DrCrFg
	    		Call SetSumItem
	    	Case C_DrCrNm  
				Call SetSumItem	
	    	Case C_ItemAmt
				.vspdData.Row = Row
				.vspdData.Col = C_ItemLocAmt
				.vspdData.Text = ""
	    		Call SetSumItem	
			Case C_ItemLocAmt
				.vspdData.Row = Row
				.vspdData.Col = C_IsLAmtChange
				.vspdData.Text = "Y"
				Call SetSumItem	
			Case C_ExchRate
				.vspdData.Row = Row
				.vspdData.Col = C_DocCur
				If UCase(Trim(.vspdData.Text)) = parent.gCurrency Then
					.vspdData.Row = Row
					.vspdData.Col = C_ExchRate
					.vspdData.Text = 1
				End If
			Case C_DocCur
				.vspdData.Row = Row
				.vspdData.Col = C_ItemLocAmt
				.vspdData.Text = ""		
				.vspdData.Col = C_DocCur
				
				If UCase(Trim(.vspdData.Text)) = parent.gCurrency Then
					.vspdData.Col = C_ExchRate
					.vspdData.Text = 1
				Else
					Call FindExchRate(UniConvDateToYYYYMMDD(.txtGLDt.text,parent.gDateFormat,""), UCase(Trim(.vspdData.Text)),.vspdData.ActiveRow)
				End If

				Call DocCur_OnChange(Row,Row)
	    End Select
	End With
End Sub

'==========================================================================================
'   Event Name : vspddata_ComboSelChange
'   Event Desc : Combo 변경 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim tmpDrCrFg
	Dim ii
	Dim iChkAcctForVat
	
	'---------- Coding part -------------------------------------------------------------
	' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1
		.vspddata.Row = Row

		Select Case Col
			Case C_DrCrNm
       			.vspddata.Col = Col		       			
				intIndex = .vspddata.Value
				.vspddata.Col = C_DrCrFg					
				.vspddata.Value = intIndex						
				tmpDrCrFg = .vspddata.text
				SetSpread2Color 	
			Case C_VatNm
				.vspddata.Col = Col		       			
			    intIndex = .vspddata.Value
				.vspddata.Col = C_VatType				
				.vspddata.Value = intIndex		
				Call InputCtrlVal(Row)'
		End Select
	End With
End Sub

'==========================================================================================
'   Event Name : txtTempGlNo_OnKeyPress
'   Event Desc : 
'==========================================================================================
Sub txtGlNo_OnKeyPress()	
	If window.event.keycode = 39 then	'Single quotation mark 입력불가 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
'   Event Name : txtTempGlNo_OnKeyPress
'   Event Desc : 
'==========================================================================================
Sub txtGlNo_OnKeyUp()	
	If Instr(1,frm1.txtGlNo.Value,"'") > 0 then
		frm1.txtGlNo.Value = Replace(frm1.txtGlNo.Value, "'", "")		
	End if
End Sub

'==========================================================================================
'   Event Name : txtTempGlNo_OnKeyPress
'   Event Desc : 
'==========================================================================================
Sub txtGlNo_onpaste()	
	Dim iStrGlNo 	
	iStrGlNo = window.clipboardData.getData("Text")
	iStrGlNo = RePlace(iStrGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrGlNo)		
End Sub

'==========================================================================================
'   Event Name : DocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub DocCur_OnChange(FromRow, ToRow)
	Dim ii
    lgBlnFlgChgValue = True

	For ii = FromRow To	ToRow
		frm1.vspdData.Row = ii
		frm1.vspdData.Col = C_DocCur

		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			Call CurFormatNumericOCX(ii)
			Call CurFormatNumSprSheet(ii)
			Call SetSumItem
		End If
	Next
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
    Dim strSelect,strFrom,strWhere
    Dim IntRetCD 
	Dim arrVal1,arrVal2
	Dim ii,jj

	With frm1
		If Trim(.txtGLDt.Text = "") Or Trim(.txtDeptCd.value) = "" Then    
			Exit sub
		End If
    
		lgBlnFlgChgValue = True

		strSelect	=			 " dept_cd, org_change_id, internal_cd "
		strFrom		=			 " b_acct_dept(NOLOCK) "
		strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.Value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")
			.txtDeptCd.Value    = ""
			.txtDeptNm.Value    = ""
			.hOrgChangeId.Value = ""
		Else
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)

			For ii = 0 To jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				.hOrgChangeId.Value = Trim(arrVal2(2))
			Next
		End If
	End With
End Sub

Sub QueryDeptCd_OnChange()
    Dim strSelect,strFrom,strWhere
    Dim IntRetCD 
	Dim arrVal1,arrVal2
	Dim ii,jj

	With frm1
		If Trim(.txtGLDt.Text = "") Then    
			Exit sub
		End If
    
		lgBlnFlgChgValue = True

		strSelect	=			 " dept_cd, org_change_id, internal_cd "
		strFrom		=			 " b_acct_dept(NOLOCK) "
		strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.Value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			.txtDeptCd.Value    = ""
			.txtDeptNm.Value    = ""
			.hOrgChangeId.Value = ""
		Else
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			jj = Ubound(arrVal1,1)

			For ii = 0 To jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				.hOrgChangeId.Value = Trim(arrVal2(2))
			Next
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : DeptCd_underChange(Byval strCode)
'   Event Desc : 
'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 

	With frm1
		If Trim(.txtGLDt.Text = "") Then    
			Exit sub
		End If

		lgBlnFlgChgValue = True

		strSelect	=			 " dept_cd, org_change_id, internal_cd "
		strFrom		=			 " b_acct_dept(NOLOCK) "
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")

			.vspdData.Col  = C_deptcd
			.vspdData.Row  = .vspdData.ActiveRow
			.vspdData.text = ""
			.vspdData.Col  = C_deptnm
			.vspdData.Row  = .vspdData.ActiveRow
			.vspdData.text = ""
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : txtGLDt_Change
'   Event Desc : 
'==========================================================================================
Sub txtGLDt_Change()
    Dim strSelect,strFrom,strWhere
    Dim IntRetCD 
	Dim arrVal1,arrVal2
	Dim ii,jj

	If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
				If LTrim(RTrim(.txtDeptCd.Value)) <> "" And Trim(.txtGLDt.Text <> "") Then
					strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
					strFrom		=			 " b_acct_dept(NOLOCK) "		
					strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(.txtDeptCd.Value)), "''", "S") 
					strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
	
					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124600","X","X","X")
						.txtDeptCd.Value = ""
						.txtDeptNm.Value = ""
						.hOrgChangeId.Value = ""
						If .vspdData.MaxRows <> 0 Then
							For ii = 1 To .vspdData.MaxRows
							.vspdData.Col = C_deptcd
						    .vspdData.Row = ii
						    .vspdData.text = ""
						    .vspdData.Col = C_deptnm
						    .vspdData.text = ""
							Next
						End If
					Else
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
						jj = Ubound(arrVal1,1)

						For ii = 0 To jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))
							.hOrgChangeId.Value = Trim(arrVal2(2))
						Next
					End If
				End If
			End With
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : cboGLType_OnChange
'   Event Desc : 
'==========================================================================================
Sub cboGLType_OnChange()
	Dim	i		
	Dim IntRetCD	

	With frm1
		ggoSpread.Source = .vspdData

		Select Case .cboGlType.Value
			Case "01"
				'입금전표로 바꾸면 차변이 입력되거나 현금계정이 입력되었는지 check한다.
				For i = 1 To  .vspdData.Maxrows
					.vspddata.Row = i
					.vspddata.col = C_Acctcd
					If .vspddata.text = lgCashAcct Then
						.cboGlType.Value = "03"
						Call CboGLType_ProtectGrid(.cboGlType.Value )
						IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
						Exit Sub
					End If
																				
					.vspddata.col = C_DrCrFg
					If  Trim(.vspddata.Value) = "2" Then
						.cboGlType.Value = "03"
						Call CboGLType_ProtectGrid(.cboGlType.Value )
						IntRetCD = DisplayMsgBox("113104", "X", "X", "X")
						Exit Sub
					End If
				Next

				For i = 1 To .vspdData.Maxrows
					.vspddata.Row = i
					.vspddata.col = C_DrCrFg
					If Trim(.vspddata.Value) <> "1"  Then					
						.vspdData.Value	= "1"							
						.vspddata.col = C_DrCrNm
						.vspdData.Value	= "1"							
					End If
					
					.vspddata.col = C_DocCur
					.vspddata.text = parent.gCurrency
				Next

				Call CboGLType_ProtectGrid(.cboGlType.Value)
			Case "02"
				'출금전표로 바꾸면 대변이 입력되거나 현금계정이 입력되었는지 check한다.	
				For i = 1 To  .vspddata.maxrows
					.vspddata.Row = i
					.vspddata.col = C_Acctcd

					If .vspddata.text = lgCashAcct Then
						.cboGlType.Value = "03"
						Call CboGLType_ProtectGrid(.cboGlType.Value )
						IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
						Exit Sub
					End If

					.vspddata.col = C_DrCrFg
					If Trim(.vspddata.Value) = "1" Then
						.cboGlType.Value = "03"
						Call CboGLType_ProtectGrid(.cboGlType.Value )
						IntRetCD = DisplayMsgBox("113105", "X", "X", "X")
						Exit Sub
					End If
				Next

				For i = 1 To .vspddata.maxrows
					.vspddata.Row = i
					.vspddata.col = C_DrCrFg

					If Trim(.vspddata.Value) <> "2"  Then					
						.vspdData.Value	= "2"							
						.vspddata.col = C_DrCrNm
						.vspdData.Value	= "2"							
					End If

					.vspddata.col = C_DocCur
					.vspddata.text = parent.gCurrency
				Next

				Call CboGLType_ProtectGrid(.cboGlType.Value )
			Case "03"
				'대체로 바꾸면 Protect를 풀어준다.		
				Call CboGLType_ProtectGrid(.cboGlType.Value )
		End Select
	End With

	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
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


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD
    Dim RetFlag
    lgstartfnc = True
    FncQuery = False																'⊙: Processing is NG
    Err.Clear																		'☜: Protect system from crashing

    '-----------------------
    'Check previous data area														 변경된 내용이 있는지 확인한다.
    '-----------------------
	With frm1
		ggoSpread.Source = .vspdData
 		' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
		' ChkField(pDoc, pStrGrp) As Boolean
		If Not chkField(Document, "1") Then												'⊙: This function check indispensable field
			Exit Function
		End If
    
		If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
			If IntRetCD = vbNo Then
		  	Exit Function
		 	End If
		End If

		'-----------------------
		'Erase contents area
		'-----------------------
		' 현재 Page의 Form Element들을 Clear한다. 
		Call ggoOper.ClearField(Document, "2")											'⊙: Condition field clear

		ggoSpread.Source = .vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = .vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = .vspdData3
		ggoSpread.ClearSpreadData

		Call InitVariables		
		'Call InitComboBoxGrid
		'Check condition area

		If .txtDeptCd.Value = "" Then
			.txtDeptNm.Value = ""
		End If

		'-----------------------
		'Query function call area
		'-----------------------
		If DbQuery = False Then															'☜: Query db data
			Exit Function
		End If

		If .vspddata.maxrows = 0 Then
			.txtGlNo.Value = ""
		End If
	End With

    FncQuery = True																		'⊙: Processing is OK
    lgstartfnc = False
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    Dim var1, var2
    
    On Error Resume Next            '☜: Protect system from crashing
    Err.Clear

    lgstartfnc = True
    FncNew = False                  '⊙: Processing is NG

	With frm1
	    ggoSpread.Source = .vspdData
	    var1 = ggoSpread.SSCheckChange
	    ggoSpread.Source = .vspdData2
	    var2 = ggoSpread.SSCheckChange

	    '-----------------------
	    'Check previous data area
	    '-----------------------
	    ' 변경된 내용이 있는지 확인한다.
	    If (lgBlnFlgChgValue = True Or var1 = True Or var2 = True) And lgBlnExecDelete <> True Then
	        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
	    
	    lgBlnExecDelete = False
	    '-----------------------
	    'Erase condition area
	    'Erase contents area
	    '-----------------------
		' 현재 Page의 Form Element들을 Clear한다. 
		' ClearField(pDoc, Optional ByVal pStrGrp)
	    Call ggoOper.ClearField(Document, "1")											'⊙: Clear Condition Field
	    Call ggoOper.ClearField(Document, "2")											'⊙: Clear Contents  Field
	    
	    ggoSpread.Source = .vspdData
	    ggoSpread.ClearSpreadData
	    ggoSpread.Source = .vspdData2
	    ggoSpread.ClearSpreadData
	    ggoSpread.Source = .vspdData3
	    ggoSpread.ClearSpreadData
	    
	    Call ggoOper.LockField(Document,  "N")											'⊙: Lock  Suitable  Field

	    SetGridFocus()
	    SetGridFocus2()

		Call SetDefaultVal
	    Call InitVariables																'⊙: Initializes local global variables
		Call SetSumItem()
	'    Call DocCur_OnChange()
	    Call SetToolbar(MENU_NEW)														'버튼 툴바 제어 

		'Call ggoOper.SetReqAttr(.txtDocCur, "N")
		Call ggoOper.SetReqAttr(.txtGlDt,"N")
		Call ggoOper.SetReqAttr(.txtdesc,"D")
	End With
	
    lgBlnFlgChgValue = False

    FncNew = True																	'⊙: Processing is OK
    lgFormLoad = True																' gldt read
    lgQueryOk = False
    lgstartfnc = False
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
    
    On Error Resume Next															'☜: Protect system from crashing
    Err.Clear                   

    FncDelete = False																'⊙: Processing is NG
    lgBlnExecDelete = True
	ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = False Then											'변경된 부분이 없을경우 
		intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")				'삭제하시겠습니까?
		If intRetCd = VBNO Then
			Exit Function
		End IF
    Else
		IntRetCD = DisplayMsgBox("900038", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 삭제하시겠습니까?
    	If IntRetCD = vbNo Then    		
      		Exit Function
    	End If
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    If  DbDelete = False Then														'☜: Delete db data
		Exit Function
	End If	

    FncDelete = True 
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    On Error Resume Next															'☜: Protect system from crashing
    Err.Clear                                                               
    
    FncSave = False                                                         
    
	'-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")							'No data changed!!
        Exit Function
    End If
   
    If CheckSpread3 = False Then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")                           '필수입력 check!!
        Exit Function
    End If
	
	If frm1.vspdData.MaxRows < 1 Then												'회계전표존재하지 않음 
		IntRetCD = DisplayMsgBox("113100", "X", "X", "X")
		Exit Function
	End If
	
	'-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") Then												'⊙: Check contents area
		Exit Function
    End If
    
    If Not ggoSpread.SSDefaultCheck Then											'⊙: Check contents area
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3
    If Not ggoSpread.SSDefaultCheck Then											'⊙: Check contents area
		Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '----------------------- 
	'Call ExchRateCheck()
    If DbSave = False Then				                                            '☜: Save db data
		Exit Function
    End If
    
    FncSave = True                                                          
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim  IntRetCD	 

	With frm1
		.vspdData.ReDraw = False
		If .vspdData.MaxRows < 1 Then Exit Function

		ggoSpread.Source = .vspdData
		ggoSpread.CopyRow
		SetSpreadColor "I", 0, .vspdData.ActiveRow, .vspdData.ActiveRow
		MaxSpreadVal .vspdData, C_ItemSeq, .vspdData.ActiveRow

		Call ReFormatSpreadCellByCellByCurrency(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,C_DocCur,C_ItemAmt ,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")
		
		'Call DocCur_OnChange(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
		Call vspdData_Change(C_AcctCd, .vspddata.activerow)
		Call SetSumItem()
	End With
End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    Dim iItemSeq
    Dim RowDocCur

    With frm1.vspdData
		If .MaxRows < 1 Then Exit Function	
	
		If .MaxRows = 1 Then Call ggoOper.SetReqAttr(frm1.cboGlType,   "N")

        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			If Len(Trim(.text)) > 0 Then 
				.Col = C_ItemSeq
				DeleteHSheet(.Text)
			End If	
        End If

        ggoSpread.Source = frm1.vspdData
        ggoSpread.EditUndo

        Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt, "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")

		If .MaxRows = 0 Then
			Call SetToolbar(MENU_NEW)
			Exit Function
		End If

        Call InitData

        .Row = .ActiveRow
        .Col = 0

		If .row = 0 Then 
			Exit Function
		End If

        If .Text = ggoSpread.InsertFlag Then
            .Col = C_AcctCd
            If Len(.Text) > 0 Then
				.Col = C_ItemSeq
				frm1.hItemSeq.Value = .Text
	            frm1.vspdData2.MaxRows = 0
		        Call DbQuery3(.ActiveRow)
            End If
        Else
            .Col = C_ItemSeq
            frm1.hItemSeq.Value = .Text
            frm1.vspdData2.MaxRows = 0
		    Call DbQuery2(.ActiveRow)            
        End If
    End With        

    Call SetSumItem()
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
	Dim iCurRowPos

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False         
   
	'-----------------------
    'Check content area
    '----------------------- 
    If Not chkField(Document, "2") Then 
        Exit Function
    End If

    If IsNumeric(Trim(pvRowCnt)) then
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
		iCurRowPos = .vspdData.ActiveRow

        For imRow2 = 1 To imRow 
            ggoSpread.InsertRow ,1

            .vspdData.row  = .vspdData.ActiveRow
			.vspdData.col   = C_deptcd
            .vspddata.text = UCase(.txtDeptCd.Value)
            .vspdData.col  = C_deptnm
            .vspddata.text = .txtDeptNm.Value
            .vspdData.col  = C_DocCur
            .vspddata.text = parent.gCurrency
            .vspdData.col  = C_ExchRate
            .vspddata.text = "1"
            .vspdData.col  = C_ItemDesc
            .vspddata.text = .txtDesc.Value

            If .cboGlType.value = "01" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 1
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 1
            Elseif .cboGlType.value = "02" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 2
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 2
            End If

            SetSpreadColor "I", 0, .vspdData.ActiveRow, .vspdData.ActiveRow
            MaxSpreadVal .vspdData, C_ItemSeq, .vspdData.ActiveRow
        Next
        
        Call ReFormatSpreadCellByCellByCurrency(.vspdData,iCurRowPos + 1,iCurRowPos + imRow,C_DocCur,C_ItemAmt ,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,iCurRowPos + 1,iCurRowPos + imRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")                
        .vspdData.ReDraw = True
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
	Dim iDelRowCnt, i
    Dim DelItemSeq

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData 

		.Row = .ActiveRow
		.Col = 0 
			
		If .MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function

        .Col = 1 
        DelItemSeq = .Text
    	
    	lDelRows = ggoSpread.DeleteRow
    End With
        
    DeleteHsheet DelItemSeq
    Call SetSumItem()
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    On Error Resume Next																'☜: Protect system from crashing
    parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    On Error Resume Next																'☜: Protect system from crashing
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
	Dim indx

	On Error Resume Next
	Err.Clear 		

	ggoSpread.Source = gActiveSpdSheet

	With frm1
		Select Case Trim(UCase(gActiveSpdSheet.Name))
			Case "VSPDDATA"
				Call PrevspdDataRestore(gActiveSpdSheet)
				Call ggoSpread.RestoreSpreadInf()
				Call InitSpreadSheet()
		        Call InitComboBoxGrid
				Call ggoSpread.ReOrderingSpreadData()
				Call InitData()
				Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,-1,C_DocCur,C_ItemAmt ,"A" ,"I","X","X")
				Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,-1,C_DocCur,C_ExchRate,"D" ,"I","X","X")			
					
				If .cboGlInputType.Value <> C_GLINPUTTYPE then
				    Call SetSpreadLock("Q", 1, 1, "")			    
				    Call SetSpread2Lock("",1,"","")			    
				Else
		            Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)
		            Call SetSpread2Color()
				End If
			Case "VSPDDATA2"
				Call PrevspdData2Restore(gActiveSpdSheet)
				Call ggoSpread.RestoreSpreadInf()
				Call InitCtrlSpread()													'관리항목 그리드 초기화			
				Call ggoSpread.ReOrderingSpreadData()
				Call InitData()

				If .cboGlInputType.Value <> C_GLINPUTTYPE Then
				    Call SetSpread2Lock("",1,"","")
				Else
		            Call SetSpread2Color()
				End If
		End Select
	
		If .vspdData2.MaxRows <= 0 Then
			Call DbQuery2(.vspdData.ActiveRow)
		End If
	End With

	Call SetSumItem()
End Sub

'=======================================================================================================
'   Function Name : PrevspdDataRestore
'   Function Desc : 
'=======================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	With frm1
		For indx = 0 To .vspdData.MaxRows
		    .vspdData.Row = indx
		    .vspdData.Col = 0
			
			If .vspdData.Text <> "" Then
				Select Case .vspdData.Text			
					Case ggoSpread.InsertFlag					
						.vspdData.Col = C_ItemSeq					
						Call DeleteHsheet(.vspdData.Text)					
					Case ggoSpread.UpdateFlag		
						For indx1 = 0 To .vspdData3.MaxRows					
							.vspdData3.Row = indx1
							.vspdData3.Col = 0
							Select Case .vspdData3.Text 
								Case ggoSpread.UpdateFlag
									.vspdData.Col = C_ItemSeq
									.vspdData3.Col = 1					
									If UCase(Trim(.vspdData.Text)) = UCase(Trim(.vspdData3.Text)) Then
										Call DeleteHsheet(.vspdData.Text)
										Call fncRestoreDbQuery2(indx, .vspdData.ActiveRow, .htxtGLNo.Value)
									End If
							End Select
						Next
						'ggoSpread.Source = .vspdData					
						'ggoSpread.EditUndo
					Case ggoSpread.DeleteFlag
						Call fncRestoreDbQuery2(indx, .vspdData.ActiveRow, .htxtGLNo.Value)
						'ggoSpread.Source = .vspdData
						'ggoSpread.EditUndo
				End Select
			End If
		Next
	End With
	ggoSpread.Source = pActiveSheetName
End Sub

'=======================================================================================================
'   Function Name : PrevspdData2Restore
'   Function Desc : 
'=======================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	With frm1
		For indx = 0 To .vspdData2.MaxRows
		    .vspdData2.Row = indx
		    .vspdData2.Col = 0

			If .vspdData2.Text <> "" Then
				Select Case .vspdData2.Text
					Case ggoSpread.InsertFlag
						.vspdData2.Col = C_HItemSeq
						For indx1 = 0 To .vspdData.MaxRows
							.vspdData.Row = indx1
							.vspdData.Col = C_ItemSeq
							If .vspdData.Text = .vspdData2.Text Then
								Call DeleteHsheet(.vspdData.Text)
								ggoSpread.Source = .vspdData	
						        ggoSpread.EditUndo							
							End If
						Next
					Case ggoSpread.UpdateFlag
						.vspdData2.Col = C_HItemSeq
						For indx1 = 0 To .vspdData.MaxRows
							.vspdData.Row = indx1
							.vspdData.Col = C_ItemSeq
							If .vspdData.Text = .vspdData2.Text Then
								Call DeleteHsheet(.vspdData.Text)
								ggoSpread.Source = .vspdData
								ggoSpread.EditUndo
								Call fncRestoreDbQuery2(indx1, .vspdData.ActiveRow, .htxtGLNo.Value)
							End If
						Next
					Case ggoSpread.DeleteFlag
				End Select
			End If
		Next
	End With

	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncRestoreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)

	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text
		
	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & strItemSeq & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.GL_NO = " & FilterVar(UCase(pInvalue1), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
'		Call SetSpread2Color()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If

'	Set gActiveElement = document.ActiveElemen
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then  
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"	
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim RetFlag

    Err.Clear																	'☜: Protect system from crashing

    DbQuery = False
    Call LayerShowHide(1)

    With frm1
		ggoSpread.Source = .vspdData3
		ggospread.ClearSpreadData()

	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtGlNo=" & UCase(Trim(.htxtGlNo.Value))		'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 
			strVal = strVal & "&txtGlNo=" & UCase(Trim(.txtGlNo.Value))			'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function


'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk()
	Dim ii
	
	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE										'Indicates that current mode is Update mode

		Call ggoOper.SetReqAttr(.txtGLDt,	"Q")
		Call ggoOper.SetReqAttr(.cboGlType,	"Q")

        If .cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetToolbar(MENU_PRT) 
			Call SetSpreadLock("Q", 1, 1, "")			
			Call ggoOper.SetReqAttr(.txtDeptCd,	"Q")													 
			Call ggoOper.SetReqAttr(.txtdesc,   "Q")
		Else
			Call SetToolbar(MENU_UPD)											'버튼 툴바 제어 
			Call SetSpreadLock("Q", 0, 1, "")
			Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)
			Call ggoOper.SetReqAttr(.txtDeptCd,	"N")													 
			Call ggoOper.SetReqAttr(.txtdesc,	"D")
		End If

        .txtCommandMode.Value = "UPDATE"

        Call InitData()

        For ii= 1 To .vspdData.MaxRows
			CurFormatNumSprSheet(ii)
		Next

        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text
            Call DbQuery2(1)
        End If
    End With

	Call SetGridFocus()
    Call SetGridFocus2()
	Call QueryDeptCd_OnChange()

	lgBlnFlgChgValue = False
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
	Dim arrTemp
	Dim Indx1
	
	'Err.Clear
	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
        If CopyFromData(.hItemSeq.Value) = True Then
			If .cboGlInputType.Value <> C_GLINPUTTYPE then
				Call SetSpread2Lock("",1,"","")
			Else
				Call SetSpread2Color()
			End If 	
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , LTrim(ISNULL(C.CTRL_VAL,'')), '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.GL_NO = " & FilterVar(UCase(.txtGLNo.Value), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
			
		frm1.vspdData2.ReDraw = False
		
'@@	txtsql.Value =  "select " & strSelect &vbcrlf & "from " & vbcrlf  & strfrom  &  vbcrlf & "where" & vbcrlf & strwhere
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = .vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2			

			For lngRows = 1 To .vspdData2.Maxrows
				.vspddata2.row = lngRows	
				.vspddata2.col = C_Tableid 
				If Trim(.vspddata2.text) <> "" Then
					.vspddata2.col = C_Tableid
					strTableid = frm1.vspddata2.text
					.vspddata2.col = C_Colid
					strColid = frm1.vspddata2.text
					.vspddata2.col = C_ColNm'XX
					strColNm = frm1.vspddata2.text	
					.vspddata2.col = C_MajorCd					
					strMajorCd = .vspddata2.text	
					
					.vspddata2.col = C_CtrlVal
					
					strNmwhere = strColid & " =  " & FilterVar(UCase(.vspddata2.text), "''", "S")
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD = " & FilterVar(strMajorCd, "''", "S") 
					End If
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						.vspddata2.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						.vspddata2.text = arrVal(0)					
					End If
				End If								
				
				strVal = strVal & Chr(11) & .hItemSeq.Value
				
				.vspdData2.Col = C_DtlSeq
				strVal = strVal & Chr(11) & .vspdData2.Text
                
				.vspdData2.Col = C_CtrlCd
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlNm
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlVal
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlPB
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlValNm
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Seq
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Tableid
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Colid
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_ColNm
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Datatype
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_DataLen
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_DRFg
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_MajorCd
				strVal = strVal & Chr(11) & .vspdData2.Text
				
				.vspdData2.Col = C_MajorCd + 1
'				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & lngRows
				
				strVal = strVal & Chr(11) & Chr(12)
			Next					

			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		intItemCnt = .vspddata.MaxRows

		If .cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetSpread2Lock("",1,"","")
		Else
			Call SetSpread2Color()
		End If

		Call LayerShowHide(0)
		.vspdData2.ReDraw = True
	End With

	DbQuery2 = True
	lgQueryOk = True
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim intIndex2 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow			
			.Col = C_DrCrFg
			intIndex = .Value
			.col = C_DrCrNm
			.Value = intindex
									
			.Col = C_VatType
			intIndex2 = .Value
			.col = C_VatNm
			.Value = intIndex2		
		Next	
	End With
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pAP010M 
    Dim lngRows , itemRows
    Dim lGrpcnt
    DIM strVal 
    Dim tempItemSeq
    Dim	intRetCd	
    Dim strNote
    Dim strItemDesc

    DbSave = False                                                          
    Call LayerShowHide(1)

    On Error Resume Next                                                   

    Call SetSumItem

	With frm1
		.txtFlgMode.Value       = lgIntFlgMode									
		.txtUpdtUserId.Value    = parent.gUsrID
		.txtInsrtUserId.Value   = parent.gUsrID
		.txtMode.Value          = parent.UID_M0002
		.txtAuthorityFlag.Value = lgAuthorityFlag               '권한관리 추가 

		'-----------------------
		'Data manipulate area
		'-----------------------
		' Data 연결 규칙 
		' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 
		lGrpCnt = 1
		strVal = ""
		ggoSpread.Source = .vspdData
	End With
    
    With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0

			If .Text <> ggoSpread.DeleteFlag Then
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
			        
			    .Col = C_ItemSeq	'1
			    strVal = strVal & Trim(.Text) & parent.gColSep
			            
			    .Col = C_deptcd	    '2
			    strVal = strVal & Trim(.Text) & parent.gColSep
			        
			    .Col = C_AcctCd		'3
			    strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_DrCrFG		'4
			    strVal = strVal & Trim(.Text) & parent.gColSep
			        
			    .Col = C_ItemAmt	'5
			    strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
			        
   				.Col = C_ItemLocAmt	'6
				strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
			        
			    .Col = C_ItemDesc	'7
				strItemDesc = Trim(.Text)
			    
				If Trim(strItemDesc) = "" Or isnull(strItemDesc) Then
					ggoSpread.Source = frm1.vspdData3
					frm1.vspdData.Col = C_ItemSeq
					tempItemSeq = .Text  
					strNote = ""
					 
					With frm1.vspdData3
					   	For itemRows = 1 To .MaxRows
					   		.Row = itemRows
					   		.Col = 1
									
					   		If .Text =  tempItemSeq Then 
					   			.Col= 9 'C_Tableid	+ 1				
					   			If .Text = "B_BIZ_PARTNER" Or .Text = "B_BANK" Or .Text = "F_DPST" Then
					   				.Col = 7 'C_CtrlValNm + 1 
					   			Else
					   				.Col = 5 'C_CtrlVal + 1 
					   			End If											
					   			strNote = strNote & C_NoteSep & Trim(.Text)
					   		End If		    
					   	Next
					   	strNote = Mid(strNote,2)
					End With

					strVal = strVal & strNote & parent.gColSep
					ggoSpread.Source = frm1.vspdData
				Else
					strVal = strVal & strItemDesc & parent.gColSep		'8
				End If

				.Col = C_ExchRate	'9
			    strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep

			    .Col = C_VatType	'10
			    strVal = strVal & Trim(.Text) & parent.gColSep

			    .Col = C_DocCur		'11
			    strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep

			    lGrpCnt = lGrpCnt + 1
			End If
		Next
    End With
	
	With frm1
		.txtMaxRows.Value = lGrpCnt-1												'Spread Sheet의 변경된 최대갯수 
		.txtSpread.Value =  strVal													'Spread Sheet 내용을 저장 

		If .txtSpread.Value = "" Then	
			intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")				'☜ 바뀐부분 

			If intRetCd = VBNO Then
				Exit Function
			End If

			If  DbDelete = False Then
				Exit Function
			End If		

			ggoSpread.Source = .vspdData
			ggoSpread.ClearSpreadData
			ggoSpread.Source = .vspdData2
			ggoSpread.ClearSpreadData
			ggoSpread.Source = .vspdData3
			ggoSpread.ClearSpreadData

			Call InitVariables
			Exit Function
		End If

		lGrpCnt = 1
		strVal = ""
		ggoSpread.Source = .vspdData3
	End With

    With frm1.vspdData3      ' Dtl 저장 
		For itemRows = 1 To frm1.vspdData.MaxRows 
 		    frm1.vspdData.Row = itemRows
		    frm1.vspdData.Col = 0

			If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
		        frm1.vspdData.Col = C_ItemSeq
			    tempItemSeq = frm1.vspdData.Text  

			    For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = 1
					If .text = tempitemseq Then
		                .Col = 0 
						strVal = strVal & "C" & parent.gColSep
						.Col = 1 		 					'ItemSEQ	
						strVal = strVal & tempitemseq & parent.gColSep
						.Col =  2 'C_DtlSeq + 1   			'Dtl SEQ
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col =  3 'C_CtrlCd + 1		 		'관리항목코드 
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col = 5 'C_CtrlVal + 1				'관리항목 Value 
						strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep	
				
						lGrpCnt = lGrpCnt + 1
					End If
		    	Next
			End If
		Next
    End With

    frm1.txtMaxRows3.Value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
    frm1.txtSpread3.Value  = strVal											'Spread Sheet 내용을 저장 

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'저장 비지니스 ASP 를 가동 

    DbSave = True                                                           
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk(Byval GlNo)												'☆: 저장 성공후 실행 로직 
	frm1.txtGlNo.Value = UCase(Trim(GlNo))
    frm1.txtCommandMode.Value = "UPDATE"

	Call ggoOper.ClearField(Document, "2")									'⊙: Condition field clear    
    Call InitVariables														'⊙: Initializes local global variables
    'Call InitComboBoxGrid
	'Call InitComboBoxGridVat
	Call DbQuery
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	Dim strVal
	
    Err.Clear
    Call LayerShowHide(1)    
	DbDelete = False																	'⊙: Processing is NG

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003								'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtGlNo=" & UCase(Trim(frm1.txtGlNo.Value))						'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtGlDt=" & ggoOper.RetFormat(frm1.txtGLDt.Text, "yyyy-MM-dd")
    strVal = strVal & "&txtDeptCd=" & UCase(Trim(frm1.txtDeptCd.Value))
	strVal = strVal & "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.Value)
    strVal = strVal & "&txtGlinputType=" & Trim(frm1.txtGlinputType.Value)
    
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True																		'⊙: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'=======================================================================================================
Function DbDeleteOk()													'삭제 성공후 실행 로직 
	Call FncNew()
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------    
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX(Row)
	With frm1
		'차변금액 
'		ggoOper.FormatFieldByObjectOfCur .txtDrAmt, .txtDocCur.Value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'대변금액 
'		ggoOper.FormatFieldByObjectOfCur .txtCrAmt, .txtDocCur.Value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet(Row)
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.Row	 = Row
		.vspdData.Col	 = C_DocCur

		Call ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,C_DocCur,C_ItemAmt ,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,C_DocCur,C_ExchRate,"D" ,"I","X","X")					
	End With
End Sub

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
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd		
	Dim ii

	lgBlnFlgChgValue = True
	With frm1
		ggoSpread.Source = .vspdData

		.vspdData.Col = C_AcctCd
		.vspdData.Row = Row
		strAcctCd = Trim(.vspdData.text)

		frm1.vspdData.Col = C_deptcd
		frm1.vspdData.Row = Row			

		Call AutoInputDetail(strAcctCd, Trim(.vspdData.text), .txtGLDt.text, Row)

		For ii = 1 To .vspdData2.MaxRows
			.vspddata2.col = C_CtrlVal
			.vspddata2.row = ii

			If Trim(.vspddata2.text) <> "" Then
				Call CopyToHSheet2(.vspdData.ActiveRow,ii)
			End If
		Next
	End With
End Sub

'=======================================================================================================
'   Event Name : GetAcctForVat
'   Event Desc :
'======================================================================================================= 
Sub GetAcctForVat()	
	Dim ii
	
	lgBlnGetAcctForVat = False

	If CommonQueryRs("acct_cd", "a_acct(nolock)", "acct_type LIKE " & FilterVar("V_", "''", "S") & " and del_fg <> " & FilterVar("Y", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		lgArrAcctForVat = Split(Mid(lgF0, 1, Len(lgF0) - 1), Chr(11))
		lgBlnGetAcctForVat = True
	End If
End Sub
	
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기초계정등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>기초번호</TD>
									<TD CLASS=TD656 NOWRAP><INPUT NAME="txtGlNo" ALT="기초번호" MAXLENGTH="18" SIZE=20 STYLE="TEXT-ALIGN: left" tag  ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefGL()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>
								<TD CLASS=TD5 NOWRAP>기초일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5106ma1_OBJECT7_txtGLDt.js'></script></TD>								
								<TD CLASS=TD5 NOWRAP>전표형태</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlType" tag="24" STYLE="WIDTH:82px:" ALT="전표형태"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>								
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X"></TD>
													 <INPUT NAME="txtInternalCd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
								<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="24" STYLE="WIDTH:82px:" ALT="전표입력경로"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
			    
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="70" tag="2X" ></TD>
							</TR>							
							<TR> 
								<TD HEIGHT="60%" COLSPAN=4>
									<script language =javascript src='./js/a5106ma1_I587860077_vspdData.js'></script>								
								</TD>
							</TR>
							<TR>
<!--								<TD CLASS=TD5 NOWRAP>차대합계(거래)</TD>
								<TD>&nbsp;<script language =javascript src='./js/a5106ma1_OBJECT1_txtDrAmt.js'></script>
									&nbsp;<script language =javascript src='./js/a5106ma1_OBJECT2_txtCrAmt.js'></script></TD>
-->										

								<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTTON>&nbsp;
								<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;<script language =javascript src='./js/a5106ma1_OBJECT3_txtDrLocAmt.js'></script>
									&nbsp;<script language =javascript src='./js/a5106ma1_OBJECT4_txtCrLocAmt.js'></script></TD>							
							</TR>
							<TR>						                 
								<TD HEIGHT="40%" COLSPAN=4>
									<script language =javascript src='./js/a5106ma1_OBJECT5_vspdData2.js'></script>									
			  			  
								</TD>
							</TR>
							<!--<TR>						                 
								<TD HEIGHT="40%" COLSPAN=4>
									<script language =javascript src='./js/a5106ma1_OBJECT5_vspdData3.js'></script>									
			  			  
								</TD>
							</TR>-->
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>&nbsp;
					</TD>					
					<TD WIDTH=* ALIGN=RIGHT>					
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_TAX_REP)">계산서수정</a>		
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		<!--<TD WIDTH="100%" HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>-->
	</TR>

</TABLE>
<script language =javascript src='./js/a5106ma1_OBJECT6_vspdData3.js'></script>
<TEXTAREA class=hidden name=txtSpread    tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtGlNo"       tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtCommandMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"     tag="24" TABINDEX="-1"><!--권한관리추가 -->
<INPUT TYPE=HIDDEN NAME="txtGlinputType" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="hItemSeq"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP"   METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"       TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"      TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"     TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date"        TABINDEX="-1">	
</FORM>
</BODY>
</HTML>
