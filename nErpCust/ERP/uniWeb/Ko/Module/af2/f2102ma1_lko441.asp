
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2102ma1_lk0441
'*  4. Program Name         : 예산정보등록(LKO441) 
'*  5. Program Desc         : Register of Budget
'*  6. Comproxy List        : FU0021, FU0028
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2008.01.02
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*   - 2001.03.09 Song, Mun Gil 예산년월에 Mask 적용 
'*   - 2001.03.20 Song, Mun Gil 스프레드에 조직변경ID 컬럼 추가 
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "f2102mb1_lko441.asp"			'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns

Dim C_BDG_CD               
Dim C_BDG_CD_PB            
Dim C_BDG_NM               
Dim C_BDG_YYYYMM           
Dim C_DEPT_CD              
Dim C_DEPT_PB              
Dim C_DEPT_NM              
Dim C_ORG_CHANGE_ID       
Dim C_BDG_CTRL_FG			'하위부서 통제 여부
Dim C_BDG_PLAN_AMT         
Dim C_BDG_AMT               
Dim C_BDG_GL_AMT           
Dim C_BDG_TEMP_AMT
Dim C_ACCT_BDG_CTRL_FG     
Dim C_GP_BDG_CTRL_FG       

Const C_SHEETMAXROWS = 100

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
'Dim lgLngCurRows

Dim lgStrPrevBdgCdKey
Dim lgStrPrevBdgYMKey
Dim lgStrPrevDeptCdKey

Dim strFrDt
Dim strToDt
 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
'Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)

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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_BDG_CD			= 1               
	C_BDG_CD_PB         = 2                  
	C_BDG_NM            = 3                  
	C_BDG_YYYYMM        = 4                  
	C_DEPT_CD           = 5                  
	C_DEPT_PB           = 6                  
	C_DEPT_NM           = 7                  
	C_ORG_CHANGE_ID     = 8    
	C_BDG_CTRL_FG       = 9			'하위부서 통제 여부
	C_BDG_PLAN_AMT      = 10                  
	C_BDG_AMT           = 11                   
	C_BDG_GL_AMT        = 12                  
	C_BDG_TEMP_AMT      = 13         	
	C_ACCT_BDG_CTRL_FG  = 14
	C_GP_BDG_CTRL_FG    = 15
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$




 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
	lgStrPrevBdgCdKey = ""
	lgStrPrevBdgYMKey = ""
	lgStrPrevDeptCdKey = ""

    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
    lgSortKey = 1
    lgPageNo  = 0

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
	strSvrDate = "<%=GetSvrDate%>"
	
	frm1.txtBdgYymmFr.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.txtBdgYymmTo.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtBdgYymmFr, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtBdgYymmTo, parent.gDateFormat, 2)
	frm1.hOrgChangeId.value = parent.gChangeOrgId
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================== 2.2.3 InitSpreadSheet() =================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

	Dim strMaskYM
	Dim		sList
		
	strMaskYM = parent.gDateFormatYYYYMM
	
	strMaskYM = Replace(strMaskYM,"YYYY"      ,"9999")
	strMaskYM = Replace(strMaskYM,"YY"        ,"99")
	strMaskYM = Replace(strMaskYM,"MM"        ,"99")
	strMaskYM = Replace(strMaskYM,parent.gComDateType,"X")

	sList = "Y" & vbTab  & "N"
	
	With frm1.vspdData
        .ReDraw = False

        .MaxCols = C_GP_BDG_CTRL_FG + 1
        .Col = .MaxCols				'☜: 공통콘트롤 사용 Hidden Column

        .MaxRows = 0

                           'patch version
        Call GetSpreadColumnPos("A")
        
        ggoSpread.SSSetEdit     C_BDG_CD,				"예산코드",   15, , , 18, 2
        ggoSpread.SSSetButton   C_BDG_CD_PB
        ggoSpread.SSSetEdit     C_BDG_NM,				"예산명",     20, , , 30
        ggoSpread.SSSetMask     C_BDG_YYYYMM,			"예산년월",   15,2, strMaskYM            

        ggoSpread.SSSetEdit     C_DEPT_CD,				"부서코드",   15, , , 10, 2
        ggoSpread.SSSetButton   C_DEPT_PB
        ggoSpread.SSSetEdit     C_DEPT_NM,				"부서명",     20, , , 40

        ggoSpread.SSSetEdit		C_ORG_CHANGE_ID,		"조직변경ID", 10, , , 5
				
		ggoSpread.SSSetCombo	C_BDG_CTRL_FG,			"하위부서 통제여부", 15, 2

        ggoSpread.SSSetFloat	C_BDG_PLAN_AMT,			"예산금액",         20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat	C_BDG_AMT,				"예산한도금액",		20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
        ggoSpread.SSSetFloat	C_BDG_GL_AMT,			"예산회계전표금액",   20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
        ggoSpread.SSSetFloat	C_BDG_TEMP_AMT,			"예산결의전표금액",   20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
		

		call ggoSpread.MakePairsColumn(C_BDG_CD,C_BDG_CD_PB)
	   	call ggoSpread.MakePairsColumn(C_DEPT_CD,C_DEPT_PB)
	   	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_BDG_CTRL_FG,C_ACCT_BDG_CTRL_FG,True)
		Call ggoSpread.SSSetColHidden(C_GP_BDG_CTRL_FG,C_GP_BDG_CTRL_FG,True)

		ggoSpread.SetCombo sList, C_BDG_CTRL_FG

		
        .ReDraw = True

    
    End With
    
	Call SetSpreadLock     
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
		.ReDraw = False

			ggoSpread.SSSetRequired C_BDG_PLAN_AMT, -1,-1
			
			ggoSpread.SpreadLock C_BDG_CD_PB,    -1,    C_BDG_CD_PB      ' ,-1
			ggoSpread.SpreadLock C_DEPT_PB,      -1,    C_DEPT_PB       ',-1
			
			ggoSpread.SpreadLock C_BDG_CD,        -1,	C_BDG_CD
			ggoSpread.SpreadLock C_BDG_YYYYMM,    -1,   C_BDG_YYYYMM
			ggoSpread.SpreadLock C_DEPT_CD,       -1,   C_DEPT_CD
			
			ggoSpread.SpreadLock C_BDG_NM,       -1,	C_BDG_NM
			ggoSpread.SpreadLock C_DEPT_NM,      -1,	C_DEPT_NM
			ggoSpread.SpreadLock C_BDG_CTRL_FG,  -1,	C_BDG_CTRL_FG
			'ggoSpread.SSSetRequired C_BDG_CTRL_FG,  -1,	-1
			ggoSpread.SpreadLock C_BDG_AMT,      -1,	C_BDG_AMT
			ggoSpread.SpreadLock C_BDG_GL_AMT,   -1,	C_BDG_GL_AMT
			ggoSpread.SpreadLock C_BDG_TEMP_AMT, -1,	C_BDG_TEMP_AMT
						
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
		.ReDraw = True

    End With

End Sub


'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1

		.vspdData.ReDraw = False

		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional Row2)
		ggoSpread.SSSetRequired  C_BDG_CD,       pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BDG_NM,       pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_BDG_YYYYMM,   pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_BDG_PLAN_AMT, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DEPT_CD,      pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DEPT_NM,      pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected C_BDG_CTRL_FG,	 pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_BDG_CTRL_FG,  pvStartRow, pvEndRow	
		ggoSpread.SSSetProtected C_BDG_AMT,      pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BDG_GL_AMT,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BDG_TEMP_AMT, pvStartRow, pvEndRow				
		ggoSpread.SSSetProtected C_TRANS_AMT,    pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DIVERT_AMT,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ADD_AMT,      pvStartRow, pvEndRow
				
		.vspdData.ReDraw = True
    
    End With

End Sub

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 

'============================================================
'예산상세내역 팝업 
'============================================================
'Function OpenPopUpDtl()
'	Dim arrRet
'	Dim arrParam(5)	
'
'	If IsOpenPop = True Then Exit Function
'	
'	With frm1.vspdData
'		If .MaxRows < 1 Then 
'			Call DisplayMsgBox("900025","X","X","X")	'선택된 항목이 없습니다.
'			Exit Function
'		End If
'		
'		.Row = .ActiveRow
'		.Col = C_BDG_CD
'		arrParam(0) = .Text
'		.Col = C_BDG_NM
'		arrParam(1) = .Text
'		.Col = C_BDG_YYYYMM
'		arrParam(2) = UNICDate(.Text)
'		.Col = C_DEPT_CD
'		arrParam(3) = .Text
'		.Col = C_DEPT_NM
'		arrParam(4) = .Text
'		.Col = C_ORG_CHANGE_ID
'		arrParam(5) = .Text
'	End With
'	
'	IsOpenPop = True
'	
'	arrRet = window.showModalDialog("f2109pa1.asp", Array(arrParam), _
'		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
'	
'	IsOpenPop = False
'	
'	If arrRet(0) = ""  Then			
'		Exit Function
'	End If
'End Function

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True	

	tmpBdgYymmddFr  =  UniConvDateAToB(frm1.txtBdgYymmFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(frm1.txtBdgYymmTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("M", +1, tmpBdgYymmddTo,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("D", -1, tmpBdgYymmddTo,parent.gServerDateFormat)	    
	
	'Company Date Type 으로 변경 
	tmpBdgYymmddFr  =  UniConvDateAToB(tmpBdgYymmddFr,parent.gServerDateFormat,gDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(tmpBdgYymmddTo,parent.gServerDateFormat,gDateFormat)

	arrParam(0) = tmpBdgYymmddFr				
   	arrParam(1) = tmpBdgYymmddTo
	arrParam(2) = lgUsrIntCd                           ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value				
	arrParam(4) = "F"										' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetDept(Byval arrRet)
	Dim strStartDt, strEndDt
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.hOrgChangeId.value=arrRet(2)

		frm1.txtBdgYymmFr.Text = UniConvDateAToB(arrRet(4),parent.gDateFormat,parent.gDateFormatYYYYMM)
		frm1.txtBdgYymmTo.Text = UniConvDateAToB(arrRet(5),parent.gDateFormat,parent.gDateFormatYYYYMM)  

		frm1.txtDeptCd.focus		
End Function
'============================================================
'공통 팝업 
'============================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	
		Case "BdgCdFr", "BdgCdTo"
			arrParam(0) = "예산코드 팝업"					' 팝업 명칭 
			arrParam(1) = "F_BDG_ACCT "    						' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "예산코드"						' 조건필드의 라벨 명칭 

			arrField(0) = "BDG_CD"	     						' Field명(0)
			arrField(1) = "GP_ACCT_NM"			    			' Field명(1)
    
			arrHeader(0) = "예산코드"						' Header명(0)
			arrHeader(1) = "예산명"							' Header명(1)
	
		Case "BdgCd_Spread"
			arrParam(0) = "예산코드 팝업"					' 팝업 명칭 
			arrParam(1) = "F_BDG_ACCT A, B_MINOR B"				' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "A.BDG_CTRL_UNIT = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("F2010", "''", "S") & " "	' Where Condition
			arrParam(5) = "예산코드"						' 조건필드의 라벨 명칭 

			arrField(0) = "A.BDG_CD"							' Field명(0)
			arrField(1) = "A.GP_ACCT_NM"		    			' Field명(1)
			arrField(2) = "B.MINOR_NM"
    
			arrHeader(0) = "예산코드"						' Header명(0)
			arrHeader(1) = "예산명"							' Header명(1)
			arrHeader(2) = "통제단위"
	
		Case "DeptCd", "DeptCd_Spread"
			arrParam(0) = "부서코드 팝업"					' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT A"							' TABLE 명칭 
			arrParam(2) = strCode    							' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(parent.gChangeOrgId , "''", "S") & ""
			
			' 권한관리 추가 
			If lgInternalCd <>  "" Then
				arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD =" & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			End If

			If lgSubInternalCd <>  "" Then
				arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			End If
			
			arrParam(5) = "부서코드"			
	
			arrField(0) = "A.DEPT_CD"						' Field명(0)
			arrField(1) = "A.DEPT_NM"						' Field명(1)

			arrHeader(0) = "부서코드"						' Header명(0)
			arrHeader(1) = "부서명"						    ' Header명(1)
		
	End Select	

	IsOpenPop = True
	
	If iWhere = "BdgCd_Spread" Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			Select Case iWhere
			Case "BdgCdFr"
				.txtBdgCdFr.value = arrRet(0)
				.txtBdgNmFr.value = arrRet(1)
				.txtBdgCdFr.focus
			Case "BdgCdTo"
				.txtBdgCdTo.value = arrRet(0)
				.txtBdgNmTo.value = arrRet(1)
				.txtBdgCdTo.focus
			Case "DeptCd"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
			
			Case "BdgCd_Spread"
				.vaSpread1.Col  = C_BDG_CD
				.vaSpread1.Text = arrRet(0)
				.vaSpread1.Col  = C_BDG_NM
				.vaSpread1.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )
				
			Case "DeptCd_Spread"
			    .vspdData.Col  = C_DEPT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_DEPT_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )	 	

			End Select
		End With
	End If	

End Function

'============================================================
'부서코드 팝업 
'============================================================
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim strYear, strMonth, strDay, strDate
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	frm1.vspdData.Col = C_BDG_YYYYMM
	Call ExtractDateFrom(frm1.vspdData.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
	
	arrParam(0) = strCode				'부서코드 
	arrParam(1) = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")		'날짜(Default:현재일)
	arrParam(2) = "1"					'부서권한(lgUsrIntCd)
	
	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	End If

	With frm1
		.vspdData.Col  = C_DEPT_CD
		.vspdData.Text = arrRet(0)
		.vspdData.Col  = C_DEPT_NM
		.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_DEPT_CD, .vspdData.Row )	 	
	End With
	
End Function

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
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
			C_BDG_CD				= iCurColumnPos(1)
			C_BDG_CD_PB             = iCurColumnPos(2)
			C_BDG_NM                = iCurColumnPos(3)    
			C_BDG_YYYYMM            = iCurColumnPos(4)
			C_DEPT_CD               = iCurColumnPos(5)
			C_DEPT_PB               = iCurColumnPos(6)
			C_DEPT_NM               = iCurColumnPos(7)
			C_ORG_CHANGE_ID         = iCurColumnPos(8)
			C_BDG_CTRL_FG			= iCurColumnPos(9)			
			C_BDG_PLAN_AMT          = iCurColumnPos(10)
			C_BDG_AMT               = iCurColumnPos(11)
			C_BDG_GL_AMT            = iCurColumnPos(12)
			C_BDG_TEMP_AMT          = iCurColumnPos(13)
			C_ACCT_BDG_CTRL_FG      = iCurColumnPos(14)
			C_GP_BDG_CTRL_FG        = iCurColumnPos(15)
    End Select    
End Sub

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

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format    	
	Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)	
	Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
	
    Call SetDefaultVal    
    
    Call InitSpreadSheet                          '⊙: Setup the Spread Sheet    
    Call InitVariables                            '⊙: Initializes local global Variables
    
    '----------  Coding part  -------------------------------------------------------------
	'Call FncSetToolBar("New")
	Call SetToolbar("1100110100101111")

    frm1.txtBdgYymmFr.focus

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

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub txtBdgYymmFr_DblClick(Button)
    If Button = 1 Then
       frm1.txtBdgYymmFr.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtBdgYymmFr.Focus       
    End If
End Sub

Sub txtBdgYymmTo_DblClick(Button)
    If Button = 1 Then
       frm1.txtBdgYymmTo.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtBdgYymmTo.Focus       
    End If
End Sub


Sub txtBdgYymmFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBdgYymmTo.focus
	   Call MainQuery
	End If   
End Sub

Sub txtBdgYymmTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBdgYymmFr.focus
	   Call MainQuery
	End If   
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(Col, Row)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If	   
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Dim strSelect, strFrom, strWhere
    Dim strYear, strMonth, strDay, strDate
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	frm1.vspdData.row = Row

	Select Case Col
	Case C_DEPT_CD, C_BDG_YYYYMM
			
		frm1.vspdData.Col = C_BDG_YYYYMM
		If Trim(frm1.vspdData.Text = "") Then	Exit sub

		frm1.vspdData.Col = C_DEPT_CD
		If Trim(frm1.vspdData.Text = "") Then	Exit sub
			'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(frm1.vspdData.Text, "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"

'		' 권한관리 추가 
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If

		frm1.vspdData.Col = C_BDG_YYYYMM

		Call ExtractDateFrom(frm1.vspdData.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
		strDate = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear, strMonth, "01")

		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(strDate, "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  

			frm1.vspdData.Col = C_DEPT_CD
			frm1.vspdData.Text = ""
			frm1.vspdData.Col = C_DEPT_NM
			frm1.vspdData.Text = ""
			frm1.vspdData.Col = C_ORG_CHANGE_ID
			frm1.vspdData.Text = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
							
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.vspdData.Col = C_ORG_CHANGE_ID
				frm1.vspdData.Text = Trim(arrVal2(2))
			Next	
					
		End If
		'----------------------------------------------------------------------------------------
	End Select
	
    lgBlnFlgChgValue = True
    
End Sub

Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)

    If frm1.vspdData.MaxRows = 0 Then							'no data일 경우 vspdData_LeaveCell no 실행 
       Exit Sub													'tab이동시에 잘못된 140318 message 방지 
    End If
    
    With frm1.vspdData
		'If Col <> NewCol  And NewCol > 0 Then
		 If NewCol > 0 Then '2002.8.13 수정 
		
			If Col = C_BDG_YYYYMM Then
				.Row = Row
				.Col = Col
			
				If .Text <> "" Then
                    If CheckDateFormat(.Text, parent.gDateFormatYYYYMM) = False  Then
						Call DisplayMsgBox("140318","X","X","X")	'년월을 올바로 입력하세요.
						.Text = ""
					End If
				End If
			End If
		
		End If

'		If Row >= NewRow Then
'		    Exit Sub
'		End If
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
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
 '----------  Coding part  -------------------------------------------------------------   
    If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevBdgCdKey <> "" and lgStrPrevBdgYMKey <> "" and lgStrPrevDeptCdKey <> "" Then	
          Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
       End If
    End if
        
End Sub

'==========================================================================================
' Event Name : vspdData_ButtonClicked
' Event Desc : 버튼 컬럼을 클릭할 경우 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	'---------- Coding part -------------------------------------------------------------
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		
		IF Row > 0 And Col = C_BDG_CD_PB Then
			.Col = Col
			.Row = Row
			Call OpenPopup(.Text, "BdgCd_Spread")
	    		
		ElseIf Row > 0 and Col = C_DEPT_PB Then
	        .Col = Col
			.Row = Row
			Call OpenPopupDept(.Text, "DeptCd_Spread")
		
		End If
		
	End With
	
End Sub

Sub txtBdgCdFr_onChange()
	frm1.txtBdgNmFr.value = ""
End Sub

Sub txtBdgCdTo_onChange()
	frm1.txtBdgNmTo.value = ""
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCD_OnChange()        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
    
    tmpBdgYymmddFr = UniConvDateAToB(frm1.txtBdgYymmFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(frm1.txtBdgYymmTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("M", +1, tmpBdgYymmddTo,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("D", -1, tmpBdgYymmddTo,parent.gServerDateFormat)			
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddFr , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddTo , "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
				'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

'		' 권한관리 추가 
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
		End If
	End IF		
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
	Dim strFrYear, strFrMonth, strFrDay 
	Dim strToYear, strToMonth, strToDay
    
    FncQuery = False          '⊙: Processing is NG
    Err.Clear                 '☜: Protect system from crashing
	
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		if IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call InitVariables							  '⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
	If Not chkField(Document, "1") Then	'⊙: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtBdgYymmFr.Text, frm1.txtBdgYymmTo.Text, frm1.txtBdgYymmFr.Alt, frm1.txtBdgYymmTo.Alt, _
						"970025", frm1.txtBdgYymmFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtBdgYymmFr.focus														'⊙: GL Date Compare Common Function
			Exit Function
	End if
    
    Call ExtractDateFrom(frm1.txtBdgYymmFr.Text,frm1.txtBdgYymmFr.UserDefinedFormat,parent.gComDateType,strFrYear,strFrMonth,strFrDay)    
    strFrDt = strFrYear & strFrMonth
        
    Call ExtractDateFrom(frm1.txtBdgYymmTo.Text,frm1.txtBdgYymmTo.UserDefinedFormat,parent.gComDateType,strToYear,strToMonth,strToDay)
    strToDt = strToYear & strToMonth
       
    frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
    frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
    
    If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus 
			Exit Function
		End If
    End If
    
    IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("800600","X",frm1.txtBdgYymmFr.alt,"X")            '⊙: Display Message(There is no changed data.)
		Exit Function
	End if

	With frm1
		.txtDeptCd.value  = UCase(Trim(.txtDeptCd.value))
		.txtBdgCdFr.value = UCase(Trim(.txtBdgCdFr.value))
		.txtBdgCdTo.value = UCase(Trim(.txtBdgCdTo.value))
	End With

	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery	    
																				'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
    FncNew = False                  '⊙: Processing is NG
    Err.Clear                       '☜: Protect system from crashing
    'On Error Resume Next            '☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ' 변경된 내용이 있는지 확인한다.
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015",parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")     '⊙: Clear Condition Field	
    Call InitVariables                         '⊙: Initializes local global variables
    Call SetDefaultVal
    
    Call FncSetToolBar("New")
    
    'SetGridFocus
    FncNew = True                              '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
    
    FncDelete = False            '⊙: Processing is NG
    Err.Clear                    '☜: Protect system from crashing
    'On Error Resume Next        '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ' Update 상태인지를 확인한다.
    If lgIntFlgMode <> parent.OPMD_UMODE Then        'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then											  '☜: Delete db data
       Exit Function                        
    End If
    
    '-----------------------
    'Erase condition area
    '-----------------------
	Call ggoOper.ClearField(Document, "1")								  '⊙: Clear Condition Field
    FncDelete = True													 '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False            '⊙: Processing is NG
    Err.Clear                  '☜: Protect system from crashing
    'On Error Resume Next       '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '☜: Save db data

	 FncSave = True                                                           '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
			SetSpreadColor .ActiveRow, .ActiveRow
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 
    'Call SetSpreadLock(frm1.vspdData.ActiveRow, "Insert")


	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

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
        for imRow2 = 1 to imRow 
            ggoSpread.Source = .vspdData
            ggoSpread.InsertRow ,1
            SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow 
            .vspdData.col = C_BDG_YYYYMM
            .vspdData.Row = .vspdData.ActiveRow
			.vspdData.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormatYYYYMM) 			'예산일자 default : today	
			'Call ggoOper.FormatDate(frm1.txtBdgYymmFr, popupparent.gDateFormat, 2)
            '.vspdData.Text= UNIMonthClientFormat("<%=GetSvrDate%>")			'예산일자 default : today	
        Next
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	    
    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With

End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call Parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
	Call ggoSpread.ReOrderingSpreadData()
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
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

	Call LayerShowHide(1)
    
    DbQuery = False
    Err.Clear                '☜: Protect system from crashing
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtBdgCdFr=" & Trim(.htxtBdgCdFr.value)						'조회 조건 데이타 
			strVal = strVal & "&txtBdgCdTo=" & Trim(.htxtBdgCdTo.value)						'조회 조건 데이타 
			strVal = strVal & "&txtBdgYymmFr=" & strFrDt 'Trim(.htxtBdgYymmFr.value )		'조회 조건 데이타 
			strVal = strVal & "&txtBdgYymmTo=" & strToDt 'Trim(.htxtBdgYymmTo.value )		'조회 조건 데이타 
			strVal = strVal & "&txtDeptCd=" & Trim(.htxtDeptCd.value)						'조회 조건 데이타 
			strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&lgStrPrevBdgCdKey=" & lgStrPrevBdgCdKey
			strVal = strVal & "&lgStrPrevBdgYMKey=" & lgStrPrevBdgYMKey
			strVal = strVal & "&lgStrPrevDeptCdKey=" & lgStrPrevDeptCdKey		
	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtBdgCdFr=" & Trim(.txtBdgCdFr.value)			'조회 조건 데이타 
			strVal = strVal & "&txtBdgCdTo=" & Trim(.txtBdgCdTo.value)			'조회 조건 데이타 
			strVal = strVal & "&txtBdgYymmFr=" & strFrDt
			strVal = strVal & "&txtBdgYymmTo=" & strToDt
			strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.value)			'조회 조건 데이타 
			strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&lgStrPrevBdgCdKey=" & lgStrPrevBdgCdKey
			strVal = strVal & "&lgStrPrevBdgYMKey=" & lgStrPrevBdgYMKey
			strVal = strVal & "&lgStrPrevDeptCdKey=" & lgStrPrevDeptCdKey		
			'strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
			strVal = strVal & "&lgPageNo=" & lgPageNo

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd				' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd					' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd				' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID					' 개인 

	    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	Call SetSpreadLock()'(-1, "Query")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE	'⊙: Indicates that current mode is Update mode
    
	' 현재 Page의 From Element들을 사용자가 입력을 받지 못하게 하거나 필수입력사항을 표시한다.	
    Call ggoOper.LockField(Document, "Q")	'⊙: This function lock the suitable field
    Call FncSetToolBar("Query")
    
    'SetGridFocus        
    Set gActiveElement = document.activeElement 
    
End Function

'========================================================================================
' Function Name : DbSave()
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal,strDel
	Dim strYear,strMonth,strDay
	Dim iColSep
	
	Call LayerShowHide(1)
	
    DbSave = False				'⊙: Processing is NG
    'On Error Resume Next		'☜: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""
		iColSep = Parent.gColSep
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
'			C_BDG_CD				= iCurColumnPos(1)
'			C_BDG_CD_PB             = iCurColumnPos(2)
'			C_BDG_NM                = iCurColumnPos(3)    
'			C_BDG_YYYYMM            = iCurColumnPos(4)
'			C_DEPT_CD               = iCurColumnPos(5)
'			C_DEPT_PB               = iCurColumnPos(6)
'			C_DEPT_NM               = iCurColumnPos(7)
'			C_ORG_CHANGE_ID         = iCurColumnPos(8)
'			C_BDG_CTRL_FG			= iCurColumnPos(9)			
'			C_BDG_PLAN_AMT          = iCurColumnPos(10)
'			C_BDG_AMT               = iCurColumnPos(11)
'			C_BDG_GL_AMT            = iCurColumnPos(12)
'			C_BDG_TEMP_AMT          = iCurColumnPos(13)
'			C_ACCT_BDG_CTRL_FG      = iCurColumnPos(14)
'			C_GP_BDG_CTRL_FG        = iCurColumnPos(15)

		    Select Case .vspdData.Text
		    
  				Case ggoSpread.InsertFlag												'☜: 신규 

					strVal = strVal & "C" & iColSep & lRow & iColSep					'☜: U=Create
				    .vspdData.Col = C_BDG_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_YYYYMM
		            
		            Call ExtractDateFrom(.vspdData.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
		            'strVal = strVal & Trim(.vspdData.Text) & iColSep
		            strVal = strVal & strYear & strMonth & iColSep
		            .vspdData.Col = C_DEPT_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ORG_CHANGE_ID						'Hidden Column 이므로 입력값 없음.
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_PLAN_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_GL_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            
		            .vspdData.Col = C_BDG_TEMP_AMT 
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            
		            .vspdData.Col = C_BDG_CTRL_FG
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1

				Case ggoSpread.UpdateFlag												'☜: 수정 

					strVal = strVal & "U" & iColSep & lRow & iColSep					'☜: U=Update
				    .vspdData.Col = C_BDG_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_YYYYMM
		            'strVal = strVal & Trim(.vspdData.Text) & iColSep
		              Call ExtractDateFrom(.vspdData.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
		            strVal = strVal & strYear & strMonth & iColSep
		            
		            .vspdData.Col = C_DEPT_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ORG_CHANGE_ID
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_PLAN_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_GL_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_TEMP_AMT
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            '.vspdData.Col = C_BDG_CTRL_FG
		            'strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            
		            
		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag												'☜: 삭제 

					strDel = strDel & "D" & iColSep & lRow & iColSep					'☜: U=Delete
				    .vspdData.Col = C_BDG_CD
		            strDel = strDel & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_YYYYMM
		              Call ExtractDateFrom(.vspdData.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
		            strDel = strDel & strYear & strMonth & iColSep
		            
		            .vspdData.Col = C_DEPT_CD
		            strDel = strDel & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ORG_CHANGE_ID
		            strDel = strDel & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_PLAN_AMT
		            strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_AMT
					strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            
		            .vspdData.Col = C_BDG_GL_AMT
		             strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		            .vspdData.Col = C_BDG_TEMP_AMT		            
		            strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            
					'.vspdData.Col = C_BDG_CTRL_FG
		            'strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            
		            
		            lGrpCnt = lGrpCnt + 1
		            
		    End Select
			            
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		
		
		 Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'☜: 비지니스 ASP 를 가동 
	
	End With

    DbSave = True                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
    
    Call InitVariables
	'frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

  
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	On Error Resume Next
End Function

'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################

Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100110100001111")
	Case "QUERY"
		Call SetToolbar("1100111100111111")
	End Select
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

'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo
 
	tmpBdgYymmddFr = UniConvDateAToB(frm1.txtBdgYymmFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(frm1.txtBdgYymmTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("M", +1, tmpBdgYymmddTo,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("D", -1, tmpBdgYymmddTo,parent.gServerDateFormat)	
	
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddFr , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddTo , "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

		' 권한관리 추가 
		If lgInternalCd <> "" Then
			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
		End If
	
		If lgSubInternalCd <> "" Then
			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
		End If


			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)			
								
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With	

End Function

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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>예산년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYymmFr" CLASS=FPDTYYYYMM tag="12XXXU" Title="FPDATETIME" ALT=시작예산년월 id=fpBdgYymmFr></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYymmTo" CLASS=FPDTYYYYMM tag="12XXXU" Title="FPDATETIME" ALT=종료예산년월 id=fpBdgYymmTo></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" MAXLENGTH="10" SIZE=10 ALT ="부서코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">&nbsp;<INPUT NAME="txtDeptNm" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="부서명" tag="24X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>시작예산</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCdFr" MAXLENGTH="18" SIZE=10 ALT ="시작예산코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBdgCdFr.Value, 'BdgCdFr')">&nbsp;<INPUT NAME="txtBdgNmFr" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="시작예산명" tag="24X">
									</TD>
									<TD CLASS="TD5" NOWRAP>종료예산</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCdTo" MAXLENGTH="18" SIZE=10 ALT ="종료예산코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBdgCdTo.Value, 'BdgCdTo')">&nbsp;<INPUT NAME="txtBdgNmTo" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="종료예산명" tag="24X">
									</TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtBdgCdFr"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtBdgCdTo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtBdgYymmFr"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtBdgYymmTo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtDeptCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="14" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

