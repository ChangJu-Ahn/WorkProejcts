
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : VAT
'*  3. Program ID           : a6117ma1
'*  4. Program Name         : 부가세수정 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004.05.10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Eun Kyung , KANG
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  --><!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css"><!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a6117mb1.asp"'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns

Dim C_VAT_NO              
Dim C_ISSUED_DT
Dim C_IO_FG            
Dim C_BP_CD 
Dim C_BP_PB          
Dim C_BP_NM 
DIM C_REG_NO    
Dim C_MADE_VAT_FG              
Dim C_VAT_TYPE      
Dim C_VAT_TYPE_NM  
Dim C_VAT_TYPE_PB     
Dim C_NET_LOC_AMT        
Dim C_VAT_LOC_AMT         
Dim C_CARD_NO 
Dim C_CARD_PB              
Dim C_REPORT_BIZ_AREA_CD  
Dim C_REPORT_BIZ_AREA_PB
Dim C_BIZ_AREA_CD   
Dim C_BIZ_AREA_PB
Dim C_GL_NO     
Dim C_TEMP_GL_NO


Dim C_issue_dt_fg_cd
Dim C_issue_dt_fg_nm
Dim C_issue_dt_kind_cd
Dim C_issue_dt_kind_nm

Const C_SHEETMAXROWS = 100

 '==========================================  1.2.2 Global 변수 선언  =====================================
'1. 변수 표준에 따름. prefix로 g를 사용함.
'2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
'Dim lgLngCurRows

Dim lgStrPrevVatKey

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
'Dim lgStrComDateType'Company Date Type을 저장(년월 Mask에 사용함.)

 '#########################################################################################################
'2. Function부 
'
'내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'           2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

   C_VAT_NO                = 1     
	C_ISSUED_DT             = 2  
	C_IO_FG                 = 3  
	C_BP_CD                 = 4  
	C_BP_PB                 = 5  
	C_BP_NM                 = 6  
	C_REG_NO                = 7  
	C_MADE_VAT_FG           = 8  
	C_VAT_TYPE              = 9  
	C_VAT_TYPE_NM           = 10
	C_VAT_TYPE_PB           = 11
	C_issue_dt_fg_cd        = 12
	C_issue_dt_fg_nm        = 13
	C_issue_dt_kind_cd      = 14
	C_issue_dt_kind_nm      = 15
	C_NET_LOC_AMT           = 16
	C_VAT_LOC_AMT           = 17
	C_CARD_NO               = 18
	C_CARD_PB               = 19
	C_REPORT_BIZ_AREA_CD    = 20
	C_REPORT_BIZ_AREA_PB    = 21
	C_BIZ_AREA_CD           = 22
	C_BIZ_AREA_PB           = 23
	C_GL_NO                 = 24
	C_TEMP_GL_NO            = 25
     					
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$




 '==========================================  2.1.1 InitVariables()  ======================================
'Name : InitVariables()
'Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevVatKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgPageNo  = 0
End Sub

 '******************************************  2.2 화면 초기화 함수  ***************************************
'기능: 화면초기화 
'설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'Name : SetDefaultVal()
'Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
Dim strSvrDate
Dim strYear, strMonth, strDay,  EndDate, StartDate

    strSvrDate = "<%=GetSvrDate%>"

	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, parent.gDateFormat)

	frm1.txtIssuedDtFr.Text = StartDate
	frm1.txtIssuedDtTo.Text = EndDate

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
	ggoSpread.Spreadinit "V20091212",,parent.gAllowDragDropSpread


    With frm1.vspdData
		
        .ReDraw = False
        '.ColHidden = True
        
       .MaxCols = C_TEMP_GL_NO + 1
        '.Col = .MaxCols'☜: 공통콘트롤 사용 Hidden Column

        .MaxRows = 0
                           'patch version
        Call GetSpreadColumnPos("A")
        
        ggoSpread.SSSetEdit      C_VAT_NO,             "계산서번호",   18, 3
        ggoSpread.SSSetDate      C_ISSUED_DT,          "발행일",       10, 2, parent.gDateFormat
        ggoSpread.SSSetEdit      C_IO_FG,              "입출구분",     8, 3
        ggoSpread.SSSetEdit      C_BP_CD,              "거래처코드",   10, 3
        ggoSpread.SSSetButton    C_BP_PB
        ggoSpread.SSSetEdit      C_BP_NM,              "거래처명",     20, 3
        ggoSpread.SSSetEdit      C_REG_NO,              "사업자번호",     12, 3
        
        ggoSpread.SSSetEdit      C_MADE_VAT_FG,        "부가세구분",   2, 3
        ggoSpread.SSSetEdit      C_VAT_TYPE,           "",                 2, 3        
        ggoSpread.SSSetEdit      C_VAT_TYPE_NM,        "부가세유형",   15, 3
        ggoSpread.SSSetButton    C_VAT_TYPE_PB
        ggoSpread.SSSetCombo     C_issue_dt_fg_cd,  "전자세금계산서발행여부",	15
		ggoSpread.SSSetCombo     C_issue_dt_fg_nm,  "전자세금계산서발행여부",	15    
        ggoSpread.SSSetCombo     C_issue_dt_kind_cd,  "전자세금계산서종류",	15
		ggoSpread.SSSetCombo     C_issue_dt_kind_nm,  "전자세금계산서종류",	15        
        ggoSpread.SSSetFloat     C_NET_LOC_AMT,        "공급가액",     20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
        ggoSpread.SSSetFloat     C_VAT_LOC_AMT,        "부가세액",     20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
        ggoSpread.SSSetEdit      C_CARD_NO,            "카드번호",     20, 3
        ggoSpread.SSSetButton    C_CARD_PB
        ggoSpread.SSSetEdit      C_REPORT_BIZ_AREA_CD, "신고사업장",   10, 3
        ggoSpread.SSSetButton    C_REPORT_BIZ_AREA_PB
        ggoSpread.SSSetEdit      C_BIZ_AREA_CD ,       "발생사업장",   10, 3
        ggoSpread.SSSetButton    C_BIZ_AREA_PB
        ggoSpread.SSSetEdit      C_GL_NO,              "전표번호",     10, 5
        ggoSpread.SSSetEdit      C_TEMP_GL_NO,         "참조번호", 10, 5

		
        

        
		Call ggoSpread.MakePairsColumn(C_BP_CD,              C_BP_PB              ,"1")
		Call ggoSpread.MakePairsColumn(C_VAT_TYPE,           C_VAT_TYPE_PB        ,"1")
		Call ggoSpread.MakePairsColumn(C_CARD_NO           , C_CARD_PB            ,"1")
		Call ggoSpread.MakePairsColumn(C_REPORT_BIZ_AREA_CD, C_REPORT_BIZ_AREA_PB ,"1")
		Call ggoSpread.MakePairsColumn(C_BIZ_AREA_CD,        C_BIZ_AREA_PB        ,"1")
		Call ggoSpread.MakePairsColumn(C_issue_dt_kind_cd, C_issue_dt_kind_nm, "1")
		Call ggoSpread.MakePairsColumn(C_issue_dt_fg_cd, C_issue_dt_fg_nm, "1")

		Call ggoSpread.SSSetColHidden(C_issue_dt_kind_cd, C_issue_dt_kind_cd, True)
		Call ggoSpread.SSSetColHidden(C_issue_dt_fg_cd, C_issue_dt_fg_cd, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        Call ggoSpread.SSSetColHidden(C_VAT_TYPE,C_VAT_TYPE,True)
        Call ggoSpread.SSSetColHidden(C_MADE_VAT_FG,C_MADE_VAT_FG,True)
		
		
        .ReDraw = True

    
    End With
    
    Call SetSpreadLock     
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock()

' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
        .ReDraw = False

        'ggoSpread.SSSetRequired C_BDG_PLAN_AMT, -1,-1
        ggoSpread.SpreadLock C_VAT_NO,      -1, C_VAT_NO
        ggoSpread.SpreadLock C_IO_FG,       -1, C_IO_FG
        ggoSpread.SpreadLock C_BP_NM,       -1, C_BP_NM      ' ,-1
        ggoSpread.SpreadLock C_REG_NO,       -1, C_REG_NO      ' ,-1
        ggoSpread.SpreadLock C_GL_NO,       -1, C_GL_NO
        ggoSpread.SpreadLock C_TEMP_GL_NO,  -1, C_TEMP_GL_NO       ',-1
        
        ggoSpread.SSSetProtected .MaxCols,  -1, -1
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
        ggoSpread.SSSetProtected  C_VAT_NO,              pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_ISSUED_DT,           pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_IO_FG,               pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_BP_CD,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_BP_NM,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_REG_NO,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_MADE_VAT_FG,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_VAT_TYPE_NM,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_NET_LOC_AMT,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_VAT_LOC_AMT,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_CARD_NO,             pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_REPORT_BIZ_AREA_CD,  pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_BIZ_AREA_CD,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_GL_NO,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_TEMP_GL_NO,          pvStartRow, pvEndRow

        .vspdData.ReDraw = True
    
    End With

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'Name : InitComboBox()
'Description : Combo Display
'========================================================================================================= 



Sub InitComboBox()

    Dim arrData
 
   ggoSpread.Source = frm1.vspdData

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("DT004", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_issue_dt_kind_cd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_issue_dt_kind_nm
 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1020", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_issue_dt_fg_cd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_issue_dt_fg_nm

    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1003", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboIoFg ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_kind ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_fg ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_kind2 ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_fg2 ,lgF0  ,lgF1  ,Chr(11))
    
    
End Sub



'========================================== 2.4.2 Open???()  =============================================
'Name : Open???()
'Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
'--------------------------------------------------------------------------------------------------------- 
'   Function Name : OpenVatNoInfo()
'   Function Desc : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenVatNoInfo(Byval strCode, Byval Cond)
	Dim iCalledAspName
	Dim arrRet
		
	If IsOpenPop = True Then Exit Function	

	iCalledAspName = AskPRAspName("a6114ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a6114ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	     
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVatNo.focus
		Exit Function
	Else
		Call SetVatNoInfo(arrRet,Cond)	
	End If	
End Function

'--------------------------------------------------------------------------------------------------------- 
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'--------------------------------------------------------------------------------------------------------- 
Function SetVatNoInfo(Byval arrRet, Byval Cond)
	Select Case Cond
		Case "VatNo"
			frm1.txtVatNo.focus
			frm1.txtVatNo.Value	= arrRet(0)
	End Select	
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'Name : OpenBp()
'Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
    Dim arrRet
    Dim arrParam(5)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
        
    arrParam(0) = strCode' Code Condition
    arrParam(1) = ""' 채권과 연계(거래처 유무)
    arrParam(2) = ""' FrDt
    arrParam(3) = ""' ToDt
    arrParam(4) = "T"' B :매출 S: 매입 T: 전체 
    arrParam(5) = ""' SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 

    arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
    "dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        frm1.txtBpCd.focus
        Exit Function
    Else
	    frm1.txtBpCd.focus
	    frm1.txtBpCd.Value    = arrRet(0)		
    	frm1.txtBpNm.Value    = arrRet(1)		
        lgBlnFlgChgValue = True
    End If
End Function

'=======================================================================================================
'    Name : OpenReportBizArea()
'    Description : Bp Cd PopUp
'=======================================================================================================
Function OpenReportBizArea()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    
    If IsOpenPop = True  Then Exit Function

    IsOpenPop = True

    arrParam(0) = "세금신고사업장 팝업"                    ' 팝업 명칭 
    arrParam(1) = "B_TAX_BIZ_AREA"                        ' TABLE 명칭 
    arrParam(2) = Trim(frm1.txtReportBizArea.Value)
    arrParam(3) = ""
    arrParam(4) = ""            
    arrParam(5) = "세금신고사업장코드"                    '조건필드의 라벨 명칭 
    
    arrField(0) = "TAX_BIZ_AREA_CD"                               ' Field명(0)
    arrField(1) = "TAX_BIZ_AREA_NM"                               ' Field명(1)
    
    arrHeader(0) = "세금신고사업장코드"                       ' Header명(0)
    arrHeader(1) = "세금신고사업장명"                       ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        frm1.txtReportBizArea.focus    
        Exit Function
    Else
        Call SetReportBizArea(arrRet)
    End If    
End Function

'=======================================================================================================
'    Name : SetReportBizArea()
'    Description : Bp Cd Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetReportBizArea(byval arrRet)
    frm1.txtReportBizArea.focus    
    frm1.txtReportBizArea.Value    = arrRet(0)        
    frm1.txtReportBizAreaNm.Value    = arrRet(1)        
    lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'    Name : OpenVatType()
'    Description : Bp Cd PopUp
'=======================================================================================================
Function OpenVatType()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
      
    If IsOpenPop = True  Then Exit Function

    IsOpenPop = True
    arrParam(0) = "부가세유형팝업"                    ' 팝업 명칭 
    arrParam(1) = "B_MINOR"                                ' TABLE 명칭 
    arrParam(2) = Trim(frm1.txtVatType.Value)
    arrParam(3) = ""
    arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "            
    arrParam(5) = "부가세코드"                    '조건필드의 라벨 명칭 
    
    arrField(0) = "MINOR_CD"                               ' Field명(0)
    arrField(1) = "MINOR_NM"                               ' Field명(1)
    
    arrHeader(0) = "부가세유형"                       ' Header명(0)
    arrHeader(1) = "부가세유형명"                       ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    If arrRet(0) = "" Then
        frm1.txtVatType.focus
        Exit Function
    Else
        Call SetVatType(arrRet)
    End If   
End Function

'=======================================================================================================
'    Name : SetVatType()
'    Description :
'=======================================================================================================
Function SetVatType(byval arrRet)
    frm1.txtVatType.focus
    frm1.txtVatType.Value   = arrRet(0)        
    frm1.txtVatTypeNm.Value = arrRet(1)        
    lgBlnFlgChgValue = True
End Function


'============================================================
'공통 팝업 
'============================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    Select Case iWhere
    
        Case "BpCd_Spread"
           arrParam(0) = strCode                                ' Code Condition
           arrParam(1) = ""                                     ' 채권과 연계(거래처 유무)
           arrParam(2) = ""                                     ' FrDt
           arrParam(3) = ""                                     ' ToDt
           arrParam(4) = "T"                                    ' B :매출 S: 매입 T: 전체 
           arrParam(5) = ""                                     ' SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 

        Case "VatType_Spread"
            arrParam(0) = "부가세유형팝업"                 ' 팝업 명칭 
            arrParam(1) = "B_MINOR "                            ' TABLE 명칭 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "                    ' Where Condition
            arrParam(5) = "부가세유형"                  ' 조건필드의 라벨 명칭 

            arrField(0) = "MINOR_CD"                            ' Field명(0)
            arrField(1) = "MINOR_NM"                            ' Field명(1)
    
            arrHeader(0) = "부가세유형"                     ' Header명(0)
            arrHeader(1) = "부가세유형명"                   ' Header명(1)
    
        Case "CardCd_Spread"
            arrParam(0) = "신용카드 팝업"                   ' 팝업 명칭 
            arrParam(1) = "B_CREDIT_CARD"                       ' TABLE 명칭 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""                                    ' Where Condition
            arrParam(5) = "신용카드"                        ' 조건필드의 라벨 명칭 

            arrField(0) = "CREDIT_NO"                           ' Field명(0)
            arrField(1) = "CREDIT_NM"                           ' Field명(1)
    
            arrHeader(0) = "신용카드번호"                   ' Header명(0)
            arrHeader(1) = "신용카드명"                     ' Header명(1)
    
        Case "ReportBizAreaCd_Spread"
            arrParam(0) = "세금신고사업장 팝업"             ' 팝업 명칭 
            arrParam(1) = "B_TAX_BIZ_AREA"                      ' TABLE 명칭 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""
            arrParam(5) = "세금신고사업장코드"            
    
            arrField(0) = "TAX_BIZ_AREA_CD"                     ' Field명(0)
            arrField(1) = "TAX_BIZ_AREA_NM"                     ' Field명(1)

            arrHeader(0) = "세금신고사업장코드"             ' Header명(0)
            arrHeader(1) = "세금신고사업장명"               ' Header명(1)
 
         Case "BizAreaCd_Spread"
            arrParam(0) = "사업장 팝업"                     ' 팝업 명칭 
            arrParam(1) = "B_BIZ_AREA"                          ' TABLE 명칭 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""
            arrParam(5) = "사업장코드"            
    
            arrField(0) = "BIZ_AREA_CD"                          ' Field명(0)
            arrField(1) = "BIZ_AREA_NM"                          ' Field명(1)

            arrHeader(0) = "사업장코드"                     ' Header명(0)
            arrHeader(1) = "사업장명"                       ' Header명(1)
       
    End Select    

    IsOpenPop = True
    
    If iWhere = "BpCd_Spread" Then
       arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
           "dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
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
            Case "BpCd_Spread"
                .vaSpread1.Col  = C_BP_CD
                .vaSpread1.Text = arrRet(0)
                .vaSpread1.Col  = C_BP_NM
                .vaSpread1.Text = arrRet(1)
                '.vaSpread1.Col  = C_REG_NO
                '.vaSpread1.Text = arrRet(2)

                Call vspdData_Change(.vspdData.Col,.vspdData.Row )

            Case "VatType_Spread"
                .vaSpread1.Col  = C_VAT_TYPE
                .vaSpread1.Text = arrRet(0)
                .vaSpread1.Col  = C_VAT_TYPE_NM
                .vaSpread1.Text = arrRet(1)
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )

            Case "CardCd_Spread"
                .vaSpread1.Col  = C_CARD_NO
                .vaSpread1.Text = arrRet(0)            
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )

            Case "ReportBizAreaCd_Spread"
                .vaSpread1.Col  = C_REPORT_BIZ_AREA_CD
                .vaSpread1.Text = arrRet(0)
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )
                
            Case "BizAreaCd_Spread"
                .vspdData.Col  = C_BIZ_AREA_CD
                .vspdData.Text = arrRet(0)
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )         

            End Select
        End With
    End If    

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
            		C_VAT_NO                = iCurColumnPos(1)      
				C_ISSUED_DT             = iCurColumnPos(2)   
				C_IO_FG                 = iCurColumnPos(3)   
				C_BP_CD                 = iCurColumnPos(4)   
				C_BP_PB                 = iCurColumnPos(5)   
				C_BP_NM                 = iCurColumnPos(6)   
				C_REG_NO                = iCurColumnPos(7)   
				C_MADE_VAT_FG           = iCurColumnPos(8)   
				C_VAT_TYPE              = iCurColumnPos(9)   
				C_VAT_TYPE_NM           = iCurColumnPos(10)  
				C_VAT_TYPE_PB           = iCurColumnPos(11)  
				C_issue_dt_fg_cd        = iCurColumnPos(12)  
				C_issue_dt_fg_nm        = iCurColumnPos(13)  
				C_issue_dt_kind_cd      = iCurColumnPos(14)  
				C_issue_dt_kind_nm      = iCurColumnPos(15)  
				C_NET_LOC_AMT           = iCurColumnPos(16)  
				C_VAT_LOC_AMT           = iCurColumnPos(17)  
				C_CARD_NO               = iCurColumnPos(18)  
				C_CARD_PB               = iCurColumnPos(19)  
				C_REPORT_BIZ_AREA_CD    = iCurColumnPos(20)  
				C_REPORT_BIZ_AREA_PB    = iCurColumnPos(21)  
				C_BIZ_AREA_CD           = iCurColumnPos(22)  
				C_BIZ_AREA_PB           = iCurColumnPos(23)  
				C_GL_NO                 = iCurColumnPos(24)  
				C_TEMP_GL_NO            = iCurColumnPos(25)  
			
            
    End Select    
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'                                                3. Event부 
'    기능: Event 함수에 관한 처리 
'    설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
 '******************************************  3.1 Window 처리  *********************************************
'    Window에 발생 하는 모든 Even 처리    
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'    Name : Form_Load()
'    Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format        
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    Call SetDefaultVal
	 
    Call InitSpreadSheet                          '⊙: Setup the Spread Sheet    
    Call InitComboBox
    Call InitVariables     
    

    '----------  Coding part  -------------------------------------------------------------
    'Call FncSetToolBar("New")
    Call SetToolbar("1100100100101111")

    frm1.txtIssuedDtFr.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'    Document의 TAG에서 발생 하는 Event 처리    
'    Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'    Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'    Window에 발생 하는 모든 Even 처리    
'********************************************************************************************************* 

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

Function txtReportBizArea_onblur()
    If frm1.txtReportBizArea.value = "" Then
        frm1.txtReportBizAreaNm.value = ""
    End If
End Function

Function txtBpCd_onblur()
    If frm1.txtBpCd.value = "" Then
        frm1.txtBpNm.value = ""
    End If
End Function

'=======================================================================================================
'   Event Name : txtIssuedDtFr_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssuedDtFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDtFr.Action = 7
          Call SetFocusToDocument("M")
        frm1.txtIssuedDtFr.Focus
    End If
End Sub

Sub txtIssuedDtTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDtTo.Action = 7
          Call SetFocusToDocument("M")
        frm1.txtIssuedDtTo.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssuedDtFr_Keypress(Key)
'   Event Desc : 조회을 한다.
'=======================================================================================================
Sub txtIssuedDtFr_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssuedDtTo.focus
        FncQuery()
    End If
End Sub

Sub txtIssuedDtTo_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssuedDtFr.focus
        FncQuery()
    End If
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_issue_dt_kind_nm
				.Col = Col
				intIndex = .Value
				.Col = C_issue_dt_kind_cd
				.Value = intIndex
			Case  C_issue_dt_fg_nm
				.Col = Col
				intIndex = .Value
				.Col = C_issue_dt_fg_cd
				.Value = intIndex
		End Select
	End With
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
    gMouseClickStatus = "SPC"    'Split 상태코드 
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
    Dim intIndex

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    frm1.vspdData.row = Row

    Select Case Col
        Case  C_issue_dt_kind_nm
            frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_kind_cd
            frm1.vspdData.Value = intIndex
            
        Case  C_issue_dt_kind_cd
              frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_kind_nm
            frm1.vspdData.Value = intIndex
            
        Case  C_issue_dt_fg_nm
            frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_fg_cd
            frm1.vspdData.Value = intIndex
        
        Case  C_issue_dt_fg_cd
              frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_fg_nm
            frm1.vspdData.Value = intIndex
                 
    End Select
    
    lgBlnFlgChgValue = True
    
End Sub

Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)

    If frm1.vspdData.MaxRows = 0 Then                            'no data일 경우 vspdData_LeaveCell no 실행 
       Exit Sub                                                    'tab이동시에 잘못된 140318 message 방지 
    End If
    
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
        If lgStrPrevVatKey <> "" then  
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
        
        IF Row > 0 And Col = C_VAT_TYPE_PB Then
            .Col = C_VAT_TYPE
            .Row = Row
            Call OpenPopup(.Text, "VatType_Spread")
                
        ElseIf Row > 0 and Col = C_CARD_PB Then
            .Col = C_CARD_NO
            .Row = Row
            Call OpenPopup(.Text, "CardCd_Spread")

        ElseIf Row > 0 and Col = C_REPORT_BIZ_AREA_PB Then
            .Col = C_REPORT_BIZ_AREA_CD
            .Row = Row
            Call OpenPopup(.Text, "ReportBizAreaCd_Spread")

        ElseIf Row > 0 and Col = C_BIZ_AREA_PB Then
            .Col = C_BIZ_AREA_CD
            .Row = Row
            Call OpenPopup(.Text, "BizAreaCd_Spread")
        
        ElseIf Row > 0 and Col = C_BP_PB Then
            .Col = C_BP_CD
            .Row = Row
            Call OpenPopup(.Text, "BpCd_Spread")
        
        End If
        
    End With
    
End Sub


'#########################################################################################################
'                                                4. Common Function부 
'    기능: Common Function
'    설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 

'#########################################################################################################
'                                                5. Interface부 
'    기능: Interface
'    설명: 각각의 Toolbar에 대한 처리를 행한다. 
'          Toolbar의 위치순서대로 기술하는 것으로 한다. 
'    << 공통변수 정의 부분 >>
'     공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'                통일하도록 한다.
'     1. 공통컨트롤을 Call하는 변수 
'           ADF (ADS, ADC, ADF는 그대로 사용)
'           - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
'     2. 공통컨트롤에서 Return된 값을 받는 변수 
'            strRetMsg
'######################################################################################################### 

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'    설명 : Fnc함수명 으로 시작하는 모든 Function
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
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")            '⊙: "Will you destory previous data"
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

    Call InitVariables                              '⊙: Initializes local global variables
   
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then    '⊙: This function check indispensable field
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtIssuedDtFr.Text, frm1.txtIssuedDtTo.Text, frm1.txtIssuedDtFr.Alt, frm1.txtIssuedDtTo.Alt, _
                        "970025", frm1.txtIssuedDtFr.UserDefinedFormat, parent.gComDateType, true) = False Then
            frm1.txtBdgYymmFr.focus                                                        '⊙: GL Date Compare Common Function
            Exit Function
    End if

    Call ExtractDateFrom(frm1.txtIssuedDtFr.Text,frm1.txtIssuedDtFr.UserDefinedFormat,parent.gComDateType,strFrYear,strFrMonth,strFrDay)    
        
    Call ExtractDateFrom(frm1.txtIssuedDtTo.Text,frm1.txtIssuedDtTo.UserDefinedFormat,parent.gComDateType,strToYear,strToMonth,strToDay)
     
   
    Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------

    Call DbQuery        
                                                                                '☜: Query db data
       
    FncQuery = True                                                                '⊙: Processing is OK
    
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
    Call DbSave                                                                  '☜: Save db data

     FncSave = True                                                           '⊙: Processing is OK
    
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
    Call Parent.FncExport(parent.C_MULTI)                                                '☜: 화면 유형 
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
'    설명 : 
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
       
            If lgIntFlgMode = parent.OPMD_CMODE Then
				IF Trim(.txtVatNo.value) <> "" Then
					lgStrPrevVatKey = Trim(.txtVatNo.value)
				END IF
			End if           
            strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
            strVal = strVal & "&txtIssuedDtFr=" & Trim(.txtIssuedDtFr.text)            '조회 조건 데이타 
            strVal = strVal & "&txtIssuedDtTo=" & Trim(.txtIssuedDtTo.text)                '조회 조건 데이타 
            strVal = strVal & "&txtReportBizArea=" & Trim(.txtReportBizArea.value)                    '조회 조건 데이타 
            strVal = strVal & "&cboIoFg=" & Trim(.cboIoFg.value)                    '조회 조건 데이타 
            strVal = strVal & "&txtVatType=" & Trim(.txtVatType.value)                    '조회 조건 데이타 
            strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)                    '조회 조건 데이타 
            strVal = strVal & "&txtissue_dt_fg_cd=" & Trim(.cboissue_dt_fg.value)
            strVal = strVal & "&txtissue_dt_kind_cd=" & Trim(.cboissue_dt_kind.value)
        
            strVal = strVal & "&lgStrPrevVatKey=" & lgStrPrevVatKey
            strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
            strVal = strVal & "&lgPageNo=" & lgPageNo

        Call RunMyBizASP(MyBizASP, strVal)        '☜: 비지니스 ASP 를 가동 
                        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()                                                        '☆: 조회 성공후 실행로직 
    
    Call SetSpreadLock()'(-1, "Query")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE    '⊙: Indicates that current mode is Update mode
    
    ' 현재 Page의 From Element들을 사용자가 입력을 받지 못하게 하거나 필수입력사항을 표시한다.    
    Call ggoOper.LockField(Document, "Q")    '⊙: This function lock the suitable field
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
    
    'Call LayerShowHide(1)

    DbSave = False                '⊙: Processing is NG
    'On Error Resume Next        '☜: Protect system from crashing
	
    With frm1
        .txtMode.value = parent.UID_M0002
        .txtUpdtUserId.value = parent.gUsrID
        
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
            
            Select Case .vspdData.Text
            
                Case ggoSpread.UpdateFlag                                                '☜: 수정 
                    strVal = strVal & "U" & iColSep & lRow & iColSep                    '☜: U=Update
                    .vspdData.Col = C_VAT_NO
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_ISSUED_DT
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_BP_CD
                    strVal = strVal & Trim(UCase(.vspdData.Text)) & iColSep
                    .vspdData.Col = C_VAT_TYPE
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_CARD_NO
                    strVal = strVal & Trim(UCase(.vspdData.Text)) & iColSep
                    .vspdData.Col = C_REPORT_BIZ_AREA_CD
                    strVal = strVal & Trim(UCase(.vspdData.Text)) & iColSep
                    .vspdData.Col = C_BIZ_AREA_CD
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_NET_LOC_AMT
                    strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
                    .vspdData.Col = C_VAT_LOC_AMT
                    strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep

					.vspdData.Col = C_issue_dt_kind_cd	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_issue_dt_fg_cd	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep  
                                        
                                   
                    
                    lGrpCnt = lGrpCnt + 1
            End Select
                        
        Next
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value =  strVal

         Call ExecMyBizASP(frm1, BIZ_PGM_ID)        '☜: 비지니스 ASP 를 가동 
    
    End With

    DbSave = True                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()                                                    '☆: 저장 성공후 실행 로직 
    
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
    Case "QUERY"
        Call SetToolbar("1100100100111111")
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


Function AcctApply()
	
	Dim lRow
	Dim str_fg,str_issue_kind
	
	str_fg = frm1.cboissue_dt_fg2.value 
	str_issue_kind = frm1.cboissue_dt_kind2.value 
     ggoSpread.Source = frm1.vspdData

	With Frm1.vspdData
       For lRow = 1 To .MaxRows
			.Row = lRow
			if str_fg<>"" then
				.Col = C_issue_dt_fg_cd : 				.Text = str_fg
				.Col = C_issue_dt_fg_nm : 				.Text = str_fg
				  ggoSpread.UpdateRow lRow
			end if	
			if str_issue_kind<>"" then	
				.Col = C_issue_dt_kind_cd : 				.Text = str_issue_kind
				call vspdData_Change(C_issue_dt_kind_cd,lRow)
				
		     end if
		Next
		
	End With

End Function
 

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->    
</HEAD>
<!-- '#########################################################################################################
'                           6. Tag부 
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>부가세수정</font></td>
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
                                    <TD CLASS="TD5" NOWRAP>발행일자</TD>
                                    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtIssuedDtFr" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="시작발행일자" id=txtIssuedDtFr></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
                                                           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtIssuedDtTo" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="종료발행일자" id=txtIssuedDtTo></OBJECT>');</SCRIPT>
                                                           
                                    </TD>
                                    <TD CLASS="TD5" NOWRAP>세금신고사업장</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportBizArea" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportBizArea" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportBizArea()">&nbsp;
                                                           <INPUT TYPE=TEXT NAME="txtReportBizAreaNm" SIZE=20 tag="14" ALT="세금신고사업장"></TD>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>입출구분</TD>
                                    <TD CLASS="TD6" NOWRAP><SELECT NAME="cboIoFg" ALT="입출구분" tag="11" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT></TD>                                        
                                                                        </TD>
                                    <TD CLASS="TD5" NOWRAP>계산서유형</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="계산서유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;
                                                           <INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="14" ALT="계산서유형"></TD>

                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>거래처</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;
                                                           <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14" ALT="거래처"></TD>                                                                        
                                    </TD>
									<TD CLASS="TD5" NOWRAP>계산서번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatNo" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="계산서번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenVatNoInfo(frm1.txtVatNo.value,'VatNo')"></TD>
                                </TR>
                                 <TR>
                                  
									<TD CLASS="TD5" NOWRAP>전자세금계산서발행여부</TD>
									<TD CLASS="TD6" NOWRAP> <SELECT NAME="cboissue_dt_fg" ALT="전자세금계산서발행여부" tag="11" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT></TD>
									  <TD CLASS="TD5" NOWRAP>전자세금계산서종류</TD>
                                    <TD CLASS="TD6" NOWRAP>
                                    <SELECT NAME="cboissue_dt_kind" ALT="전자세금계산서종류" tag="11" STYLE="WIDTH: 170px"  ><OPTION VALUE=""></OPTION></SELECT>
                                    </TD>
                                </TR>
                                
                                
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                
                <TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>전자세금계산서발행여부</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboissue_dt_fg2" ALT="전자세금계산서발행여부" tag="11" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT>
									 </TD>
									 
									<TD CLASS=TD5 NOWRAP>전자세금계산서종류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboissue_dt_kind2" ALT="전자세금계산서종류" tag="11" STYLE="WIDTH: 170px"  ><OPTION VALUE=""></OPTION></SELECT>
														 <BUTTON NAME="btnApply" style="height:20px" CLASS="CLSSBTN" ONCLICK="vbscript:AcctApply()">적용</BUTTON> </TD>
								
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
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

