<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% Option Explicit %>
<% session.CodePage=949 %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Tax
'*  2. Function Name        : 
'*  3. Program ID           : W1111MA1
'*  4. Program Name         : 제3호의3(3) 부속명세서-임대원가명세서 
'*  5. Program Desc         : 제3호의3(3) 부속명세서-임대원가명세서 등록/조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/12/29
'*  8. Modified date(Last)  : 2004/12/30
'*  9. Modifier (First)     : LSHSAT
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'
'서식간 검증 
'      14.접대비 >0 인경우 47호(갑) (117)의 ③ < 0 이면 오류	오류	WC0024	법인세 제47호(갑) 주요계정명세서의 접대비중 ③회사계상금액이 존재하지 않습니다.	%1의 %2중 %3이 존재하지 않습니다.
'***********************************************************************k*********************** -->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '☜: indicates that All variables must be declared in advance 


'********************************************  1.2 Global 변수/상수 선언  *********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global 상수 선언  ====================================
'==========================================================================================================

Const BIZ_PGM_ID 		= "W1111MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 		= "W1111MB2.asp"											 '☆: 비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 100	                                      '☜: Visble row

'========================================================================================================= 
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim lgOldRow

Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        

Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

Dim C_GP_CD
Dim C_PAR_GP_CD
Dim C_GP_NM
Dim C_FORM_REP_NO
Dim C_AMT
Dim C_SUM_FG
Dim C_GP_LVL


'============================================  초기화 함수  ====================================
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()
	
	C_GP_CD = 1
	C_PAR_GP_CD = 2
	C_GP_NM = 3
	C_FORM_REP_NO = 4
	C_AMT = 5
	C_SUM_FG = 6
	C_GP_LVL = 7
	
end sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'☆: 사용자 변수 초기화 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""

    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow = 0

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False

    lgRefMode = False

End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub



'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub


'============================================  그리드 함수  ====================================
'========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_GP_CD						= iCurColumnPos(1)
			C_PAR_GP_CD					= iCurColumnPos(2)
			C_GP_NM						= iCurColumnPos(3)
			C_FORM_REP_NO				= iCurColumnPos(4)
			C_AMT						= iCurColumnPos(5)
			C_SUM_FG					= iCurColumnPos(6)
			C_GP_LVL					= iCurColumnPos(7)
    End Select    
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gForbidDragDropSpread    

		.ReDraw = false
		.MaxCols   = C_GP_LVL + 1                                          ' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:
		.MaxRows = 0

		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 

		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetEdit     C_GP_CD,				"계정코드",				10
		ggoSpread.SSSetEdit     C_PAR_GP_CD,			"상위계정코드",				10
		ggoSpread.SSSetEdit     C_GP_NM,				"계정과목",				25
		ggoSpread.SSSetEdit     C_FORM_REP_NO,			"코드",					7, 2
		ggoSpread.SSSetFloat    C_AMT,					"금액",					15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,  "1", True,   "Z"
		ggoSpread.SSSetEdit		C_SUM_FG,				"계산여부",				10
		ggoSpread.SSSetEdit		C_GP_LVL,				"레벨",					5

		Call ggoSpread.SSSetColHidden(C_GP_CD,C_GP_CD,True)	
		Call ggoSpread.SSSetColHidden(C_PAR_GP_CD,C_PAR_GP_CD,True)	
		Call ggoSpread.SSSetColHidden(C_SUM_FG,C_SUM_FG,True)	
		Call ggoSpread.SSSetColHidden(C_GP_LVL,C_GP_LVL,True)	

		Call SetSpreadLock 
		.ReDraw = true
	    
	End With
	
	
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
'      ggoSpread.Source = frm1.vspdData
'      ggoSpread.SpreadLockWithOddEvenRowColor()

   With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_GP_NM, -1, C_GP_NM
    ggoSpread.SpreadLock C_FORM_REP_NO, -1, C_FORM_REP_NO
	ggoSpread.SSSetProtected C_AMT, 18, 18

    'ggoSpread.SSSetRequired C_SUMMARY_DESC, -1, -1
    'ggoSpread.SpreadLock C_COMPANY_NM, -1, C_COMPANY_NM
    'ggoSpread.SSSetRequired C_STOCK_RATE, -1, -1
    'ggoSpread.SSSetRequired C_ACQUIRE_AMT, -1, -1
    'ggoSpread.SSSetRequired C_COMPANY_TYPE, -1, -1
    'ggoSpread.SpreadLock C_OWN_RGST_NO, -1, C_OWN_RGST_NO
	'ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub


'============================================  조회조건 함수  ====================================
Sub CheckFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		sFISC_START_DT = CDate(lgF0)
	Else
		sFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		sFISC_END_DT = CDate(lgF1)
	Else
		sFISC_END_DT = ""
	End if
	
End Sub


'============================================  폼 함수  ====================================

'========================================================================================================= 
Sub Form_Load()
    Call InitVariables																'⊙: Initializes local global variables
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox

    Call SetToolBar("1100100000010111")

	Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	Call ggoOper.FormatDate(frm1.txtFISC_YEAR_Body, parent.gDateFormat,3)

    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    
    Call ggoOper.ClearField(Document, "2")
    Call InitData
    Call FncQuery()
    
End Sub

'==========================================================================================
'==========================================================================================
Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call CheckFISC_DATE
End Sub



'============================================  툴바지원 함수  ====================================

'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
'    If lgBlnFlgChgValue = True Then
'		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'    End If

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

  '-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolBar("1100100000010111")

    If DbQuery2 = False Then
		Call RestoreToolBar()
        Exit Function
    End If
    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
	Dim IntRetCD
	
    FncSave = False                                                         
    
    Err.Clear                                                               
    'On Error Resume Next                                                   

	'-----------------------
	'Condition copy to Check Field
	'-----------------------
	If Not chkField(Document, "1") Then                             '⊙: Check indispensable field
	   Exit Function
	End If
	Frm1.txtFISC_YEAR_Body.Value = Frm1.txtFISC_YEAR.Text
	Frm1.txtREP_TYPE_Body.Value = Frm1.cboREP_TYPE.Value
'	Frm1.txtBS_PL_FG.Value
	

	'-----------------------
	'Check content area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
       Exit Function
    End If

	If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If

<%  '-----------------------
    'Precheck area
    '----------------------- %>
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '⊙: No data changed!!
	    Exit Function
	End If
    

	'-----------------------
	' 필수입력 금액 확인 
	'-----------------------
	If SubMakeSum <> True Then
		Exit Function
	End If

	If Verification = False Then Exit Function

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function


' ---------------------- 서식내 검증 -------------------------
Function  Verification()
	Dim iSchRow, iSumAmt
	
	Verification = False

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "17", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = UNICdbl(Frm1.vspdData.text)

	'17 <= 0	오류	WC0007	입력된 DATA가 없습니다.
    If iSumAmt <= 0 Then
        Call DisplayMsgBox("WC0007", "X", "X", "X")                          
        SubMakeSum = False
        Exit Function
    End If

	Verification = True	
End Function

'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode

     ' 조건부 필드를 삭제한다. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
	lgBlnFlgChgValue = True

'    frm1.txtCO_CD_Body.value = ""

'    frm1.txtCO_CD_Body.focus
    
End Function


'========================================================================================
Function FncCancel()
     On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF

    response.write lgPrevNo

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtco_cd =" & lgPrevNo

	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						  '☜: 비지니스 처리 ASP의 상태값 
    strVal = strVal & "&txtco_cd=" & lgNextNo

	Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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
	Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub


'============================================  DB 억세스 함수  ====================================

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
    strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtBS_PL_FG=" & Trim(frm1.txtBS_PL_FG.value)				'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function


'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
    strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 


	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
 '   Call LayerShowHide(0)
End Function

'========================================================================================
Function DbQuery2()

    Err.Clear

    DbQuery2 = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 


	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery2 = True
 '   Call LayerShowHide(0)
End Function

'========================================================================================
Function DbQueryOk()
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	Call SetToolbar("1101100000010111")												'⊙: Set ToolBar
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
    Call SetSpreadLock

	Frm1.vspdData.focus

End Function

'========================================================================================
Function DbQueryOk2()
	lgIntFlgMode      =  parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
    Call SetSpreadLock

	Frm1.vspdData.focus

End Function

'========================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With Frm1
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
			strVal = strVal & "C"  &  Parent.gColSep

            .vspdData.Col = C_GP_CD				: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
            .vspdData.Col = C_GP_NM     	    : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
            .vspdData.Col = C_FORM_REP_NO		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
            .vspdData.Col = C_AMT				: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep

            lGrpCnt = lGrpCnt + 1
                    
       Next
		.txtMode.value        =  Parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
		.txtSpread.value      = strVal
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
''    frm1.txtCO_CD.value = frm1.txtCO_CD_Body.value 
    lgBlnFlgChgValue = False
    Call FncQuery
End Function


'============================================  이벤트 함수  ====================================

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
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
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
    lgBlnFlgChgValue = True

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	Call SubMakeSum
	
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)

	'Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    If Row = 0 Then
       Exit Sub
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
    
   	If lgOldRow <> Row Then
		
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row
	
		lgOldRow = Row
		  		
	End If
       frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If
	
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = NewRow
	
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
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
	End If  
End Sub



'============================== 그리드 데이타 처리 정의 함수  ========================================

'========================================================================================================
'   Event Name : SubMakeSum
'   Event Desc : This function is Sum Amt with spread sheet
'	FORM_REP_NO	:	01 ~ 16		사용자가 숫자로 직접입력함.
'					17		1 ~ 16 의 합계를 입력함.
'========================================================================================================
Function SubMakeSum()
	Dim iSchRow
	Dim iSumAmt
	DIm iRowCnt

	iSumAmt = 0
	
	'17		1 ~ 16 의 합계를 입력함.
	iSumAmt = 0
	For iRowCnt = 1 to 18
		If iRowCnt <> 17 Then
			iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, Right("0" + CStr(iRowCnt), 2), 0)
	   		Frm1.vspdData.Row = iSchRow
			Frm1.vspdData.Col = C_AMT
			iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)
		End If
	Next

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "17", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iSchRow
	
	SubMakeSum = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white>제3호의3(3) 부속명세 임대원가</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
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
					<TD WIDTH="100%" valign=top>
						<TABLE  CLASS="TB3" CELLSPACING=0>
							<TR>
								<TD HEIGHT="100%" NOWRAP COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtCO_CD_Body" tag="24" tabindex="-1">
<INPUT TYPE=hidden name=txtFISC_YEAR_Body  tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtREP_TYPE_Body" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtBS_PL_FG" tag="24" tabindex="-1">
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="strUrl" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

