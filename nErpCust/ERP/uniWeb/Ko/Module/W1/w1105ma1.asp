<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% session.CodePage=949 %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Tax
'*  2. Function Name        : 
'*  3. Program ID           : W1105MA1
'*  4. Program Name         : 제3호의3(1)(2)표준손익계산서 
'*  5. Program Desc         : 제3호의3(1)(2)표준손익계산서 등록/조회 
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
' 음수처리 미비 
' 불러오기 미비.
'서식간 검증 
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
<SCRIPT LANGUAGE="VBScript"  SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '☜: indicates that All variables must be declared in advance 


'********************************************  1.2 Global 변수/상수 선언  *********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global 상수 선언  ====================================
'==========================================================================================================

Const BIZ_MNU_ID		= "W1105MA1"
Const BIZ_PGM_ID 		= "W1105MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 		= "W1105MB2.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W1105MB3.asp"

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
Dim C_FISC_CD
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
	C_FISC_CD = 4
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
            
			C_GP_CD						= iCurColumnPos(1)
			C_PAR_GP_CD					= iCurColumnPos(2)
			C_GP_NM						= iCurColumnPos(3)
			C_FISC_CD					= iCurColumnPos(4)
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
		ggoSpread.SSSetEdit     C_GP_NM,				"계정과목",				25,,,50,1
		ggoSpread.SSSetEdit     C_FISC_CD,				"코드",					7, 2
		ggoSpread.SSSetFloat    C_AMT,					"금액",					15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,  "1", True,   ""
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
	Dim iSchROw
    
    ggoSpread.Source = frm1.vspdData

	With frm1

		.vspdData.ReDraw = False
    
		ggoSpread.SpreadLock C_GP_NM, -1, C_GP_NM
		ggoSpread.SpreadLock C_FISC_CD, -1, C_FISC_CD
 
		For iSchRow = 1 to .vspdData.MaxRows
			.vspdData.Row = iSchRow
			.vspdData.Col = C_SUM_FG
			If "X" = .vspdData.text Then
				ggoSpread.SSSetProtected C_AMT, iSchRow, iSchRow
			End If
		Next
		
		' -- 200603 개정: 사용자정의 코드명 추가 
		' -- 일반법인 
		If Frm1.txtCOMP_TYPE2.Value = "1" Then	
			ggoSpread.SpreadUnLock C_GP_NM, 40, C_GP_NM, 42
			ggoSpread.SpreadUnLock C_GP_NM, 62, C_GP_NM, 64
			ggoSpread.SpreadUnLock C_GP_NM, 85, C_GP_NM, 87
		Else
		' -- 금융법인 
			ggoSpread.SpreadUnLock C_GP_NM, 17, C_GP_NM, 19
			ggoSpread.SpreadUnLock C_GP_NM, 52, C_GP_NM, 54
			ggoSpread.SpreadUnLock C_GP_NM, 57, C_GP_NM, 59
			ggoSpread.SpreadUnLock C_GP_NM, 70, C_GP_NM, 72
			ggoSpread.SpreadUnLock C_GP_NM, 83, C_GP_NM, 85
		End If		

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

    Call SetToolBar("1101100000010111")

	Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	Call ggoOper.FormatDate(frm1.txtFISC_YEAR_Body, parent.gDateFormat,3)

    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    
    Call ggoOper.ClearField(Document, "2")
    Call InitData
    Call fncQuery()
    
End Sub

'==========================================================================================
'==========================================================================================
Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	'Call CheckFISC_DATE	' -- 2006-01-23 제거 : cyt
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
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

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


	Call DbQuery2

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
	Dim iSchRow
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
'	Frm1.txtCOMP_TYPE2.Value
	

	'-----------------------
	'Check content area
	'-----------------------
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
	' XII.법인세비용이 제15-1호 서식의 합계금액 보다 크면 오류																										
	' 12.기부금 금액이 500만원 이상인 경우 제22호 기부금 합계금액이 0보다 작거나 같으면 오류																										
	'-----------------------
'	iSchRow = 0
'	For iSchRow = 1 to Frm1.vspdData.MaxRows
'		Frm1.vspdData.Row = iSchRow
'		Frm1.vspdData.Col = C_GP_LVL

'		If Trim(Frm1.vspdData.text) <> "1" Then
'			Call SubMakeSum(iSchRow)
'		End IF
'	Next
	Call SetMakeSum()

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
       Exit Function
    End If


    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          

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
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
	strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 


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
            .vspdData.Col = C_FISC_CD			: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
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
Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD


    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If
    

    Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE)
    call CommonQueryRs("COMP_TYPE2"," TB_COMPANY_HISTORY "," CO_CD= '" & varCo_Cd & "' AND FISC_YEAR='" & varFISC_YEAR & "' AND REP_TYPE='" & varREP_TYPE & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if unicdbl(lgF0) = 1  then
      	 EBR_RPT_ID	    = "W1105OA1"

    else
         EBR_RPT_ID	    = "W1105OA2"

    end if
    
   
    StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
	StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE

     ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
     if  strPrintType = "VIEW" then
	 Call FncEBRPreview(ObjName, StrUrl)
     else
	 Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
     

   

End Function 
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
	
'	Call SubMakeSum(Row)
	Call SetMakeSum()
	
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
   	
'    If Row = 0 Then
'        ggoSpread.Source = frm1.vspdData
'        If lgSortKey = 1 Then
'            ggoSpread.SSSort
'            lgSortKey = 2
'        Else
'            ggoSpread.SSSort ,lgSortKey
'            lgSortKey = 1
'        End If
'    End If
    
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
'   Event Desc : This function is calcuralte spread data
'========================================================================================================
Sub SubMakeSum(ByVal Row )
	Dim iStrParGpCd
	Dim iSchRow
	Dim iSumAmt

   	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = C_PAR_GP_CD
	iStrParGpCd = Frm1.vspdData.text
	
	If iStrParGpCd = "" Then
		Exit Sub
	End If

	iSumAmt = 0
	
	For iSchRow = 1 to Frm1.vspdData.MaxRows
	   	Frm1.vspdData.Row = iSchRow
		Frm1.vspdData.Col = C_PAR_GP_CD
		If iStrParGpCd = Frm1.vspdData.text Then
			Frm1.vspdData.Col = C_AMT
			iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)
		End If
	Next

	iSchRow = Frm1.vspdData.SearchCol(C_GP_CD, 0, Frm1.vspdData.MaxRows, iStrParGpCd, 0)

   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iSchRow
	
	Frm1.vspdData.Col = C_GP_LVL
	If Trim(Frm1.vspdData.text) <> "1" Then
		Call SubMakeSum(iSchRow)
	End IF
	
End Sub

Sub SetMakeSum()
	Dim iStrRow, iEndRow, iTarRow
	Dim iSumAmt
	
	If Frm1.txtCOMP_TYPE2.Value = "1" Then
		With Frm1.vspdData
			.Redraw = False
		
			' 01 = sum(02 ~ 08)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "01", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "02", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "08", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 10 = 11 + 12 - 13 + 91
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "11", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "12", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "13", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "91", 0)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "10", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 14 = 15 + 16 - 17 - 18
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "15", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "16", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "17", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "18", 0)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "14", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 09 = 10 + 14
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "10", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "09", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 19 = 01 - 09
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "01", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text) - iSumAmt
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "19", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' -- 추가 계정 
			' 35 = sum(201 ~ 204)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "35", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "201", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "204", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 20 = sum(21 ~ 35)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "20", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "21", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "35", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 36 = 19 - 20
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "19", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "20", 0)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "36", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' -- 추가 계정 
			' 52 = sum(211 ~ 214)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "52", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "211", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "214", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 37 = sum(38 ~ 52)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "37", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "38", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "52", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' -- 추가 계정 
			' 70 = sum(221 ~ 224)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "70", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "221", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "224", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 53 = sum(54 ~ 70)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "53", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "54", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "70", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 71 = 36 + 37 - 53
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "36", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "37", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "53", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "71", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 72 = sum(73 ~ 76)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "72", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "73", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "76", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 77 = sum(78 ~ 79)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "77", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "78", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "79", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 80 = 71 + 72 - 77
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "71", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "72", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "77", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "80", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 82 = 80 - 81
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "80", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "81", 0)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "82", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			.Redraw = True
		End With
	Else
		With Frm1.vspdData
		
			' 02 = sum(03 ~ 08)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "02", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "03", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "08", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' -- 추가 계정 
			' 16 = sum(201 ~ 204)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "16", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "201", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "204", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 01 = 02 + sum(09 ~ 16)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "01", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "09", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "16", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "01", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "02", 0)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "01", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 18 = sum(19 ~ 23)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "18", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "19", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "23", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' -- 추가 계정 
			' 44 = sum(211 ~ 214)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "44", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "211", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "214", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' -- 추가 계정 
			' 45 = sum(221 ~ 224)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "45", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "221", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "224", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 32 = sum(33 ~ 44)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "32", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "33", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "44", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 17 = 18 + sum(24 ~ 32) + 45
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "17", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "24", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "32", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "17", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "18", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "45", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 46 = 01 - 17
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "01", 0)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "17", 0)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "46", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' -- 추가 계정 
			' 53 = sum(231 ~ 234)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "53", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "231", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "234", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 47 = sum(48 ~ 53)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "47", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "48", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "53", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' -- 추가 계정 
			' 61 = sum(241 ~ 244)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "61", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "241", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "244", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 54 = sum(55 ~ 61)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "54", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "55", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "61", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 
		
			' 62 = 46 + 47 - 54
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "46", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "47", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "54", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "62", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 63 = sum(64 ~ 67)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "63", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "64", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "67", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 68 = sum(69 ~ 70)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "68", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "69", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "70", 0)
		
			Call FncSumSheet(Frm1.vspdData, C_AMT, iStrRow, iEndRow, true, iTarRow, C_AMT, "V")	' 합계 

			' 71 = 62 + 63 - 68
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "62", 0)
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "63", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "68", 0)
			.Row = iTarRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = iSumAmt + UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "71", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

			' 73 = 71 - 72
			iStrRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "71", 0)
			iEndRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "72", 0)
			.Row = iStrRow : .Col = C_AMT : iSumAmt = UNICDbl(.Text)
			.Row = iEndRow : .Col = C_AMT : iSumAmt = iSumAmt - UNICDbl(.Text)
			iTarRow = .SearchCol(C_FISC_CD, 0, .MaxRows, "73", 0)
			.Row = iTarRow : .Col = C_AMT : .Text = iSumAmt

		End With
	End If

End Sub

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim sMesg

	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
		
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtCO_CD="      	 & Frm1.txtCO_CD.Value	      '☜: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
    	
End Function

Function GetRefOK(ByVal pStrData)
	Dim arrRowVal, arrColVal
	Dim lLngMaxRow, iDx, iSchRow

	If pStrData <> "" Then
		lgBlnFlgChgValue = True
		arrRowVal = Split(pStrData, Parent.gRowSep)                                 '☜: Split Row    data
		lLngMaxRow = UBound(arrRowVal)

		For iDx = 1 To lLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), Parent.gColSep)   
		    
		    With Frm1.vspdData
				.Redraw = False
				iSchRow = .SearchCol(C_GP_CD, 0, .MaxRows, arrColVal(1), 0)
			   	.Row = iSchRow

				.Col	= C_AMT	:	.Text	= arrColVal(2)
				.Redraw = True
			End With
			Call SetMakeSum()
		Next
	End IF
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">금액불러오기</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			
			
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" STYLE="DISPLAY:NONE"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtCO_CD_Body" tag="24" tabindex="-1">
<INPUT TYPE=hidden name=txtFISC_YEAR_Body  tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtREP_TYPE_Body" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtCOMP_TYPE2" tag="24" tabindex="-1">
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
