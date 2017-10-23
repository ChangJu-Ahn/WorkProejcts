
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 소득금액조정 
'*  3. Program ID           : W5105RA1
'*  4. Program Name         : W5105RA1.asp
'*  5. Program Desc         : 소득금액합계표 조회팝업 
'*  6. Modified date(First) : 2005/02/14
'*  7. Modified date(Last)  : 2005/02/14
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
 -->
<HTML>
<HEAD>
<TITLE>소득금액합계표 조회</TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID 		= "W5105RB1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 18												'☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                          
Dim lgPopUpR                                              

Dim  IsOpenPop                                                  '☜: 마크                                  

Dim arrReturn
Dim arrParent
Dim arrParam

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(0)

	 '------ Set Parameters from Parent ASP ------ 
Dim C_W_TYPE
Dim C_SEQ_NO
Dim C_W1
Dim C_W1_BT
Dim C_W1_NM
Dim C_W2
Dim C_W3_NM
Dim C_W3
Dim C_W4

	top.document.title = "소득금액합계표 조회"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    Redim arrReturn(0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

	Self.Returnvalue = arrReturn
 
End Sub

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_W_TYPE	= 1
	C_SEQ_NO	= 2
	C_W1		= 3
	C_W1_BT		= 4
	C_W1_NM		= 5
	C_W2		= 6
	C_W3_NM		= 7
	C_W3		= 8
	C_W4		= 9
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()


    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
End Sub


'======================================================================================================
'   Function Name : EscPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtDeptCd.focus
			Case 1
				.txtDealBpCd.focus
		End Select
	End With
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtDeptCd.focus
			Case 1
				.txtDealBpCd.value = arrRet(0)
				.txtDealBpNm.value = arrRet(1)
				.txtDealBpCd.focus
		End Select
	End With
End Function
'========================================  2.3 LoadInfTB19029()  =========================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
			
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Dim ret
    Call initSpreadPosVariables()  
    
	With frm1.vspdData
				
		ggoSpread.Source = frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20041222_1",,PopupParent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
	
  		'헤더를 3줄로    
		.ColHeaderRows = 3
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W_TYPE,	"데이타구분",		5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W1,		"(1)과목",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)과목명",		15,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)금액",			15,		PopupParent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		PopupParent.gComNum1000,		PopupParent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_W3_NM,	"처분",		10,,,50,1
	    ggoSpread.SSSetEdit		C_W3,		"코드",		10,,,50,1
		ggoSpread.SSSetEdit		C_W4,		"(4)조정내용",	20,,,100,1

		' 그리드 헤더 합침 
		ret = .AddCellSpan(0		, -1000, 1, 3)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1		, -1000, 7, 1)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1		, -999, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1_BT	, -999, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1_NM	, -999, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W2		, -999, 1, 2)	
		ret = .AddCellSpan(C_W3_NM	, -999, 2, 1)
		ret = .AddCellSpan(C_W4 	, -999, 1, 2)
    
    
    
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W1		: .Text = "익금산입 및 손금불산입"
		
		' 첫번째 헤더 출력 글자 
		.Row = -999
		.Col = C_W1_NM	: .Text = "(1)과목"
		.Col = C_W2		: .Text = "(2)금액"
		.Col = C_W3_NM	: .Text = "(3)소득처분"
		
		.Row = -998
		.Col = C_W3_NM	: .Text = "처분"
		.Col = C_W3		: .Text = "코드"
								
		.rowheight(-999) = 15	' 높이 재지정 
		.rowheight(-998) = 15	' 높이 재지정 
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W1_BT,C_W1_BT,True)
		Call ggoSpread.SSSetColHidden(C_W4,C_W4,True)
				
		Call SetSpreadLock()

		.ReDraw = true	
				
	End With

	' 2번 그리드 

	With frm1.vspdData2
				
		ggoSpread.Source = frm1.vspdData2
		'patch version
		ggoSpread.Spreadinit "V20041222_2",,PopupParent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
	
  		'헤더를 3줄로    
		.ColHeaderRows = 3
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W_TYPE,	"데이타구분",		5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W1,		"(1)과목",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)과목명",		15,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)금액",			15,		PopupParent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		PopupParent.gComNum1000,		PopupParent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_W3_NM,	"코드",		10,,,50,1
	    ggoSpread.SSSetEdit		C_W3,		"처분",		10,,,50,1
		ggoSpread.SSSetEdit		C_W4,		"(4)조정내용",	20,,,100,1

		' 그리드 헤더 합침 
		' 그리드 헤더 합침 
		ret = .AddCellSpan(0		, -1000, 1, 3)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1		, -1000, 7, 1)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1		, -999, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1_BT	, -999, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1_NM	, -999, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W2		, -999, 1, 2)	
		ret = .AddCellSpan(C_W3_NM	, -999, 2, 1)
		ret = .AddCellSpan(C_W4 	, -999, 1, 2)
    
    
    
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W1		: .Text = "손금산입 및 익금불산입"
		
		' 첫번째 헤더 출력 글자 
		.Row = -999
		.Col = C_W1_NM	: .Text = "(1)과목"
		.Col = C_W2		: .Text = "(2)금액"
		.Col = C_W3_NM	: .Text = "(3)소득처분"
		
		.Row = -998
		.Col = C_W3_NM	: .Text = "처분"
		.Col = C_W3		: .Text = "코드"
								
		.rowheight(-999) = 15	' 높이 재지정 
		.rowheight(-998) = 15	' 높이 재지정 
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W1_BT,C_W1_BT,True)
		Call ggoSpread.SSSetColHidden(C_W4,C_W4,True)

		Call SetSpreadLock()
				
		.ReDraw = true	
				
	End With

    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
    	ggoSpread.Source = .vspdData

	    .vspdData.ReDraw = False
	    
		ggoSpread.SpreadLock C_W_TYPE,   -1, C_W4
	
	    .vspdData.ReDraw = True

    	ggoSpread.Source = .vspdData2

	    .vspdData2.ReDraw = False
	    
		ggoSpread.SpreadLock C_W_TYPE,   -1, C_W4
	
	    .vspdData2.ReDraw = True

    End With
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
Function OKClick()
'	If frm1.vspdData.ActiveRow > 0 Then 				
'		Redim arrReturn(1)
'		frm1.vspdData.Row	= frm1.vspdData.ActiveRow
'		frm1.vspdData.Col	= GetKeyPos("A",1)		
'		arrReturn(0)		= frm1.vspdData.Text
'	End if			

	arrReturn(0)		= ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()
    Call ggoOper.FormatField(Document, "1", PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, PopupParent.gDateFormat,3)
    Call ggoOper.LockField(Document, "N")
    
	Call InitSpreadSheet()
    Call InitComboBox()
    Call InitVariables()														
	Call SetDefaultVal()	
	Call FncQuery()
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
	Dim IntRetCD
    FncQuery = False                                            
    
    Err.Clear                                                   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables() 											
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData		'Buffer Clear
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & PopupParent.UID_M0001						         
        strVal = strVal     & "&txtCO_CD="       	& Frm1.txtCO_CD.Value      '☜: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  						
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
	If frm1.vspdData2.MaxRows > 0 Then
		Frm1.vspdData2.Row = frm1.vspdData2.MaxRows
		Frm1.vspdData2.Col = C_W1_NM
		Frm1.vspdData2.Text = "합계"
		Frm1.vspdData2.TypeHAlign = 2
		
	End If
		
	If frm1.vspdData.MaxRows > 0 Then
		Frm1.vspdData.Row = frm1.vspdData.MaxRows
		Frm1.vspdData.Col = C_W1_NM
		Frm1.vspdData.Text = "합계"
		Frm1.vspdData.TypeHAlign = 2
		
		frm1.vspdData.Focus
	End If
		
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = gMethodText
  
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
      
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(GetSQLSortFieldCD("A"),GetSQLSortFieldNM("A"),arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell을 dblclick했거나....
		Exit Function
	End If
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Function
	End If
	Call OKClick
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim ii
	
    gMouseClickStatus = "SPC"   
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
'    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		    
End Sub
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	'If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    '	If lgPageNo <> "" Then								
    '       If DbQuery = False Then
    ''          Exit Sub
    '       End if
    '	End If
    'End If
    
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5">사업연도</TD>
						<TD CLASS="TD6"><script language =javascript src='./js/w5105ra1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
							     <TD WIDTH="50%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5105ra1_vspdData_vspdData.js'></script>
							    </TD>
							     <TD WIDTH="50%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5105ra1_vspdData2_vspdData2.js'></script>
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
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()"></IMG>&nbsp;
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=  <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtFrArDt"	tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToArDt"	tag="24">
<INPUT TYPE=HIDDEN NAME="htxtFrArNo"	tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToArNo"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtdeptcd"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDealBpCd"  tag="24">
<INPUT		TYPE=hidden	 NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
