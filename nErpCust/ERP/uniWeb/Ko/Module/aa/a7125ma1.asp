
<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7102ma1
'*  4. Program Name         : 고정자산취득상세내역등록 
'*  5. Program Desc         : 고정자산별 취득 상세 내역을 등록,수정,삭제,조회 
'*  6. Comproxy List        : +As0021
'                             +As0029
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/19
'*  9. Modifier (First)     : 김희정 
'* 10. Modifier (Last)      : 김희정 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/003/30 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* 
 -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
'==============================================  1.1.1 Style Sheet  ======================================
'=========================================================================================================
 -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit									'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "a7125mb1.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID_Q1 = "a7125mb2.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID_Q2 = "a7125mb3.asp"			'☆: 비지니스 로직 ASP명 


'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

'''자산master
Dim C_Deptcd
Dim C_DeptNm
Dim C_AcctCd
Dim C_AcctNm
Dim C_AsstNo
Dim C_AsstNm
Dim C_AcqAmt
Dim C_AcqLocAmt
Dim C_AcqQty
Dim C_ResAmt
Dim C_RefNo
Dim C_Desc

Const C_SHEETMAXROWS = 30

''취득상세내역 
Dim C_Seq_2
Dim C_Desc_2
Dim C_Amt_2
Dim C_AsstNo_2
Dim C_LocAmt_2


Const C_SHEETMAXROWS_2  = 30	


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
Dim lgStrPrevKey_m

'Dim lgLngCurRows
'Dim lgKeyStream
Dim lgKeyStream_m

'Dim lgSortKey

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  --------------------------------------------------------------
Dim IsOpenPop        
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
 
 
 
'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################

'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"	
			C_Deptcd		= 1
			C_DeptNm		= 2
			C_AcctCd		= 3
			C_AcctNm		= 4
			C_AsstNo		= 5
			C_AsstNm		= 6
			C_AcqAmt		= 7
			C_AcqLocAmt	= 8
			C_AcqQty		= 9
			C_ResAmt		= 10
			C_RefNo		= 11
			C_Desc		= 12
		Case "B"			
			C_Seq_2			= 1
			C_Desc_2		= 2
			C_Amt_2			= 3
			C_AsstNo_2		= 4
			C_LocAmt_2		= 5
	End Select
End Sub


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
	
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
frm1.txtAcqNo.focus
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey_m = 0                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
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
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>

End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
   Select Case pOpt
       Case "Q"
                  lgKeyStream = Frm1.txtAcqNo.Value  & Parent.gColSep       'You Must append one character(Parent.gColSep)
       Case "M"
                  lgKeyStream = Frm1.htxtAcqNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   End Select 
                   
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)
    
    Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

				.ReDraw = false
				.MaxCols = C_Desc +1							'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols								    '☜: 공통콘트롤 사용 Hidden Column
				.ColHidden = True
				.MaxRows = 0
	
				Call GetSpreadColumnPos(pvSpdNo)

   				'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar
				ggoSpread.SSSetEdit		C_DeptCd,  "부서코드", 8, , , 10
				ggoSpread.SSSetEdit		C_DeptNm,  "부서명",   10

				ggoSpread.SSSetEdit		C_AcctCd,  "계정코드", 10, , , 20
				ggoSpread.SSSetEdit		C_AcctNm,  "계정명",   20
				ggoSpread.SSSetEdit		C_AsstNo, "자산번호", 15, , , 18
			    ggoSpread.SSSetEdit		C_AsstNm, "자산명",   20, , , 40
			    
				ggoSpread.SSSetFloat    C_AcqAmt,   "취득금액",      15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetFloat    C_AcqLocAmt,"취득금액(자국)",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				Call AppendNumberPlace("6","3","0")

			    ggoSpread.SSSetFloat    C_AcqQty,   "취득수량",      15,"6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetFloat    C_ResAmt,"잔존가액(자국)",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetEdit		C_RefNo, "참조번호", 30, , , 30
				ggoSpread.SSSetEdit		C_Desc,  "적요",     30, , , 128

				.ReDraw = true

			End With
		Case "B"
		
			With frm1.vspdData2
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 
				
				.ReDraw = false
				.MaxCols = C_LocAmt_2 +1							'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols								    '☜: 공통콘트롤 사용 Hidden Column
				.ColHidden = True
				.MaxRows = 0

				Call GetSpreadColumnPos(pvSpdNo)

				'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar
				ggoSpread.SSSetFloat	C_Seq_2,     "순번", 14, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,1,True  ,, "1","999"
				ggoSpread.SSSetEdit		C_Desc_2,  "내역"		,53, , , 40
				ggoSpread.SSSetFloat    C_Amt_2,   "금액"		,25, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetEdit		C_AsstNo_2, "자산번호", 15, , , 18			'Asset_no를 Hidden으로 가지고 간다.
				ggoSpread.SSSetFloat    C_LocAmt_2,"금액(자국)"	,25, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		        Call ggoSpread.SSSetColHidden(C_AsstNo_2,C_AsstNo_2,True)

				.ReDraw = true
					
			End With

	End Select
	
    Call SetSpreadLock(pvSpdNo)	
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData
	
				ggoSpread.Source = frm1.vspdData

				.ReDraw = False		

					ggoSpread.SpreadLock C_DeptCd,   -1
					ggoSpread.SpreadLock C_DeptNm,   -1
					ggoSpread.SpreadLock C_AcctCd,   -1
					ggoSpread.SpreadLock C_AcctNm,   -1
					ggoSpread.SpreadLock C_AcqAmt,   -1
					ggoSpread.SpreadLock C_AcqLocAmt,   -1
					ggoSpread.SpreadLock C_AcqQty,   -1
					ggoSpread.SpreadLock C_ResAmt,   -1
					ggoSpread.SpreadLock C_RefNo,   -1
					ggoSpread.SpreadLock C_Desc,   -1
						
				.ReDraw = True

			End With    
		Case "B"	
			With frm1.vspdData2
	
				ggoSpread.Source = frm1.vspdData2
						
				.ReDraw = False		
				
					ggoSpread.SpreadLock C_Seq_2,   -1
					ggoSpread.SpreadUnLock C_Desc_2,   -1
					ggoSpread.SpreadUnLock C_Amt_2,   -1
					ggoSpread.SpreadUnLock C_LocAmt_2,   -1
					
					ggoSpread.SSSetProtected C_LocAmt_2 +1, -1,C_LocAmt_2 +1
				
				.ReDraw = True

			End With    
		End Select

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow , ByVal pvEndRow)
	With frm1.vspdData2
		.Redraw = False
		ggoSpread.Source = frm1.vspdData2			
		ggoSpread.SSSetRequired C_Seq_2, pvStartRow, pvEndRow
'		.Col = 2											'컬럼의 절대 위치로 이동 
'		.Row = .ActiveRow
'		.Action = 0                         
'		.EditMode = True		
		.Redraw = True		
    End With		
End Sub


'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_Deptcd	= iCurColumnPos(1)
			C_DeptNm	= iCurColumnPos(2)
			C_AcctCd	= iCurColumnPos(3)
			C_AcctNm	= iCurColumnPos(4)
			C_AsstNo	= iCurColumnPos(5)
			C_AsstNm	= iCurColumnPos(6)
			C_AcqAmt	= iCurColumnPos(7)
			C_AcqLocAmt	= iCurColumnPos(8)
			C_AcqQty	= iCurColumnPos(9)
			C_ResAmt	= iCurColumnPos(10)
			C_RefNo		= iCurColumnPos(11)
			C_Desc		= iCurColumnPos(12)

		Case "B"
			ggoSpread.Source = frm1.vspdData2

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)							
			
			C_Seq_2		= iCurColumnPos(1)
			C_Desc_2	= iCurColumnPos(2)
			C_Amt_2		= iCurColumnPos(3)
			C_AsstNo_2	= iCurColumnPos(4)
			C_LocAmt_2	= iCurColumnPos(5)

	End select
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
'======================================================================================================
'   Function Name : OpenAcqNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenAcqNoInfo()
	Dim arrRet
	Dim arrParam(3)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a7102ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7102ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True	
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent,arrParam), _
		     "dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetAcqNoInfo(arrRet)
	End If	

End Function

'======================================================================================================
'   Function Name : SetAcqNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetAcqNoInfo(Byval arrRet)

	With frm1
		.txtAcqNo.value  = arrRet(0)
		
		.txtAcqNo.focus
	End With

End Function



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

    Call LoadInfTB19029()                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field

    Call InitSpreadSheet("A")                                                     'Setup the Spread sheet

    Call InitSpreadSheet("B")                                                     'Setup the Spread sheet
    
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    Call SetDefaultVal    
	Call SetToolbar("1100000000000111")        
	
	frm1.txtAcqNo.focus
	
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
    Dim var_i, var_m
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    var_i = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var_i = True or var_m = True    Then    
		IntRetCD = DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X") '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData		'Buffer Clear
        
    Call InitVariables															'⊙: Initializes local global variables
'    Call InitSpreadSheet("A")                                                     'Setup the Spread sheet
'    Call InitSpreadSheet("B")                                                     'Setup the Spread sheet
    
    '-----------------------
    'Check condition area
    '-----------------------

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery("Q") = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '☜: Processing is OK
	   
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    Dim var_i
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
'	if frm1.vspdData2.MaxRows < 1 then'
'		IntRetCD = DisplayMsgBox("900001","X","X","X")  ''자산세부내역을 입력하십시오.
'		Exit Function
'	end if

		
    ggoSpread.Source = frm1.vspdData2
    var_i = ggoSpread.SSCheckChange
	
    If lgBlnFlgChgValue = False and var_i = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")  '☜ 바뀐부분 
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then	
		Exit Function
    End if
	    
    Call DbSave()				                                                
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    Dim IntRetCD
    
	frm1.vspdData2.ReDraw = False

	if frm1.vspdData2.MaxRows < 1 then Exit Function
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.CopyRow

	SetSpreadColor frm1.vspdData2.ActiveRow , frm1.vspdData2.ActiveRow
    
	frm1.vspdData2.Col  = C_Seq_2
	frm1.vspdData2.Text = ""
		
	frm1.vspdData2.ReDraw = True
	
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
    if frm1.vspdData2.MaxRows < 1 then	 Exit Function

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.EditUndo                                                  '☜: Protect system from crashing

End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(Byval pvRowCnt) 
	Dim varMaxRow
	Dim strDoc
	Dim varXrate
	Dim imRow
	
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

		If ImRow="" then
			Exit Function
		End If
	End If
		
	with frm1
		varMaxRow = .vspdData2.MaxRows 

		.vspdData2.focus
		
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False
		
		ggoSpread.InsertRow ,imRow

		
		frm1.vspdData.row = frm1.vspdData.activeRow
		frm1.vspdData.Col = C_asstNo

		.vspdData2.row = .vspdData2.ActiveRow
		.vspdData2.Col = C_asstNo_2
		.vspdData2.value = frm1.vspdData.value
		
		.vspdData2.Col = C_Amt_2
		.vspdData2.value = 0
		.vspdData2.Col = C_LocAmt_2
		.vspdData2.value = 0
		
		.vspdData2.ReDraw = True

		SetSpreadColor .vspdData2.ActiveRow , frm1.vspdData2.ActiveRow

	end with
	
'	Call SetToolbar("1100111100111111")

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 

	frm1.vspdData2.focus
   	ggoSpread.Source = frm1.vspdData2

	if frm1.vspdData2.MaxRows < 1 then Exit Function
	
	lDelRows = ggoSpread.DeleteRow    

End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Parent.fncPrint()    
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
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False	
		
	If lgBlnFlgChgValue = True then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 

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
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	Call Detail_Sum
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    frm1.txtpDirect.value = pDirect
    
    DbQuery = False                                                              '☜: Processing is NG

'    Call DisableToolBar(TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream(pDirect)

    Select Case pDirect
       Case "M" 
          With Frm1
                 strVal = BIZ_PGM_ID_Q1  & "?txtMode="         & parent.UID_M0001						         
                 strVal = strVal      & "&txtKeyStream="    & lgKeyStream           '☜: Query Key
                 strVal = strVal      & "&txtMaxRows="      & .vspdData.MaxRows
                 strVal = strVal      & "&lgStrPrevKey="    & lgStrPrevKey          '☜: Next key tag
          End With
       Case "Q"
                 strVal = BIZ_PGM_ID_Q1 & "?txtMode="          & parent.UID_M0001            '☜: Query
                 strVal = strVal      & "&txtKeyStream="     & lgKeyStream          '☜: Query Key
    End Select    
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                  '☜:  Run biz logic

    DbQuery = True                                                      '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
        
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()													'☆: 조회 성공후 실행로직	
		
    lgIntFlgMode =  parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
   ' Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field    	
	Call SetToolbar("1100111100111111")	
	
	lgBlnFlgChgValue = False

	IF frm1.txtpDirect.value = "M" Then
		Exit Function
	End IF
	
	Call dbquery2(1,1,"Q")
	
End Function

 '========================================================================================
'    Function Name : InitData()
'    Function Desc : 
'   ======================================================================================== 
Sub InitData()

End Sub

'========================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery2(pRow, pCol,pDirect) 

	Dim strVal
	Dim IntRetCD
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next

    DbQuery2 = False                                                         '⊙: Processing is NG

	If pDirect = "Q" Then 

		ggoSpread.Source = frm1.vspdData2
    
		If ggoSpread.SSCheckChange = True Then    
			IntRetCD = DisplayMsgBox("990027", "X","X","X") '☜ 바뀐부분 
			frm1.vspdData.row = frm1.txtActiveRows.value
			frm1.vspdData.Col = frm1.txtActiveCols.value
			frm1.vspdData.action = 0
			Exit Function
		End If    
		frm1.vspdData2.MaxRows = 0
	End IF
	        
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

	With frm1.vspdData
		.Row = pRow
		.Col = C_AsstNo
		frm1.txtKeyStream_m.value = .value
        lgKeyStream_m = .Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
    End with
    
    
    Select Case pDirect
       Case "M" 
          With Frm1
                 strVal = BIZ_PGM_ID_Q2  & "?txtMode="         & Parent.UID_M0001						         
                 strVal = strVal      & "&txtKeyStream_m="    & lgKeyStream_m           '☜: Query Key
                 strVal = strVal      & "&txtMaxRows_2="      & .vspdData2.MaxRows
                 strVal = strVal      & "&lgStrPrevKey_m="    & lgStrPrevKey_m          '☜: Next key tag
          End With
       Case "Q"
                 strVal = BIZ_PGM_ID_Q2 & "?txtMode="          & Parent.UID_M0001            '☜: Query
                 strVal = strVal      & "&txtKeyStream_m="     & lgKeyStream_m          '☜: Query Key
    End Select    

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Call RunMyBizASP(MyBizASP, strVal)                                  '☜:  Run biz logic

    DbQuery2 = True                                                      '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk2()													'☆: 조회 성공후 실행로직 

    lgIntFlgMode = Parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
	Call SetToolbar("1100111100111111")	
	
	Call Detail_Sum
	
	With frm1.vspdData
		frm1.txtActiveRows.value = .activerow
		.Row = frm1.txtActiveRows.value
'		.Col = C_Next
'		.value = lgStrPrevKey_m
	End with
    
End Function


Function Detail_Sum()
	Dim i
	Dim Sum
	Dim LocSum

	Sum = 0 
	LocSum = 0
	
	With frm1.vspdData2
		for i = 1 to .Maxrows
			.row = i
			.col = C_Amt_2
			
			Sum = UNICDbl(Sum) + UNICDbl(.text)
			.Col = C_LocAmt_2
			
			LocSum = UNICDbl(LocSum) + UNICDbl(.text)
		Next
	End With
	frm1.txtSum.text  = UNIFormatNumber(Sum, Parent.ggAmtOfMoney.DecPoint, -2, 0, Parent.ggAmtOfMoney.RndPolicy, Parent.ggAmtOfMoney.RndUnit)
	frm1.txtLocSum.text  = UNIFormatNumber(LocSum, Parent.ggAmtOfMoney.DecPoint, -2, 0, Parent.ggAmtOfMoney.RndPolicy, Parent.ggAmtOfMoney.RndUnit)
'	frm1.txtSum.text  = Sum
'	frm1.txtLocSum.text = LocSum

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows 
    Dim lGrpcnt 
	Dim strVal, strDel
	Dim IntLocAmt
	
    DbSave = False                                                          '⊙: Processing is NG    

	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value    = Parent.UID_M0002										'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태			
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 
	frm1.txtSum.text  = "0"
	frm1.txtLocSum.text  = "0"

    lGrpCnt = 1    
	strVal = ""
	strDel = ""
    
    '-----------------------------
    '   Acq item Part
    '-----------------------------
    With frm1.vspdData2
	    
    For IntRows = 1 To .MaxRows
    	
		.Row = IntRows
		.Col = 0		
		
		Select Case .Text		    
		        
		    Case ggoSpread.DeleteFlag

		        strDel = strDel & "D" & Parent.gColSep & IntRows & Parent.gColSep

				.Col = C_AsstNo_2
				strDel = strDel & Trim(.value) & Parent.gColSep
					
				.Col = C_Seq_2
				strDel = strDel & Trim(.Text) & Parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 
					
				lGrpcnt = lGrpcnt + 1            
		    
		    Case ggoSpread.UpdateFlag

				strVal = strVal & "U" & Parent.gColSep & IntRows & Parent.gColSep

				.Col = C_AsstNo_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Seq_2
				strVal = strVal & Trim(.Text) & Parent.gColSep
					
				.Col = C_Desc_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Amt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0)  & Parent.gColSep
				
				.Col = C_LocAmt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0)  & Parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 

		        lGrpCnt = lGrpCnt + 1
		        
		    Case ggoSpread.InsertFlag

				strVal = strVal & "C" & Parent.gColSep & IntRows & Parent.gColSep

				.Col = C_AsstNo_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Seq_2
				strVal = strVal & Trim(.Text) & Parent.gColSep
					
				.Col = C_Desc_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Amt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0) & Parent.gColSep
					
				.Col = C_LocAmt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0) & Parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 

		        lGrpCnt = lGrpCnt + 1

		End Select

    Next

	End With
	
	frm1.txtMaxRows_2.value  = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread_m.value = strDel & strVal									'☜: Spread Sheet 내용을 저장 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'☜: 저장 비지니스 ASP 를 가동 

    DbSave = True                                                           ' ⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
   ' Call InitVariables	
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    'lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey_m = 0                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    '-----------------------
    'Erase contents area
    '-----------------------
    frm1.vspdData2.MaxRows = 0
	call dbquery2(frm1.txtActiveRows.value,1,"Q")
	
End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

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
	Dim indx

	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()

		Case "VSPDDATA2"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
	End Select
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

Sub vspdData_Change(Col , Row)


End Sub


'========================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_Change(Col , Row)

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
	If Col = C_Amt_2 or Col = C_LocAmt_2 Then
		Call Detail_Sum
	End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	If  lgIntFlgMode =  parent.OPMD_UMODE Then
		Call SetPopUpMenuItemInf("1111111111")
    Else
		Call SetPopUpMenuItemInf("0000111111")    
    End if
    
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
    End If

    
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData2_Click(Col, Row)
	If  lgIntFlgMode =  parent.OPMD_UMODE Then
		Call SetPopUpMenuItemInf("1111111111")
    Else
		Call SetPopUpMenuItemInf("0000111111")    
    End if

    gMouseClickStatus = "SP2C"
    Set gActiveSpdSheet = frm1.vspdData2
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
    End If

    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'======================================================================================================
'   Event Name : vspdData2_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("B")
End Sub



'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================


Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData
	
		If col < 1 or Row < 1 or NewCol < 1 or NewRow < 1 Then
			Exit Sub
		End IF
				
		If Row = NewRow Then
		    Exit Sub
		End If
		 
			frm1.txtActiveRows.value = NewCol
			frm1.txtActiveCols.value = NewRow

		Call Dbquery2(NewRow, NewCol, "Q")
		
    End With
    
End Sub

Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

Sub  vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData2
	
		If Newcol < 1 or NewRow < 1 Then
			frm1.txtActiveRows_m.value = NewRow
			Exit Sub
		End If
				
		If Row = NewRow Then
		    Exit Sub
		End If
		
    End With
    
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If



    if frm1.vspdData.MaxRows < (NewTop + VisibleRowCnt(frm1.vspdData,NewTop)) Then	
    	If lgStrPrevKey <> "" Then  
           If DbQuery("M") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData2.MaxRows < (NewTop + VisibleRowCnt(frm1.vspdData2,NewTop)) Then	 
    	If lgStrPrevKey_m <> 0 Then   
           If DbQuery2(frm1.txtActiveRows.value,frm1.txtActiveCols.value, "M") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고정자산상세취득내역등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
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
									<TD CLASS="TD5" NOWRAP>취득번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtAcqNo" TYPE="Text" MAXLENGTH=18 tag="12XXXU" ALT="취득번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenAcqNoInfo"></TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%></TD>
				</TR>
				<TR HEIGHT=100%>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" MAXLENGTH=3 SIZE=10 tag="24XXXU" ></TD>
							    <TD CLASS="TD5" NOWRAP>취득일자</TD>																							    
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/a7125ma1_fpDateTime1_txtAcqDt.js'></script>											    
								</TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="거래처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU">
													<INPUT NAME="txtBpNm" TYPE="Text" SIZE = 22 tag="24">
								</TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a7125ma1_fpDoubleSingle1_txtXchRate.js'></script>
	                            </TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=4></TD>
							</TR>
							
							<TR>
								<TD WIDTH="100%" HEIGHT=45% COLSPAN=4>
									<script language =javascript src='./js/a7125ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=4></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT=40% COLSPAN=4>
									<script language =javascript src='./js/a7125ma1_vspdData2_vspdData2.js'></script>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=4></TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100% COLSPAN=4>
									<FIELDSET CLASS="CLSFLD">
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>상세내역합계</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/a7125ma1_fpDoubleSingle2_txtSum.js'></script>&nbsp;
												</TD>
												<TD CLASS="TD5" NOWRAP>상세내역합계(자국)</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/a7125ma1_fpDoubleSingle3_txtlocSum.js'></script>&nbsp;
												</TD>
											</TR>
										</TABLE>
									</FIELDSET>
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
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="htxtAcqNo"    tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMode"      tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtActiveRows" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtActiveCols" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtActiveRows_m" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows_2" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows_3" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream_m"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpread_m"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"   tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtpDirect"   tag="24" TABINDEX = "-1" >

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

