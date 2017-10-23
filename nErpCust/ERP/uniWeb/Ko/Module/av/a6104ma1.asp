
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!--*******************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : 부가세관리 
'*  3. Program ID        : a6104ma1
'*  4. Program 이름      : 부가세내역조회 
'*  5. Program 설명      : 부가세를 건별로 조회한다.
'*  6. Comproxy 리스트   : a6104ma1
'*  7. 최초 작성년월일   : 2000/04/22
'*  8. 최종 수정년월일   : 2001/01/17
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'*                         -2000/04/22 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncServer.asp"  -->
<!--========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/AdoQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a6104mb1.asp"			'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns
Const C_IssuedDT  = 1
Const C_IOFGCD    = 2
Const C_IOFGNM    = 3
Const C_BPCD      = 4
Const C_BPNM      = 5
Const C_OwnRGSTNo = 6
Const C_NetAmt    = 7
Const C_VatAmt    = 8
Const C_VatTypeCD = 9
Const C_VatTypeNM = 10

Const C_SHEETMAXROWS = 50		' : 한 화면에 보여지는 최대갯수*1.5

<%
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	EndDate = GetSvrDate
	Call ExtractDateFrom(EndDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)

	StartDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, "01")
	EndDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)
%>

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
Dim lgStrPrevKeyISSUEDT
Dim lgStrPrevKeyGLNO

'Dim lgLngCurRows

Dim lgBlnStartFlag				' 메세지 관련하여 프로그램 시작시점 Check Flag

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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

    lgIntFlgMode = OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = 0                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
	
	lgSortKey = 1
	
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

	frm1.txtIssueDT1.Text = "<%=StartDate%>"
	frm1.txtIssueDT2.Text = "<%=EndDate%>"

	
	frm1.txtBizAreaCD.value	= gBizArea
	lgBlnStartFlag = False
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call loadInfTB19029(gCurrency, "Q", "A")%>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
        
	With frm1.vspdData
	
		.MaxCols = C_VatTypeNM + 1
		.Col = .MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
		.Col = C_IOFGCD
		.ColHidden = True
		.Col = C_BPCD
		.ColHidden = True
		.Col = C_VatTypeCD
		.ColHidden = True
		.MaxRows = 0

		ggoSpread.Source = frm1.vspdData

		.ReDraw = False

		ggoSpread.SpreadInit 

		ggoSpread.SSSetDate  C_IssuedDT, "계산서일", 11, 2, gDateFormat
		ggoSpread.SSSetCombo C_IOFGCD,   "", 10
		ggoSpread.SSSetCombo C_IOFGNM,   "입출", 10, 2
<%
		Call InitComboBoxDtl("2", "A1003")		' 입출구분 
%>
		ggoSpread.SSSetEdit  C_BPCD,      "",   20, , , 40
		ggoSpread.SSSetEdit  C_BPNM,      "거래처명", 20, , , 40
	    ggoSpread.SSSetEdit  C_OwnRGSTNo, "사업자등록번호", 20, , , 20
	    ggoSpread.SSSetFloat C_NetAmt,    "공급가", 19, ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
	    ggoSpread.SSSetFloat C_VatAmt,    "부가세", 19, ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetCombo C_VatTypeCD, "계산서유형", 18
		ggoSpread.SSSetCombo C_VatTypeNM, "계산서유형", 18
<%
		Call InitComboBoxDtl("3", "B9001")		' 부가세유형 
%>
		.ReDraw = True

		Call SetSpreadLock                                              '바뀐부분 
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()

    With frm1.vspdData
		.ReDraw = False

		ggoSpread.SpreadLock C_IssuedDT, -1, C_VatTypeNM
		
		.ReDraw = True
    End With

End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal lRow)
End Sub


'=============================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'============================================================================================================ 
Function InitComboBox()
<%
		Call InitComboBoxDtl("0", "A1003")		' 입출구분 
		Call InitComboBoxDtl("1", "B9001")		' 부가세유형 
%>

End Function

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
<%
Function InitComboBoxDtl(Byval Index, Byval MajorCd)

   ' Dim B1a028
    Dim intMaxRow
    Dim intLoopCnt
	Dim strListCd
	Dim strListNm
    
    Err.Clear                                                               '☜: Clear error no
	On Error Resume Next

	'Set B1a028 = Server.CreateObject("B1a028.B1a028ListMinorCode")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
'	If Err.Number <> 0 Then
'		Set B1a028 = Nothing												'☜: ComProxy Unload
'		Call MessageBox(Err.description, I_INSCRIPT)						'⊙:
'		Response.End														'☜: 비지니스 로직 처리를 종료함 
'	End If

 '   B1a028.ImportBMajorMajorCd = Trim(MajorCd)									'⊙: Major Code
  '  B1a028.ServerLocation = ggServerIP
    
  '  B1a028.ComCfg = gConnectionString
  '  B1a028.Execute															'☜:
    
    '-----------------------
    'Com action result check area(DB,internal)
    '-----------------------
 '   If Not (B1a028.OperationStatusMessage = MSG_OK_STR) Then
'		Call MessageBox(B1a028.OperationStatusMessage, I_INSCRIPT)         '☆: you must release this line if you change msg into code
'		Set B1a028 = Nothing												'☜: ComProxy Unload
'		Response.End														'☜: 비지니스 로직 처리를 종료함 
 '   End If

'	intMaxRow = B1a028.ExportGroupCount
	strListCd = ""
	strListNm = ""
	
	Select Case Index
		Case "0"	' 입출구분 
			
%>
				Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				Call SetCombo2(frm1.cboIOFlag ,lgF0  ,lgF1  ,Chr(11))
<%
			'Next
		Case "1"	' 부가세유형 
			'For intLoopCnt = 1 To intMaxRow
%>
	    		 Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("B9001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
				Call SetCombo2(frm1.cboVatType ,lgF0  ,lgF1  ,Chr(11))
<%
			'Next
		Case "2"	' 입출구분 
%>		
			'For intLoopCnt = 1 To intMaxRow
			'	If intLoopCnt <> intMaxRow Then
			'		strListCd = strListCd & B1a028.ExportItemBMinorMinorCd(intLoopCnt) & vbtab
			'		strListNm = strListNm & B1a028.ExportItemBMinorMinorNm(intLoopCnt) & vbtab
			'	ElseIf intLoopCnt = intMaxRow Then
			'		strListCd = strListCd & B1a028.ExportItemBMinorMinorCd(intLoopCnt)
			'		strListNm = strListNm & B1a028.ExportItemBMinorMinorNm(intLoopCnt)
			'	End If
			'Next  
			
			Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		    ggoSpread.Source = frm1.vspdData
			ggoSpread.SetCombo lgF0, C_IOFGCD		' 문자 
			ggoSpread.SetCombo lgF1, C_IOFGNM		' 문자 
<%
		Case "3"	' 부가세유형 
			'For intLoopCnt = 1 To intMaxRow
			'	If intLoopCnt <> intMaxRow Then
			'		strListCd = strListCd & B1a028.ExportItemBMinorMinorCd(intLoopCnt) & vbtab
			'		strListNm = strListNm & B1a028.ExportItemBMinorMinorNm(intLoopCnt) & vbtab
			'	ElseIf intLoopCnt = intMaxRow Then
			'		strListCd = strListCd & B1a028.ExportItemBMinorMinorCd(intLoopCnt)
			'		strListNm = strListNm & B1a028.ExportItemBMinorMinorNm(intLoopCnt)
			'	End If
			'Next  
%>
		    ggoSpread.Source = frm1.vspdData
			ggoSpread.SetCombo "<%=strListCd%>", C_VatTypeCD		' 문자 
			ggoSpread.SetCombo "<%=strListNm%>", C_VatTypeNM		' 문자 
<%
	End Select

	Set B1a028 = Nothing                        

End Function
%>

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

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"	 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(0)
    
			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(0)
		Case 1
			arrParam(0) = "거래처 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래처"					' 조건필드의 라벨 명칭 

			arrField(0) = "BP_CD"						' Field명(0)
			arrField(1) = "BP_NM"						' Field명(1)
    
			arrHeader(0) = "거래처코드"				' Header명(0)
			arrHeader(1) = "거래처명"				' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' 세무서 
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
				
				.txtBizAreaCD.focus
			Case 1		' 거래처 
				.txtBPCd.value = UCase(Trim(arrRet(0)))
				.txtBPNM.value = arrRet(1)
				
				.txtBPCd.focus
		End Select
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

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)

	' 현재 Page의 Form Element들을 Clear한다. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	' 현재 Page의 From Element들을 사용자가 입력을 받지 못하게 하거나 필수입력사항을 표시한다.
	' LockField(pDoc, pACode)
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    
    Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
    Call InitVariables                            '⊙: Initializes local global Variables
    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox

	Call SetDefaultVal

	' [Main Menu ToolBar]의 각 버튼을 [Enable/Disable] 처리하는 부분 
    Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어 

    frm1.txtIssueDT1.focus
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
Sub txtIssueDt1_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub txtIssueDt2_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			
			' 입출 
			.Col = C_IOFGCD
			intIndex = .value
			.col = C_IOFGNM
			.value = intindex
			' 부가세유형 
			.Col = C_VatTypeCD
			intIndex = .value
			.col = C_VatTypeNM
			.value = intindex
					
		Next	
	End With
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt2_Change()
    'lgBlnFlgChgValue = True
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)
    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    
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
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------   

    End With

End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> 0 Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
    End if

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
    
    FncQuery = False          '⊙: Processing is NG
    Err.Clear                 '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
	' 현재 Page의 Form Element들을 Clear한다. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear
    Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
    Call InitVariables							'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
	' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then	'⊙: This function check indispensable field
       Exit Function
    End If
 
    If UniCDate(frm1.txtIssueDt1.text) > UniCDate(frm1.txtIssueDt2.text) Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If
 
	If frm1.txtBPCd.value = "" Then
		frm1.txtBPNm.value = ""
	End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
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
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
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
    Call Parent.FncExport(C_MULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
	
	iColumnLimit = 10
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
	If lgBlnStartFlag = True Then
		' 변경된 내용이 있는지 확인한다.
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
	
			If IntRetCD = vbNo Then
				Exit Function
			End If
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

    DbQuery = False
    Err.Clear                '☜: Protect system from crashing
    
    With frm1
    
		Call LayerShowHide(1)
	
	    If lgIntFlgMode = OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001
			strVal = strVal & "&txtIssueDT1=" & (Trim(.hIssueDT1.value))
			strVal = strVal & "&txtIssueDT2=" & (Trim(.hIssueDT2.value))
			strVal = strVal & "&cboVatType=" & Trim(.hVatType.value)
			strVal = strVal & "&cboIOFlag=" & Trim(.hIOFlag.value)
			strVal = strVal & "&txtBizAreaCd=" & UCase(Trim(.hBizAreaCd.value))
			strVal = strVal & "&txtBPCd=" & UCase(Trim(.hBPCd.value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001
			strVal = strVal & "&txtIssueDT1=" & (Trim(.txtIssueDT1.text))
			strVal = strVal & "&txtIssueDT2=" & (Trim(.txtIssueDT2.text))
			strVal = strVal & "&cboVatType=" & Trim(.cboVatType.value)
			strVal = strVal & "&cboIOFlag=" & Trim(.cboIOFlag.value)
			strVal = strVal & "&txtBizAreaCd=" & UCase(Trim(.txtBizAreaCd.value))
			strVal = strVal & "&txtBPCd=" & UCase(Trim(.txtBPCd.value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
		    
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = OPMD_UMODE	'⊙: Indicates that current mode is Update mode
    
	lgBlnFlgChgValue = False
	
	lgBlnStartFlag = True		' 메세지 관련하여 프로그램 시작시점 Check Flag
	
	' 현재 Page의 From Element들을 사용자가 입력을 받지 못하게 하거나 필수입력사항을 표시한다.
	' LockField(pDoc, pACode)
    Call ggoOper.LockField(Document, "Q")	'⊙: This function lock the suitable field

    Call InitData	' Combo의 Name을 Code를 기준으로 맞춤 

    Call SetToolbar("1100000000011111")										'⊙: 버튼 툴바 제어 

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete()
	On Error Resume Next
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>부가세내역조회</font></td>
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
									<TD CLASS="TD5">발행일</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6104ma1_fpDateTime2_txtIssueDt1.js'></script>&nbsp;~&nbsp;
													<script language =javascript src='./js/a6104ma1_fpDateTime2_txtIssueDt2.js'></script></TD>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="신고사업장" tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
													<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="신고사업장" tag="14X" ></TD>
									<TD CLASS="TD5">거래처</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBPCd" NAME="txtBPCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBPCd.Value, 1)">&nbsp;
													<INPUT TYPE=TEXT ID="txtBPNm" NAME="txtBPNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">입출구분</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag" NAME="cboIOFlag" ALT="입출구분" STYLE="WIDTH: 98px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5">부가세유형</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT ID="cboVatType" NAME="cboVatType" ALT="부가세유형" STYLE="WIDTH: 130px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN=7 >
								<script language =javascript src='./js/a6104ma1_vaSpread1_vspdData.js'></script></TD>
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100% COLSPAN=7></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>매입처  :</TD>
								<TD CLASS="TD18">매수합계</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingle1_txtCntSumI.js'></script></TD>
								<TD CLASS="TD18">공급가합계</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingle1_txtAmtSumI.js'></script></TD>
								<TD CLASS="TD18">부가세합계</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingle1_txtVatSumI.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>매출처  :</TD>
								<TD CLASS="TD18">매수합계</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingleO_txtCntSumO.js'></script></TD>
								<TD CLASS="TD18">공급가합계</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingleO_txtAmtSumO.js'></script></TD>
								<TD CLASS="TD18">부가세합계</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingleO_txtVatSumO.js'></script></TD>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="hIssueDT1" tag="24">
<INPUT TYPE=HIDDEN NAME="hIssueDT2" tag="24">
<INPUT TYPE=HIDDEN NAME="hVatType" tag="24">
<INPUT TYPE=HIDDEN NAME="hIOFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="hBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBPCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
