
<%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1102ma2.asp
'*  4. Program Name         : Calendar Adjustment
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/09
'*  8. Modified date(Last)  : 2002/05/09
'*  9. Modifier (First)     : Mr  KimGyoungDon
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!--
========================================================================================================
=                          1.1.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<STYLE TYPE="text/css">
	.Header {height:24; font-weight:bold; text-align:center; color:darkblue}
	.Day {height:22;cursor:Hand;
		font-size:17; font-weight:bold; Border:0; text-align:right}
	.DummyDay {height:22;cursor:;
		font-size:12; font-weight:; Border:0; text-align:right}
</STYLE>
<MAP NAME="CalButton">
	<AREA SHAPE=RECT COORDS="1, 1, 20, 20" ALT="Year -" onClick="ChangeMonth(-12)">
	<AREA SHAPE=RECT COORDS="20, 1, 40, 20" ALT="Month -" onClick="ChangeMonth(-1)">
	<AREA SHAPE=RECT COORDS="40, 1, 60, 20" ALT="Month +" onClick="ChangeMonth(1)">
	<AREA SHAPE=RECT COORDS="60, 1, 80, 20" ALT="Year +" onClick="ChangeMonth(12)">
</MAP>

<!--==========================================  1.1.2 공통 Include   ======================================
==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit            '☜: indicates that All variables must be declared in advance

Dim BaseDate
Dim StartDate
Dim strYear
Dim strMonth
DIm strDay


<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, Parent.gServerDateFormat, Parent.gDateFormat)
Call ExtractDateFrom(BaseDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

Const BIZ_PGM_QRY_ID = "p1102mb1.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "p1102mb3.asp"											'☆: 비지니스 로직 ASP명 

Const CChnageColor = "#f0fff0"
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim lgLastDay
Dim lgStartIndex
Dim lgArrDate(31, 3)
Dim IsOpenPop
Dim lgChgCboYear
Dim lgChgCboMonth          

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

    lgIntFlgMode = Parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    lgChgCboYear = False
    lgChgCboMonth = False
    '----------  Coding part  -------------------------------------------------------------

	Dim iRow, iCol
	For iRow = 1 To 6
		For iCol = 1 To 7
			If frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = CChnageColor Then
				frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = "white"
				frm1.txtDate((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtDesc((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
			End If
		Next
	Next

End Sub

'=============================== 2.1.2 LoadInfTB19029() =================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029() 
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call LoadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub 

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : ComboBox 초기화 
'=========================================================================================================
Sub InitComboBox()
	Dim i, ii
	Dim oOption
	
	For i = (StrYear - 10) To (StrYear + 20)
		Call SetCombo(frm1.cboYear, i, i)
	Next

    frm1.cboYear.value = StrYear
    
	For i=1 To 12
		ii = Right("0" & i, 2)
		Call SetCombo(frm1.cboMonth, ii, ii)
	Next

    frm1.cboMonth.value = Right("0" & StrMonth, 2)
End Sub
'==========================================  2.2.2 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'*********************************************************************************************************


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

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "칼렌다 타입 팝업"			<%' 팝업 명칭 %>
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtClnrType.Value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "칼렌다 타입"					<%' TextBox 명칭 %>
	
    arrField(0) = "CAL_TYPE"						<%' Field명(0)%>
    arrField(1) = "CAL_TYPE_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "칼렌다 타입"				<%' Header명(0)%>
    arrHeader(1) = "칼렌다 타입명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'------------------------------------------  SetCalType()  -----------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================

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
'**********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11000000000001")
    Call SetDefaultVal
    Call InitComboBox															'⊙: Initialize combobox at load time
    Call InitVariables		
    
    Call ggoOper.SetReqAttr(frm1.txtClnrType,"N")
    Call ggoOper.SetReqAttr(frm1.txtClnrTypeNm,"Q") 
    
    frm1.txtClnrType.focus
    Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : DescChange
'   Event Desc : Remark Change
'==========================================================================================

Sub DescChange(iDate)
	Dim strDesc
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If
	
	strDesc = frm1.txtDesc(index).value
	frm1.txtDesc(index).value = ""
	
	frm1.txtDesc(index).value = strDesc
	frm1.txtDesc(index).title = strDesc

	Call SetChange(iDate)
End Sub

'==========================================================================================
'   Event Name : HoliChange
'   Event Desc : Holiday Change
'==========================================================================================

Sub HoliChange(iDate)

	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If
	
	'If UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtYear.value, frm1.txtMonth.value, frm1.txtDate(index).value) < StartDate Then	
	'If UniConvDateAToB((frm1.txtYear.value & "-" & frm1.txtMonth.value & "-" & frm1.txtDate(index).value), Parent.gServerDateFormat, Parent.gDateFormat) < StartDate Then
	If UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtYear.value, frm1.txtMonth.value, frm1.txtDate(index).value) < BaseDate Then	
		Call DisplayMsgBox("180215","X","X","X")
		Exit Sub
	End If

	If frm1.txtHoli(index).value = "0" Then
		frm1.txtDate(index).style.color = "black"
		frm1.txtHoli(index).value = "2"
	ElseIf frm1.txtHoli(index).value = "1" Then
		frm1.txtDate(index).style.color = "red"
		frm1.txtHoli(index).value = "0"
	Else
		frm1.txtDate(index).style.color = "blue"
		frm1.txtHoli(index).value = "1"
	End if

	Call SetChange(iDate)
End Sub

'==========================================================================================
'   Event Name : SetChange
'   Event Desc : Color Change
'==========================================================================================

Sub SetChange(iDate)
	Dim index
	index = iDate - 1

	lgBlnFlgChgValue = True
	
	frm1.All.tblCal.Rows(Int((index+7)/7)).Cells(index Mod 7).Style.backgroundColor = CChnageColor
	frm1.txtDate(index).Style.backgroundColor = CChnageColor
	frm1.txtDesc(index).Style.backgroundColor = CChnageColor
End Sub

'==========================================================================================
'   Event Name : ChangeMonth
'   Event Desc : 화살표 클릭 
'==========================================================================================

Sub ChangeMonth(i)
    Dim strVal
    Dim dtDate
    Dim IntRetCD
    Dim StrYear1
    Dim StrMonth1
    Dim StrDay1
	   
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X") 
		If IntRetCD = vbNo Then
			Exit Sub
		End If
    End If

    Call InitVariables						
	
	On Error Resume Next
	Err.Clear
	
    dtDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, frm1.txtYear.value, frm1.txtMonth.value, "01")
    
    If Err.Number <> 0 Then                        
        Err.Clear
		Call DisplayMsgBox("900002","X","X","X")
        Exit Sub
    End If

	dtDate = UNIDateAdd("m", i, dtDate, Parent.gDateFormat)
	Call ExtractDateFrom(dtDate, Parent.gDateFormat, Parent.gComDateType, strYear1, strMonth1, strDay1)
	
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
    strVal = strVal & "&txtClnrType=" & Trim(frm1.txtClnrType.value)
    strVal = strVal & "&txtYear=" & StrYear1					'☆: 조회 조건 데이타 
    strVal = strVal & "&txtMonth=" & StrMonth1		'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	LayerShowHide(1)
											'⊙: 작업진행중 표시	
	Call RunMyBizASP(MyBizASP, strVal)

End Sub

'==========================================================================================
'   Event Name : CboYear_OnChange
'   Event Desc : Combo Change
'==========================================================================================

Function CboYear_OnChange()
	Dim IntRetCD

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			frm1.cboYear.value = frm1.txtYear.value
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
		
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
End Function

'==========================================================================================
'   Event Name : CboMonth_OnChange
'   Event Desc : Combo Change
'==========================================================================================

Function CboMonth_OnChange()
    Dim IntRetCD
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			frm1.cboMonth.value = frm1.txtMonth.value 
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
		
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                        '⊙: Processing is NG
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
		
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     											'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
        
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                            '⊙: No data changed!!        
        Exit Function
       
    End If
    
   If Not chkField(Document, "2") Then                             
       Exit Function
    End If												

    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     								                                                  '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()     
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)								  '☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                          '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSpliteColumn
' Function Desc : This function is related to FncSpliteColumn menu item of Main menu
'========================================================================================
Function FncSpliteColumn() 
       On Error Resume Next
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
       On Error Resume Next
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
       On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
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
    DbQuery = False                                                         '⊙: Processing is NG
    
    LayerShowHide(1)
		
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
    strVal = strVal & "&txtClnrType=" & Trim(frm1.txtClnrType.value)	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtYear=" & Trim(frm1.cboYear.value)	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtMonth=" & Trim(frm1.cboMonth.Value)	'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()													'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False  
    
    Call SetToolbar("11001000000101")
    
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
	
	Dim ColCol
    
    Err.Clear																	'☜: Protect system from crashing

	DbSave = False																'⊙: Processing is NG
	
	LayerShowHide(1)
											'⊙: 작업진행중 표시 
	'-----------------------
    'Check content area
    '-----------------------

	'-------------------------------------------------------------
	' 현재일 이전은 disable되어 있어 biz asp로 넘어가지 않는다.
	' 따라서 임시로 disable을 enable로 변경시킨다.
	'-------------------------------------------------------------
	For ColCol = 0 To 41
		If frm1.txtDate(ColCol).className <> "DummyDay" Then
			frm1.txtDesc(ColCol).disabled = False
		End If
	Next	

	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	frm1.txtUpdtUserId.value = Parent.gUsrID
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)									'☜: 비지니스 ASP 를 가동 
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

    Call InitVariables
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>생산칼렌다수정</font></td>
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
					<TD>
						<TABLE ID="tbTitle" WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="center">
							<TR>
								<TD CLASS=TD5 NOWRAP>칼렌다 타입</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="12XXXU" ALT="칼렌다 타입"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=40 tag="14"></TD>
								<TD STYLE="TEXT-ALIGN: RIGHT"><IMG SRC="../../../CShared/image/CalButton.gif" WIDTH=80 HEIGHT=20 style="cursor:Hand" ISMAP USEMAP="#CalButton"></IMG>&nbsp;</TD>
								<TD WIDTH=10% STYLE="TEXT-ALIGN:RIGHT"><SELECT Name="cboYear" STYLE="WIDTH=60"></SELECT>&nbsp;<SELECT Name="cboMonth" STYLE="WIDTH=40"></SELECT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD>
						<TABLE ID="tblCal" WIDTH=100% HEIGHT=100% BORDER=1 CELLSPACING=0 CELLPADDING=0 ALIGN="center">
							<THEAD CLASS="Header">
								<TR>
									<TD>일요일</TD>
									<TD>월요일</TD>
									<TD>화요일</TD>
									<TD>수요일</TD>
									<TD>목요일</TD>
									<TD>금요일</TD>
									<TD>토요일</TD>
					            </TR>
				        	</THEAD>
							<TBODY>
								<%
								Dim i, j, k
								k = 1
								For i=1 To 6
								%>
					            <TR>
									<%
										For j=1 To 7
									%>
									<TD ALIGN="Center">
										<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="Center">
											<TR>
												<TD ALIGN="Left">
													<INPUT type="hidden" name="txtHoli" size=1 maxlength=1 disabled>
													<INPUT type="text" name="txtDate" class="DummyDay" size=2 maxlength=2  
														tabindex=-1 readonly disabled onclick="HoliChange(<%=k%>)">
												</TD>
											</TR>
											<TR>
												<TD ALIGN="Left">
													<INPUT type="text" name="txtDesc"  MaxLength=20 Style="Width:100%;Border:0;text-align:center" disabled tag=2 onchange="DescChange(<%=k%>)" ALT="비고">
												</TD>
											</TR>
										</TABLE>
									</TD>
									<%
											k = k + 1
										Next
									%>
								</TR>
								<%
								Next
								%>
							</TBODY>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_01%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtYear" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMonth" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
