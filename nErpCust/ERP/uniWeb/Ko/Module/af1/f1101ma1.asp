<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f1101ma1
'*  4. Program Name         : 예산통제기간등록 
'*  5. Program Desc         : Register of Control Period
'*  6. Comproxy List        : FB0011, FB0019
'*  7. Modified date(First) : 2000.09.18
'*  8. Modified date(Last)  : 2003.06.13
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Oh, Soo Min
'* 11. Comment              :
'* - 2003/06/13 Oh, Soo Min ClearField속성 변경 
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################%>
<% '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit                                   '☜: indicates that All variables must be declared in advance %>


'********************************************  1.2 Global 변수/상수 선언  *********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'============================================  1.2.1 Global 상수 선언  ====================================
'==========================================================================================================
Const BIZ_PGM_ID = "f1101mb1.asp"											 '☆: 비지니스 로직 ASP명 

'============================================  1.2.2 Global 변수 선언  ===================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2. Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

Dim IsOpenPop       
Dim strCtrlYear
Dim strCtrlMonth
Dim strCtrlDay

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

    if frm1.cboCtrlUnit.length > 0 then
       frm1.cboCtrlUnit.selectedindex = 0
    end if

	frm1.fpCtrlYR.Text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat) 
		
	Call ggoOper.FormatDate(frm1.txtCtrlYR,  Parent.gDateFormat, 3)
    Call ggoOper.FormatDate(frm1.txt1stFrYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt1stToYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt2ndFrYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt2ndToYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt3rdFrYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt3rdToYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt4thFrYR, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txt4thToYR, Parent.gDateFormat, 2)
End Sub

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'☆: 사용자 변수 초기화 
		
End Sub


'============================================= 2.1.2 LoadInfTB19029() ====================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================================= 
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I", "*","NOCOOKIE" ,"MA") %>
	
End Sub


'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 



'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'********************************************************************************************************* 

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox()
	
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2010", "''", "S") & "  AND MINOR_CD <> " & FilterVar("M", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboctrlunit ,lgF0  ,lgF1  ,Chr(11))
	

End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 


'###########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################


'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'**********************************************************************************************************
'=======================================================================================================
'   Event Name : txt1StPrRdpDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtCtrlYR_DblClick(Button)
    If Button = 1 Then
		frm1.txtCtrlYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtCtrlYR.Focus       
    End If
End Sub

Sub txt1stFrYR_DblClick(Button)
    If Button = 1 Then
        frm1.txt1stFrYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt1stFrYR.Focus     

    End If
End Sub

Sub txt1stToYR_DblClick(Button)
    If Button = 1 Then
        frm1.txt1stToYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt1stToYR.Focus             
    End If
End Sub

Sub txt2ndFrYR_DblClick(Button)
    If Button = 1 Then
		 frm1.txt2ndFrYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt2ndFrYR.Focus      
    End If
End Sub

Sub txt2ndToYR_DblClick(Button)
    If Button = 1 Then
		frm1.txt2ndToYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt2ndToYR.Focus      	            
    End If
End Sub

Sub txt3rdFrYR_DblClick(Button)
    If Button = 1 Then
       	frm1.txt3rdFrYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt3rdFrYR.Focus      	         
    End If
End Sub

Sub txt3rdToYR_DblClick(Button)
    If Button = 1 Then
		frm1.txt3rdToYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt3rdToYR.Focus         
    End If
End Sub

Sub txt4thFrYR_DblClick(Button)
    If Button = 1 Then
		frm1.txt4thFrYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt4thFrYR.Focus                
    End If
End Sub

Sub txt4thToYR_DblClick(Button)
    If Button = 1 Then
		frm1.txt4thToYR.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txt4thToYR.Focus             
        
    End If
End Sub

'=======================================================================================================
'   Event Name : KeyDown 이벤트 
'   Event Desc : 엔터키를 치면 FncQuery 호출 
'=======================================================================================================
Sub txtCtrlYR_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

'=======================================================================================================
'   Event Name : Change 이벤트 
'   Event Desc : 데이터 변경플래그 수정 
'=======================================================================================================
Sub txt1stFrYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt1stToYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt2ndFrYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt2ndToYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt3rdFrYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt3rdToYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt4thFrYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt4thToYR_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub cboctrlunit_OnChange()
    'Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    
    frm1.txt1stFrYR.value =""
    frm1.txt1stToYR.value =""
    frm1.txt2ndFrYR.value =""
    frm1.txt2ndToYR.value =""
    frm1.txt3rdFrYR.value =""
    frm1.txt3rdToYR.value =""
    frm1.txt4thFrYR.value =""
    frm1.txt4thToYR.value =""
    
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    lgBlnFlgChgValue = False
    lgIntFlgMode = Parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    
End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'==========================================================================================================
Sub Form_Load()

    
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	      
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    Call InitComboBox
    Call SetDefaultVal
    '----------  Coding part  -------------------------------------------------------------
    Call FncSetToolBar("New")
    Call InitVariables																'⊙: Initializes local global variables
    
    frm1.txtCtrlYR.focus
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
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


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ExtractDateFrom(frm1.txtCtrlYR.Text,frm1.txtCtrlYR.UserDefinedFormat,Parent.gComDateType,strCtrlYear,strCtrlMonth,strCtrlDay)
            
    '-----------------------
    'Erase contents area
    '----------------------- 
    'Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not ChkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    Call InitVariables															'⊙: Initializes local global variables
 '   Call SetDefaultVal  
    Call FncSetToolBar("New")
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    Set gActiveElement = document.activeElement
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                     '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	    
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Contents  Field        
  
    Call SetDefaultVal 
    Call InitVariables															'⊙: Initializes local global variables
    Call cboCtrlUnit_Change()
    
    Call FncSetToolBar("New")
  
    frm1.txtCtrlYR.focus
  
    FncNew = True																'⊙: Processing is OK
    Set gActiveElement = document.activeElement  
    
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD
    
    FncDelete = False														'⊙: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                
        Exit Function
    End If
    
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO,"X","X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
' Modify Date : 2001-12-04
' Modify Contents : date format 표준 반영 
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strStartYear,strStartMonth,strStartDay
    Dim TempFiscStart,TempFiscEnd
    
    On Error Resume Next

    TempFiscStart = UniConvDateToYYYYMMDD(Parent.gFiscStart,Parent.gAPDateFormat,Parent.gServerDateType)
    TempFiscEnd   = UniConvDateToYYYYMMDD(Parent.gFiscEnd  ,Parent.gAPDateFormat,Parent.gServerDateType)

    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '⊙: No data changed!!
        Exit Function
    End If
 
    '-----------------------
    'Check content area
    '-----------------------
	Call ExtractDateFrom(Parent.gFiscStart,Parent.gAPDateFormat, Parent.gAPDateSeperator,strStartYear,strStartMonth,strStartDay)
	Call ExtractDateFrom(frm1.txtCtrlYR.Text,frm1.txtCtrlYR.UserDefinedFormat,Parent.gComDateType,strCtrlYear,strCtrlMonth,strCtrlDay)
		
	frm1.txthCtrlYr.Value = strCtrlYear

	TempFiscStart = strCtrlYear + "-" + mid(TempFiscStart,6,2) + "-" + mid(TempFiscStart,9,2)
	TempFiscEnd   = strCtrlYear + "-" + mid(TempFiscEnd,6,2) + "-" + mid(TempFiscEnd,9,2)
	
'	If frm1.txthCtrlYr.Value <> strStartYear Then
'		Call DisplayMsgBox("140127","X",frm1.txtCtrlYR.Alt,"X")
'		frm1.txtCtrlYR.focus
'		Exit Function
'	End If

	If frm1.txt1stFrYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt1stFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt1stFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
           Call DisplayMsgBox("140127","X", frm1.txt1stFrYR.Alt , "X")
           Exit Function
        End If   
	End If 
	
	If frm1.txt1stToYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt1stToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt1stToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
           Call DisplayMsgBox("140127","X", frm1.txt1stToYR.Alt, "X")
           Exit Function
        End If   
	End If 
	If frm1.txt2ndFrYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt2ndFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt2ndFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
           Call DisplayMsgBox("140127","X", frm1.txt2ndFrYR.Alt, "X")
           Exit Function
        End If   
	End If 

	If frm1.txt2ndToYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt2ndToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt2ndToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
           Call DisplayMsgBox("140127","X",frm1.txt2ndToYR.Alt, "X")
           Exit Function
        End If   
	End If       

	If frm1.txt3rdFrYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt3rdFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt3rdFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
			Call DisplayMsgBox("140127","X", frm1.txt3rdFrYR.Alt, "X")
          Exit Function
	    End If   
	End If 

	If frm1.txt3rdToYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt3rdToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt3rdToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
	    Call DisplayMsgBox("140127","X", frm1.txt3rdToYR.Alt, "X")
          Exit Function
       End If   
	End If 

	If frm1.txt4thFrYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt4thFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt4thFrYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
           Call DisplayMsgBox("140127","X", frm1.txt4thFrYR.Alt, "X")
           Exit Function
        End If   
	End If 

	If frm1.txt4thToYR.Text <> "" Then
	    If Not(TempFiscStart <= UniConvDateAToB(frm1.txt4thToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) And UniConvDateAToB(frm1.txt4thToYR,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat) <= TempFiscEnd) Then
           Call DisplayMsgBox("140127","X", frm1.txt4thToYR.Alt, "X")
           Exit Function
        End If   
	End If 
	

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then				                                                '☜: Save db data 
       Exit Function
    End If
    
    FncSave = True                                                          '⊙: Processing is OK
    Set gActiveElement = document.activeElement
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
    Call parent.FncExport(Parent.C_SINGLE)												'☜: 화면 유형 
    Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
    Set gActiveElement = document.activeElement
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
    Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
    Err.Clear         
    
                                                             '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCtrlYR=" & strCtrlYear							'☜: 삭제 조건 데이타 
    strVal = strVal & "&cboctrlunit=" & Trim(frm1.cboctrlunit.value)		'☜: 삭제 조건 데이타 
    
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	Call FncNew()
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim strVal
    
    
    Err.Clear																	'☜: Protect system from crashing
    
    DbQuery = False																'⊙: Processing is NG
    
    Call LayerShowHide(1)														'☜: Protect system from crashing
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCtrlYR=" & strCtrlYear									'☆: 조회 조건 데이타 
    strVal = strVal & "&cboctrlunit=" & Trim(frm1.cboctrlunit.value)			'☆: 조회 조건 데이타 
    
        
    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True																'⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 	
    '-----------------------
    'Reset variables area
    '-----------------------
    
    Call InitVariables															'⊙: Initializes local global variables
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	Call SetToolbar("1111100000011111")
	
	frm1.txt1stFrYR.focus

	Set gActiveElement = document.activeElement 

	
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 

    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG

    Dim strVal
    Call LayerShowHide(1)                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
	lgBlnFlgChgValue = False
	Call MainQuery
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'######################################################################################## 
    '----------  Coding part  -------------------------------------------------------------

Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1110100000001111")
	Case "QUERY"
		Call SetToolbar("1111100000011111")
		
	End Select
End Function

Function cboCtrlUnit_Change()
	select case frm1.cboCtrlUnit.value
		Case "B"
			lblTitle1.innerHTML = "기간1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = "기간2"
			lblHyphen2.innerHTML = "~"
			lblTitle3.innerHTML = "기간3"
			lblHyphen3.innerHTML = "~"
			lblTitle4.innerHTML = "기간4"
			lblHyphen4.innerHTML = "~"
			Call ElementVisible(frm1.fp1stFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp1stToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp3rdFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp3rdToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp4thFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp4thToYR, 1)	'InVisible
		Case "H"
			lblTitle1.innerHTML = "기간1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = "기간2"
			lblHyphen2.innerHTML = "~"
			lblTitle3.innerHTML = ""
			lblHyphen3.innerHTML = ""
			lblTitle4.innerHTML = ""
			lblHyphen4.innerHTML = ""
			Call ElementVisible(frm1.fp1stFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp1stToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp3rdFrYR, 0)	'InVisible
			Call ElementVisible(frm1.fp3rdToYR, 0)	'InVisible
			Call ElementVisible(frm1.fp4thFrYR, 0)	'InVisible
			Call ElementVisible(frm1.fp4thToYR, 0)	'InVisible
		Case "Y"
			lblTitle1.innerHTML = "기간1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = ""
			lblHyphen2.innerHTML = ""
			lblTitle3.innerHTML = ""
			lblHyphen3.innerHTML = ""
			lblTitle4.innerHTML = ""
			lblHyphen4.innerHTML = ""
			Call ElementVisible(frm1.fp1stFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp1stToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndFrYR, 0)	'InVisible
			Call ElementVisible(frm1.fp2ndToYR, 0)	'InVisible
			Call ElementVisible(frm1.fp3rdFrYR, 0)	'InVisible
			Call ElementVisible(frm1.fp3rdToYR, 0)	'InVisible
			Call ElementVisible(frm1.fp4thFrYR, 0)	'InVisible
			Call ElementVisible(frm1.fp4thToYR, 0)	'InVisible
		Case Else
			lblTitle1.innerHTML = "기간1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = "기간2"
			lblHyphen2.innerHTML = "~"
			lblTitle3.innerHTML = "기간3"
			lblHyphen3.innerHTML = "~"
			lblTitle4.innerHTML = "기간4"
			lblHyphen4.innerHTML = "~"
			Call ElementVisible(frm1.fp1stFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp1stToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp2ndToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp3rdFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp3rdToYR, 1)	'InVisible
			Call ElementVisible(frm1.fp4thFrYR, 1)	'InVisible
			Call ElementVisible(frm1.fp4thToYR, 1)	'InVisible
	End Select
	
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>예산통제기간등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>년도</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtCtrlYR" CLASS=FPDTYYYY tag="12" Title="FPDATETIME" ALT="예산통제년도" id=fpCtrlYR></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>예산통제단위</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboCtrlUnit" ALT="예산통제단위" STYLE="WIDTH: 100px" tag="12" ONCHANGE="vbscript:Call cboCtrlUnit_Change()"><!--<OPTION VALUE=""></OPTION></SELECT>--></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP><!-- 첫번째 탭 내용  -->
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 ID="lblTitle1"NOWRAP>기간1</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt1stFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간1 시작년월" id=fp1stFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen1">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt1stToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간1 종료년월" id=fp1stToYR></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 ID="lblTitle2" NOWRAP>기간2</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt2ndFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간2 시작년월" id=fp2ndFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen2">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt2ndToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간2 종료년월" id=fp2ndToYR></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 ID="lblTitle3" NOWRAP>기간3</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt3rdFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간3 시작년월" id=fp3rdFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen3">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt3rdToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간3 종료년월" id=fp3rdToYR></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 ID="lblTitle4" NOWRAP>기간4</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt4thFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간4 시작년월" id=fp4thFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen4">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt4thToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="기간4 종료년월" id=fp4thToYR></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<% Call SubFillRemBodyTD5656(22) %>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<!-- Batch Button  -->
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO TABINDEX="-1" oresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24">
<INPUT TYPE=hidden NAME="txthCtrlYr" tag="24"CLASS=FPDTYYYY Title="FPDATETIME">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

