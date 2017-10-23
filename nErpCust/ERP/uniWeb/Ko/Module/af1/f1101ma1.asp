<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f1101ma1
'*  4. Program Name         : ���������Ⱓ��� 
'*  5. Program Desc         : Register of Control Period
'*  6. Comproxy List        : FB0011, FB0019
'*  7. Modified date(First) : 2000.09.18
'*  8. Modified date(Last)  : 2003.06.13
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Oh, Soo Min
'* 11. Comment              :
'* - 2003/06/13 Oh, Soo Min ClearField�Ӽ� ���� 
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
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


Option Explicit                                   '��: indicates that All variables must be declared in advance %>


'********************************************  1.2 Global ����/��� ����  *********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'============================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================
Const BIZ_PGM_ID = "f1101mb1.asp"											 '��: �����Ͻ� ���� ASP�� 

'============================================  1.2.2 Global ���� ����  ===================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2. Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 

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
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
		
End Sub


'============================================= 2.1.2 LoadInfTB19029() ====================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================================= 
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I", "*","NOCOOKIE" ,"MA") %>
	
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 



'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'********************************************************************************************************* 

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
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


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 


'###########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################


'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************
'=======================================================================================================
'   Event Name : txt1StPrRdpDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Name : KeyDown �̺�Ʈ 
'   Event Desc : ����Ű�� ġ�� FncQuery ȣ�� 
'=======================================================================================================
Sub txtCtrlYR_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

'=======================================================================================================
'   Event Name : Change �̺�Ʈ 
'   Event Desc : ������ �����÷��� ���� 
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
    'Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    
    frm1.txt1stFrYR.value =""
    frm1.txt1stToYR.value =""
    frm1.txt2ndFrYR.value =""
    frm1.txt2ndToYR.value =""
    frm1.txt3rdFrYR.value =""
    frm1.txt3rdToYR.value =""
    frm1.txt4thFrYR.value =""
    frm1.txt4thToYR.value =""
    
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    lgBlnFlgChgValue = False
    lgIntFlgMode = Parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    
End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    
    Call LoadInfTB19029																'��: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	      
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    Call InitComboBox
    Call SetDefaultVal
    '----------  Coding part  -------------------------------------------------------------
    Call FncSetToolBar("New")
    Call InitVariables																'��: Initializes local global variables
    
    frm1.txtCtrlYR.focus
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ExtractDateFrom(frm1.txtCtrlYR.Text,frm1.txtCtrlYR.UserDefinedFormat,Parent.gComDateType,strCtrlYear,strCtrlMonth,strCtrlDay)
            
    '-----------------------
    'Erase contents area
    '----------------------- 
    'Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not ChkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    Call InitVariables															'��: Initializes local global variables
 '   Call SetDefaultVal  
    Call FncSetToolBar("New")
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    Set gActiveElement = document.activeElement
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                     '��: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	    
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    
    Call ggoOper.ClearField(Document, "A")                                      '��: Clear Contents  Field        
  
    Call SetDefaultVal 
    Call InitVariables															'��: Initializes local global variables
    Call cboCtrlUnit_Change()
    
    Call FncSetToolBar("New")
  
    frm1.txtCtrlYR.focus
  
    FncNew = True																'��: Processing is OK
    Set gActiveElement = document.activeElement  
    
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD
    
    FncDelete = False														'��: Processing is NG
    
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
    
    
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
' Modify Date : 2001-12-04
' Modify Contents : date format ǥ�� �ݿ� 
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strStartYear,strStartMonth,strStartDay
    Dim TempFiscStart,TempFiscEnd
    
    On Error Resume Next

    TempFiscStart = UniConvDateToYYYYMMDD(Parent.gFiscStart,Parent.gAPDateFormat,Parent.gServerDateType)
    TempFiscEnd   = UniConvDateToYYYYMMDD(Parent.gFiscEnd  ,Parent.gAPDateFormat,Parent.gServerDateType)

    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '��: No data changed!!
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
    If DbSave = False Then				                                                '��: Save db data 
       Exit Function
    End If
    
    FncSave = True                                                          '��: Processing is OK
    Set gActiveElement = document.activeElement
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
    Call parent.FncExport(Parent.C_SINGLE)												'��: ȭ�� ���� 
    Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		
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
    
                                                             '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCtrlYR=" & strCtrlYear							'��: ���� ���� ����Ÿ 
    strVal = strVal & "&cboctrlunit=" & Trim(frm1.cboctrlunit.value)		'��: ���� ���� ����Ÿ 
    
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncNew()
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim strVal
    
    
    Err.Clear																	'��: Protect system from crashing
    
    DbQuery = False																'��: Processing is NG
    
    Call LayerShowHide(1)														'��: Protect system from crashing
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCtrlYR=" & strCtrlYear									'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&cboctrlunit=" & Trim(frm1.cboctrlunit.value)			'��: ��ȸ ���� ����Ÿ 
    
        
    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True																'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 	
    '-----------------------
    'Reset variables area
    '-----------------------
    
    Call InitVariables															'��: Initializes local global variables
    lgIntFlgMode = Parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	Call SetToolbar("1111100000011111")
	
	frm1.txt1stFrYR.focus

	Set gActiveElement = document.activeElement 

	
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 

    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG

    Dim strVal
    Call LayerShowHide(1)                                                   '��: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()															'��: ���� ������ ���� ���� 
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
			lblTitle1.innerHTML = "�Ⱓ1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = "�Ⱓ2"
			lblHyphen2.innerHTML = "~"
			lblTitle3.innerHTML = "�Ⱓ3"
			lblHyphen3.innerHTML = "~"
			lblTitle4.innerHTML = "�Ⱓ4"
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
			lblTitle1.innerHTML = "�Ⱓ1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = "�Ⱓ2"
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
			lblTitle1.innerHTML = "�Ⱓ1"
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
			lblTitle1.innerHTML = "�Ⱓ1"
			lblHyphen1.innerHTML = "~"
			lblTitle2.innerHTML = "�Ⱓ2"
			lblHyphen2.innerHTML = "~"
			lblTitle3.innerHTML = "�Ⱓ3"
			lblHyphen3.innerHTML = "~"
			lblTitle4.innerHTML = "�Ⱓ4"
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���������Ⱓ���</font></td>
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
									<TD CLASS="TD5" NOWRAP>�⵵</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtCtrlYR" CLASS=FPDTYYYY tag="12" Title="FPDATETIME" ALT="���������⵵" id=fpCtrlYR></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>������������</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboCtrlUnit" ALT="������������" STYLE="WIDTH: 100px" tag="12" ONCHANGE="vbscript:Call cboCtrlUnit_Change()"><!--<OPTION VALUE=""></OPTION></SELECT>--></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP><!-- ù��° �� ����  -->
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 ID="lblTitle1"NOWRAP>�Ⱓ1</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt1stFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ1 ���۳��" id=fp1stFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen1">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt1stToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ1 ������" id=fp1stToYR></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 ID="lblTitle2" NOWRAP>�Ⱓ2</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt2ndFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ2 ���۳��" id=fp2ndFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen2">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt2ndToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ2 ������" id=fp2ndToYR></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 ID="lblTitle3" NOWRAP>�Ⱓ3</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt3rdFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ3 ���۳��" id=fp3rdFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen3">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt3rdToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ3 ������" id=fp3rdToYR></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 ID="lblTitle4" NOWRAP>�Ⱓ4</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt4thFrYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ4 ���۳��" id=fp4thFrYR></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen4">~</SPAN>&nbsp;
													 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txt4thToYR" CLASS=FPDTYYYYMM tag="21X1" Title="FPDATETIME" ALT="�Ⱓ4 ������" id=fp4thToYR></OBJECT>');</SCRIPT>
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

