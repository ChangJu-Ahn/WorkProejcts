
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
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
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

<!--==========================================  1.1.2 ���� Include   ======================================
==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit            '��: indicates that All variables must be declared in advance

Dim BaseDate
Dim StartDate
Dim strYear
Dim strMonth
DIm strDay


<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, Parent.gServerDateFormat, Parent.gDateFormat)
Call ExtractDateFrom(BaseDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

Const BIZ_PGM_QRY_ID = "p1102mb1.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "p1102mb3.asp"											'��: �����Ͻ� ���� ASP�� 

Const CChnageColor = "#f0fff0"
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim lgLastDay
Dim lgStartIndex
Dim lgArrDate(31, 3)
Dim IsOpenPop
Dim lgChgCboYear
Dim lgChgCboMonth          

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
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

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : ComboBox �ʱ�ȭ 
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
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
End Sub

'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'*********************************************************************************************************


'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
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

	arrParam(0) = "Į���� Ÿ�� �˾�"			<%' �˾� ��Ī %>
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtClnrType.Value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "Į���� Ÿ��"					<%' TextBox ��Ī %>
	
    arrField(0) = "CAL_TYPE"						<%' Field��(0)%>
    arrField(1) = "CAL_TYPE_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "Į���� Ÿ��"				<%' Header��(0)%>
    arrHeader(1) = "Į���� Ÿ�Ը�"				<%' Header��(1)%>

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
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029																'��: Load table , B_numeric_format
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11000000000001")
    Call SetDefaultVal
    Call InitComboBox															'��: Initialize combobox at load time
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
'   Event Desc : ȭ��ǥ Ŭ�� 
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
	
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'��: 
    strVal = strVal & "&txtClnrType=" & Trim(frm1.txtClnrType.value)
    strVal = strVal & "&txtYear=" & StrYear1					'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtMonth=" & StrMonth1		'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
	LayerShowHide(1)
											'��: �۾������� ǥ��	
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"		
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
		
    Call InitVariables															'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"		
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
		
    Call InitVariables															'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
End Function

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
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                        '��: Processing is NG
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"		
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
		
    Call InitVariables															'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     											'��: Query db data
       
    FncQuery = True																'��: Processing is OK
        
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                            '��: No data changed!!        
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
    End If     								                                                  '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     On Error Resume Next                                                   '��: Protect system from crashing
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
    Call parent.FncExport(Parent.C_SINGLE)								  '��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                          '��:ȭ�� ����, Tab ���� 
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    DbQuery = False                                                         '��: Processing is NG
    
    LayerShowHide(1)
		
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001							'��: 
    strVal = strVal & "&txtClnrType=" & Trim(frm1.txtClnrType.value)	'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtYear=" & Trim(frm1.cboYear.value)	'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtMonth=" & Trim(frm1.cboMonth.Value)	'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()													'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False  
    
    Call SetToolbar("11001000000101")
    
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
	
	Dim ColCol
    
    Err.Clear																	'��: Protect system from crashing

	DbSave = False																'��: Processing is NG
	
	LayerShowHide(1)
											'��: �۾������� ǥ�� 
	'-----------------------
    'Check content area
    '-----------------------

	'-------------------------------------------------------------
	' ������ ������ disable�Ǿ� �־� biz asp�� �Ѿ�� �ʴ´�.
	' ���� �ӽ÷� disable�� enable�� �����Ų��.
	'-------------------------------------------------------------
	For ColCol = 0 To 41
		If frm1.txtDate(ColCol).className <> "DummyDay" Then
			frm1.txtDesc(ColCol).disabled = False
		End If
	Next	

	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	frm1.txtUpdtUserId.value = Parent.gUsrID
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)									'��: �����Ͻ� ASP �� ���� 
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����Į���ټ���</font></td>
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
								<TD CLASS=TD5 NOWRAP>Į���� Ÿ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="12XXXU" ALT="Į���� Ÿ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=40 tag="14"></TD>
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
									<TD>�Ͽ���</TD>
									<TD>������</TD>
									<TD>ȭ����</TD>
									<TD>������</TD>
									<TD>�����</TD>
									<TD>�ݿ���</TD>
									<TD>�����</TD>
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
													<INPUT type="text" name="txtDesc"  MaxLength=20 Style="Width:100%;Border:0;text-align:center" disabled tag=2 onchange="DescChange(<%=k%>)" ALT="���">
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
