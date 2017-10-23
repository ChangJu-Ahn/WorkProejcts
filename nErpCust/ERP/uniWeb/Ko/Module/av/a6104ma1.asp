
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!--*******************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : �ΰ������� 
'*  3. Program ID        : a6104ma1
'*  4. Program �̸�      : �ΰ���������ȸ 
'*  5. Program ����      : �ΰ����� �Ǻ��� ��ȸ�Ѵ�.
'*  6. Comproxy ����Ʈ   : a6104ma1
'*  7. ���� �ۼ������   : 2000/04/22
'*  8. ���� ���������   : 2001/01/17
'*  9. ���� �ۼ���       : ����ȯ 
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'*                         -2000/04/22 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncServer.asp"  -->
<!--========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/AdoQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "a6104mb1.asp"			'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns
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

Const C_SHEETMAXROWS = 50		' : �� ȭ�鿡 �������� �ִ밹��*1.5

<%
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	EndDate = GetSvrDate
	Call ExtractDateFrom(EndDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)

	StartDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, "01")
	EndDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)
%>

 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
Dim lgStrPrevKeyISSUEDT
Dim lgStrPrevKeyGLNO

'Dim lgLngCurRows

Dim lgBlnStartFlag				' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag

 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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

    lgIntFlgMode = OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = 0                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
	
	lgSortKey = 1
	
End Sub

 '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
		.Col = .MaxCols				'��: ������Ʈ�� ��� Hidden Column
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

		ggoSpread.SSSetDate  C_IssuedDT, "��꼭��", 11, 2, gDateFormat
		ggoSpread.SSSetCombo C_IOFGCD,   "", 10
		ggoSpread.SSSetCombo C_IOFGNM,   "����", 10, 2
<%
		Call InitComboBoxDtl("2", "A1003")		' ���ⱸ�� 
%>
		ggoSpread.SSSetEdit  C_BPCD,      "",   20, , , 40
		ggoSpread.SSSetEdit  C_BPNM,      "�ŷ�ó��", 20, , , 40
	    ggoSpread.SSSetEdit  C_OwnRGSTNo, "����ڵ�Ϲ�ȣ", 20, , , 20
	    ggoSpread.SSSetFloat C_NetAmt,    "���ް�", 19, ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
	    ggoSpread.SSSetFloat C_VatAmt,    "�ΰ���", 19, ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetCombo C_VatTypeCD, "��꼭����", 18
		ggoSpread.SSSetCombo C_VatTypeNM, "��꼭����", 18
<%
		Call InitComboBoxDtl("3", "B9001")		' �ΰ������� 
%>
		.ReDraw = True

		Call SetSpreadLock                                              '�ٲ�κ� 
    
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
		Call InitComboBoxDtl("0", "A1003")		' ���ⱸ�� 
		Call InitComboBoxDtl("1", "B9001")		' �ΰ������� 
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
    
    Err.Clear                                                               '��: Clear error no
	On Error Resume Next

	'Set B1a028 = Server.CreateObject("B1a028.B1a028ListMinorCode")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
'	If Err.Number <> 0 Then
'		Set B1a028 = Nothing												'��: ComProxy Unload
'		Call MessageBox(Err.description, I_INSCRIPT)						'��:
'		Response.End														'��: �����Ͻ� ���� ó���� ������ 
'	End If

 '   B1a028.ImportBMajorMajorCd = Trim(MajorCd)									'��: Major Code
  '  B1a028.ServerLocation = ggServerIP
    
  '  B1a028.ComCfg = gConnectionString
  '  B1a028.Execute															'��:
    
    '-----------------------
    'Com action result check area(DB,internal)
    '-----------------------
 '   If Not (B1a028.OperationStatusMessage = MSG_OK_STR) Then
'		Call MessageBox(B1a028.OperationStatusMessage, I_INSCRIPT)         '��: you must release this line if you change msg into code
'		Set B1a028 = Nothing												'��: ComProxy Unload
'		Response.End														'��: �����Ͻ� ���� ó���� ������ 
 '   End If

'	intMaxRow = B1a028.ExportGroupCount
	strListCd = ""
	strListNm = ""
	
	Select Case Index
		Case "0"	' ���ⱸ�� 
			
%>
				Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				Call SetCombo2(frm1.cboIOFlag ,lgF0  ,lgF1  ,Chr(11))
<%
			'Next
		Case "1"	' �ΰ������� 
			'For intLoopCnt = 1 To intMaxRow
%>
	    		 Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("B9001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
				Call SetCombo2(frm1.cboVatType ,lgF0  ,lgF1  ,Chr(11))
<%
			'Next
		Case "2"	' ���ⱸ�� 
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
			ggoSpread.SetCombo lgF0, C_IOFGCD		' ���� 
			ggoSpread.SetCombo lgF1, C_IOFGNM		' ���� 
<%
		Case "3"	' �ΰ������� 
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
			ggoSpread.SetCombo "<%=strListCd%>", C_VatTypeCD		' ���� 
			ggoSpread.SetCombo "<%=strListNm%>", C_VatTypeNM		' ���� 
<%
	End Select

	Set B1a028 = Nothing                        

End Function
%>

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
			arrParam(0) = "����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA"	 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "������ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"					' Field��(0)
			arrField(1) = "BIZ_AREA_NM"					' Field��(0)
    
			arrHeader(0) = "������ڵ�"				' Header��(0)
			arrHeader(1) = "������"				' Header��(0)
		Case 1
			arrParam(0) = "�ŷ�ó �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�ŷ�ó"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BP_CD"						' Field��(0)
			arrField(1) = "BP_NM"						' Field��(1)
    
			arrHeader(0) = "�ŷ�ó�ڵ�"				' Header��(0)
			arrHeader(1) = "�ŷ�ó��"				' Header��(1)
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
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ������ 
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
				
				.txtBizAreaCD.focus
			Case 1		' �ŷ�ó 
				.txtBPCd.value = UCase(Trim(arrRet(0)))
				.txtBPNM.value = arrRet(1)
				
				.txtBPCd.focus
		End Select
	End With
End Function


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
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                           '��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)

	' ���� Page�� Form Element���� Clear�Ѵ�. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	' ���� Page�� From Element���� ����ڰ� �Է��� ���� ���ϰ� �ϰų� �ʼ��Է»����� ǥ���Ѵ�.
	' LockField(pDoc, pACode)
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
    
    Call InitSpreadSheet                          '��: Setup the Spread Sheet
    Call InitVariables                            '��: Initializes local global Variables
    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox

	Call SetDefaultVal

	' [Main Menu ToolBar]�� �� ��ư�� [Enable/Disable] ó���ϴ� �κ� 
    Call SetToolbar("1100000000001111")										'��: ��ư ���� ���� 

    frm1.txtIssueDT1.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
Sub txtIssueDt1_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub txtIssueDt2_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================

Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			
			' ���� 
			.Col = C_IOFGCD
			intIndex = .value
			.col = C_IOFGNM
			.value = intindex
			' �ΰ������� 
			.Col = C_VatTypeCD
			intIndex = .value
			.col = C_VatTypeNM
			.value = intindex
					
		Next	
	End With
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_Change()
    'lgBlnFlgChgValue = True
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)
    
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
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
'   Event Desc : Spread Split �����ڵ� 
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
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ 
		If lgStrPrevKey <> 0 Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
    End if

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


'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
	Dim IntRetCD 
    
    FncQuery = False          '��: Processing is NG
    Err.Clear                 '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
	' ���� Page�� Form Element���� Clear�Ѵ�. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "2")      '��: Condition field clear
    Call InitSpreadSheet                          '��: Setup the Spread Sheet
    Call InitVariables							'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
	' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then	'��: This function check indispensable field
       Exit Function
    End If
 
    If UniCDate(frm1.txtIssueDt1.text) > UniCDate(frm1.txtIssueDt2.text) Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'��: "Will you destory previous data"
		Exit Function
    End If
 
	If frm1.txtBPCd.value = "" Then
		frm1.txtBPNm.value = ""
	End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
    Call Parent.FncExport(C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
		' ����� ������ �ִ��� Ȯ���Ѵ�.
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"X","X")			'��: "Will you destory previous data"
	
			If IntRetCD = vbNo Then
				Exit Function
			End If
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
Dim strVal
Dim RetFlag

    DbQuery = False
    Err.Clear                '��: Protect system from crashing
    
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
		Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 
		    
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = OPMD_UMODE	'��: Indicates that current mode is Update mode
    
	lgBlnFlgChgValue = False
	
	lgBlnStartFlag = True		' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag
	
	' ���� Page�� From Element���� ����ڰ� �Է��� ���� ���ϰ� �ϰų� �ʼ��Է»����� ǥ���Ѵ�.
	' LockField(pDoc, pACode)
    Call ggoOper.LockField(Document, "Q")	'��: This function lock the suitable field

    Call InitData	' Combo�� Name�� Code�� �������� ���� 

    Call SetToolbar("1100000000011111")										'��: ��ư ���� ���� 

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
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


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag�� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ΰ���������ȸ</font></td>
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
									<TD CLASS="TD5">������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6104ma1_fpDateTime2_txtIssueDt1.js'></script>&nbsp;~&nbsp;
													<script language =javascript src='./js/a6104ma1_fpDateTime2_txtIssueDt2.js'></script></TD>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű�����</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="�Ű�����" tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
													<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="�Ű�����" tag="14X" ></TD>
									<TD CLASS="TD5">�ŷ�ó</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBPCd" NAME="txtBPCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBPCd.Value, 1)">&nbsp;
													<INPUT TYPE=TEXT ID="txtBPNm" NAME="txtBPNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">���ⱸ��</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag" NAME="cboIOFlag" ALT="���ⱸ��" STYLE="WIDTH: 98px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5">�ΰ�������</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT ID="cboVatType" NAME="cboVatType" ALT="�ΰ�������" STYLE="WIDTH: 130px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
								<TD CLASS="TD18"><FONT COLOR=Blue>����ó  :</TD>
								<TD CLASS="TD18">�ż��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingle1_txtCntSumI.js'></script></TD>
								<TD CLASS="TD18">���ް��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingle1_txtAmtSumI.js'></script></TD>
								<TD CLASS="TD18">�ΰ����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingle1_txtVatSumI.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>����ó  :</TD>
								<TD CLASS="TD18">�ż��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingleO_txtCntSumO.js'></script></TD>
								<TD CLASS="TD18">���ް��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6104ma1_fpDoubleSingleO_txtAmtSumO.js'></script></TD>
								<TD CLASS="TD18">�ΰ����հ�</TD>
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
