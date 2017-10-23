<%@ LANGUAGE="VBSCRIPT" %>
<!--===================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4231ma1
'*  4. Program Name         : ������������ 
'*  5. Program Desc         : Register of Loan Change
'*  6. Comproxy List        : FL0091, FL0098
'*  7. Modified date(First) : 2002-04-02
'*  8. Modified date(Last)  : 2003-05-19
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--########################################################################################################
'												1. �� �� �� 
'###########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "f4231mb1_ko441.asp"			'��: �����Ͻ� ���� ASP�� 
'��: Jump Program ID ASP�� 
Const JUMP_PGM_ID_LOAN_ENTRY = "f4201ma1"	 '���Աݵ�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns

Dim C_SEQ		
Dim C_CHG_DT	
Dim C_INT_RATE	
Dim C_DESC		
Dim C_COL_END	 

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
'Dim lgLngCurRows

 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey
Dim SvrDate
SvrDate = <%=GetSvrDate%>

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

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    'lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False	
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
	lgPageNo  = ""
    lgSortKey = 1
    
End Sub

 '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	C_SEQ		= 1
	C_CHG_DT	= 2
	C_INT_RATE	= 3
	C_DESC		= 4
	C_COL_END	= 5 
	
End Sub



'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	' ���� Page�� Form Element���� Clear�Ѵ�. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear

End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "COOKIE", "MA") %>

End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	
	Call initSpreadPosVariables
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021206",,parent.gAllowDragDropSpread
	
	With frm1.vspdData

		.MaxCols = C_COL_END
		
		.Col = .MaxCols				'��: ������Ʈ�� ��� Hidden Column
		.ColHidden = True
		.MaxRows = 0

		.ReDraw = False

		'���н� �ӹ������ ��û���� ����...kbs..20090831
		Call AppendNumberPlace("6","4","6")

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit  C_SEQ     , "����"      , 08, 2, , 3
		ggoSpread.SSSetDate  C_CHG_DT  , "��������"  , 20, 2, parent.gDateFormat		

		'���н� �ӹ������ ��û���� ����...kbs..20090831
	       'ggoSpread.SSSetFloat C_INT_RATE, "������"    , 20, parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	       'ggoSpread.SSSetFloat C_INT_RATE, "����������", 20, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_INT_RATE, "������"    , 20, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_INT_RATE, "����������", 20, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"



		ggoSpread.SSSetEdit  C_DESC    , "���泻��"  , 47,  , , 128		
		
		.ReDraw = True
		Call ggoSpread.SSSetColHidden(C_SEQ ,C_SEQ	,True)
		Call ggoSpread.SSSetColHidden(C_COL_END ,C_COL_END	,True)
		
		Call SetSpreadLock                                              
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()

    With frm1.vspdData
		.ReDraw = False

		ggoSpread.SpreadLock C_SEQ   , -1, C_SEQ
		ggoSpread.SpreadLock C_CHG_DT, -1, C_CHG_DT	
		ggoSpread.SSSetRequired C_INT_RATE, -1
'		ggoSpread.SpreadDefault C_DESC  , -1, C_DESC
		
		.ReDraw = True

    End With

End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow)

    With frm1

		.vspdData.ReDraw = False

		' �ʼ� �Է� �׸����� ���� 
		' SSSetRequired(ByVal Col, ByVal Row, Optional Row2)
'		ggoSpread.SSSetRequired C_SEQ, lRow, lRow			    ' ���� 
		ggoSpread.SSSetRequired C_CHG_DT, lRow, lRow			' ������ 
		ggoSpread.SSSetRequired C_INT_RATE, lRow, lRow
'		ggoSpread.SSSetDefault C_DESC, lRow, lRow
		
		.vspdData.ReDraw = True
    
    End With

End Sub

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
            
			C_SEQ		= iCurColumnPos(1)
			C_CHG_DT	= iCurColumnPos(2)
			C_INT_RATE	= iCurColumnPos(3)
			C_DESC		= iCurColumnPos(4)
			C_COL_END	= iCurColumnPos(5) 
    End Select    
    
End Sub

'==============================================================
'���Աݹ�ȣ �˾� 
'==============================================================
Function OpenPopupLoan()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("f4232ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4232ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		frm1.txtLoanNo.focus
		Exit Function
	Else		
		frm1.txtLoanNo.value = arrRet(0)
		frm1.txtLoanNm.value = arrRet(1)
	End If
	
	frm1.txtLoanNo.focus
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCalendar(Byval RetCal, Byval iWhere)

	With frm1
		Select Case iWhere
		
			Case 1		
				.vspdData.Col = C_CHG_DT
				.vspdData.Text = RetCal
			Case 2		
								
		End Select
		
		Call vspdData_Change(.vspdData.Col,.vspdData.Row )	

		lgBlnFlgChgValue = True

	End With

End Function



'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)

'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp

	Select Case Kubun		
		Case "FORM_LOAD"

			strTemp = ReadCookie("LOAN_NO")
			Call WriteCookie("LOAN_NO", "")

			If strTemp = "" then Exit Function
						
			frm1.txtLoanNo.value = strTemp
			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("LOAN_NO", "")
				Exit Function 
			End If
					
			Call MainQuery()
		Case JUMP_PGM_ID_LOAN_ENTRY
			Call WriteCookie("LOAN_NO", frm1.txtLoanNo.value)
	
		Case Else
			Exit Function
	End Select
End Function	

'========================================================================================================
'	Desc : ȭ���̵� 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD

	'-----------------------
	'Check previous data area
	'------------------------ 
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		if IntRetCD = parent.vbNo Then
			Exit Function
		End If
    End If

    Call CookiePage(strPgmId)
    Call PgmJump(strPgmId)
End Function

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

    Call LoadInfTB19029                            '��: Load table , B_numeric_format
	' ���� Page�� From Element���� ����ڰ� �Է��� ���� ���ϰ� �ϰų� �ʼ��Է»����� ǥ���Ѵ�.
	' LockField(pDoc, pACode)

	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	'���н� �������� ��û���� ����...20090831...kbs
	call ggoOper.FormatNumber(frm1.txtIntRate, "99.999999", "0", False, 6)			'������
	
	Call InitSpreadSheet                          '��: Setup the Spread Sheet
	Call InitVariables                            '��: Initializes local global Variables
    
    Call CookiePage("FORM_LOAD")
    '----------  Coding part  -------------------------------------------------------------
	Call FncSetToolBar("New")
    Call SetDefaultVal
	Call FncNew()

    frm1.txtLoanNo.focus 
    Set gActiveElement = document.activeElement
    
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
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
    
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
    
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If


    End With

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_SEQ Or NewCol <= C_SEQ Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
    
End Sub


'==========================================================================================
' Event Name : vspdData_ButtonClicked
' Event Desc : ��ư �÷��� Ŭ���� ��� 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	'---------- Coding part -------------------------------------------------------------
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_CHG_DT_PB Then
			.Col = Col
			.Row = Row
			
		Call OpenCalendar(1)			
	    
		End If
		
	End With
	
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
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		if IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call InitVariables							  '��: Initializes local global variables

	frm1.vspdData.MaxRows = 0
    
	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
	If Not chkField(Document, "1") Then	  '��: This function check indispensable field
       Exit Function
    End If
    
    Call FncSetToolBar("New")
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
	Dim IntRetCD 
    FncNew = False                  '��: Processing is NG
    Err.Clear                       '��: Protect system from crashing
    'On Error Resume Next            '��: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ' ����� ������ �ִ��� Ȯ���Ѵ�.
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015",parent.VB_YES_NO,"X","X")
		If IntRetCD = parent.vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
	
    Call ggoOper.ClearField(Document, "1")     '��: Clear Condition Field
	Call ggoOper.LockField(Document, "N")      '��: Lock  Suitable  Field
    Call InitVariables                         '��: Initializes local global variables
    Call SetDefaultVal
	frm1.vspdData.MaxRows = 0

    'SetGridFocus
    
    Call FncSetToolBar("New")
    FncNew = True                              '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	Dim IntRetCD 
    
    FncDelete = False            '��: Processing is NG
    Err.Clear                    '��: Protect system from crashing
    'On Error Resume Next        '��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ' Update ���������� Ȯ���Ѵ�.
    If lgIntFlgMode <> Parent.OPMD_UMODE Then        'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	
	If IntRetCD = parent.vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then											  '��: Delete db data
       Exit Function                        
    End If
    
    '-----------------------
    'Erase condition area
    '-----------------------
	Call ggoOper.ClearField(Document, "1")								  '��: Clear Condition Field
    FncDelete = True													 '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False            '��: Processing is NG
    Err.Clear                  '��: Protect system from crashing
    'On Error Resume Next       '��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '��: Check required field(Multi area)
       Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
	If Not chkField(Document, "1") Then								  '��: Check contents area
       Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '��: Save db data

	 FncSave = True                                                           '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	Dim  IntRetCD
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow

End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        for imRow2 = 1 to imRow 
            ggoSpread.Source = .vspdData
            ggoSpread.InsertRow ,1

			.vspdData.col = C_CHG_DT
			.vspdData.Text= UniConvDateAToB("<%=GetSvrdate%>",Parent.gServerDateFormat,Parent.gDateFormat)	'������ default : today		
			.vspdData.Col = C_INT_RATE
			.vspdData.Text= "0"				

            Call SetSpreadColor(.vspdData.ActiveRow) 

        Next
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	    
    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With

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
	Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
	Call SetSpreadLock

End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
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
	Dim strVal
	
	Call DisableToolBar(Parent.TBC_QUERY)
	Call LayerShowHide(1)
    
    DbQuery = False
    Err.Clear                '��: Protect system from crashing
    
    With frm1
        
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode="	& Parent.UID_M0001
			strVal = strVal & "&txtLoanNo="		& Trim(.htxtLoanNo.value)	'��ȸ ���� ����Ÿ 
		Else
			strVal = BIZ_PGM_ID & "?txtMode="	& Parent.UID_M0001
			strVal = strVal & "&txtLoanNo="		& Trim(.txtLoanNo.value)	'��ȸ ���� ����Ÿ 
		End If
			strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey
			strVal = strVal & "&lgPageNo="		& lgPageNo         
			strVal = strVal & "&txtMaxRows="	& .vspdData.MaxRows
			
		Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 
    End With
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	Call SetSpreadLock
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE	'��: Indicates that current mode is Update mode

	' ���� Page�� From Element���� ����ڰ� �Է��� ���� ���ϰ� �ϰų� �ʼ��Է»����� ǥ���Ѵ�.
	' LockField(pDoc, pACode)
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	Call FncSetToolBar("Query")
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 	
    Call ggoOper.LockField(Document, "Q")	'��: This function lock the suitable field
	
	'SetGridFocus
	If frm1.vspdData.MaxRows > 0 Then
		Frm1.vspdData.Focus
	Else
		frm1.txtLoanNo.focus
	End If
	
	
	Set gActiveElement = document.activeElement 
	
End Function


'========================================================================================
' Function Name : DbSave()
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
	Dim lRow        
	Dim lGrpCnt     
	Dim retVal      
	Dim boolCheck   
	Dim lStartRow   
	Dim lEndRow     
	Dim lRestGrpCnt 
	Dim strVal,strDel, iColSep, iRowSep

	Call LayerShowHide(1)
	
    DbSave = False				'��: Processing is NG
    'On Error Resume Next		'��: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""
		iColSep = Parent.gColSep
		iRowSep = Parent.gRowSep
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
			
		    Select Case .vspdData.Text
		    
		        Case ggoSpread.InsertFlag											'��: �ű� 
					strVal = strVal & "C" & iColSep & lRow & iColSep				'��: C=Create
		            .vspdData.Col = C_SEQ
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_CHG_DT
		            strVal = strVal & UniConvDate(Trim(.vspdData.Text)) & iColSep		            
		            .vspdData.Col = C_INT_RATE
		            strVal = strVal & UNICdbl(Trim(.vspdData.Text)) & iColSep
		            .vspdData.Col = C_DESC
		            strVal = strVal & Trim(.vspdData.Text) & iRowSep
		             					
		            lGrpCnt = lGrpCnt + 1

				Case ggoSpread.UpdateFlag												'��: ���� 

					strVal = strVal & "U" & iColSep & lRow & iColSep					'��: U=Update
				    .vspdData.Col = C_SEQ
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_CHG_DT
		            strVal = strVal & UniConvDate(Trim(.vspdData.Text)) & iColSep		            
		            .vspdData.Col = C_INT_RATE
		            strVal = strVal & UNICdbl(Trim(.vspdData.Text)) & iColSep
		            .vspdData.Col = C_DESC
		            strVal = strVal & Trim(.vspdData.Text) & iRowSep		            
		            
		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag												'��: ���� 

					strDel = strDel & "D" & iColSep & lRow & iColSep					'��: U=Delete
		            .vspdData.Col = C_SEQ
		            strDel = strDel & Trim(.vspdData.Text) & iRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		
		 Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'��: �����Ͻ� ASP �� ���� 
	
	End With

    DbSave = True                           '��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
    
    ggoSpread.SSDeleteFlag 1 
	
	Call InitVariables
	Call MainQuery
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

'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100111100111111")
	End Select
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!--########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>���Աݹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP Colspan=3><INPUT NAME="txtLoanNo" MAXLENGTH="18" SIZE=15  ALT ="���Աݹ�ȣ" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupLoan()">
														   <INPUT NAME="txtLoanNm" MAXLENGTH="20" SIZE=40   ALT  ="���Աݳ���" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanDt" ALT="������" SIZE = "10" MAXLENGTH="10" STYLE="TEXT-ALIGN: Center" tag="24X1"></TD>
								<TD CLASS="TD5" NOWRAP>��ȯ������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDueDt" ALT="��ȯ������" SIZE = "10" MAXLENGTH="10" STYLE="TEXT-ALIGN: Center" tag="24X1"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���Աݾ�</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpLoanAmt name=txtLoanAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ܾ�" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								<!--					<INPUT NAME="txtDocCur" ALT="��ȭ" SIZE = "10" MAXLENGTH="3" STYLE="TEXT-ALIGN: Left" tag="24X"> -->

								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntRate name=txtIntRate CLASS=FPDS115 title=FPDOUBLESINGLE ALT="������" tag="24X5Z" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp; %</TD>
							    <!-- ���佺 �ӹ������ ��û���� ����...20090831...kbs
								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=OBJECT5 name=txtIntRate CLASS=FPDS90 title=FPDOUBLESINGLE ALT="������" tag="24X5Z" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp; %</TD>
							     -->
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_LOAN_ENTRY)">���Աݵ��</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="htxtLoanNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>