<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : C3970MA1.asp
'*  4. Program Name         : MCS
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-02-19
'*  8. Modified date(Last)  : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "c3970mb1.asp"							'��: �����Ͻ� ���� ASP�� 

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

' Grid 1(vspdData1) 
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_MCSItem
Dim C_MCSDTLItem
Dim C_MCSDTLItemNm
Dim C_Amount
Dim C_AcctSeq
Dim C_Seq
Dim C_Type

dim	strFromYYYYMM 
Dim strToYyyyMm 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgIsOpenPop

Dim lgStrPrevKey1
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2
Dim lgRadio

Dim strDate
Dim iDBSYSDate


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
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	Dim strYear,strMonth,strDay
	
	LocSvrDate = "<%=GetSvrDate%>"
	
	lgRadio	= "S"

	Call ggoOper.FormatDate(frm1.txtFromYYYYMM, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToYYYYMM, Parent.gDateFormat, 2)

	frm1.txtFromYYYYMM.text	= UniConvDateAToB(LocSvrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtFromYYYYMM, Parent.gDateFormat, 2)
	frm1.txtToYYYYMM.text	= UniConvDateAToB(LocSvrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtToYYYYMM, Parent.gDateFormat, 2)

	
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet()


	Call InitSpreadPosVariables()


	With frm1.vspdData1 
			
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
		.ReDraw = false
					
		.MaxCols = C_Type + 1    
		.MaxRows = 0    
			

		Call GetSpreadColumnPos()


		ggoSpread.SSSetEdit 	C_ItemAcct,		"ǰ�����"	,10 
		ggoSpread.SSSetEdit 	C_ItemAcctNm,   "ǰ�������",10
		ggoSpread.SSSetEdit 	C_MCSItem,      "�׸�"		,25 
		ggoSpread.SSSetEdit 	C_MCSDTLItem,   "�����׸�"	,10
		ggoSpread.SSSetEdit 	C_MCSDTLItemNm, "�����׸��",25
		ggoSpread.SSSetFloat 	C_Amount,		"�ݾ�"		,30,parent.ggAmtofMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_AcctSeq,		"SEQ"		,5
		ggoSpread.SSSetEdit 	C_Seq,			"SUB_SEQ"	,5
		ggoSpread.SSSetEdit 	C_Type,			"Type"	,5	

	
		Call ggoSpread.MakePairsColumn(C_ItemAcct, C_ItemAcctNm )

		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
		
		Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
		Call ggoSpread.SSSetColHidden( C_MCSDTLItemNM, C_MCSDTLItemNm, True)			
		
		Call ggoSpread.SSSetColHidden( C_AcctSeq, C_AcctSeq, True)			
		Call ggoSpread.SSSetColHidden( C_Seq, C_Seq, True)
		Call ggoSpread.SSSetColHidden( C_Type, C_Type, True)
					
		ggoSpread.SSSetSplit2(4)
			
		Call SetSpreadLock()
			
		.ReDraw = true    
    
	End With
	


    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData1
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub



'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()

		' Grid 1(vspdData1) - Order Header

	C_ItemAcct		= 1
	C_ItemAcctNm	= 2
	C_MCSItem		= 3
	C_MCSDTLItem	= 4
	C_MCSDTLItemNm 	= 5		
	C_Amount		= 6
	C_AcctSeq		= 7
	C_Seq			= 8
	C_Type			= 9

End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos()
	Dim iCurColumnPos
 	ggoSpread.Source = frm1.vspdData1
		
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	
	C_ItemAcct			= iCurColumnPos(1)
	C_ItemAcctNm		= iCurColumnPos(2)
	C_MCSItem			= iCurColumnPos(3)
	C_MCSDTLItem		= iCurColumnPos(4)
	C_MCSDTLItemNm		= iCurColumnPos(5)
	C_Amount			= iCurColumnPos(6)
	C_AcctSeq			= iCurColumnPos(7)
	C_Seq				= iCurColumnPos(8)
	C_Type				= iCurColumnPos(9)
	

End Sub    


'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
'=======================================================================================================
'	Name : OpenWorkStep()
'	Description : Condition Plant PopUp
'=======================================================================================================
Function OpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
		Case 1
			arrParam(0) = "ǰ������˾�"			'�˾� ��Ī 
			arrParam(1) = "(select minor_cd,minor_nm from B_MINOR where MAJOR_CD =" & FilterVar("P1001", "''", "S") & " union all select minor_cd,minor_nm from b_minor where major_cd =" & FilterVar("C2111", "''", "S") & " and minor_cd in (" & FilterVar("MC", "''", "S") & "," & FilterVar("BFO", "''", "S") & "," & FilterVar("EFO", "''", "S") & ")) a "						'TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = ""							'Where Condition
			arrParam(5) = "ǰ�����"				'TextBox ��Ī 
	
			arrField(0) = "a.minor_cd"					'Field��(0)
			arrField(1) = "a.minor_nm"					'Field��(1)
    
			arrHeader(0) = "ǰ�����"				'Header��(0)
			arrHeader(1) = "ǰ�������"				'Header��(1)

		Case 2
			arrParam(0) = "���������˾�"			'�˾� ��Ī 
			arrParam(1) = "B_MINOR"						'TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtMovTypeCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = "major_cd = " & FilterVar("I0001", "''", "S") & " "							'Where Condition
			arrParam(5) = "��������"				'TextBox ��Ī 
	
			arrField(0) = "minor_cd"					'Field��(1)
			arrField(1) = "minor_nm"					'Field��(1)
			
			arrHeader(0) = "��������"				'Header��(1)
			arrHeader(1) = "����������"				'Header��(2)
			

	End Select
    

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(arrRet,iWhere)
	End If
		
End Function


Function SetPopup(byval arrRet,byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtItemAcctCd.Value = arrRet(0)		
				.txtItemAcctNm.Value = arrRet(1)		
			Case 2
				.txtMovTypeCd.Value = arrRet(0)		
				.txtMovTypeNm.Value = arrRet(1)		
		End Select

		lgBlnFlgChgValue = True
	End With
End Function

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
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
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet()                                               '��: Setup the Spread sheet
   
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetToolBar("11000000000111")										'��: ��ư ���� ���� 
 
    frm1.txtFromYyyymm.focus
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

'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromYyyymm_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFromYyyymm.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtFromYyyymm.focus
	End If 
End Sub

Sub txtToYyyymm_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToYyyymm.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtToYyyymm.focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromYyyymm_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

Sub txtToYyyymm_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then							'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    
    lgOldRow = frm1.vspdData1.ActiveRow
			

					
    
End Sub



'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub



'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub



'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos()
End Sub


'==========================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

End Sub


'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData1 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub



Function Radio1_onChange()
	
	IF lgRadio = "S" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, True)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()
	
	lgRadio = "S"
	
	lgBlnFlgChgValue = True
End Function

Function Radio2_onChange()

	IF lgRadio = "D" Then
		Exit Function
	ENd IF

	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, False)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, False)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()


	lgRadio = "D"
	
	lgBlnFlgChgValue = True
End Function

Function Radio3_onChange()
	
	IF lgRadio = "S1" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, True)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()
	
	lgRadio = "S1"
	
	lgBlnFlgChgValue = True
End Function

Function Radio4_onChange()

	IF lgRadio = "D1" Then
		Exit Function
	ENd IF

	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, False)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, False)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()


	lgRadio = "D1"
	
	lgBlnFlgChgValue = True
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

    FncQuery = False														'��: Processing is NG
    Err.Clear																'��: Protect system from crashing

	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    Call InitVariables														'��: Initializes local global variables
	
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'��: This function check indispensable field
       Exit Function
    End If
    
    
    If CompareDateByFormat(frm1.txtFromYYYYMM.Text,frm1.txtToYYYYMM.Text,frm1.txtFromYYYYMM.Alt,frm1.txtToYYYYMM.Alt, _
	 "970024", frm1.txtFromYYYYMM.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtFromYYYYMM.Focus
		Exit Function
	End If    
    '-----------------------
    'Query function call area
    '-----------------------



    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'��: Query db data
	End If
	
    FncQuery = True															'��: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
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
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
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
	FncExit = True
End Function

'******************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  **************************
'	���� : 
'**************************************************************************************** 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal
    Dim strYear,strMonth,strDay
    
    DbQuery = False

	Call LayerShowHide(1)
    
    Call ExtractDateFrom(frm1.txtFromYyyyMm.Text,frm1.txtFromYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strFromYYYYMM = strYear & strMonth

    Call ExtractDateFrom(frm1.txtToYyyyMm.Text,frm1.txtToYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strToYYYYMM = strYear & strMonth

	With frm1
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtFromYyyymm="		& strFromYYYYMM			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToYyyymm="		& strToYYYYMM			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemAcctCd="		& .txtItemAcctCd.Value
			strVal = strVal & "&txtMovTypeCd="		& .txtMovTypeCd.Value
			strVal = strVal & "&txtRadio="			& lgRadio
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtFromYyyymm="		& strFromYYYYMM			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToYyyymm="		& strToYYYYMM			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemAcctCd="		& .txtItemAcctCd.Value
			strVal = strVal & "&txtMovTypeCd="		& .txtMovTypeCd.Value
			strVal = strVal & "&txtRadio="			& lgRadio
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
		End If
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 
	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	
		
End Function


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag�� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��������������ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
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
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�۾����</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/c3970ma1_fpDateTime1_txtFromYYYYMM.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/c3970ma1_fpDateTime2_txtToYYYYMM.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_Sum Checked tag = 2 value="01" onclick=radio1_onchange()><LABEL FOR=Rb_Sum>����</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Dtl tag = 2 value="02" onclick=radio2_onchange()><LABEL FOR=Rb_Dtl>��</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Sum1 tag = 2 value="03" onclick=radio3_onchange()><LABEL FOR=Rb_Sum1>����Sim</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Dtl1 tag = 2 value="04" onclick=radio4_onchange()><LABEL FOR=Rb_Dtl1>��Sim</LABEL></TD>										        							
								</TR>
								<TR>								
									<TD CLASS="TD5">ǰ�����</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemAcctCd" SIZE=9 MAXLENGTH=10 tag="11XXXU" ALT="ǰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
										 <INPUT TYPE=TEXT ID="txtItemAcctNm" NAME="txtItemAcctNm" SIZE=25 tag="14X">
									</TD>
									<TD CLASS="TD5">��������</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtMovTypeCd" SIZE=9 MAXLENGTH=10 tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovTypeCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2)">
										 <INPUT TYPE=TEXT ID="txtMovTypeNm" NAME="txtMovTypeNm" SIZE=25 tag="14X">
									</TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=top COLSPAN=4>
						<script language =javascript src='./js/c3970ma1_vaSpread1_vspdData1.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMovTypeCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

