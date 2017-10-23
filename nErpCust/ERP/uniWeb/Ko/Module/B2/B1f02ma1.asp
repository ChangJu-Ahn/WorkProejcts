

<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Translation of Unit for Item)
'*  3. Program ID           : B1f02ma1.asp
'*  4. Program Name         : B1f02ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/15
'*  7. Modified date(Last)  : 2002/12/12
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
		Response.Expires = -1
%>
<HTML>
<HEAD>
<TITLE> <% =Request("strASPMnuMnuNm") %> </TITLE>
<%'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

<%'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
Const BIZ_PGM_ID = "B1f02mb1.asp"			'��: �����Ͻ� ���� ASP�� 
<%'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================%>
Dim C_Item
Dim C_ItemPopup
Dim C_ItemNm
Dim C_Unit
Dim C_Sep1
Dim C_ToUnit
Dim C_Equal
Dim C_Factor
Dim C_Sep2
Dim C_ToFactor

Const C_SHEETMAXROWS = 30 'Sheet Max Rows

<% '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= %>

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrQueryFlag		'"":New, "P":Prev, "N":Next, "R":After Save ...Query

Dim lgStrPrevUnit
Dim lgStrPrevToUnit
'Dim lgLngCurRows

<% '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
<% '----------------  ���� Global ������ ����  ----------------------------------------------------------- %>
'Dim lgSortKey
Dim IsOpenPop        
<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>

<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
    C_Item          = 1
    C_ItemPopup     = 2
    C_ItemNm        = 3
    C_Unit          = 4
    C_Sep1          = 5
    C_ToUnit        = 6
    C_Equal         = 7
    C_Factor        = 8
    C_Sep2          = 9
    C_ToFactor      = 10
End Sub

<% '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= %>
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
	lgStrPrevUnit = ""
	lgStrPrevToUnit = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

<% '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* %>
<% '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= %>
Sub SetDefaultVal()
End Sub

<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

<%
'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
%>
Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_ToFactor + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetEdit C_Item, "ǰ���ڵ�", 19,,,18 '1
    ggoSpread.SSSetButton C_ItemPopup '2
    ggoSpread.SSSetEdit C_ItemNm, "ǰ���", 30 '3
    ggoSpread.SSSetEdit C_Unit, "���ش���", 15,,,3 '4
    ggoSpread.SSSetEdit C_Sep1, ":", 2, 2 '5
    ggoSpread.SSSetEdit C_ToUnit, "��ȯ����", 15,,,3 '6
    ggoSpread.SSSetEdit C_Equal, "=", 2, 2 '7
    
    'SetSpreadFloat C_Factor, "���ذ��", 14, 1, 3 '8
    ggoSpread.SSSetFloat C_Factor,"���ذ��",14,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
    
    ggoSpread.SSSetEdit C_Sep2, ":", 2, 2 '9
    
    'SetSpreadFloat C_ToFactor, "��ȯ���", 14, 1, 3 '10
    ggoSpread.SSSetFloat C_ToFactor,"��ȯ���",14,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
    
    call ggoSpread.MakePairsColumn(C_Item,C_ItemPopup)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With

End Sub

<%
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
%>
Sub SetSpreadLock()
    With frm1
	    .vspdData.ReDraw = False
		ggoSpread.SpreadLock C_Item, -1, C_Item
		ggoSpread.SpreadLock C_ItemPopup, -1, C_ItemPopup
		ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
		ggoSpread.SpreadLock C_Unit, -1, C_Unit
		ggoSpread.SpreadLock C_Sep1, -1, C_Sep1
		ggoSpread.SpreadLock C_ToUnit, -1, C_ToUnit
		ggoSpread.SpreadLock C_Equal, -1, C_Equal
		ggoSpread.SpreadLock C_Sep2, -1, C_Sep2
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
	    .vspdData.ReDraw = True
    End With
End Sub

<%
'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
%>
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_Item, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ToUnit, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Sep1, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Sep2, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Equal, pvStartRow, pvEndRow
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
            C_Item       = iCurColumnPos(1)   
            C_ItemPopup  = iCurColumnPos(2)
            C_ItemNm     = iCurColumnPos(3)
            C_Unit       = iCurColumnPos(4)
            C_Sep1       = iCurColumnPos(5)
            C_ToUnit     = iCurColumnPos(6)
            C_Equal      = iCurColumnPos(7)
            C_Factor     = iCurColumnPos(8)
            C_Sep2       = iCurColumnPos(9)
            C_ToFactor   = iCurColumnPos(10)
    End Select    
End Sub


<% '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= %>
Sub InitComboBox()
    Dim strCboData 
    Dim strCboData2
    Dim i 
    
    strCboData = ""
    strCboData2 = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0005", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboDimension, lgF0, lgF1, Chr(11))
	
    strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)  

End Sub

<% '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* %>

<% '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= %>
<% '----------------------------------------  OpenUnit()  ------------------------------------------
'	Name : OpenUnit()
'	Description : Country PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"					<%' �˾� ��Ī %>
	arrParam(1) = "b_unit_of_measure"			<%' TABLE ��Ī %>
	arrParam(2) = frm1.txtUnit.value			<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = "dimension= " & FilterVar(frm1.cboDimension.value, "''", "S") & "" 	<%' Where Condition%>
	arrParam(5) = "����"						<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "unit"						<%' Field��(0)%>
    arrField(1) = "unit_nm"						<%' Field��(1)%>
    
    arrHeader(0) = "����"						<%' Header��(0)%>
    arrHeader(1) = "������"						<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUnit(arrRet)
	End If	
	
End Function


<% '----------------------------------------  OpenItem()  ------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ�� �˾�"				<%' �˾� ��Ī %>
	arrParam(1) = "b_item"					<%' TABLE ��Ī %>
	arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>
	arrParam(3) = ""						<%' Name Cindition%>
	arrParam(4) = ""						<%' Where Condition%>
	arrParam(5) = "ǰ��"					<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "item_cd"					<%' Field��(0)%>
    arrField(1) = "item_nm"					<%' Field��(1)%>
    
    arrHeader(0) = "ǰ���ڵ�"				<%' Header��(0)%>
    arrHeader(1) = "ǰ���"					<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItem(arrRet)
	End If	
	
End Function

<% '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= %>
<% '------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetUnit(Byval arrRet)
	With frm1
		.txtUnit.value = arrRet(0)
		.txtUnitNm.value = arrRet(1)
	End With
End Function

<% '------------------------------------------  SetItem()  --------------------------------------------------
'	Name : SetItem()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetItem(Byval arrRet)
	With frm1
		.vspdData.Col = C_Item
		.vspdData.Text = arrRet(0)
		
		.vspdData.Col = C_ItemNm
		.vspdData.Text = arrRet(1)
		
		lgBlnFlgChgValue = True
	End With
End Function

<% '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<% '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################%>
<% '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>
<% '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= %>
Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100110111101111")										<%'��ư ���� ���� %>    
    frm1.cboDimension.focus
    
End Sub
<%
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
%>
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

<% '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* %>

<% '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

<%'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================%>
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
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
End Sub


<%
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
%>
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	With frm1.vspdData 
		If Row > 0 And Col = C_ItemPopUp Then
		    .Row = Row
		    .Col = C_Item

		    Call OpenItem()
		End If
    End With
End Sub

<%
'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
%>
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

<%
'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 '��: Indicates that value changed
End Sub

<%
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
%>
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'��: ��ȸ���̸� ���� ��ȸ ���ϵ��� üũ 
        Exit Sub
	End If
	    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
	And Not(lgStrPrevKey = "" And lgStrPrevUnit = "" And lgStrPrevToUnit = "") Then '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
		Call DisableToolBar(parent.TBC_QUERY)					'�� : Query ��ư�� disable ��Ŵ.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
		End If 
    End if
    
End Sub

Sub cboDimension_onchange()
	frm1.txtUnit.value = ""
	frm1.txtUnitNm.value = ""
End Sub

<% '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### %>


<% '#########################################################################################################
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
'######################################################################################################### %>
<% '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
		
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    frm1.txtUnitNm.value = ""

    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
    FncQuery = True
            
End Function

<%
'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
%>
Function FncNew() 
End Function

<%
'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
%>
Function FncDelete() 
End Function

<%
'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
%>
Function FncSave() 
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then    'Not chkField(Document, "2") OR     '��: Check contents area
       Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

<%
'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
%>
Function FncCopy()
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

    With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
			
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
			'Key field clear
			.Col = C_Item
			.Text = ""

			.Col = C_ItemNm
			.Text = ""
			
			.ReDraw = True
		End If
    End With

    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
    
End Function

<%
'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
%>
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

<%
'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
%>
Function FncInsertRow(ByVal pvRowCnt)				'020325 �������� �ѷ��ִ� ���� 
    Dim IntRetCD
    Dim imRow
    Dim iRow
    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
    
	If Trim(frm1.txtTUnit.value)  = "" Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002", "X", "X", "X")                                '��:
        Exit Function
	End If

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		'.vspdData.EditMode = True
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    
	    .vspdData.Row = .vspdData.ActiveRow    ' 020322â �������� 

        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
		    .vspdData.Row = iRow
		    .vspdData.Col = C_Unit
		    .vspdData.Value = .txtFUnit.value
  
		    .vspdData.Col = C_ToUnit
		    .vspdData.Text = .txtTUnit.value
		    										'vb ����׹�� msgbox "2===>"&.txtFUnit.value
		    .vspdData.Col = C_Sep1
		    .vspdData.Text = ":"

		    .vspdData.Col = C_Equal
		    .vspdData.Text = "="

		    .vspdData.Col = C_Sep2
		    .vspdData.Text = ":"
        Next
		.vspdData.ReDraw = True

    End With
End Function

<%
'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
%>
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1.vspdData 
	    .focus
		ggoSpread.Source = frm1.vspdData 
    
		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
    End With
End Function

<%
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
%>
Function FncPrint() 
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

<%
'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
%>
Function FncPrev() 
	DIm IntRetCD
		
	If frm1.txtFUnit.value = "" Or frm1.txtTUnit.value = "" Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
	End If
		
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    	
	lgStrQueryFlag = "P"
	If DbQuery = False Then Exit Function
End Function

<%
'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
%>
Function FncNext() 
	Dim IntRetCD

	If frm1.txtFUnit.value = "" Or frm1.txtTUnit.value = "" Then
        Call DisplayMsgBox("900002", "X", "X", "X") 
        Exit Function    
	End If
	
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
	
	lgStrQueryFlag = "N"

	If DbQuery = False Then Exit Function
End Function

<%
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
%>
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												<%'��: ȭ�� ���� %>
End Function

<%
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
%>
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         <%'��:ȭ�� ����, Tab ���� %>
End Function

<%
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
%>
Function FncExit()
	Dim IntRetCD
	FncExit = False
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    FncExit = True
End Function

<% '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
	If lgStrQueryFlag <> "P" And lgStrQueryFlag <> "N" And lgStrQueryFlag <> "R" Then
	    strVal = BIZ_PGM_ID  & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal      & "&txtUnit=" & Trim(.txtUnit.value)
	ElseIf lgStrQueryFlag = "R"  Then
	    strVal = BIZ_PGM_ID  & "?txtMode="    & "R"							'��: 
		strVal = strVal      & "&txtFUnit="    & Trim(.txtFUnit.value)
		strVal = strVal      & "&txtFUnitNm=" & Trim(.txtFUnitNm.value)
		strVal = strVal      & "&txtToUnit=" & Trim(.txtTUnit.value)
		strVal = strVal      & "&txtTUnitNm=" & Trim(.txtTUnitNm.value)
		
	Else
	    strVal = BIZ_PGM_ID & "?txtMode="   & lgStrQueryFlag
		strVal = strVal     & "&txtUnit="   & Trim(.txtFUnit.value)
		strVal = strVal     & "&txtToUnit=" & Trim(.txtTUnit.value)

	End If
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = strVal & "&txtDim="          & .hDimension.value 
		strVal = strVal & "&lgStrPrevKey="    & lgStrPrevKey
		strVal = strVal & "&lgStrPrevUnit="   & lgStrPrevUnit
		strVal = strVal & "&lgStrPrevToUnit=" & lgStrPrevToUnit	
		strVal = strVal & "&txtMaxRows="      & frm1.vspdData.MaxRows
    Else
		strVal = strVal & "&txtDim="          & Trim(.cboDimension.value)
		strVal = strVal & "&lgStrPrevKey="    & lgStrPrevKey
		strVal = strVal & "&lgStrPrevUnit="   & lgStrPrevUnit
		strVal = strVal & "&lgStrPrevToUnit=" & lgStrPrevToUnit		
		strVal = strVal & "&txtMaxRows="      & frm1.vspdData.MaxRows
    End If	
			
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    End With
    
	lgStrQueryFlag = ""
    DbQuery = True
    
End Function

<%'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================%>
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									<%'This function lock the suitable field%>

	Call SetToolbar("1100111111111111")										<%'��ư ���� ���� %>
		
End Function

<%
'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data save
'========================================================================================
%>
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim a, b
	
    DbSave = False                                                          '��: Processing is NG
    
    Call LayerShowHide(1)
    
    'On Error Resume Next                                                   '��: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep '��: C=Create, Row��ġ ���� 
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep '��: U=Update, Row��ġ ���� 
			End Select			

		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'��: �ű�, ���� 
		            .vspdData.Col = C_Item	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Unit			'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ToUnit		'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_Factor		'9
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep		            
		            
		            If unicdbl(.vspdData.Text) <= 0 Then
		            	.vspdData.Row = 0		            	
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")

						.vspdData.Row = lRow
						.vspdData.Action = 0 'ActionActiveCell
						.vspdData.EditMode = True

						Call LayerShowHide(0)
						Exit Function
					End If

		            .vspdData.Col = C_ToFactor		'10
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            If unicdbl(.vspdData.Text) <= 0 Then
		            	.vspdData.Row = 0
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")

						.vspdData.Row = lRow
						.vspdData.Action = 0 'ActionActiveCell
						.vspdData.EditMode = True
						Call LayerShowHide(0)
						Exit Function
					End If

		            lGrpCnt = lGrpCnt + 1
                    

		        Case ggoSpread.DeleteFlag								'��: ���� 

					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep	'��: D=Delete

		            .vspdData.Col = C_Item	'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Unit			'4
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ToUnit		'6
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
  
  		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
End Function

<%
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
%>
Function DbSaveOk()													'��: ���� ������ ���� ���� 

	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
	lgStrQueryFlag = "R" 'After Save....Query
	
	Call DBQuery()	
End Function

<%
'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
%>
Function DbDelete() 
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### %>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�� ����ȯ��</font></td>
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
									<TD CLASS="TD5">Dimension</TD>
									<TD CLASS="TD6">
										<SELECT Name="cboDimension" STYLE="WIDTH:150" tag="12"></SELECT>
									</TD>
									<TD CLASS="TD5">���ش���</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtUnit" SIZE=10 MAXLENGTH=3 tag="12xxxU"  ALT="���ش���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnit" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit">
										<INPUT TYPE=TEXT NAME="txtUnitNm" tag="14X">
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
					<TD HEIGHT=100% WIDTH=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% HEIGHT=*>
									<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD CLASS="TD5">���ش���</TD>
											<TD CLASS="TD6">
												<INPUT TYPE=TEXT NAME="txtFUnit" SIZE=12.5 MAXLENGTH=10 tag="24X">
												<INPUT TYPE=TEXT NAME="txtFUnitNm" tag="24X">
											</TD>
											<TD CLASS="TD5">��ȯ����</TD>
											<TD CLASS="TD6">
												<INPUT TYPE=TEXT NAME="txtTUnit" SIZE=12.5 MAXLENGTH=10 tag="24X">
												<INPUT TYPE=TEXT NAME="txtTUnitNm" tag="24X">
											</TD>
										</TR>									
										<TR>
											<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
												
												
												
												<script language =javascript src='./js/b1f02ma1_vaSpread1_vspdData.js'></script>
												
												
												
											</TD>
										</TR>
									</TABLE>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1f02mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hDimension" tag="24">
<INPUT TYPE=HIDDEN NAME="hFUnit" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
