<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : p6
'*  2. Function Name        : ��������������ȸ(HB)
'*  3. Program ID           : p6220QA1
'*  4. Program Name         : ��������������ȸ(HB)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/05/03
'*  8. Modified date(Last)  : 2005/07/20
'*  9. Modifier (First)     : Yoo Myung Sik
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

<%'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
Const BIZ_PGM_ID = "P6220QB1.asp"												'��: �����Ͻ� ���� ASP�� 

<%'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================%>

Dim C_FAC_CAST_CD			'= 1
Dim C_CAST_NM				'= 2
Dim C_WORK_DT				'= 3
Dim C_MINOR_NM			'= 4
Dim C_INSP_TEXT			'= 5
Dim C_BP_NM				'= 6
Dim C_NAME				'= 7
Dim C_BIGO				'= 8

Const C_SHEETMAXROWS = 30

<% '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= %>
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey
<% '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
<% '----------------  ���� Global ������ ����  ----------------------------------------------------------- %>
Dim IsOpenPop 
Dim lsDnNo 
Dim iDBSYSDate
Dim EndDate, StartDate,ACT_ROW,selChk,EndDate_,StartDate_

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = DateAdd("d", -7, EndDate)


<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>
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
    lgLngCurRows = 0                            'initializes Deleted Rows Count
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

	frm1.txtReqdlvyFromDt.text = StartDate
	frm1.txtReqdlvyToDt.text = Enddate
	Call BtnDisabled(1)
	selChk=false

End Sub

<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
<% '== ��ȸ,��� == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029(gCurrency, "I", "*") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
		
	ggoSpread.Source = frm1.vspdData
		
	ggoSpread.Spreadinit	"V20021108",, parent.gAllowDragDropSpread    
		
	Call AppendNumberPlace("6", "5", "0")
		
	With frm1.vspdData
			
		.ReDraw = False
				  
		.MaxCols = C_BIGO + 1
		.MaxRows = 0
				
				
		Call ggoSpread.ClearSpreadData()	
				
		Call GetSpreadColumnPos("A")
	
		.ReDraw = false

	    ggoSpread.Source = frm1.vspdData			 

		ggoSpread.SSSetEdit		C_FAC_CAST_CD			, "�����ڵ�"		, 10	
		ggoSpread.SSSetEdit		C_CAST_NM				, "������"			, 20
		ggoSpread.SSSetDate		C_WORK_DT				, "�۾�����"		, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_MINOR_NM				, "��������"		, 10
		ggoSpread.SSSetEdit		C_INSP_TEXT				, "���˳���"		, 40 
		ggoSpread.SSSetEdit		C_BP_NM					, "�ŷ�ó"			, 15
		ggoSpread.SSSetEdit		C_NAME					, "�۾���"			, 10
		ggoSpread.SSSetEdit		C_BIGO					, "���"			, 10
		
				
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		

		ggoSpread.SpreadLockWithOddEvenRowColor()

		
		.ReDraw = True
    
    End With
    
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  	
	
	C_FAC_CAST_CD			= 1
	C_CAST_NM				= 2
	C_WORK_DT				= 3
	C_MINOR_NM				= 4
	C_INSP_TEXT				= 5
	C_BP_NM					= 6
	C_NAME					= 7
	C_BIGO					= 8

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
		
			C_FAC_CAST_CD			= iCurColumnPos(1)
			C_CAST_NM				= iCurColumnPos(2)
			C_WORK_DT				= iCurColumnPos(3)
			C_MINOR_NM				= iCurColumnPos(4)
			C_INSP_TEXT				= iCurColumnPos(5)			
			C_BP_NM					= iCurColumnPos(6)
			C_NAME					= iCurColumnPos(7)		
			C_BIGO					= iCurColumnPos(8)
		
    End Select    
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

		ggoSpread.spreadlock		C_FAC_CAST_CD				, -1		
		ggoSpread.spreadlock		C_CAST_NM					, -1
		ggoSpread.spreadlock		C_WORK_DT					, -1
		ggoSpread.spreadlock		C_MINOR_NM					, -1
		ggoSpread.spreadlock		C_INSP_TEXT					, -1
		ggoSpread.spreadlock		C_BP_NM						, -1
		ggoSpread.spreadlock		C_NAME						, -1
		ggoSpread.spreadlock		C_BIGO						, -1
	
    .vspdData.ReDraw = True

    End With

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
<% '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++%>

Function OpenRequried(ByVal iRequried)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried

	Case 1												
	
		arrParam(0) = "�����ڵ���ȸ"					<%' �˾� ��Ī %>
		arrParam(1) ="Y_CAST"	<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtCastCd.value)		<%' Code Condition%>
		'arrParam(3) = Trim(frm1.txtDn_TypeNm.value)		<%' Name Cindition%>
		arrParam(4) = " " 
		arrParam(5) = "�����ڵ�"			  	   <%' TextBox ��Ī %>

		arrField(0) = "CAST_CD"							<%' Field��(0)%>
		arrField(1) = "CAST_NM"							<%' Field��(1)%>

		arrHeader(0) = "�����ڵ�"					<%' Header��(0)%>
		arrHeader(1) = "������Ī"					<%' Header��(1)%>

			 
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRequried(arrRet,iRequried)
	End If	
	
End Function

<% '------------------------------------------  SetRequried()  --------------------------------------------------
'	Name : SetRequried()
'	Description : �ŷ�ó Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetRequried(Byval arrRet,ByVal iRequried)

	Select Case iRequried
	Case 1
		
		frm1.txtCastCd.value = Trim(arrRet(0))
		frm1.txtCastNM.value = Trim(arrRet(1))	
			
	End Select
	
	lgBlnFlgChgValue=true
	

End Function


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

	Err.Clear

    Call LoadInfTB19029	

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec) 'condition
    
    Call ggoOper.LockField(Document, "N")      
    
    
	'----------  Coding part  -------------------------------------------------------------

	Call InitSpreadSheet

	Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 

    Call InitVariables                                                      '��: Initializes local global variables

	Call SetDefaultVal
	
	frm1.txtCastCd.focus
	
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
<%
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
%>


<%
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
     '----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 	End If
End Sub

<%
'==========================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_MouseDown(Button , Shift , x , y)


    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

<%
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_Change(ByVal Col , ByVal Row )


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

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
    Call ggoSpread.ReOrderingSpreadData
    
End Sub 

<%
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
%>
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ 
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			DbQuery
		End If
    End if
    
End Sub

<%
'==========================================================================================
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() �� Calendar Popup
'==========================================================================================
%>


<%
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
%>

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
Function FncQuery() 
	Call BtnDisabled(1)
	Dim IntRetCD
	
	selChk=false

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 				
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
 										'��: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then
        Call RestoreToolBar()
        Exit Function
    End If      																'��: Query db data
    
    FncQuery = True																'��: Processing is OK

End Function

<%
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
%>
Function FncPrint() 
    Call parent.FncPrint()
End Function

<%
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
%>
Function FncExcel() 
	Call parent.FncExport(C_MULTI)
End Function

<%
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
%>
Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function

<%
'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
%>
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    
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
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
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

    DbQuery = False
    
    Err.Clear																	 '��: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
    
    
    With frm1

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001		
	
		strVal = strVal & "&txtCastCd=" & Trim(.txtCastCd.value)
		strVal = strVal & "&txtReqdlvyFromDt=" & Trim(.txtReqdlvyFromDt.text)
		strVal = strVal & "&txtReqdlvyToDt=" & Trim(.txtReqdlvyToDt.text)

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	        
    End With
	
    DbQuery = True

End Function

<%
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
%>

Sub txtReqdlvyToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqdlvyToDt.Action = 7
	End If
End Sub

Sub txtReqdlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call fncquery()
End Sub

Sub txtReqdlvyFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqdlvyFromDt.Action = 7
	End If
End Sub

Sub txtReqdlvyFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call fncquery()
End Sub


Function DbQueryOk()														'��: ��ȸ ������ ������� 

	lgIntFlgMode = parent.OPMD_UMODE	
	lgBlnFlgChgValue = False
    '-----------------------
    'Reset variables area
    '-----------------------
	Call SetToolbar("11000000000111")										'��: ��ư ���� ���� 
    Call ggoOper.LockField(Document, "2")									'��: This function lock the suitable field
End Function


'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	
    Dim strEbrFile
    Dim objName
    
	Dim var1

	dim strUrl
	dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)
	
	If frm1.vspdData.ActiveRow < 1 Then
		msgbox "���� ����� �����ڵ带 Ŭ���Ͻʽÿ�"
		exit function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_FAC_CAST_CD
		var1 = Trim(.Text)
	End With
	
	strUrl = "cast_cd|" & var1
	
	strEbrFile = "P6220OA1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()                                                    '��: Protect system from crashing
    
    Dim strEbrFile
    Dim objName
    
	Dim var1

	
	dim strUrl
	dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)

	If frm1.vspdData.ActiveRow < 1 Then
		msgbox "���� ����� �����ڵ带 Ŭ���Ͻʽÿ�"
		exit function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_FAC_CAST_CD
		var1 = Trim(.Text)
	End With
	
	strUrl = "cast_cd|" & var1 
	
	ObjName = AskEBDocumentName("P6220OA1","ebr")

	call FncEBRPreview(objName, strUrl)

	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

<%
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
%>

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### %>
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
					<TD CLASS="CLSMTABP" colspan=2>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������������ȸ</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH= align=right></TD><TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>�۾�����</TD>
									<TD CLASS=TD6 NOWRAP><TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/p6220qa1_fpDateTime1_txtReqdlvyFromDt.js'></script>
											&nbsp;~&nbsp;
											<script language =javascript src='./js/p6220qa1_fpDateTime1_txtReqdlvyToDt.js'></script>
											</TD>
										</TR>
													</TABLE></TD>
									<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCastCd" ALT="�����ڵ�" TYPE="Text" MAXLENGTH="13" SIZE=10 tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtCastNM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% >
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/p6220qa1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA Class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPostFlag" tag="14">

<INPUT TYPE=HIDDEN NAME="txtHDn_Type" tag="24">
<INPUT TYPE=HIDDEN NAME="hcastcd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSo_no" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHShip_to_party" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHReqGiDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHReqGiDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTrans_meth" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPostGiFlag" tag="24">


</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>