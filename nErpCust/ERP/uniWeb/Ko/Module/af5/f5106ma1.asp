
<%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module��          : �ڱ� 
'*  2. Function��        : 
'*  3. Program ID        : f5106ma1
'*  4. Program �̸�      : �����Ϻ�������ȸ 
'*  5. Program ����      : �����Ϻ��� ������ ���� ��ȸ��� 
'*  6. Comproxy ����Ʈ   : fn0018_List_Note_Svr
'*                         fn0028_List_Note_Dtl_Svr
'*  7. ���� �ۼ������   : 2000/10/12
'*  8. ���� ���������   : 2000/10/16
'*  9. ���� �ۼ���       : Hersheys
'* 10. ���� �ۼ���       : Hersheys
'* 11. ��ü comment      :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'=============================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "f5106mb1.asp"								'��: �����Ͻ� ���� ASP�� 


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns

Const C_MaxKey       = 2										'�١١١�: Max key value

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'==========================================  1.2.4 User Defined Variable�� ����  =========================

'Dim lgPageNo
Dim intItemCnt
Dim lgIsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 


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

    lgIntFlgMode = parent.OPMD_CMODE							'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
    
    '---- Coding part--------------------------------------------------------------------
    lgPageNo = ""
	lgIsOpenPop = False									'��: ����� ���� �ʱ�ȭ		
    lgSortKey = 1
    
    
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

    Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	toDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	frm1.txtFrDt.Text = toDt
	frm1.txtToDt.Text = toDt
	Call ggoOper.FormatDate(frm1.txtFrDt,  parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtToDt,  parent.gDateFormat, 1)	
	frm1.txtFrDt.focus	

End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'======================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
        Call SetZAdoSpreadSheet("f5106ma1","S","A","V20021215",Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
	    Call SetSpreadLock("A") 
End Sub


'=======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(Byval iOpt)
    If iOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		End With
    End If
End Sub

'========================== 2.2.5 SetSpreadColor() ======================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow)

    With frm1

		.vspdData.ReDraw = False
		.vspdData.ReDraw = True

    End With

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
		
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))

End Sub

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

'++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If lgIsOpenPop = True Then Exit Function
	
	
	lgIsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If
End Function

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function		

	Select Case iWhere
		Case 0		' ���� 
			arrParam(0) = "���� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK"			 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "����"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BANK_CD"						' Field��(0)
			arrField(1) = "BANK_NM"					' Field��(1)
    
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)
		
		Case 2		' �������� 

			If frm1.txtStsCd.className = parent.UCN_PROTECTED Then Exit Function

				arrParam(0) = "���������˾�"
				arrParam(1) = "B_MINOR "								'popup�� sql�� 
				arrParam(2) = strCode
				arrParam(3) = ""
				
				arrParam(4) = "MAJOR_CD = " & FilterVar("F1003", "''", "S") & "  " 
				arrParam(5) = "��������"
	
				arrField(0) = "MINOR_CD"					' form1�� ������ minor_cd,nmǥ�� 
				arrField(1) = "MINOR_NM"
				    
				arrHeader(0) = "���������ڵ�"	
				arrHeader(1) = "�������¸�"
			
		Case Else
			Exit Function
	End Select
    
	lgIsOpenPop = True		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False		
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function 


'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
'------------------------------------------  EscPopUp()  --------------------------------------------------
'	Name : EscPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function EscPopUp(Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' ���� 
				.txtBankCd.focus
			Case 1		'�ŷ�ó 
				.txtBpCd.focus
			Case 2		' �������� 
				.txtStsCd.focus
		End Select

	End With
	
End Function
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' ���� 
				.txtBankCd.value = arrRet(0)
				.txtBankNM.value = arrRet(1)
				.txtBankCd.focus
			Case 1		'�ŷ�ó 
				.txtBpCd.value	= arrRet(0)
				.txtBpNM.value	= arrRet(1)
				.txtBpCd.focus
			Case 2		' �������� 
				.txtStsCd.value = arrRet(0)
				.txtStsNm.value = arrRet(1)
				.txtStsCd.focus
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function


'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value = arrRet(0)
			frm1.txtBizAreaNm1.Value = arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

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

'*****************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call InitVariables                                                      '��: Initializes local global variables
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitComboBox
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 

	' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing

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


'***********************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 


'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData_onfocus()
End Sub


'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData2_onfocus()
End Sub


'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.cboNoteFg.focus
		frm1.txtFrDt.focus
	   Call MainQuery
	End If   
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.cboNoteFg.focus
		frm1.txtToDt.focus
	   Call MainQuery
	End If   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
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
		Exit Sub
	End If

	If frm1.vspdData.MaxRows <= 0 Then
		Exit Sub
	End If

	If Row = frm1.vspdData.ActiveRow Then
		Exit Sub
	End If
	
	ggoSpread.Source = frm1.vspdData
	
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	
    gMouseClickStatus = "SP2C"	'Split �����ڵ� 
        
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
End Sub

'========================================================================================== 
' Event Name : vspdData_LeaveCell 
' Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If

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
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then						'��: ������ üũ 
		If lgPageNo <> "" Then
	       Call DisableToolBar(parent.TBC_QUERY)
	       If DbQuery = False Then
				Call RestoreToolBar()
	          Exit Sub
	       End if
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
'******************************************************************************************************** 


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    'Dim RetFlag
    
    FncQuery = False									'��: Processing is NG
    Err.Clear											'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    '-----------------------
    'Erase contents area
    '-----------------------
	' ���� Page�� Form Element���� Clear�Ѵ�. 
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData
    Call InitVariables									'��: Initializes local global variables
	frm1.vspdData.MaxRows = 0
	
    '-----------------------
    'Check condition area
    '-----------------------
	' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
    If Not chkField(Document, "1") Then					'��: This function check indispensable field
		Exit Function
    End If

	If Trim(frm1.txtBizAreaCd.value) <> "" And Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
    
    If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If frm1.txtBizAreaCd1.value = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'��: Query db data

    FncQuery = True																'��: Processing is OK
    Set gActiveElement = document.activeElement    
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement    
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
	Set gActiveElement = document.activeElement    
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                     '��:ȭ�� ����, Tab ���� 
	Set gActiveElement = document.activeElement    
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
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'*****************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'******************************************************************************************************** 


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal
	'Dim RetFlag

    Err.Clear                '��: Protect system from crashing
    DbQuery = False

    Call LayerShowHide(1)

    With frm1
		strVal = BIZ_PGM_ID
		
	    If lgIntFlgMode = parent.OPMD_UMODE Then			'parent.OPMD_UMODE-->include file::ccm.vbs
	    
			strVal = strVal & "?txtMode	="			& parent.UID_M0001
			strVal = strVal & "&txtFrDt		="		& Trim(.hFrDt.value)			
			strVal = strVal & "&txtToDt		="		& Trim(.hToDt.value)
			strVal = strVal & "&cboNoteFg	="		& Trim(.hcboNoteFg.value)
			strVal = strVal & "&txtBankCd	="		& Trim(.htxtBankCd.value)
			strVal = strVal & "&txtBpCd="			& Trim(.htxtBpCd.value)					'��: ��ȸ ���� ����Ÿ(�ŷ�ó�ڵ�)
			strVal = strVal & "&txtStsCd="			& Trim(.htxtStsCd.value)				'��: ��ȸ ���� ����Ÿ(����)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.htxtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="	& Trim(frm1.txtBizAreaCd1.alt)
	    Else
			strVal = strVal & "?txtMode	="			& parent.UID_M0001 
			strVal = strVal & "&txtFrDt		="		& Trim(.txtFrDt.text)			
			strVal = strVal & "&txtToDt		="		& Trim(.txtToDt.text)
			strVal = strVal & "&cboNoteFg	="		& Trim(.cboNoteFg.value)
			strVal = strVal & "&txtBankCd	="		& Trim(.txtBankCd.value)
			strVal = strVal & "&txtBpCd="			& Trim(frm1.txtBpCd.value)				'��: ��ȸ ���� ����Ÿ	
			strVal = strVal & "&txtStsCd="			& Trim(frm1.txtStsCd.value)				'��: ��ȸ ���� ����Ÿ(����)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="	& Trim(frm1.txtBizAreaCd1.alt)
			
		End If

		strVal = strVal & "&lgPageNo="			& lgPageNo								'��: Next key tag
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))
   
		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 
		    
    End With
    
    DbQuery = True

End Function


'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk()
	
	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
		lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is [[[Update mode]]]
		lgBlnFlgChgValue = False
		    
		Call ggoOper.LockField(Document, "I")									'This function lock the suitable field

		Call SetToolbar("11000000000111")										'��: ��ư ���� ���� 
    End With

End Function


Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
	
		For intRow = 1 To .MaxRows
			.Row = intRow

			.Col = C_NoteSts
			intIndex = .value
			.col = C_NoteStsNm
			.value = intindex
		Next	
		
	End With
	
End Sub


Sub InitDataDtl()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData2
	
		For intRow = 1 To .MaxRows
			
			.Row = intRow

			.Col = C_DtlNoteSts
			intIndex = .value
			.col = C_DtlNoteStsNm
			.value = intindex

		Next	
		
	End With
	
End Sub


'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################

    '----------  Coding part  -------------------------------------------------------------
    
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
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="���۸�����" id=fpDateTime2></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="���Ḹ����" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP>
										<SELECT ID="cboNoteFg"  NAME="cboNoteFg"  ALT="��������" STYLE="WIDTH: 105px" tag="12"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID="txtBankNm" NAME="txtBankNm" SIZE=20 MAXLENGTH=20 tag="14XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd1.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10   tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value, 1)"> &nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="14X" ALT="�ŷ�ó"> </TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtStsCd" NAME="txtStsCd" SIZE=10 MAXLENGTH=10   tag="11XXXU" ALT="���౸��" value=""><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtStsCd.Value, 2)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtStsNM" NAME="txtStsNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="14X" ALT="���౸��"> </TD>
									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=10>
									<FIELDSET CLASS="CLSFLD">
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
												<TD CLASS="TD5" NOWRAP>�����ݾ��հ�</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNoteAmtSum" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����ݾ��հ�" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</FIELDSET>
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
		<TD WIDTH="100%" HEIGHT=20>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread  tag="24"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3 tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"        tag="24">
<INPUT TYPE=hidden NAME="txtCommandMode" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows"     tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     tag="24">
<INPUT TYPE=hidden NAME="hFrDt"          tag="24">
<INPUT TYPE=hidden NAME="hToDt"          tag="24">
<!--<INPUT TYPE=hidden NAME="hNoteFg"        tag="24">-->
<INPUT TYPE=HIDDEN NAME="hcboNoteFg"     tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBankCd"     tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"       tag="2">
<INPUT TYPE=HIDDEN NAME="htxtStsCd"      tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">


<INPUT TYPE=hidden NAME="txtMaxRows3" tag="24">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 tag="2" width="100%"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

