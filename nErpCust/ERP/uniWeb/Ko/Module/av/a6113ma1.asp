<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : A6113MA1
'*  4. Program Name         : ���Ը�������ȸ 
'*  5. Program Desc         : Query of Account Code
'*  6. Component List       : ADO
'*  7. Modified date(First) : 2001.11.15
'*  8. Modified date(Last)  : 2002.09,11
'* 10. Modifier (Last)      : Lee Hye young
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
'Dim lgStrPrevKey                                            '��: Next Key tag                          
'Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim lgIsOpenPop                                          
Dim IsOpenPop                                               '��: Popup status                           
Dim lgMark                                                  '��: ��ũ                                  

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "A6113MB1.asp"
'Dim lsPoNo								                       '��: Jump�� Cookie�� ���� Grid value
Const C_MaxKey          = 2                                    '�١١١�: Max key value
'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
'--------------- ������ coding part(�������,Start)--------------------------------------------------

	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
    EndDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

	frm1.txtIssueDt1.Text = StartDate
	frm1.txtIssueDt2.Text = EndDate


	'frm1.txtBizAreaCD.value	= parent.gBizArea
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A6113MA1","S","A","V20021211",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
        With frm1

        .vspdData.ReDraw = False
        ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
        .vspdData.ReDraw = True

        End With
    End if
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Function InitComboBox()

    '���ⱸ�� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Call SetCombo2(frm1.cboIOFlag ,lgF0  ,lgF1  ,Chr(11))
   
    '�ΰ������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("B9001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Call SetCombo2(frm1.cboVatType ,lgF0  ,lgF1  ,Chr(11))
	
	'�Ű��� 
   	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Call SetCombo2(frm1.cboReportFg ,lgF0  ,lgF1  ,Chr(11))


		'Call InitComboBoxDtl("0", "A1003")		' ���ⱸ�� 
		'Call InitComboBoxDtl("1", "B9001")		' �ΰ������� 
		'Call InitComboBoxDtl("5", "A1020")		' �Ű��� 


End Function
<%
Function InitComboBoxDtl(Byval Index, Byval MajorCd)
                     

End Function
%>

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
			frm1.txtBPCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
		lgBlnFlgChgValue = True
	End If
	

End Function
 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
		arrParam(0) = "���ݽŰ����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_TAX_BIZ_AREA"	 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "TAX_BIZ_AREA_CD"					' Field��(0)
			arrField(1) = "TAX_BIZ_AREA_NM"					' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"				' Header��(0)
			arrHeader(1) = "���ݽŰ������"				' Header��(0)
	
			
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
		Select Case iWhere
			Case 0		' ������ 
				frm1.txtBizAreaCD.focus				
			Case 1		' �ŷ�ó 
				frm1.txtBPCd.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ������ 
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
				
			Case 1		' �ŷ�ó 
				.txtBPCd.focus
				.txtBPCd.value = UCase(Trim(arrRet(0)))
				.txtBPNM.value = arrRet(1)
		End Select
	End With
End Function

'============================================================
'�μ��ڵ� �˾� 
'============================================================
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = parent.UCN_PROTECTED Then Exit Function
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrParam(0) = strCode				'�μ��ڵ� 
	arrParam(1) = frm1.txtLoanDt.Text	'��¥(Default:������)
	arrParam(2) = "1"					'�μ�����(lgUsrIntCd)
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If
	frm1.txtDeptCd.focus
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	
	lgBlnFlgChgValue = True
End Function


'----------------------------------------  OpenAcctCd()  -------------------------------------------------
'	Name : OpenAcctCd()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"	
	arrParam(1) = " A_ACCT A "
	arrParam(2) = Trim(frm1.txtAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "�����ڵ�"			
	
	arrField(0) = "A.ACCT_CD"						' Field��(0)
	arrField(1) = "A.ACCT_NM"						' Field��(1)

    
	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "������"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,2)
	End If	
	
End Function

'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "VSPDDATA2"                  
	            gPos = "B"
	       End Select     
	       
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						 'Cookie Split String : CookiePage Function Use

	If Kubun = 1 Then								 'Jump�� ȭ���� �̵��� ��� 

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie "PoNo" , lsPoNo					 'Jump�� ȭ���� �̵��Ҷ� �ʿ��� Cookie �������� 
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							 'Jump�� ȭ���� �̵��� ������� 

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

'--------------- ������ coding part(�������,Start)---------------------------------------------------
		 '�ڵ���ȸ�Ǵ� ���ǰ��� �˻����Ǻ� Name�� Match 
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("��������")
				frm1.txtPoType.value =  arrVal(iniSep + 1)
			Case UCase("�������¸�")
				frm1.txtPoTypeNm.value =  arrVal(iniSep + 1)
			Case UCase("����ó")
				frm1.txtSpplCd.value =  arrVal(iniSep + 1)
			Case UCase("����ó��")
				frm1.txtSpplNm.value =  arrVal(iniSep + 1)
			Case UCase("���ű׷�")
				frm1.txtPurGrpCd.value =  arrVal(iniSep + 1)
			Case UCase("���ű׷��")
				frm1.txtPurGrpNm.value =  arrVal(iniSep + 1)
			Case UCase("ǰ��")
				frm1.txtItemCd.value =  arrVal(iniSep + 1)
			Case UCase("ǰ���")
				frm1.txtItemNm.value =  arrVal(iniSep + 1)
			Case UCase("Tracking No.")
				frm1.txtTrackNo.value =  arrVal(iniSep + 1)
			End Select
		Next
'--------------- ������ coding part(�������,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call FncQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

	Call InitVariables														'��: Initializes local global variables
    Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Call FncSetToolBar("New")
'	Call CookiePage(0)
'--------------- ������ coding part(�������,End)------------------------------------------------------

frm1.txtIssueDt1.focus
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


'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'==========================================================================================

Sub txtIssueDt1_DblClick(Button)
	if Button = 1 then
		frm1.txtIssueDt1.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
	End if
End Sub

Sub txtIssueDt2_DblClick(Button)
	if Button = 1 then
		frm1.txtIssueDt2.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
	End if
End Sub

Sub txtIssueDt1_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtIssueDt2.focus
		Call FncQuery
	End If
End Sub

Sub txtIssueDt2_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtIssueDt1.focus
		Call FncQuery
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
    If frm1.vspdData.MaxRows = 0 then
        Exit Sub
    End If
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub
    End If
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
'--------------- ������ coding part(�������,End)------------------------------------------------------
    'Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
End Sub
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery
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
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    Call SetToolbar("1100000000011111")
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtIssueDt1.text,frm1.txtIssueDt2.text,frm1.txtIssueDt1.Alt,frm1.txtIssueDt2.Alt, _
        	               "970025",frm1.txtIssueDt1.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtIssueDt1.focus
	   Exit Function
	End If
	
	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery															'��: Query db data

    FncQuery = True		
    Call SetToolbar("1100000000011111")
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
	Call parent.FncExport(parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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

'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(ByVal arrRet,ByVal field_fg) 
	With frm1	
		Select case field_fg
			case 1
				.txtBizAreaCd.focus
				.txtBizAreaCd.Value		= arrRet(0)
				.txtBizAreaNm.Value		= arrRet(1)
			case 2
				.txtAcctCd.focus
				.txtAcctCd.Value		= arrRet(0)
				.txtAcctNm.Value		= arrRet(1)
				
				 'Call DbPopUpQuery()
			case 3
				'.txtSubLedger1.value	= arrRet(0)
				'.txtSubLedger3.value	= arrRet(1)
			case 4											'OpenSubledger2
				'.txtSubLedger2.value	= arrRet(0)
				'.txtSubLedger4.value	= arrRet(1)
		End select	
	End With

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

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtIssueDT1=" & Trim(.txtIssueDT1.Text)
		strVal = strVal & "&txtIssueDT2=" & Trim(.txtIssueDT2.Text)
		strVal = strVal & "&txtBizAreaCd=" & UCase(Trim(.txtBizAreaCd.value))
		strVal = strVal & "&cboReportFg=" & Trim(.cboReportFg.value)
		strVal = strVal & "&txtBPCd=" & UCase(Trim(.txtBPCd.value))
		strVal = strVal & "&cboIOFlag=" & Trim(.cboIOFlag.value)
		strVal = strVal & "&cboVatType=" & Trim(.cboVatType.value)
		

'--------------- ������ coding part(�������,End)------------------------------------------------

		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
			
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����       
        	
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
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

'	IF trim(frm1.txtdeptcd.value) = "" then
'		frm1.txtdeptnm.value = ""
'	end if	
	Call FncSetToolBar("Query")
	'SetGridFocus
		
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################



'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function
'=========================================================================================================


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

<!--########################################################################################################
'       					6. Tag�� 
#########################################################################################################-->

<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���Ը�������ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT></TD>					
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
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6113ma1_fpDateTime2_txtIssueDt1.js'></script>
													&nbsp;~&nbsp;
													<script language =javascript src='./js/a6113ma1_fpDateTime2_txtIssueDt2.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="���ݽŰ�����" tag="11XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=25 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="���ݽŰ������" tag="24X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�Ű���</TD>
									<TD CLASS="TD6"><SELECT ID="cboReportFg" NAME="cboReportFg" ALT="�Ű���" STYLE="WIDTH: 98px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
													
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBPCd" NAME="txtBPCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBPCd.Value, 1)">&nbsp;<INPUT TYPE=TEXT ID="txtBPNm" NAME="txtBPNm" SIZE=25 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ�ó��" tag="24X" ></TD>
													
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ⱸ��</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag" NAME="cboIOFlag" ALT="���ⱸ��" STYLE="WIDTH: 98px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>�ΰ�������</TD>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN=7>
								<script language =javascript src='./js/a6113ma1_vaSpread1_vspdData.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>����ó  :</TD>
								<TD CLASS="TD18">�ż��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6113ma1_fpDoubleSingle1_txtCntSumI.js'></script></TD>
								<TD CLASS="TD18">���ް��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6113ma1_fpDoubleSingle1_txtAmtSumI.js'></script></TD>
								<TD CLASS="TD18">�ΰ����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6113ma1_fpDoubleSingle1_txtVatSumI.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>����ó  :</TD>
								<TD CLASS="TD18">�ż��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6113ma1_fpDoubleSingleO_txtCntSumO.js'></script></TD>
								<TD CLASS="TD18">���ް��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6113ma1_fpDoubleSingleO_txtAmtSumO.js'></script></TD>
								<TD CLASS="TD18">�ΰ����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6113ma1_fpDoubleSingleO_txtVatSumO.js'></script></TD>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabIndex="-1">
</TEXTAREA><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" tabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>

