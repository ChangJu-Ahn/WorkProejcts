<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4108ma1
'*  4. Program Name         : ���Աݻ�ȯ��ȹ��� 
'*  5. Program Desc         : Report of Loan Repay Plan
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002.04.25
'*  8. Modified date(Last)  : 2004/02/03
'*  9. Modifier (First)     : Jang, Yoon ki
'* 10. Modifier (Last)      : U & I (Kim Chang Jin)
'* 11. Comment              :
'======================================================================================================= -->
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

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'��: �����Ͻ� ���� ASP�� 
'Const BIZ_PGM_ID = "f5111mb1.asp"			'��: �����Ͻ� ���� ASP�� 


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
'Dim lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 



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

    '---- Coding part--------------------------------------------------------------------    
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

DIm strYear, strMonth, strDay
Dim frDt, toDt
Dim TempDate
Dim strSvrDate

	strSvrDate = "<%=GetSvrDate%>"
	TempDate = UNIDateAdd("M", 1, strSvrDate, parent.gServerDateFormat)
	TempDate = UNIGetLastDay(TempDate, parent.gServerDateFormat)

	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
	frDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)	

	Call ExtractDateFrom(TempDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
	toDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
   
	frm1.txtDateFr.Text = frDt   
	frm1.txtDateTo.Text = toDt
	
	frm1.Rb_Fg1.checked = True	


End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","OA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","OA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================

Sub SetSpreadLock()
End Sub


'================================== 2.2.5 SetSpreadColor() ================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================

Sub SetSpreadColor(ByVal lRow)
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1012", "''", "S") & "  AND MINOR_CD IN(" & FilterVar("U", "''", "S") & " ," & FilterVar("C", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboApSts ,lgF0  ,lgF1  ,Chr(11))
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub



'==========================================  2.2.7 SetCheckBox()  =======================================
'	Name : SetCheckBox()
'	Description : üũ�ڽ� ���� ó��(1���� ���õǵ��� ��)
'========================================================================================================= 
Function SetCheckBox(objCheckBox)
	Dim idx
	
	For idx = 0 To Document.All.Length - 1
		Select Case Document.All(idx).TagName
		Case "INPUT"
			If UCase(Document.All(idx).Type) = "CHECKBOX" Then
				Document.All(idx).Checked = False
			End If
		End Select
	Next
	
	objCheckBox.Checked = True
End Function


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
 '------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtLoanPlcCd.className) = "PROTECTED" Then Exit Function

	
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
		frm1.txtDateFr.Focus
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	

End Function
'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		case 0
			If frm1.txtLoanPlcCd.className = parent.UCN_PROTECTED Then Exit Function	
			If frm1.txtLoanPlcfg1.Checked = true Then
				arrParam(0) = "�����˾�"
				arrParam(1) = "B_BANK A"
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "�����ڵ�"

				arrField(0) = "A.BANK_CD"
				arrField(1) = "A.BANK_NM"
						    
				arrHeader(0) = "�����ڵ�"
				arrHeader(1) = "�����"
			Else
				Call OpenBp(strCode, iWhere)
				exit function
			End If
        
        Case 1	
			arrParam(0) = "���Կ뵵�˾�"			' �˾� ��Ī 
			arrParam(1) = "b_minor" 				    ' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "major_cd=" & FilterVar("f1000", "''", "S") & " "	        ' Where Condition
			arrParam(5) = "���Կ뵵"				' �����ʵ��� �� ��Ī 

			arrField(0) = "minor_cd"						' Field��(0)
			arrField(1) = "minor_nm"						' Field��(1)
    
			arrHeader(0) = frm1.txtLoanType.Alt				' Header��(0)
			arrHeader(1) = frm1.txtLoanTypeNm.Alt				    ' Header��(1)
		Case 2
			arrParam(0) = "�ŷ���ȭ�˾�"								' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY"	 									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' �����ʵ��� �� ��Ī 

		    arrField(0) = "CURRENCY"										' Field��(0)
		    arrField(1) = "CURRENCY_DESC"									' Field��(1)

		    arrHeader(0) = "��ȭ�ڵ�"									' Header��(0)
			arrHeader(1) = "��ȭ�ڵ��"									' Header��(1)
			
		Case 3, 4
			arrParam(0) = "������ڵ� �˾�"			' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 					' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"					' Field��(0)
			arrField(1) = "BIZ_AREA_NM"					' Field��(1)

			arrHeader(0) = "������ڵ�"				' Header��(0)
			arrHeader(1) = "������"				' Header��(1)
				
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDateFr.Focus
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetReturnPopUp()  --------------------------------------------------
'	Name : SetReturnPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' �ŷ�ó 
				frm1.txtLoanPlcCd.value = arrRet(0)
				frm1.txtLoanPlcNm.value = arrRet(1)
				frm1.txtLoanPlcCd.focus
				
			Case 1		'���Կ뵵 
				frm1.txtLoanType.value = arrRet(0)
				frm1.txtLoanTypeNm.value = arrRet(1)
				frm1.txtLoanType.focus
				
			Case 2
				frm1.txtDocCur.value = arrRet(0)
				frm1.txtDocCur.focus
			
			Case 3	'������ڵ� 
				frm1.txtBizAreaCd.value = arrRet(0)
				frm1.txtBizAreaNm.value = arrRet(1)
				frm1.txtBizAreaCd.focus
					
			Case 4	'������ڵ� 
				frm1.txtBizAreaCd1.value = arrRet(0)
				frm1.txtBizAreaNm1.value = arrRet(1)
				frm1.txtBizAreaCd1.focus
		End Select

	End With
	
End Function


Function OpenPopupLoan()

	Dim arrRet
	Dim arrParam(3)	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
   
	arrRet = window.showModalDialog("f4202ra1.asp", Array(arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = ""  Then
		frm1.txtLoanNo.focus
		Exit Function
	Else		
		frm1.txtLoanNo.value = arrRet(0)
		frm1.txtLoanNm.value = arrRet(1)
		frm1.txtLoanNo.focus
	End If
	
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
'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029                           '��: Load table , B_numeric_format
	Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
	
	Call InitVariables                            '��: Initializes local global Variables
	Call SetDefaultVal
	Call txtLoanPlcfg_onchange()

    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call SetToolbar("1000000000001111")				'��: ��ư ���� ���� 
    
	frm1.txtDateFr.focus

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
'	Event�� �浹�� �����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 
'======================================================================================================
'   Event Name : txtLoanPlcfg_onchange
'   Event Desc : 
'=======================================================================================================
Function txtLoanPlcfg_onchange()
	If frm1.txtLoanPlcfg0.checked = true then
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "Q")
		frm1.txtLoanPlcCd.value = ""
		frm1.txtLoanPlcNm.value = ""
	Else
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "D")
	End If
End Function



'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Event ó��	
'********************************************************************************************************* 

'======================================================================================================
'   Event Name : txtDateFr_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================

Sub txtDateFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDateFr.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtDateFr.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtDateTo_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'========================================================================================================
Sub txtDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDateto.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtDateto.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtBankCd_Onchange()
'   Event Desc : �����ڵ带 �����Է��Ұ�쿡 �����ڵ���� �������ش�.
'========================================================================================================
sub txtBankCd_Onchange()
	Dim strCd

	strCd = frm1.txtBankCd.value
	Call CommonQueryRs("A.BANK_NM","B_BANK A","A.BANK_cd = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtBankNM.value = ""
	else
		frm1.txtBankNM.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub

'========================================================================================================
'   Event Name : txtLoanType_Onchange()
'   Event Desc : �����ڵ带 �����Է��Ұ�쿡 �����ڵ���� �������ش�.
'========================================================================================================
sub txtLoanType_Onchange()
	Dim strCd

	strCd = frm1.txtLoanType.value
	Call CommonQueryRs("A.MINOR_NM","B_MINOR A","A.MAJOR_CD = " & FilterVar("F1000", "''", "S") & "  and A.minor_cd = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtLoanTypeNm.value = ""
	else
		frm1.txtLoanTypeNm.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub

'========================================================================================================
'   Event Name : txtBankCd_Onchange()
'   Event Desc : �����ڵ带 �����Է��Ұ�쿡 �����ڵ���� �������ش�.
'========================================================================================================
sub txtLoanPlcCd_Onchange()
	Dim strCd

	strCd = frm1.txtLoanPlcCd.value
	
	If frm1.txtLoanPlcfg1.checked = true Then
		Call CommonQueryRs("A.BANK_NM","B_BANK A","A.BANK_CD = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	ElseIf frm1.txtLoanPlcfg2.checked = true Then
		Call CommonQueryRs("A.BP_NM","B_BIZ_PARTNER A","A.BP_CD = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	End If
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtLoanPlcNm.value = ""
	else
		frm1.txtLoanPlcNm.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub



'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange()
'   Event Desc : ������ڵ带 �����Է��Ұ�쿡 ������ڵ���� �������ش�.
'========================================================================================================
sub txtBizAreaCd_Onchange()
	Dim strCd

	strCd = frm1.txtBizAreaCd.value
	Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtBizAreaNm.value = ""
	else
		frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub


'========================================================================================================
'   Event Name : txtBizAreaCd1_Onchange()
'   Event Desc : ������ڵ带 �����Է��Ұ�쿡 ������ڵ���� �������ش�.
'========================================================================================================
sub txtBizAreaCd1_Onchange()
	Dim strCd

	strCd = frm1.txtBizAreaCd1.value
	Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtBizAreaNm1.value = ""
	else
		frm1.txtBizAreaNm1.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(condvar, StrEbrFile)

	Dim VarPlanFrDT, VarPlanToDt
    Dim VarBizAreaCd, VarBizAreaCd1, VarLoanFg, VarLoanType, VarLoanPlcFg, VarLoanPlcCd, VarDocCur, VarConfFg, VarConfFg1, VarConfFg2, VarRdpClsFg, VarLoanPlcField

	Dim	strAuthCond

	If frm1.Rb_Fg1.checked = True Then	 '���Թ�ȣ�� 		
			StrEbrFile = "f4222oa1a"
	ElseIf frm1.Rb_Fg2.checked = True Then	 '�������ں� 
			StrEbrFile = "f4222oa1c"
	End If
	
	VarLoanFg		= "%"
	VarLoanType		= "%"
	VarLoanPlcFg	= "%"
	VarLoanPlcCd	= "%"
	VarDocCur		= "%"
	VarConfFg1		= "%"
	VarConfFg2		= "%"
	VarRdpClsFg		= "%"
	VarLoanPlcField	= "%"

	VarPlanFrDT = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,parent.gDateFormat, parent.gServerDateType)
	VarPlanToDt = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,parent.gDateFormat, parent.gServerDateType)
	
	If frm1.cboLoanFg.value <> "" Then VarLoanFg	= frm1.cboLoanFg.value
	If Trim(frm1.txtLoanType.value) <> "" THen VarLoanType = Filtervar(Trim(frm1.txtLoanType.value), "", "SNM")
	If frm1.txtLoanPlcfg1.checked = true Then
		VarLoanPlcFg = "BK"
		VarLoanPlcField = "F_LN_INFO.LOAN_BANK_CD"
	ElseIf frm1.txtLoanPlcfg2.checked = true Then
		VarLoanPlcFg = "BP"
		VarLoanPlcField = "F_LN_INFO.BP_CD"
	Else 
		VarLoanPlcFg = "%"
		VarLoanPlcField = "F_LN_INFO.LOAN_NO"
	End If
	If Trim(frm1.txtLoanPlcCd.value) <> "" Then VarLoanPlcCd = Filtervar(Trim(frm1.txtLoanPlcCd.value), "", "SNM")
	If Trim(frm1.txtDocCur.value) <> "" Then VarDocCur	= Filtervar(Trim(frm1.txtDocCur.value), "", "SNM")
	If frm1.cboConfFg.value <> "" Then VarConfFg	= frm1.cboConfFg.value
	If frm1.cboApSts.value <> "" Then VarRdpClsFg = frm1.cboApSts.value
	If VarConfFg = "C" Then
		VarConfFg1 = "C"
		VarConfFg2 = "E"
	ElseIf VarConfFg = "U" Then
		VarConfFg1 = "U"
		VarConfFg2 = "U"
	End If
	
	If frm1.txtBizAreaCd.value = "" then 
		frm1.txtBizAreaNm.value = ""
		VarBizAreaCd = ""
	else 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
	end if
	
	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		VarBizAreaCd1 = "ZZZZZZZZZZ"
	else 
		VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
	end if
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	
	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	condvar = "DateFr|" & VarPlanFrDT
	condvar = condvar & "|DateTo|"			& VarPlanToDt
	condvar = condvar & "|LoanFg|"			& VarLoanFg
	condvar = condvar & "|LoanType|"		& VarLoanType
	condvar = condvar & "|LoanPlcFg|"		& VarLoanPlcFg
	condvar = condvar & "|LoanPlcCd|"		& VarLoanPlcCd
	condvar = condvar & "|DocCur|"			& VarDocCur
	condvar = condvar & "|ConfFg1|"			& VarConfFg1
	condvar = condvar & "|ConfFg2|"			& VarConfFg2
	condvar = condvar & "|RdpClsFg|"		& VarRdpClsFg
	condvar = condvar & "|LoanPlcField|"	& VarLoanPlcField
	condvar = condvar & "|BizAreaCd|"		& VarBizAreaCd
	condvar = condvar & "|BizAreaCd1|"		& VarBizAreaCd1

	condvar = condvar & "|strAuthCond|"		& strAuthCond

End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim condvar
    Dim StrEbrFile
    Dim ObjName
    
    
    'On Error Resume Next       	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, parent.gComDateType, true) = False Then
		frm1.txtDateFr.focus											'��: GL Date Compare Common Function
		Exit Function
	End if
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	Call SetPrintCond(condvar, StrEbrFile)
    
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
    
	Call FncEBRPrint(EBAction,ObjName,Condvar)
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
    
    Dim condvar
    Dim StrEbrFile    
    Dim ObjName
    
 	'On Error Resume Next                                                    '��: Protect system from crashing
   
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, parent.gComDateType, true) = False Then
		frm1.txtDateFr.focus											'��: GL Date Compare Common Function
		Exit Function
	End if
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	Call SetPrintCond(condvar, StrEbrFile)

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
    
	Call FncEBRPreview(ObjName,Condvar)	
			
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


'********************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function

Function FncQuery()
	Exit Function
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function




'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	

</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD HEIGHT=* WIDTH=100%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>��±���</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg1 Checked ><LABEL FOR=Rb_Fg1>���Թ�ȣ��</LABEL>&nbsp;
										<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg2 ><LABEL FOR=Rb_Fg2>�������ں�</LABEL>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���޿�������</TD>
									<TD CLASS=TD6 NOWRAP colspan=2 >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���޿���������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���޿���������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���ۻ����" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,3)"> 
														   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="������">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="��������" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,4)"> 
														   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="������"></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>��ܱⱸ��</TD>
									<TD CLASS="TD6" NOWRAP colspan=2><SELECT NAME="cboLoanFg" ALT="��ܱⱸ��" STYLE="WIDTH: 135px" tag="11X"><OPTION VALUE=""></OPTION></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���Կ뵵</TD>												
								    <TD CLASS=TD6 NOWRAP colspan=2><INPUT NAME="txtLoanType" MAXLENGTH="18" SIZE=10  ALT ="���Կ뵵�ڵ�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.Value, 1)">
									                     <INPUT NAME="txtLoanTypeNm" MAXLENGTH="25" ALT ="���Կ뵵��" tag="14X"></TD>  
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����ó����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg0 VALUE="" Checked tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg0>����+�ŷ�ó</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg1 VALUE="BK" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg1>����</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg2 VALUE="BP" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg2>�ŷ�ó</LABEL></TD>
                              	</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanPlcCd" ALT="����ó" SIZE="10" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanPlcCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanPlcCd.Value, 0)">
															<INPUT NAME="txtLoanPlcNm" ALT="����ó��" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 2)">
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���λ���</TD>
									<TD CLASS="TD6" NOWRAP><SELECT ID="cboConfFg" NAME="cboConfFg" ALT="���λ���" STYLE="WIDTH: 135px" tag="1XN"><OPTION VALUE="" selected></OPTION></SELECT></TD>
                              	</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����Ȳ</TD>
									<TD CLASS="TD6" NOWRAP><SELECT ID="cboApSts" NAME="cboApSts" ALT="�����Ȳ" STYLE="WIDTH: 135px" tag="1XN"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>