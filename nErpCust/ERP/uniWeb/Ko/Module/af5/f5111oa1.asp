
<%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module��          : ȸ��-�ڱݰ���-���� 
'*  2. Function��        : 
'*  3. Program ID        : f5111ma1.asp
'*  4. Program �̸�      : ���޾���������� 
'*  5. Program ����      : ���޾���������� 
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2003/01/08
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : Kim Chang Jin
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
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

'Dim  lgBlnFlgChgValue           ' Variable is for Dirty flag 
'Dim  lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim  IsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================
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
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtDateFr.Text = frDt
	frm1.txtDateTo.Text = toDt
		
	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	frm1.Rb_Dt1.checked = True	 '������ 
	Call Radio_Dt_Click
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","Q") %>
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================

Sub InitComboBox()
	'�������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))
End Sub

'==========================================  2.2.7 SetCheckBox()  =======================================
'	Name : SetCheckBox()
'	Description : ����������ǥ ��¹� üũ�ڽ� ���� ó��(1���� ���õǵ��� ��)
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
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	
End Function


'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 5
			arrParam(0) = "������ڵ� �˾�"								' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"										' Field��(0)
			arrField(1) = "BIZ_AREA_NM"										' Field��(1)
    
			arrHeader(0) = "������ڵ�"									' Header��(0)
			arrHeader(1) = "������"									' Header��(1)
			
		Case 2
			arrParam(0) = "���� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK"			 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�����ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BANK_CD"						' Field��(0)
			arrField(1) = "BANK_NM"						' Field��(1)
    
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)
		Case 3,4
			arrParam(0) = "������ȣ"	' �˾� ��Ī 
			arrParam(1) = "f_note"			 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = " 1=1 "							' Where Condition


			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			End If

			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			End If

			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			End If

			arrParam(5) = "������ȣ"					' �����ʵ��� �� ��Ī 

			arrField(0) = "note_no"						' Field��(0)
    
			arrHeader(0) = "������ȣ"					' Header��(0)
			
		Case Else
			Exit Function
	End Select
    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	

End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 


'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Account Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
	Select Case iWhere
		Case 0		'������ڵ� 
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
			frm1.txtBizAreaCd.focus
		Case 1		'�ŷ�ó�ڵ� 
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNM.value = arrRet(1)
			frm1.txtBpCd.focus
		Case 2		'�����ڵ� 
			frm1.txtBankCd.value = arrRet(0)
			frm1.txtBankNM.value = arrRet(1)
			frm1.txtBankCd.focus
		Case 3		'������ȣ 
			frm1.txtNoteNoFr.value = arrRet(0)
			frm1.txtNoteNoFr.focus
		Case 4		'������ȣ 
			frm1.txtNoteNoTo.value = arrRet(0)
			frm1.txtNoteNoTo.focus
		Case 5		'������ڵ� 
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
			frm1.txtBizAreaCd1.focus
		Case Else
	End select	

End Function

'------------------------------------------  EscPopUp()  ---------------------------------------------
'	Name : EscPopUp()
'	Description : Account Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function EscPopUp(ByVal iWhere)
	
	Select Case iWhere
		Case 0		'������ڵ� 
			frm1.txtBizAreaCd.focus
		Case 1		'�ŷ�ó�ڵ� 
			frm1.txtBpCd.focus
		Case 2		'�����ڵ� 
			frm1.txtBankCd.focus
		Case 3		'������ȣ 
			frm1.txtNoteNoFr.focus
		Case 4		'������ȣ 
			frm1.txtNoteNoTo.focus
		Case 5		'������ڵ� 
			frm1.txtBizAreaCd1.focus
		Case Else
	End select	

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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
    
    Call InitVariables                            '��: Initializes local global Variables
    Call SetDefaultVal
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	
    Call SetToolbar("1000000000001111")				'��: ��ư ���� ���� 
	frm1.txtBizAreaCd.focus
	
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



'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Event ó��	
'********************************************************************************************************* 

'======================================================================================================
'   Event Name : Radio_Dt
'   Event Desc : �������� ���� ���� 
'=======================================================================================================
Sub Radio_Dt_Click()
	With frm1
		If .Rb_Dt1.checked = True Then	 '������ 
			lblTitle1.innerHTML = "��������"
			lblHyphen.innerHTML = "~"
			Call ElementVisible(frm1.fpDateTime2, 1)
		ElseIf .Rb_Dt2.checked = True Then	 '������ 
			lblTitle1.innerHTML = "��������"
			lblHyphen.innerHTML = "~"
			Call ElementVisible(frm1.fpDateTime2, 1)
		Else	 '������ 
			lblTitle1.innerHTML = "��������"
			lblHyphen.innerHTML = ""
			Call ElementVisible(frm1.fpDateTime2, 0)
		End If
	End With
End Sub

'======================================================================================================
'   Event Name : txtDateFr_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.fpDateTime1.Focus
    End If
End Sub

Sub txtDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.fpDateTime2.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpDateTime2.Focus
	End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrUrl, StrEbrFile)
	Dim StrDt, StrFg, VarBizAreaCd, VarBizAreaCd1, VarDateFr, VarDateTo, VarBpCd, VarBankCd, VarNoteSts
	Dim varNoteNoFr, varNoteNoTo
	Dim	strAuthCond
	
	If frm1.Rb_Dt1.checked = True Then
		StrDt = "a"
	ElseIf frm1.Rb_Dt2.checked = True Then
		StrDt = "b"
	Else
		StrDt = "c"
	End If

	If frm1.Rb_Fg1.checked = True Then
		StrFg = "1"
	ElseIf frm1.Rb_Fg2.checked = True Then
		StrFg = "2"
	ElseIf frm1.Rb_Fg3.checked = True Then
		StrFg = "3"
	Else
		StrFg = "4"
	End If
	
	StrEbrFile = "f5111ma1" & StrDt & StrFg

	VarBizAreaCd = "%"
	VarBpCd      = "%"
	VarBankCd    = "%"
	VarNoteSts   = "%"
	varNoteNoFr  = "0"
	varNoteNoTo  = "ZZZZZZZZZZZZZZZZZZ"
	
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, Parent.gServerDateType)
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, Parent.gServerDateType)
	
	If Trim(frm1.txtBizAreaCd.value) <> "" Then 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCd.value,"","SNM")
	else
		VarBizAreaCd = ""
	end if
	
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
	
	If Trim(frm1.txtBpCd.value)		<> "" Then VarBpCd = FilterVar(Trim(frm1.txtBpCd.value), "", "SNM")
	If Trim(frm1.txtBankCd.value)	<> "" Then VarBankCd = FilterVar(Trim(frm1.txtBankCd.value), "", "SNM")
	If Trim(frm1.cboNoteSts.value)	<> "" Then VarNoteSts = Trim(frm1.cboNoteSts.value)
	If Trim(frm1.txtNoteNoFr.value)	<> "" Then varNoteNoFr = Trim(frm1.txtNoteNoFr.value)
	If Trim(frm1.txtNoteNoTo.value)	<> "" Then varNoteNoTo = Trim(frm1.txtNoteNoTo.value)
	
	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_NOTE.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_NOTE.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_NOTE.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_NOTE.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	StrUrl = StrUrl & "BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	StrUrl = StrUrl & "|DateFr|"		& VarDateFr
	StrUrl = StrUrl & "|DateTo|"		& VarDateTo
	StrUrl = StrUrl & "|BpCd|"			& VarBpCd
	StrUrl = StrUrl & "|BankCd|"		& VarBankCd
	StrUrl = StrUrl & "|NoteSts|"		& VarNoteSts
	StrUrl = StrUrl & "|NoteNoFr|"		& varNoteNoFr
	StrUrl = StrUrl & "|NoteNoTo|"		& varNoteNoTo

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile	
    Dim ObjName
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If (frm1.Rb_Dt1.checked = True) Or (frm1.Rb_Dt2.checked = True) Then
		If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'��: GL Date Compare Common Function
			Exit Function
		End if
	End If

	Call SetPrintCond(StrUrl, StrEbrFile)
	
'    On Error Resume Next                                                    '��: Protect system from crashing
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '��: Protect system from crashing
    
    Dim StrUrl, StrUrl2
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile	    
    Dim ObjName
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
	If (frm1.Rb_Dt1.checked = True) Or (frm1.Rb_Dt2.checked = True) Then
		If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'��: GL Date Compare Common Function
			Exit Function
		End if
	End If
	
	Call SetPrintCond(StrUrl, StrEbrFile)
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)	
		
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

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    FncQuery = True
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
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

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
					<TD WIDTH=100% HEIGHT=20%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>��±���</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg1 checked><LABEL FOR=Rb_Fg1>������</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg2 ><LABEL FOR=Rb_Fg2>������</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg3 ><LABEL FOR=Rb_Fg3>����</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg4 ><LABEL FOR=Rb_Fg4>������ȣ</LABEL>&nbsp;
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��¥����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Dt ID=Rb_Dt1 checked ONCLICK="vbscript:Call Radio_Dt_Click()"><LABEL FOR=Rb_WK1>������</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Dt ID=Rb_Dt2 ONCLICK="vbscript:Call Radio_Dt_Click()"><LABEL FOR=Rb_WK2>������</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Dt ID=Rb_Dt3 ONCLICK="vbscript:Call Radio_Dt_Click()"><LABEL FOR=Rb_WK3>������</LABEL>&nbsp;
									</TD>
								</TR>
							</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=*>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)">
														   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="������">&nbsp;~
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,5)">
									<INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="������">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle1">��������</SPAN></TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=�������� id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen">~</SPAN>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=�������� id=fpDateTime2></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="�ŷ�ó�ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value, 1)">
														   <INPUT TYPE="Text" NAME="txtBpNM" SIZE=25 MAXLENGTH=40  tag="14X" ALT="�ŷ�ó��">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="Text" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 2)">
														   <INPUT CLASS="Text" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=25 MAXLENGTH=30  tag="14X" ALT="�����">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="Text" TYPE=TEXT ID="txtNoteNoFr" NAME="txtNoteNoFr" SIZE=15 MAXLENGTH=18   tag="11XXXU" ALT="������ȣ"><IMG SRC="../../image/btnPopup.gif" NAME="btnNoteNoFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNoFr.Value, 3)">
														   ~ <INPUT CLASS="Text" TYPE=TEXT ID="txtNoteNoTo" NAME="txtNoteNoTo" SIZE=15 MAXLENGTH=18   tag="11XXXU" ALT="������ȣ"><IMG SRC="../../image/btnPopup.gif" NAME="btnNoteNoTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNoTo.Value, 4)"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="��������" STYLE="WIDTH: 132px" tag="11X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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
 
