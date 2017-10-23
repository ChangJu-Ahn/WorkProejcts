

<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2111ma1
'*  4. Program Name         : ���������� 
'*  5. Program Desc         : Report of Budget Result
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001.01.06
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
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

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
Dim lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop


'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim strSvrDate

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


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
	strSvrDate = "<%=GetSvrDate%>"
	
	frm1.hOrgChangeId.value = parent.gChangeOrgId
'	frm1.fpDateTime1.Text = UNIDateClientFormat(strSvrDate)
'	frm1.fpDateTime2.Text = UNIDateClientFormat(strSvrDate)
	frm1.fpDateTime1.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.fpDateTime2.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.fpDateTime1, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.fpDateTime2, parent.gDateFormat, 2)

	frm1.Rb_Fg1.checked = True
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================

Sub InitComboBox()

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

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
'		Case 0, 1
'			arrParam(0) = "�μ��ڵ� �˾�"								' �˾� ��Ī 
'			arrParam(1) = "B_ACCT_DEPT" 									' TABLE ��Ī 
'			arrParam(2) = strCode											' Code Condition
'			arrParam(3) = ""												' Name Cindition
'			arrParam(4) = "ORG_CHANGE_ID = '" & parent.gChangeOrgId & "'"			' Where Condition
'			arrParam(5) = "�μ��ڵ�"									' �����ʵ��� �� ��Ī 
'
'			arrField(0) = "DEPT_CD"											' Field��(0)
'			arrField(1) = "DEPT_NM"											' Field��(1)
 '   
'			arrHeader(0) = "�μ��ڵ�"									' Header��(0)
'			arrHeader(1) = "�μ���"										' Header��(1)

		Case 2, 3
			arrParam(0) = "�����ڵ� �˾�"			' �˾� ��Ī 
			arrParam(1) = "F_BDG_ACCT"		 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�����ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BDG_CD"						' Field��(0)
			arrField(1) = "GP_ACCT_NM"					' Field��(1)
    
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)

		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
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
'		Case 0		'���ۺμ��ڵ� 
'			frm1.txtDeptCdFr.value = arrRet(0)
'			frm1.txtDeptNmFr.value = arrRet(1)
'		Case 1		'����μ��ڵ� 
'			frm1.txtDeptCdTo.value = arrRet(0)
'			frm1.txtDeptNmTo.value = arrRet(1)
		Case 2		'���ۿ����ڵ� 
			frm1.txtBdgCdFr.value = arrRet(0)
			frm1.txtBdgNmFr.value = arrRet(1)
		Case 3		'���Ό���ڵ� 
			frm1.txtBdgCdTo.value = arrRet(0)
			frm1.txtBdgNmTo.value = arrRet(1)
		Case Else
	End select	

End Function
'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup( ByVal iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim strYear,strMonth,strDay

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0)	= UniConvDateAToB(frm1.txtDymFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	arrParam(1)	= UniConvDateAToB(frm1.txtDymTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	arrParam(1)	= UNIDateAdd("M", +1, arrParam(1),parent.gServerDateFormat)
	arrParam(1)	= UNIDateAdd("D", -1, arrParam(1),parent.gServerDateFormat)	    

	arrParam(0)	=  UniConvDateAToB(arrParam(0),parent.gServerDateFormat,gDateFormat)
	arrParam(1)	=  UniConvDateAToB(arrParam(1),parent.gServerDateFormat,gDateFormat)

'	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCdFr.value
	arrParam(4) = "F"									' �������� ���� Condition  

	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(iWhere,arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept( ByVal iWhere,Byval arrRet)
	Select Case iWhere
		Case 0		'���ۺμ��ڵ� 

			'frm1.hOrgChangeId.value=arrRet(2)
			
			frm1.txtDeptCdFr.value = arrRet(0)
			frm1.txtDeptNmFr.value = arrRet(1)		
		Case 1		'���ۺμ��ڵ� 
			'frm1.hOrgChangeId.value=arrRet(2)
			
			frm1.txtDeptCdTo.value = arrRet(0)
			frm1.txtDeptNmTo.value = arrRet(1)		
	End Select 

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
	' ���� Page�� Form Element���� Clear�Ѵ�. 
		
    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call SetDefaultVal
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
    
    'Call InitSpreadSheet                          '��: Setup the Spread Sheet
    Call InitVariables                            '��: Initializes local global Variables
    
    '----------  Coding part  -------------------------------------------------------------
	'Call InitComboBox
	
	' [Main Menu ToolBar]�� �� ��ư�� [Enable/Disable] ó���ϴ� �κ� 
    Call SetToolbar("1000000000001111")				'��: ��ư ���� ���� 
    
	frm1.txtDymFr.focus

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

Sub Radio_Fg_Click()
	If frm1.Rb_Fg1.checked = True Then	'�μ��� 
		lblTitle1.innerHTML = "������"
		lblHyphen.innerHTML = "~"
		Call ElementVisible(frm1.fpDateTime2, 1)	'Visible
		Call ggoOper.FormatDate(frm1.txtDymFr, parent.gDateFormat, 2)	'��� 
	ElseIf frm1.Rb_Fg2.checked = True Then	'�����ڵ庰 
		lblTitle1.innerHTML = "������"
		lblHyphen.innerHTML = "~"
		Call ElementVisible(frm1.fpDateTime2, 1)	'Visible
		Call ggoOper.FormatDate(frm1.txtDymFr, parent.gDateFormat, 2)	'��� 
	ElseIf frm1.Rb_Fg3.checked = True Then	'�Ⱓ 
		lblTitle1.innerHTML = "����⵵"
		lblHyphen.innerHTML = ""
		Call ElementVisible(frm1.fpDateTime2, 0)	'InVisible
		Call ggoOper.FormatDate(frm1.txtDymFr, parent.gDateFormat, 3)	'�⵵ 
	End If
End Sub

'======================================================================================================
'   Event Name : 
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtDymFr_DblClick(Button)
    If Button = 1 Then
		frm1.fpDateTime1.Action = 7
		Call SetFocusToDocument("M")	
		frm1.fpDateTime1.Focus       
        
    End If
End Sub

Sub txtDymTo_DblClick(Button)
    If Button = 1 Then
		frm1.fpDateTime2.Action = 7
		Call SetFocusToDocument("M")	
		frm1.fpDateTime2.Focus               
    End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, StrUrl)
	Dim StrFg, VarDeptCdFr, VarDeptCdTo, VarBdgCdFr, VarBdgCdTo, VarDymFr, VarDymTo
	Dim strYear, strMonth, strDay
	Dim strYear1, strMonth1, strDay1

	Dim strAuthCond

	If frm1.Rb_Fg1.checked = True Then	 '�μ��� 
		StrEbrFile = "f2111ma1a"
	ElseIf frm1.Rb_Fg2.checked = True Then	 '�����ڵ庰 
		StrEbrFile = "f2111ma1b"
	ElseIf frm1.Rb_Fg3.checked = True Then	 '�Ⱓ 
		StrEbrFile = "f2111ma1c"
	End If

	VarDeptCdFr	= " "
	VarDeptCdTo	= "ZZZZZZZZZZ"
	VarBdgCdFr	= " "
	VarBdgCdTo	= "ZZZZZZZZZZZZZZZZZZ"
	
	Call ExtractDateFrom(frm1.fpDateTime1.Text,frm1.fpDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	VarDymFr = strYear & strMonth
	
	Call ExtractDateFrom(frm1.fpDateTime2.Text,frm1.fpDateTime2.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	VarDymTo = strYear1 & strMonth1

	If frm1.Rb_Fg3.checked = True Then	 '�Ⱓ�� ���, FromDate�� �⵵�� ��� 
		Call ExtractDateFrom(frm1.fpDateTime2.Text,frm1.fpDateTime2.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	    VarDymTo = strYear1 & strMonth1
				
		VarDymFr = Trim(frm1.fpDateTime1.Text)
		
	End If
	
	If Trim(frm1.txtDeptCdFr.value) <> ""	Then VarDeptCdFr = FilterVar(Trim(frm1.txtDeptCdFr.value),"","SNM")
	If Trim(frm1.txtDeptCdTo.value) <> ""	Then VarDeptCdTo = FilterVar(Trim(frm1.txtDeptCdTo.value),"","SNM")
	If Trim(frm1.txtBdgCdFr.value) <> ""	Then VarBdgCdFr = FilterVar(Trim(frm1.txtBdgCdFr.value),"","SNM")
	If Trim(frm1.txtBdgCdTo.value) <> ""	Then VarBdgCdTo = FilterVar(Trim(frm1.txtBdgCdTo.value),"","SNM")
	
	'-----------------------------------------------------------------------------------
	
	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_BDG.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_BDG.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_BDG.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_BDG.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	StrUrl = StrUrl & "DeptCdFr|"	& VarDeptCdFr
	StrUrl = StrUrl & "|DeptCdTo|"	& VarDeptCdTo
	StrUrl = StrUrl & "|BdgCdFr|"	& VarBdgCdFr
	StrUrl = StrUrl & "|BdgCdTo|"	& VarBdgCdTo
	StrUrl = StrUrl & "|DymFr|"		& VarDymFr
	StrUrl = StrUrl & "|DymTo|"		& VarDymTo
	StrUrl = StrUrl & "|DYear|"		& Trim(VarDymFr) & "__"
	
	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond
	

	'-----------------------------------------------------------------------------------
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
    On Error Resume Next
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile	
	Dim ObjName
    	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtDymFr.Text, frm1.txtDymTo.Text, frm1.txtDymFr.Alt, frm1.txtDymTo.Alt, _
				"970025", frm1.txtDymFr.UserDefinedFormat, parent.gComDateType, true) = False Then
		frm1.txtDymFr.focus											'��: GL Date Compare Common Function
		Exit Function
	End if	
	
	
	frm1.txtDeptCdFr.value = Trim(frm1.txtDeptCdFr.value)
	frm1.txtDeptCdTo.value = Trim(frm1.txtDeptCdTo.value)
	If frm1.txtDeptCdFr.value <> "" And frm1.txtDeptCdTo.value <> "" Then
		If frm1.txtDeptCdFr.value > frm1.txtDeptCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtDeptCdFr.Alt, frm1.txtDeptCdTo.Alt)
			frm1.txtDeptCdFr.focus 
			Exit Function
		End If
	End If
	
		
	frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
	frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
	If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus 
			Exit Function
		End If
	End If
		
	Call SetPrintCond(StrEbrFile, StrUrl)
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	

End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	Dim StrFg
    Dim StrUrl, StrUrl2
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile	
	Dim ObjName
        
	If frm1.Rb_Fg1.checked = True Then	 '�μ��� 
		StrFg = "a"
	ElseIf frm1.Rb_Fg2.checked = True Then	 '�����ڵ庰 
		StrFg = "b"
	ElseIf frm1.Rb_Fg3.checked = True Then	 '�Ⱓ 
		StrFg = "c"
	End If
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If StrFg <> "c" Then
		If CompareDateByFormat(frm1.txtDymFr.Text, frm1.txtDymTo.Text, frm1.txtDymFr.Alt, frm1.txtDymTo.Alt, _
					"970025", frm1.txtDymFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtDymFr.focus											'��: GL Date Compare Common Function
			Exit Function	
		End if	
	End If
	
	frm1.txtDeptCdFr.value = Trim(frm1.txtDeptCdFr.value)
	frm1.txtDeptCdTo.value = Trim(frm1.txtDeptCdTo.value)
	If frm1.txtDeptCdFr.value <> "" And frm1.txtDeptCdTo.value <> "" Then
		If frm1.txtDeptCdFr.value > frm1.txtDeptCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtDeptCdFr.Alt, frm1.txtDeptCdTo.Alt)
			frm1.txtDeptCdFr.focus 
			Exit Function
		End If
	End If
	
	frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
	frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
	If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus 
			Exit Function
		End If
	End If
	
	Call SetPrintCond(StrEbrFile, StrUrl)
	
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


'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************** 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
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
					<TD WIDTH=100%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>��±���</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg1 ONCLICK="vbscript:Call Radio_Fg_Click()"><LABEL FOR=Rb_Fg1>�μ���</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg2 ONCLICK="vbscript:Call Radio_Fg_Click()"><LABEL FOR=Rb_Fg2>�����ڵ庰</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg3 ONCLICK="vbscript:Call Radio_Fg_Click()"><LABEL FOR=Rb_Fg3>�Ⱓ</LABEL>&nbsp;
									</TD>
								</TR>
								<TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle1">������</SPAN></TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDymFr" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=���ۿ����� id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen">~</SPAN>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDymTo" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=���Ό���� id=fpDateTime2></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�μ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCdFr" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���ۺμ��ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup(0)">&nbsp;<INPUT TYPE="Text" NAME="txtDeptNmFr" SIZE=25 tag="14X" ALT="���ۺμ���">&nbsp;~
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCdTo" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="����μ��ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup(1)">&nbsp;<INPUT TYPE="Text" NAME="txtDeptNmTo" SIZE=25 tag="14X" ALT="����μ���">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBdgCdFr" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT ="���ۿ����ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBdgCdFr.Value, 2)">&nbsp;<INPUT TYPE="Text" NAME="txtBdgNmFr" SIZE=25 tag="14X" ALT="���ۿ����ڵ��">&nbsp;~
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBdgCdTo" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT ="���Ό���ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBdgCdTo.Value, 3)">&nbsp;<INPUT TYPE="Text" NAME="txtBdgNmTo" SIZE=25 tag="14X" ALT="���Ό���ڵ��">
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TabIndex="-1">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TabIndex="-1">	
</FORM>
</BODY>
</HTML>

