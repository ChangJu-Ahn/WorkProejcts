
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4242oa1
'*  4. Program Name         : ���������ڿ���곻����� 
'*  5. Program Desc         : Report of Loan Repay Plan
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003.05.19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ahn, do hyun
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">				</SCRIPT>

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


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop
Dim varDateFrist, varDateToto



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

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    '---- Coding part--------------------------------------------------------------------    
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	Dim FirstDt, NextDtTo
	DIm strYear, strMonth, strDay

	Dim FirstDate, LastDate
 
    FirstDate	= UNIGetFirstDay("<%=GetSvrDate%>",	Parent.gServerDateFormat)				'Get Server DB Date
	LastDate	= UNIDateAdd("M",6,FirstDate,Parent.gServerDateFormat)
	LastDate	= UNIDateAdd("D",-1,LastDate,Parent.gServerDateFormat)


	Call ggoOper.FormatDate(frm1.txtDateFr, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtDateTo, Parent.gDateFormat, 2)
	
	Call ExtractDateFrom(FirstDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)

	Call InitComboBox

	frm1.txtDateFr.Year = strYear
	frm1.txtDateFr.Month = strMonth
	frm1.txtDateFr.Day = strDay
	
	Call ExtractDateFrom(LastDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtDateTo.Year = strYear
	frm1.txtDateTo.Month = strMonth
	frm1.txtDateTo.Day = strDay

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>
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

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
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
		Select Case iWhere
			
			Case 1		'���Կ뵵 
				frm1.txtLoanType.focus
				
			Case 2
				frm1.txtDocCur.focus
				
			Case 3
				frm1.txtBizAreaCd.focus
				
			Case 4
				frm1.txtBizAreaCd1.focus
		End Select
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
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

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
		frm1.txtDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateTo.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtLoanType_Onchange()
'   Event Desc : �����ڵ带 �����Է��Ұ�쿡 �����ڵ���� �������ش�.
'========================================================================================================
sub txtLoanType_Onchange()
	Dim strCd

	strCd = frm1.txtLoanType.value
	Call CommonQueryRs("A.minor_NM","B_MINOR A","A.major_cd=" & FilterVar("f1000", "''", "S") & "  AND A.minor_cd = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtLoanTypeNm.value = ""
	else
		frm1.txtLoanTypeNm.value = Trim(Replace(lgF0,Chr(11),""))
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

	Dim i
	Dim VarLoanType, VarDocCur
	Dim varDateFr(6), varDateTo(5)
	Dim strYear, strMonth, strDay
	Dim VarBizAreaCd, VarBizAreaCd1
	
	Dim	strAuthCond
	
	StrEbrFile = "f4242oa1"

	VarLoanType = "%"
	VarDocCur = "%"

	If Trim(frm1.txtLoanType.value) <> "" Then VarLoanType = FilterVar(Trim(frm1.txtLoanType.value)	, "", "SNM")
	If Trim(frm1.txtDocCur.value) <> "" Then VarDocCur = FilterVar(Trim(frm1.txtDocCur.value)	, "", "SNM")

	varDateFr(0) = varDateFrist
	For i = 1 to 6
		varDateFr(i) = UNIDateAdd("M",i,varDateFr(0),Parent.gServerDateFormat)
		varDateTo(i-1) = UNIDateAdd("D",-1,varDateFr(i),Parent.gServerDateFormat)
		If varDateToto < varDateTo(i-1) Then
			varDateFr(i) = varDateFr(i-1)
			varDateTo(i-1) = varDateToto
		End If
	Next
	
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

	
	condvar = "LoanType|" & VarLoanType
	condvar = condvar & "|DocCur|"			& VarDocCur
	condvar = condvar & "|DateFr1|"			& varDateFr(0)
	condvar = condvar & "|DateTo1|"			& varDateTo(0)
	condvar = condvar & "|DateFr2|"			& varDateFr(1)
	condvar = condvar & "|DateTo2|"			& varDateTo(1)
	condvar = condvar & "|DateFr3|"			& varDateFr(2)
	condvar = condvar & "|DateTo3|"			& varDateTo(2)
	condvar = condvar & "|DateFr4|"			& varDateFr(3)
	condvar = condvar & "|DateTo4|"			& varDateTo(3)
	condvar = condvar & "|DateFr5|"			& varDateFr(4)
	condvar = condvar & "|DateTo5|"			& varDateTo(4)
	condvar = condvar & "|DateFr6|"			& varDateFr(5)
	condvar = condvar & "|DateTo6|"			& varDateTo(5)
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
    
	varDateFrist = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateFr.Year, frm1.txtDateFr.Month, "01")
	varDateToto = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateTo.Year, frm1.txtDateTo.Month, "01")
	varDateToto = UNIGetLastDay (varDateToto,Parent.gServerDateFormat)

	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'��: GL Date Compare Common Function
			Exit Function
	End if

	
	'6�������� 
	If UNICDbl(DateDiff("M",varDateFrist,varDateToto)) > 5 Then 
		Call DisplayMsgBox("141158","x", frm1.txtDateFr.Alt,frm1.txtDateTo.Alt)					 '��: Data is changed.  Do you want to display it? 
		Exit Function
		frm1.txtDateFr.focus
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
    
    'On Error Resume Next       	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
  
	Call SetPrintCond(condvar, StrEbrFile)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	Call FncEBRPrint(EBAction,ObjName,condvar)	
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
    
    Dim condvar
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile    
    Dim ConDateFr, ConDateTo
	'On Error Resume Next                                                    '��: Protect system from crashing
   
	varDateFrist = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateFr.Year, frm1.txtDateFr.Month, "01")
	varDateToto = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateTo.Year, frm1.txtDateTo.Month, "01")
	varDateToto = UNIGetLastDay (varDateToto,Parent.gServerDateFormat)

	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'��: GL Date Compare Common Function
			Exit Function
	End if

	'6�������� 
	If UNICDbl(DateDiff("M",varDateFrist,varDateToto)) > 5 Then 
		Call DisplayMsgBox("141158","x", frm1.txtDateFr.Alt,frm1.txtDateTo.Alt)					 '��: Data is changed.  Do you want to display it? 
		Exit Function
		frm1.txtDateFr.focus
	End If
    
    If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
	Call SetPrintCond(condvar, StrEbrFile)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	Call FncEBRPreview(ObjName,condvar)
			
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
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP colspan=2 >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMM title=FPDATETIME ALT="���۱������ڿ�" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMM title=FPDATETIME ALT="�����������ڿ�" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
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
									<TD CLASS="TD5" NOWRAP>���Կ뵵</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtLoanType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���Կ뵵�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtLoanType.Value,1)">
														   <INPUT TYPE="Text" NAME="txtLoanTypeNm" SIZE=25 tag="14X" ALT="���Կ뵵��">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 2)">
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

