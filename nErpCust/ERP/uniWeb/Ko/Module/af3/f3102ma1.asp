
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3102ma1
'*  4. Program Name         : ��������ȸ 
'*  5. Program Desc         : Query of Deposit
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.11
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  �����ڵ�, ���¹�ȣ ���� Check ���� �߰� 
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->				
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance
<!--#Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim IsOpenPop                                               '��: Popup status                           
Dim lgIsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

<%

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
'  Call GetAdoFiledInf("F3102MA1","S", "A")						'��: spread sheet �ʵ����� query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
%>

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f3102mb1.asp"						'��: �����Ͻ� ���� ASP�� 

'Dim lsPoNo                                                '��: Jump�� Cookie�� ���� Grid value
Const C_MaxKey          = 3                                    '�١١١�: Max key value
Const C_GL_NO			= 9
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
    lgPageNo         = 0
    lgSortKey        = 1

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
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
		

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "F","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","QA") %>

End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
        Call SetZAdoSpreadSheet("f3102ma1","S","A","V20030410",Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
        Call SetSpreadLock("A") 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval iOpt)
	  If iOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		End With
      End If
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()
<%	
	Dim arrData
	
'	arrData = InitCombo("F3011", "frm1.cboDpstFg")
'	arrData = InitCombo("F3014", "frm1.cboTransSts")
%>
End Sub
 
<%
Function InitCombo(ByVal strMajorCd, ByVal objCombo)

    Dim pB1a028
    Dim intMaxRow
    Dim intLoopCnt
    Dim strCodeList
    Dim strNameList
        
    Err.Clear                                                               '��: Clear error no
	On Error Resume Next

	Set pB1a028 = Server.CreateObject("B1a028.B1a028ListMinorCode")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pB1a028 = Nothing												'��: ComProxy Unload
		Call MessageBox(Err.description, I_INSCRIPT)						'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	pB1a028.ImportBMajorMajorCd = strMajorCd									'��: Major Code
    pB1a028.ServerLocation = ggServerIP
    pB1a028.ComCfg = gConnectionString
    pB1a028.Execute															'��:
    
    '-----------------------
    'Com action result check area(DB,internal)
    '-----------------------
    If Not (pB1a028.OperationStatusMessage = Parent.MSG_OK_STR) Then
		Call MessageBox(pB1a028.OperationStatusMessage, I_INSCRIPT)         '��: you must release this line if you change msg into code
		Set pB1a028 = Nothing												'��: ComProxy Unload
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
    End If

	intMaxRow = pB1a028.ExportGroupCount
	
	For intLoopCnt = 1 To intMaxRow
%>
		Call SetCombo(<%=objCombo%>, "<%=pB1a028.ExportItemBMinorMinorCd(intLoopCnt)%>", "<%=pB1a028.ExportItemBMinorMinorNm(intLoopCnt)%>")		'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
<%
		strCodeList = strCodeList & vbtab & pB1a028.ExportItemBMinorMinorCd(intLoopCnt)
		strNameList = strNameList & vbtab & pB1a028.ExportItemBMinorMinorNm(intLoopCnt)
	Next
	
	InitCombo = Array(strCodeList, strNameList)
		
	Set pB1a028 = Nothing

End Function
%>

 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0
			arrParam(0) = "���� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BANK A(NOLOCK), F_DPST B(NOLOCK)" 			' TABLE ��Ī 
			arrParam(2) = Trim(strCode)					' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD"		' Where Condition

'			' ���Ѱ��� �߰� 
'			If lgAuthBizAreaCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
'			End If
'			If lgInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
'			End If
'			If lgSubInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")
'			End If
'			If lgAuthUsrID <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
'			End If

			arrParam(5) = "�����ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "A.BANK_CD"					' Field��(0)
			arrField(1) = "A.BANK_NM"					' Field��(1)
			arrField(2) = "B.BANK_ACCT_NO"				' Field��(2)
    
			arrHeader(0) = "�����ڵ�"				' Header��(0)
			arrHeader(1) = "�����"					' Header��(1)
			arrHeader(2) = "���¹�ȣ"				' Header��(1)
			
		Case 1
			arrParam(0) = "���¹�ȣ �˾�"			' �˾� ��Ī 
			arrParam(1) = "B_BANK A(NOLOCK), F_DPST B(NOLOCK)" 			' TABLE ��Ī 
			arrParam(2) = Trim(strCode)					' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD"		' Where Condition

'			' ���Ѱ��� �߰� 
'			If lgAuthBizAreaCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
'			End If
'			If lgInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
'			End If
'			If lgSubInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")
'			End If
'			If lgAuthUsrID <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
'			End If

			arrParam(5) = "���¹�ȣ"				' �����ʵ��� �� ��Ī 
	
			arrField(0) = "B.BANK_ACCT_NO"				' Field��(0)
			arrField(1) = "B.BANK_CD"					' Field��(1)
			arrField(2) = "A.BANK_NM"					' Field��(2)
    
			arrHeader(0) = "���¹�ȣ"				' Header��(0)
			arrHeader(1) = "�����ڵ�"				' Header��(1)
			arrHeader(2) = "�����"					' Header��(1)

		Case 2
			arrParam(0) = "��ȭ�ڵ� �˾�"				' �˾� ��Ī 
			arrParam(1) = " B_CURRENCY A"					' TABLE ��Ī 
			arrParam(2) = Trim(strCode)						' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "��ȭ�ڵ�"					' �����ʵ��� �� ��Ī 
	
			arrField(0) = "A.CURRENCY"						' Field��(0)
			arrField(1) = "A.CURRENCY_DESC"					' Field��(1)
    
			arrHeader(0) = "��ȭ�ڵ�"					' Header��(0)
			arrHeader(1) = "��ȭ��"						' Header��(1)
 		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True
	
	If iWhere = 2 Then 
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If 
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Select Case iWhere
		Case 0
			frm1.txtBankCd.value = arrRet(0)
			frm1.txtBankNm.value = arrRet(1)
			frm1.txtBankAcctNo.value = arrRet(2)
			frm1.txtBankAcctNo.focus
		Case 1
			frm1.txtBankAcctNo.value = arrRet(0)
			frm1.txtBankCd.value = arrRet(1)
			frm1.txtBankNm.value = arrRet(2)
			frm1.txtBankAcctNo.focus
		Case 2
			frm1.txtDocCur.value = arrRet(0)
			frm1.txtDocCur.focus
		End Select
	End If	
	
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
	arrParam(2) = Trim(strCode)					' Code Condition
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
Function SetReturnVal(byval arrRet, iWhere)
	Select Case iWhere
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value	= arrRet(0)
			frm1.txtBizAreaNm1.Value	= arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'============================================================
'ȸ����ǥ �˾� 
'============================================================
Function OpenPopupGL()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	
	Dim arrField
	Dim ii

	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_GL_NO
			arrParam(0) = Trim(.Text)	'ȸ����ǥ��ȣ 
			arrParam(1) = ""			'Reference��ȣ 
		Else
			Call DisplayMsgBox("900025","X","X","X")
			Exit Function
		End If
	End With
	
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
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
    Call LoadInfTB19029														'��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	'Call InitComboBox()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Call FncSetToolBar("New")
	frm1.txtBankCd.focus 
	
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
	
'--------------- ������ coding part(�������,End)------------------------------------------------------
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

Sub txtDateFr_DblClick(Button)
	if Button = 1 then
		frm1.txtDateFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateFr.Focus
	End if
End Sub

Sub txtDateTo_DblClick(Button)
	if Button = 1 then
		frm1.txtDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateTo.Focus
	End if
End Sub

Sub txtDateFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtDateTo.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtDateTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtDateFr.Focus
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
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
'	lsPoNo=frm1.vspdData.Text
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
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
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData										'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
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
	
	If Trim(frm1.txtBizAreaCd.value) = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If Trim(frm1.txtBizAreaCd1.value) = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    
    If DbQuery	= False Then
       Exit Function
    End If

    FncQuery = True		
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
	Call parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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
		strVal = BIZ_PGM_ID & "?txtBankCd="		& Trim(.txtBankCd.value)
		strVal = strVal & "&txtBankAcctNo="		& Trim(.txtBankAcctNo.value)
		strVal = strVal & "&txtDateFr="			& Trim(.txtDateFr.Text)
		strVal = strVal & "&txtDateTo="			& Trim(.txtDateTo.Text)
		strVal = strVal & "&txtDocCur="			& Trim(.txtDocCur.value)
		strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
		strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
		strVal = strVal & "&txtBankCd_Alt="		& Trim(.txtBankCd.Alt)
		strVal = strVal & "&txtBankAcctNo_Alt="	& Trim(.txtBankAcctNo.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt1="	& Trim(.txtBizAreaCd1.Alt)
'--------------- ������ coding part(�������,End)------------------------------------------------

        strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgPageNo="			& lgPageNo                      '��: Next key tag		
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

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
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call FncSetToolBar("Query")
	Call CurFormatNumericOCX()
	
	Set gActiveElement = document.activeElement 
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	Dim IntRetCD
	Dim strGBCurrency
	Dim strBankCd
	Dim strBankAcctNo
	
	With frm1
        strBankCd =  frm1.txtBankCd.value
        strBankAcctNo =  frm1.txtBankAcctNo.value
         
		If Trim(.txtDocCur.value) = "" Then

		  intRetCD = CommonQueryRs("top 1 doc_cur"," f_dpst_item "," bank_cd =  " & FilterVar(strBankCd , "''", "S") & " and bank_acct_no =  " & FilterVar(strBankAcctNo , "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
          If intRetCD = True Then	
		  	strGBCurrency = Trim(Replace(lgF0,Chr(11),""))
		  Else
		    strGBCurrency = 	parent.gCurrency
		  End If					 

			ggoOper.FormatFieldByObjectOfCur .txtPreAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtRcptAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		Else
			ggoOper.FormatFieldByObjectOfCur .txtPreAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtRcptAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		End If
	End With

End Sub

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
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>
						<TABLE CELLSPACING=0 CELLPADDING=0 align=right>
							<TR>
								<td>
									<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</a>
								</td>
						    </TR>
						</TABLE>
					</TD>
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
							<TABLE WIDTH=100% CELLSPACING="0" CELLPADDING="0">
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="12XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 0)">
										<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=25 MAXLENGTH=30  tag="24X" ALT="�����">
									</TD>
									<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtBankAcctNo" NAME="txtBankAcctNo" SIZE=20 MAXLENGTH=30 tag="12XXXU" ALT="���¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.Value, 1)">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ȸ�Ⱓ</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��ȸ������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��ȸ������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>��ȭ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="��ȭ" SIZE = "10" MAXLENGTH="3"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 2)"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenBizAreaCd(frm1.txtBizAreaCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenBizAreaCd(frm1.txtBizAreaCd1.Value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100%>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS="TD5" NOWRAP>�̿��ݾ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPreAmt" title=FPDOUBLESINGLE ALT="�̿��ݾ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
											
											<TD CLASS="TD5" NOWRAP>�̿��ݾ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPreLocAmt" title=FPDOUBLESINGLE ALT="�̿��ݾ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>�Ա��հ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtRcptAmt" title=FPDOUBLESINGLE ALT="�Ա��հ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>�Ա��հ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtRcptLocAmt" title=FPDOUBLESINGLE ALT="�Ա��հ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>����հ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPaymAmt" title=FPDOUBLESINGLE ALT="����հ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>����հ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPaymLocAmt" title=FPDOUBLESINGLE ALT="����հ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>�ܾ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtBalAmt" title=FPDOUBLESINGLE ALT="�ܾ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>�ܾ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtBalLocAmt" title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
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
		<TD <%=HGIEHT_TYPE_01%>></td>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

