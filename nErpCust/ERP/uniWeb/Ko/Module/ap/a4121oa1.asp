
<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : A/P
'*  3. Program ID           : a4113ma1
'*  4. Program Name         : ä������� 
'*  5. Program Desc         : Report of A/P Detail
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001.01.29
'*  8. Modified date(Last)  : 2004.01.12
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Kim Chang Jin
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

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
	Dim strYear, strMonth, strDay,  EndDate, StartDate

<%  
	Dim dtToday
	dtToday = GetSvrDate
%>	
	Call ExtractDateFrom("<%=dtToday%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

	frm1.txtDateMid.Text = EndDate

	frm1.hOrgChangeId.value = parent.gChangeOrgId
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
		Case 0, 5
			arrParam(0) = "������ڵ� �˾�"								' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"										' Field��(0)
			arrField(1) = "BIZ_AREA_NM"										' Field��(1)
    
			arrHeader(0) = "������ڵ�"									' Header��(0)
			arrHeader(1) = "������"									' Header��(1)
			
			
		Case 3, 4
						If frm1.rdoPayBp.checked = False then
				arrParam(0) = "����ó�˾�"
				arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AP B " 
				'arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.DEAL_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AP_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
				 arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.DEAL_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") 
				IF frm1.txtDateMid.Text<>"" THEN		arrParam(1) = arrParam(1) & " AND AP_DT <= " & FilterVar(UNIConvDate(frm1.txtDateMid.Text), "''", "S") & ""
				arrParam(1) = arrParam(1) & ") TMP"
			
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "����ó"			
	
				arrField(0) = "TMP.BP_CD"	
				arrField(1) = "TMP.BP_NM"	

				arrHeader(0) = "����ó"                                     ' Header��(0)
				arrHeader(1) = "����ó��"                                   ' Header��(1)
			
			ELSE
				arrParam(0) = "����ó�˾�"
				arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AP B " 
				'arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AP_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
				 arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") 
				IF frm1.txtDateMid.Text<>"" THEN		arrParam(1) = arrParam(1) & " AND AP_DT <= " & FilterVar(UNIConvDate(frm1.txtDateMid.Text), "''", "S") & ""
				arrParam(1) = arrParam(1) & ") TMP"
			
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "����ó"			
	
				arrField(0) = "TMP.BP_CD"	
				arrField(1) = "TMP.BP_NM"	

				arrHeader(0) = "����ó"                                     ' Header��(0)
				arrHeader(1) = "����ó��"                                   ' Header��(1)

			End IF
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				frm1.txtBizAreaCd.focus
			Case 3
				frm1.txtDealBpCdFr.focus
			Case 4
				frm1.txtDealBpCdTo.focus
			Case 5
				frm1.txtBizAreaCd1.focus

			Case Else
		End Select
		Exit Function
	Else
		Select Case iWhere
			Case 0
				frm1.txtBizAreaCd.value = arrRet(0)
				frm1.txtBizAreaNm.value = arrRet(1)
				frm1.txtBizAreaCd.focus

			Case 3
				frm1.txtDealBpCdFr.value = arrRet(0)
				frm1.txtDealBpNmFr.value = arrRet(1)
				frm1.txtDealBpCdFr.focus
			Case 4
				frm1.txtDealBpCdTo.value = arrRet(0)
				frm1.txtDealBpNmTo.value = arrRet(1)
				frm1.txtDealBpCdTo.focus
			Case 5
			    frm1.txtBizAreaCd1.value = arrRet(0)
			    frm1.txtBizAreaNm1.value = arrRet(1)
				frm1.txtBizAreaCd1.focus
			Case Else
		End Select
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
    Call LoadInfTB19029()                           '��: Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
    
    Call InitVariables()                            '��: Initializes local global Variables
    Call SetDefaultVal()
    
    Call SetToolbar("1000000000000111")				'��: ��ư ���� ���� 
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
'	Event�� �浹�� �����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 



'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Event ó��	
'********************************************************************************************************* 
Function rdoDealBp_OnClick() 
	if frm1.rdoDealBp.checked = True then
		BP_Cd.innerHTML = "����ó"
	end if
End Function
Function rdoPayBp_OnClick() 
	if frm1.rdoPayBp.checked = True then
		BP_Cd.innerHTML = "����ó"
	end if
End Function
'======================================================================================================
'   Event Name : txtDateMid_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtDateMid_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateMid.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtDateMid.Focus 
    End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Function SetPrintCond(StrEbrFile, StrUrl)

	Dim	VarBizAreaCd, VarBizAreaCd1, VarDateMid, VarDealBpCdFr, VarDealBpCdTo, VarBpLabel
	Dim strSelect, strFrom ,  strWhere ,iFlag,iRs    

	Dim	strAuthCond	


	SetPrintCond = False


	If Not chkField(Document, "1") Then									'��: This function check indispensable field
	   Exit Function
	End If

	frm1.txtDealBpCdFr.value = UCase(Trim(frm1.txtDealBpCdFr.value))
	frm1.txtDealBpCdTo.value = UCase(Trim(frm1.txtDealBpCdTo.value))
		
	If frm1.txtDealBpCdFr.value <> "" And frm1.txtDealBpCdTo.value <> "" Then
		strSelect = ""
		strSelect = "  case when  " & FilterVar(frm1.txtDealBpCdFr.value, "''", "S") & " " 
		strSelect = strSelect & "  >  " & FilterVar(frm1.txtDealBpCdTo.value, "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
		strSelect = strSelect & " when  " & FilterVar(frm1.txtDealBpCdFr.value, "''", "S") & " " 
		strSelect = strSelect & "  <=  " & FilterVar(frm1.txtDealBpCdTo.value, "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  end "
		strfrom  = ""
		strwhere = ""
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , iRs) = True Then	
			iFlag = Split(iRs, Chr(11))	
			If Trim(iFlag(1)) = "N" Then
				Call DisplayMsgBox("970025", "X", frm1.txtDealBpCdFr.Alt, frm1.txtDealBpCdTo.Alt)
				frm1.txtDealBpCdFr.focus
				Exit Function			
			End If				
		Else
			Call DisplayMsgBox("970025", "X", frm1.txtDealBpCdFr.Alt, frm1.txtDealBpCdTo.Alt)
			frm1.txtDealBpCdFr.focus
			Exit Function						
		End If			
	End If


	StrEbrFile = "a4121oa1"
	VarBpLabel= "DEAL_BP_CD"
		
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
		
	VarDateMid		= UNIConvDate(frm1.txtDateMid.Text)
	VarDealBpCdFr	= UCase(Trim(frm1.txtDealBpCdFr.value))
	VarDealBpCdTo	= UCase(Trim(frm1.txtDealBpCdTo.value))
		
	If VarDealBpCdFr = "" Then VarDealBpCdFr = ""
	If VarDealBpCdTo = "" Then VarDealBpCdTo = "ZZZZZZZZZZ"
	If frm1.rdoPayBp.checked = True then VarBpLabel = "PAY_BP_CD"

	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
		
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	StrUrl = StrUrl & "BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	StrUrl = StrUrl & "|DateMid|"		& VarDateMid
	StrUrl = StrUrl & "|DealBpCdFr|"	& VarDealBpCdFr
	StrUrl = StrUrl & "|DealBpCdTo|"	& VarDealBpCdTo
	StrUrl = StrUrl & "|BpLabel|"		& VarBpLabel

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	SetPrintCond = True
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrEbrFile, StrUrl,objName
	
	On Error Resume Next                                                    '��: Protect system from crashing
	Err.Clear 
    
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	objName = AskEBDocumentName(StrEbrFile,"ebr")
	
	Call FncEBRPrint(EBAction,objName,StrUrl)
	
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	Dim StrEbrFile, StrUrl,objName
	
	On Error Resume Next                                                    '��: Protect system from crashing
	Err.Clear 
    
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	objName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPreview(objName,StrUrl)
	
	Set gActiveElement = document.activeElement
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
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

'Function FncQuery() 
'End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                        '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                        '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
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


'*******************************  5.2 Fnc�Լ������� ȣ��Ǵ� ���� Function  *******************************
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
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateMid" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=�������� id=fpDateMid></OBJECT>');</SCRIPT>&nbsp;&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=RADIO CLASS="RADIO" NAME="rdoBpLabel" ID="rdoDealBp" VALUE="S" TAG="11" Checked><LABEL FOR="rdoReport1">����ó</LABEL>&nbsp;&nbsp
													 <INPUT TYPE=RADIO CLASS="RADIO" NAME="rdoBpLabel" ID="rdoPayBp" VALUE="D" TAG="11"><LABEL FOR="rdoReport2">����ó</LABEL>
								</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" ID= BP_Cd  NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCdFr" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���۰���ó�ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDealBpCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDealBpCdFr.Value, 3)">
													   <INPUT TYPE="Text" NAME="txtDealBpNmFr" SIZE=25 tag="14X" ALT="���۰���ó��">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCdTo" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�������ó�ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDealBpCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDealBpCdTo.Value, 4)">
													   <INPUT TYPE="Text" NAME="txtDealBpNmTo" SIZE=25 tag="14X" ALT="�������ó��">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)">
													   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="������">
								</TD>
							</TR>
							<TR>
	                            <TD CLASS="TD5" NOWRAP></TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,5)">
                                                       <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="������">
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
					<TD>
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>	
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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

  