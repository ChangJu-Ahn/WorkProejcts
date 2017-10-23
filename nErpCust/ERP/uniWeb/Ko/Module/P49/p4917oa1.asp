<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : PRODUCTION
'*  2. Function Name        :
'*  3. Program ID           : p4917oa1
'*  4. Program Name         : �۾�������� 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-01-18
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Yoon, Jeong Woo
'* 10. Modifier (Last)      :
'* 11. Comment              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--=======================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--=======================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--��:Print Program needs this vbs file-->

<SCRIPT LANGUAGE=VBSCRIPT>

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
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim  lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim  lgIntFlgMode               ' Variable is for Operation Status
Dim  lgIntGrpCount              ' initializes Group View Size
Dim  IsOpenPop
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim  LocSvrDate
Dim  BaseMm
Dim  ToDate

LocSvrDate = "<%=GetSvrDate%>"
BaseMm = UniConvDateAtoB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormatYYYYMM)     	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
'BaseMm = UNIConvDateAToB(UniDateAdd("m", 0, "<%=BaseDate%>",parent.gServerDateFormat),parent.gServerDateFormat,parent.gDateFormat)     	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
'ToDate = UNIDateAdd("D",7,BaseMm,parent.gDateFormat)							    '��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

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
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtBaseMm.Text	= cstr(BaseMm)
'	frm1.txtBaseMm.Text = BaseMm
End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()
    Dim  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P1017' AND MINOR_CD IN ('RL','ST') ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
End Sub

'=======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

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

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

    Call InitComboBox
                                       '��: Lock  Suitable  Field
	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables

    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
	Call ggoOper.FormatDate(frm1.txtBaseMm, Parent.gDateFormat, 2)

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Ucase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
'		frm1.txtOrderNo.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó�� 
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó�� 
'*********************************************************************************************************
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
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncSave()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncNew()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncDelete()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncInsertRow()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncDeleteRow()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncCopy()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncCancel()
	On Error Resume Next                                                    '��: Protect system from crashing
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
	Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	Dim  strEbrFile
    Dim  objName

	Dim  var1
	Dim  var2
	Dim  var3
	Dim  var4
	Dim  var5
	Dim  var6
	Dim  var7
	Dim  var8
	Dim  var9
	Dim  var11
	Dim  var12

	Dim  strUrl
	Dim  arrParam, arrField, arrHeader

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If

	If frm1.txtFromWcCd.value = "" Then
		frm1.txtFromWcNm.value = ""
	End If

	If frm1.txtToWcCd.value = "" Then
		frm1.txtToWcNm.value = ""
	End If

	Call BtnDisabled(1)

	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)
       Exit Function
    End If

	var1 = UCase(Trim(frm1.txtPlantCd.value))

	If frm1.txtFromWcCd.value = "" Then
		var3 = "0"
	Else
		var3 = Trim(frm1.txtFromWcCd.value)
	End If

	If frm1.txtToWcCd.value = "" Then
		var4 = "zzzzzzz"
	Else
		var4 = Trim(frm1.txtToWcCd.value)
	End If

	var7 = UniConvDateAtoB(frm1.txtBaseMm.Text,parent.gDateFormat,parent.gDateFormatYYYYMM)
'	var8 = UniConvDateAtoB(frm1.txtEndDt.Text,parent.gDateFormat,parent.gServerDateFormat)

	strUrl = strUrl & "plant_cd|" & var1
'	strUrl = strUrl & "|fr_prod_order_no|" & var2
'	strUrl = strUrl & "|to_prod_order_no|" & var11
	strUrl = strUrl & "|from_wc_cd|" & var3
	strUrl = strUrl & "|to_wc_cd|" & var4
'	strUrl = strUrl & "|from_item_cd|" & var5
'	strUrl = strUrl & "|to_item_cd|" & var6
	strUrl = strUrl & "|start_date|" & var7
'	strUrl = strUrl & "|end_date|" & var8
'	strUrl = strUrl & "|slip_print_flg|" & var9
'	strUrl = strUrl & "|status|" & var12

	strEbrFile = "p4917oa1"
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
    Dim  strEbrFile
    Dim  objName

	Dim  var1
	Dim  var2
	Dim  var3
	Dim  var4
	Dim  var5

	Dim  strUrl
	Dim  arrParam, arrField, arrHeader

	Dim strYear1,strMonth1,strDay1
	Dim strDate1

	Call ExtractDateFrom(frm1.txtBaseMm.Text,frm1.txtBaseMm.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")

	Call BtnDisabled(1)

	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)
       Exit Function
    End If

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If

	If frm1.txtFromWcCd.value = "" Then
		frm1.txtFromWcNm.value = ""
	End If

	If frm1.txtToWcCd.value = "" Then
		frm1.txtToWcNm.value = ""
	End If

	var1 = Trim(frm1.txtPlantCd.value)

	var2 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")
	var3 = UniDateAdd("M",1, var2, parent.gServerDateFormat)

	If frm1.txtFromWcCd.value = "" Then
		var4 = "0"
	Else
		var4 = Trim(frm1.txtFromWcCd.value)
	End If

	If frm1.txtToWcCd.value = "" Then
		var5 = "zzzzzzz"
	Else
		var5 = Trim(frm1.txtToWcCd.value)
	End If

	strUrl = strUrl & "plant_cd|" & var1
	strUrl = strUrl & "|start_date|" & var2
	strUrl = strUrl & "|end_date|" & var3
	strUrl = strUrl & "|from_wc_cd|" & var4
	strUrl = strUrl & "|to_wc_cd|" & var5

	strEbrFile = "p4917oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRPreview(objName, strUrl)

	Call BtnDisabled(0)

	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement

End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================

Function FncExit()
	Dim  IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

Function OpenPlantCd()

	Dim  arrRet
	Dim  arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"		' �˾� ��Ī 
	arrParam(1) = "B_PLANT"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "����"			' TextBox ��Ī 

    arrField(0) = "PLANT_CD"			' Field��(0)
    arrField(1) = "PLANT_NM"			' Field��(1)

    arrHeader(0) = "����"			' Header��(0)
    arrHeader(1) = "�����"			' Header��(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus

End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'---------------------------------------------------------------------------------------------------------
Function OpenFromWcCd()
	Dim  arrRet
	Dim  arrParam(6), arrField(6), arrHeader(6)

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"
	arrParam(1) = "P_WORK_CENTER"
	arrParam(2) = frm1.txtFromWcCd.value
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S")

	arrParam(5) = "�۾���"

    arrField(0) = "WC_CD"
    arrField(1) = "WC_NM"
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"
    'arrField(3) = "VALID_FROM_DT"
    'arrField(4) = "VALID_TO_DT"


    arrHeader(0) = "�۾���"
    arrHeader(1) = "�۾����"
    arrHeader(2) = "�۾���Ÿ��"
    arrHeader(3) = "�۾�������"
    'arrHeader(3) = "������"
    'arrHeader(4) = "������"


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetFromWcCd(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtFromWcCd.focus

End Function
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'---------------------------------------------------------------------------------------------------------
Function OpenToWcCd()
	Dim  arrRet
	Dim  arrParam(6), arrField(6), arrHeader(6)

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"
	arrParam(1) = "P_WORK_CENTER"
	arrParam(2) = frm1.txtToWcCd.value
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S")

	arrParam(5) = "�۾���"

    arrField(0) = "WC_CD"
    arrField(1) = "WC_NM"
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"
    'arrField(3) = "VALID_FROM_DT"
    'arrField(4) = "VALID_TO_DT"


    arrHeader(0) = "�۾���"
    arrHeader(1) = "�۾����"
    arrHeader(2) = "�۾���Ÿ��"
    arrHeader(3) = "�۾�������"
    'arrHeader(3) = "������"
    'arrHeader(4) = "������"


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetToWcCd(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtToWcCd.focus

End Function

Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)
End Function

Function SetFromWcCd(ByVal arrRet)
	frm1.txtFromWcCd.value = arrRet(0)
	frm1.txtFromWcNm.value = arrRet(1)
End Function

Function SetToWcCd(ByVal arrRet)
	frm1.txtToWcCd.value = arrRet(0)
	frm1.txtToWcNm.value = arrRet(1)
End Function

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtBaseMm_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseMm.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseMm.Focus
    End If
End Sub

'-------------------------------------------------------------------------------
' Function Name : txtPlantCd_OnChange()
' Function Desc :
'-------------------------------------------------------------------------------
Sub txtPlantCd_OnChange()
	Dim strPlant
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	If Trim(frm1.txtPlantCd.value) <> "" Then
		Call CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		strPlant = Replace(lgF0, Chr(11), "")

		If Trim(strPlant) = "" Then
			frm1.txtPlantNm.Value = ""
		Else
			frm1.txtPlantNm.Value = Trim(strPlant)
		End If
	End If
End Sub

Sub txtFromWcCd_OnChange()
	Dim strFrWc
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Sub
	End If

	If Trim(frm1.txtFromWcCd.value) <> "" Then
		Call CommonQueryRs("WC_NM", "P_WORK_CENTER", "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " AND WC_CD = " & FilterVar(frm1.txtFromWcCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		strFrWc = Replace(lgF0, Chr(11), "")

		If Trim(strFrWc) = "" Then
			frm1.txtFromWcNm.Value = ""
		Else
			frm1.txtFromWcNm.Value = Trim(strFrWc)
		End If
	End If
End Sub


Sub txtToWcCd_OnChange()
	Dim strToWc
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Sub
	End If

	If Trim(frm1.txtToWcCd.value) <> "" Then
		Call CommonQueryRs("WC_NM", "P_WORK_CENTER", "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " AND WC_CD = " & FilterVar(frm1.txtToWcCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		strToWc = Replace(lgF0, Chr(11), "")

		If Trim(strToWc) = "" Then
			frm1.txtToWcNm.Value = ""
		Else
			frm1.txtToWcNm.Value = Trim(strToWc)
		End If
	End If
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5 colspan="2">&nbsp;<% ' ���� ���� %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�۾��������</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" colspan="2">
			<TABLE CLASS="BasicTB" CELLSPACING=0 >
	    		<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="x2xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�����">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4917oa1_fpDateTime1_txtBaseMm.js'></script>
										<!--OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConFromDt" CLASS=FPDTYYYYMM tag="12X1" Alt="��ȹ�������" Title="FPDATETIME"></OBJECT-->
									</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>�۾���</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromWcCd" SIZE=10 MAXLENGTH=7 tag="x1xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromWcNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�۾����">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToWcCd" SIZE=10 MAXLENGTH=7 tag="x1xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToWcNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�۾����">&nbsp;
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>