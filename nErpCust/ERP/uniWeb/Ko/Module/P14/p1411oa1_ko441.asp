<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : p1411oa1.asp
'*  4. Program Name         : BOM Change History ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-03-12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Woo Guen
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = VBSCRIPT>

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
<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
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
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"							'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	
	frm1.txtStartDt.text = UniConvDateAToB(UNIDateAdd ("D", -10, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtEndDt.text   = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	
	frm1.txtBomNo.value = "1"
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
End Sub	

 '========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox()
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
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
    
	Call ggoOper.FormatField(Document, "X",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables														'��: Initializes local global variables
    'Call InitComboBox	
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 

	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtOrderNo.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
   
End Sub

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)									'��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                         '��: Protect system from crashing
    Call parent.FncPrint()
End Function

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

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	Dim var9
	Dim var10
	Dim var11
	
	Dim strUrl, strEbrFile, objName
	
    Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
    
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If
	
	If frm1.txtChildItemCd.value = "" Then
		frm1.txtChildItemNm.value = "" 
	End If	

    
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.txtBOMNo.value))
	

	var3 = UniConvDateAtoB(frm1.txtStartDt.Text,parent.gDateFormat,parent.gServerDateFormat) 
	var4 = UniConvDateAtoB(frm1.txtEndDt.Text,parent.gDateFormat,parent.gServerDateFormat)
		
	var11 = UNIDateAdd ("D", 1, var4, Parent.gServerDateFormat)

	If frm1.txtECNCd.value = "" Then
		var5 = "0"
		var6 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var5 = Trim(frm1.txtECNCd.value)
		var6 = Trim(frm1.txtECNCd.value)
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		var7 = "0"
	Else
		var7 = Trim(frm1.txtFromItemCd.value)
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var8 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var8 = Trim(frm1.txtToItemCd.value)
	End If
		
	If frm1.txtChildItemCd.value = "" Then
		var9 = "0"
		var10 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var9 = Trim(frm1.txtChildItemCd.value)
		var10 = Trim(frm1.txtChildItemCd.value)
	End If
	
	strUrl = strUrl & "plant_cd|"			& var1 
	strUrl = strUrl & "|bom_no|"			& var2 
	strUrl = strUrl & "|ecn_start_dt|"		& var3
	strUrl = strUrl & "|ecn_end_dt|"		& var4
	strUrl = strUrl & "|fr_ecn_no|"			& var5
	strUrl = strUrl & "|to_ecn_no|"			& var6 
	strUrl = strUrl & "|fr_item_cd|"		& var7
	strUrl = strUrl & "|to_item_cd|"		& var8 
	strUrl = strUrl & "|fr_child_item_cd|"	& var9 
	strUrl = strUrl & "|to_child_item_cd|"	& var10 
	strUrl = strUrl & "|ecn_finish_dt|"		& var11
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	strEbrFile = "p1411oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

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
Function BtnPreview() 
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	Dim var9
	Dim var10
	Dim var11	
	
	Dim strUrl, strEbrFile, objName
	
	Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
    
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If
	
	If frm1.txtChildItemCd.value = "" Then
		frm1.txtChildItemNm.value = "" 
	End If	

	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.txtBOMNo.value))
	
	var3 = UniConvDateAtoB(frm1.txtStartDt.Text,parent.gDateFormat,parent.gServerDateFormat) 
	var4 = UniConvDateAtoB(frm1.txtEndDt.Text,parent.gDateFormat,parent.gServerDateFormat)
		
	var11 = UNIDateAdd ("D", 1, var4, Parent.gServerDateFormat)

	If frm1.txtECNCd.value = "" Then
		var5 = "0"
		var6 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var5 = Trim(frm1.txtECNCd.value)
		var6 = Trim(frm1.txtECNCd.value)
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		var7 = "0"
	Else
		var7 = Trim(frm1.txtFromItemCd.value)
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var8 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var8 = Trim(frm1.txtToItemCd.value)
	End If
		
	If frm1.txtChildItemCd.value = "" Then
		var9 = "0"
		var10 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var9 = Trim(frm1.txtChildItemCd.value)
		var10 = Trim(frm1.txtChildItemCd.value)
	End If
	
	strUrl = strUrl & "plant_cd|"			& var1 
	strUrl = strUrl & "|bom_no|"			& var2 
	strUrl = strUrl & "|ecn_start_dt|"		& var3
	strUrl = strUrl & "|ecn_end_dt|"		& var4
	strUrl = strUrl & "|fr_ecn_no|"			& var5
	strUrl = strUrl & "|to_ecn_no|"			& var6 
	strUrl = strUrl & "|fr_item_cd|"		& var7
	strUrl = strUrl & "|to_item_cd|"		& var8 
	strUrl = strUrl & "|fr_child_item_cd|"	& var9 
	strUrl = strUrl & "|to_child_item_cd|"	& var10	 
	strUrl = strUrl & "|ecn_finish_dt|"		& var11

'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	strEbrFile = "p1411oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRPreview(objName, strUrl)
'----------------------------------------------------------------	

	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================

'------------------------------------------  OpenPlantCd()  ----------------------------------------------
'	Name : OpenPlantCd()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtPlantCd.Focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtBomNo.Focus
	
End Function

'------------------------------------------  OpenECNCd()  ----------------------------------------------
'	Name : OpenECNCd()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenECNCd()

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtECNCD.value)	' ECNNo
	arrParam(1) = ""							' ReasonCd
	arrParam(2) = ""							' Status
	arrParam(3) = ""							' EBomFlg
	arrParam(4) = ""							' MBomFlg

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetECNCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtECNCD.Focus
	
End Function

'------------------------------------------  OpenFromItemCd()  -------------------------------------------------
'	Name : OpenFromItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenFromItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtfromItemCd.Value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetFromItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtfromItemCd.Focus

End Function

'------------------------------------------  OpenToItemCd()  -------------------------------------------------
'	Name : OpenToItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenToItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtToItemCd.value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetToItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtToItemCd.Focus

End Function

'------------------------------------------  OpenChildItemCd()  -------------------------------------------------
'	Name : OpenChildItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenChildItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtChildItemCd.value		' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetChildItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtChildItemCd.Focus

End Function

'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetPlantCd()  -----------------------------------------------
'	Name : SetPlantCd()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1) 
End Function

'------------------------------------------  SetBomNo()  -------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetECNCd()  ------------------------------------------------
'	Name : SetECNCd()
'	Description : ECN Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetECNCd(byval arrRet)
	frm1.txtECNCd.Value    = arrRet(0)	
	frm1.txtECNDesc.Value    = arrRet(1) 
End Function

'------------------------------------------  SetFromItemCd()  --------------------------------------------
'	Name : SetFromItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetFromItemCd(ByVal arrRet)
	frm1.txtFromItemCd.value = arrRet(0)
	frm1.txtFromItemNm.value = arrRet(1) 
End Function

'------------------------------------------  SetToItemCd()  --------------------------------------------
'	Name : SetToItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetToItemCd(ByVal arrRet)
	frm1.txtToItemCd.value = arrRet(0)
	frm1.txtToItemNm.value = arrRet(1)  
End Function

'------------------------------------------  SetChildItemCd()  --------------------------------------------
'	Name : SetChildItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetChildItemCd(ByVal arrRet)
	frm1.txtChildItemCd.value = arrRet(0)
	frm1.txtChildItemNm.value = arrRet(1)  
End Function

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtEndDt.Focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>BOM�̷����</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="X4" ALT="�����"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>BOM Type</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBOMNo" SIZE=6 MAXLENGTH=3 tag="X2XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBOMNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBOMNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1411oa1_I202342957_txtStartDt.js'></script>								
										&nbsp;~&nbsp; 
										<script language =javascript src='./js/p1411oa1_I822757589_txtEndDt.js'></script>								
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>���躯���ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtECNCd" SIZE=20 MAXLENGTH=18 tag="X1XXXU" ALT="���躯���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnECNCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtECNDesc" SIZE=40 tag="X4" ALT="���躯�泻��"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=20 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=40 tag="X4" ALT="ǰ���">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=20 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=40 tag="X4" ALT="ǰ���"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>��ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtChildItemCd" SIZE=20 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChildItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenChildItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemNm" SIZE=40 tag="X4" ALT="��ǰ���"></TD>
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
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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

