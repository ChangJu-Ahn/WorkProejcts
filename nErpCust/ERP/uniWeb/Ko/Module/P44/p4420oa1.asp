
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4420oa1.asp
'*  4. Program Name         : Month/Day Prod.Report   
'*  5. Program Desc         : Month/Day Prod.Report
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/04/12
'*  8. Modified date(Last)  : 2002/04/20
'*  9. Modifier (First)     : Kang HK
'* 10. Modifier (Last)      : Kang HK
'* 11. Comment              : 
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************** -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!-- '==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--��:Print Program needs this vbs file-->
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit														'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global ���� ����  ======================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'==========================================================================================================
'==========================================  1.2.2 Global ���� ����  ======================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'==========================================================================================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

Dim StartDate
Dim EndDate
Dim strYear, strMonth, strDay
Dim strDate

Call ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)	    	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")	        '��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 
strDate = UniConvDateAToB("<%=GetsvrDate%>",parent.gServerDateFormat,parent.gClientDateFormat) 	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

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
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtStartDt.Text = StartDate
	frm1.txtEndDt.Text = EndDate
	'frm1.rdoFlag1.checked = True
	
	frm1.txtMonthDt.Text = UNIFormatMonth(strDate)    
    
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
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
'**********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "x",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtMonthDt, parent.gDateFormat, "2")
                                       
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtFromItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
	End If
	Set gActiveElement = document.activeElement

    '�����ʱ�ȭ��Ŵ - rdoclick���ִ��� 
    frm1.txtStartDt.value = ""
    frm1.txtEndDt.value = ""
	'Call ggoOper.SetReqAttr(frm1.txtMonthDt,"N")    'cboProdMgr   'HTML���� �̹� tag="x2"�� �����Ƿ� ���⼱ �ʿ����.
	Call ggoOper.SetReqAttr(frm1.txtStartDt,"Q")      'txtPurOrg
	Call ggoOper.SetReqAttr(frm1.txtEndDt,"Q")      'txtPurOrg
        
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

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

'******************************  3.2.1 Object Tag ó��  **************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMonthDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtMonthDt.Action = 7 
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtMonthDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtMonthDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function ********************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'=========================================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'=========================================================================================================
Function BtnPrint()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	Dim var7
	dim strUrl
	dim arrParam, arrField, arrHeader
	Dim strEbrFile
    Dim objName
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then							'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))         'plantcd�� var1
	
	If frm1.rdoPrintType1.checked = True Then         '�������ý� ������ ebr  var2,var3,var6
		
		var2 = UniConvYYYYMMDDToDate(parent.gServerDateFormat,frm1.txtMonthDt.Year,frm1.txtMonthDt.Month,"01")      '�������� = �״޸� �����ǹǷ� ������ 
		var3 = "p4420oa1"
		var6 = UniDateAdd("M",1,UniConvDate(var2),parent.gServerDateFormat) 
		var7 = frm1.txtMonthDt.text
		
	Elseif frm1.rdoPrintType2.checked = True Then       '�Ϻ����ý� ������ ebr  var2,var3,var6
		
		var2 = UniConvDate(frm1.txtStartDt.text)
		var3 = "p4420oa2"
		var6 = UniConvDate(frm1.txtEndDt.text)
		var7 = ""
	End If
	
	
    If frm1.txtFromItemCd.value = "" Then            'ǰ�� txtFromItemCd   var4
		var4 = "0"
	Else
		var4 = Trim(frm1.txtFromItemCd.value)
	End If
	
	If frm1.txtToItemCd.value = "" Then              'ǰ�� txtToItemCd   var5
		var5 = "zzzzzzzzzzzzzzzzzz"
	Else
		var5 = Trim(frm1.txtToItemCd.value)
	End If
	
	strUrl = strUrl & "plant_cd|" & var1 
	strUrl = strUrl & "|Start_Dt|" & var2 
	strUrl = strUrl & "|End_Dt|" & var6 
	strUrl = strUrl & "|fr_item_cd|" & var4 
	strUrl = strUrl & "|to_item_cd|" & var5 
	strUrl = strUrl & "|Month_Text|" & var7 
	
	strEbrFile = var3
    objName = AskEBDocumentName(strEbrFile,"ebr")
	'msgbox objName & strUrl
'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
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
Function BtnPreview() 
    
    Dim strEbrFile
    Dim objName
    
    Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	Dim var7
	
	dim strUrl
	dim arrParam, arrField, arrHeader
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then							'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))         'plantcd�� var1
	
	If frm1.rdoPrintType1.checked = True Then         '�������ý� ������ ebr  var2,var3,var6
		
		var2 = UniConvYYYYMMDDToDate(parent.gServerDateFormat,frm1.txtMonthDt.Year,frm1.txtMonthDt.Month,"01")      '�������� = �״޸� �����ǹǷ� ������ 
		var3 = "p4420oa1"
		var6 = UniDateAdd("M",1,UniConvDate(var2),parent.gServerDateFormat) 
		var7 = frm1.txtMonthDt.text
		
	Elseif frm1.rdoPrintType2.checked = True Then       '�Ϻ����ý� ������ ebr  var2,var3,var6
		
		var2 = UniConvDate(frm1.txtStartDt.text)
		var3 = "p4420oa2"
		var6 = UniConvDate(frm1.txtEndDt.text)
		var7 = ""
	End If
	
    If frm1.txtFromItemCd.value = "" Then            'ǰ�� txtFromItemCd   var4
		var4 = "0"
	Else
		var4 = Trim(frm1.txtFromItemCd.value)
	End If
	
	If frm1.txtToItemCd.value = "" Then              'ǰ�� txtToItemCd   var5
		var5 = "zzzzzzzzzzzzzzzzzz"
	Else
		var5 = Trim(frm1.txtToItemCd.value)
	End If
	
	strUrl = strUrl & "plant_cd|" & var1 
	strUrl = strUrl & "|Start_Dt|" & var2 
	strUrl = strUrl & "|End_Dt|" & var6 
	strUrl = strUrl & "|fr_item_cd|" & var4 
	strUrl = strUrl & "|to_item_cd|" & var5 
	strUrl = strUrl & "|Month_Text|" & var7 
		
	strEbrFile = var3
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
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
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

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

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
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemCd()
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

    If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
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
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetFromItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtfromItemCd.focus

End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
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
	
   If frm1.txtPlantCd.value= "" Then
		'Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
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
    
   	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetToItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtToItemCd.focus

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

Function SetFromItemCd(ByVal arrRet)
	frm1.txtFromItemCd.value = arrRet(0)
	frm1.txtFromItemNm.value = arrRet(1)  
End Function

Function SetToItemCd(ByVal arrRet)
	frm1.txtToItemCd.value = arrRet(0)
	frm1.txtToItemNm.value = arrRet(1)  
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

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
Sub rdoPrintType1_OnClick()
    frm1.txtStartDt.value = ""
    frm1.txtEndDt.value = ""
	Call ggoOper.SetReqAttr(frm1.txtMonthDt,"N")    'cboProdMgr
	Call ggoOper.SetReqAttr(frm1.txtStartDt,"Q")      'txtPurOrg
	Call ggoOper.SetReqAttr(frm1.txtEndDt,"Q")      'txtPurOrg	
	frm1.txtMonthDt.Text = UNIFormatMonth(strDate)    
End Sub

Sub rdoPrintType2_OnClick()
	frm1.txtMonthDt.value = "" 
	Call ggoOper.SetReqAttr(frm1.txtMonthDt,"Q")   'cboProdMgr
	Call ggoOper.SetReqAttr(frm1.txtStartDt,"N")     'txtPurOrg
	Call ggoOper.SetReqAttr(frm1.txtEndDt,"N")     'txtPurOrg
	frm1.txtStartDt.Text = StartDate
	frm1.txtEndDt.Text = EndDate
End Sub



</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��ǰ�����/�Ϻ����</font></td>
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
								    <TD CLASS="TD5" NOWRAP>�������</TD>
									<TD CLASS="TD6" NOWRAP>
								    <script language =javascript src='./js/p4420oa1_I812183463_txtMonthDt.js'></script>
									</TD>
								</TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4420oa1_I774504381_txtStartDt.js'></script>
										&nbsp;~&nbsp; 
										<script language =javascript src='./js/p4420oa1_I279063844_txtEndDt.js'></script>
									</TD>
								</TR>
								
								<TR>	
								    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="ǰ���">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="ǰ���">&nbsp;
									</TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
	    		<TR>
				<TD HEIGHT=10 WIDTH=100%>
					    <FIELDSET CLASS="CLSFLD">
					        <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>��¹��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType1" CLASS="RADIO" tag="x1" CHECKED>
									                     <LABEL FOR="rdoPrintType1">����</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType2" CLASS="RADIO"  tag="x1">
													     <LABEL FOR="rdoPrintType2">�Ϻ�</LABEL>
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
