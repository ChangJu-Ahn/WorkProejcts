
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b2501ma1.asp
'*  4. Program Name         : Plant Management
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "b2501mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "b2501mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID = "b2501mb3.asp"											 '��: �����Ͻ� ���� ASP�� 

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgBlnFlgConChg				'��: Condition ���� Flag
Dim IsOpenPop          

Dim lgRdoOldVal1

Dim BaseDate, StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                        '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size

    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	
	With frm1
	
	 .txtValidFromDt.text	= StartDate
	 .txtValidToDt.text		= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	 
	 frm1.txtInvOpenDt.text		= UNIFormatMonth(BaseDate)
	 frm1.txtInvClsDt.text		= UNIFormatMonth(BaseDate)
	 .txtPlngHrzn.value		= 0
	 .txtPtfForMps.value	= 0
	 .txtDtfForMps.value	= 0
	 .txtPtfForMrp.value	= 0	 
	 .txtPlngHrzn.Value     = 0
	 .cboSOFlag.value = "MTS"
	 lgRdoOldVal1 = 1
	End With
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd1.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd1.focus
	
End Function

'------------------------------------------  OpenBizArea()  -------------------------------------------------
'	Name : OpenBizArea()
'	Description : BizArea PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBizAreaCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������˾�"	
	arrParam(1) = "B_BIZ_AREA"				
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "�����"			
	
    
    arrField(0) = "BIZ_AREA_CD"	
    arrField(1) = "BIZ_AREA_NM"	
    
    arrHeader(0) = "�����"		
    arrHeader(1) = "������"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetBizArea(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBizAreaCd.focus
	
End Function

'------------------------------------------  OpenCurrency()  -------------------------------------------------
'	Name : OpenCurrency()
'	Description : Currency Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtCountryCd.className) = UCase(parent.UCN_PROTECTED)  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ȭ�˾�"					' �˾� ��Ī 
	arrParam(1) = "B_CURRENCY"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtCurCd.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "��ȭ"						' TextBox ��Ī 
	
    arrField(0) = "CURRENCY"						' Field��(0)
    arrField(1) = "CURRENCY_DESC"					' Field��(1)
    
    arrHeader(0) = "��ȭ"						' Header��(0)
    arrHeader(1) = "��ȭ��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCurrency(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCurCd.focus
	
End Function

'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtClnrType.className) = UCase(parent.UCN_PROTECTED) Then  Exit Function

	IsOpenPop = True

	arrParam(0) = "Į���� Ÿ�� �˾�"					' �˾� ��Ī 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtClnrType.Value)				' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "Į���� Ÿ��"							' TextBox ��Ī 
	
    arrField(0) = "CAL_TYPE"								' Field��(0)
    arrField(1) = "CAL_TYPE_NM"							' Field��(1)
    
    arrHeader(0) = "Į���� Ÿ��"						' Header��(0)
    arrHeader(1) = "Į���� Ÿ�Ը�"					' Header��(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
    
    Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'------------------------------------------  OpenCountry()  -------------------------------------------
'	Name : OpenCountry()
'	Description : Country PopUp
'------------------------------------------------------------------------------------------------------
Function OpenCountry()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"				<%' �˾� ��Ī %>
	arrParam(1) = "B_COUNTRY"					<%' TABLE ��Ī %>
	arrParam(2) = frm1.txtCountryCd.value		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "����"					<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "COUNTRY_CD"					<%' Field��(0)%>
    arrField(1) = "COUNTRY_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "�����ڵ�"				<%' Header��(0)%>
    arrHeader(1) = "����"					<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCountry(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCountryCd.focus
	
End Function
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd1.Value    = arrRet(0)		
	frm1.txtPlantNm1.Value    = arrRet(1)		
End Function

'------------------------------------------  SetBizArea()  --------------------------------------------------
'	Name : SetBizArea()
'	Description : BizArea Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizArea(byval arrRet)
	frm1.txtBizAreaCd.Value    = arrRet(0)		
	frm1.txtBizAreaNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetCurrency()  --------------------------------------------------
'	Name : SetCurrency()
'	Description : ��ȭ�ڵ� 
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrency(byval arrRet)
	frm1.txtCurCd.Value    = UCase(arrRet(0))		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetCalType()  --------------------------------------------------
'	Name : SetCalType()
'	Description : Į���� Ÿ�� 
'--------------------------------------------------------------------------------------------------------- 
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetCountry()  --------------------------------------------
'	Name : SetCountry()
'	Description : Country Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetCountry(Byval arrRet)
	With frm1
		.txtCountryCd.value = arrRet(0)
		.txtCountryNm.value = arrRet(1)
	End With
	lgBlnFlgChgValue = True
End Function

Function ChkValidData()
	'------------------------
	'Date Validation Check
	'------------------------
	ChkValidData = False
	With frm1

		If UNICDbl(.txtPtfForMps.Text) < UNICDbl(.txtDtfForMps.Text) Then
			Call DisplayMsgBox("972002", VBOKOnly, "MPS ��ȹ�Ⱓ(PTF)", "MPS Ȯ���Ⱓ(DTF)")
			.txtPtfForMps.Focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If
	
		If UNICDbl(.txtPlngHrzn.Text) < UNICDbl(.txtPtfForMps.Text) Then
			Call DisplayMsgBox("972002", "X", "Planning Horizon", "MPS ��ȹ�Ⱓ(PTF)")
			.txtPlngHrzn.focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If

		If UNICDbl(.txtPlngHrzn.Text) < UNICDbl(.txtPtfForMrp.Text) Then
			Call DisplayMsgBox("972002", "X", "Planning Horizon", "MRP Ȯ���Ⱓ")
			.txtPlngHrzn.focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If

		If lgIntFlgMode = parent.OPMD_CMODE Then
			If  UNIConvDate(.hInvOpenDt.value) > UNIConvDate(.hInvClsDt.value) Then 
				Call DisplayMsgBox("972002", "X", "�����������", "���ʽ��۳��")
				.txtInvClsDt.focus
				Set gActiveElement = document.activeElement   
				Exit Function
			End If
		End IF
		
		If ValidDateCheck(.txtValidFromDt, .txtValidToDt) = False Then Exit Function

		If lgIntFlgMode = parent.OPMD_CMODE Then
			If UNIConvDate(.hInvOpenDt.value) < UNIConvDate(.txtValidFromDt.text) Then 

				Call DisplayMsgBox("972002", "X", "���ʽ��۳��", "��ȿ�Ⱓ ������")
				.txtInvOpenDt.focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If
		End If

		If UNIConvDate(.hInvClsDt.value) > UNIConvDate(.txtValidToDt.Text) Then 

			Call DisplayMsgBox("972002", "X", "��ȿ�Ⱓ ������", "�����������")
			.txtValidToDt.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If

	End With
	
	ChkValidData = True
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	
    Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6", "4", "0")
	Call AppendNumberPlace("7", "3", "0")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
	Call ggoOper.FormatDate(frm1.txtInvOpenDt, parent.gDateFormat, "2")
	Call ggoOper.FormatDate(frm1.txtInvClsDt, parent.gDateFormat, "2")
	
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal
    Call InitVariables
	frm1.txtPlantCd1.focus
	Set gActiveElement = document.activeElement  
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtInvClsDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtInvClsDt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtInvClsDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtInvClsDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInvClsDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtInvClsDt_Change()
	Dim strYear1
	Dim strMonth1
	Dim InvClsDt

	strYear1 = frm1.txtInvClsDt.Year
	strMonth1 = frm1.txtInvClsDt.Month
	
	InvClsDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear1, strMonth1, "01")
	frm1.hInvClsDt.value = UNIDateAdd("d", -1, UNIDateAdd("m", 1, InvClsDt, parent.gDateFormat), parent.gDateFormat)	
	
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInvOpenDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtInvOpenDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInvOpenDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtInvOpenDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInvClsDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtInvOpenDt_Change()
	Dim strYear2
	Dim strMonth2
	Dim InvOpenDt

	strYear2 = frm1.txtInvOpenDt.Year
	strMonth2 = frm1.txtInvOpenDt.Month
	
	InvOpenDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear2, strMonth2, "01")
	frm1.hInvOpenDt.value = UNIDateAdd("d", -1, UNIDateAdd("m", 1, InvOpenDt, parent.gDateFormat), parent.gDateFormat)

    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPlngHrzn_Change()
		lgBlnFlgChgValue = True	
End Sub	

Sub txtDtfForMps_Change()
		lgBlnFlgChgValue = True	
End Sub	

Sub txtPtfForMps_Change()
		lgBlnFlgChgValue = True	
End Sub	
	
Sub txtPtfForMrp_Change()
		lgBlnFlgChgValue = True	
End Sub	

Sub cboSOFlag_OnChange() 
	lgBlnFlgChgValue = True		
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False															'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    
	If frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
	End If
    
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables

	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then   
		Exit Function           
    End If     														'��: Query db data
       
    FncQuery = True																'��: Processing is OK
        
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False																'��: Processing is NG
    
	'-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")					'��: "����Ÿ�� ����Ǿ����ϴ�. �ű��Է��� �Ͻðڽ��ϱ�?"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                      '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call SetToolbar("11101000000011")
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
    
    frm1.txtPlantCd2.focus
    Set gActiveElement = document.activeElement   
    
    FncNew = True																'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim intRetCD
    
    FncDelete = False														'��: Processing is NG
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
	'-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then   
		Exit Function           
    End If     											'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                       '��: No data changed!!
        Exit Function
    End If
    
	'-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then										'��: Check contents area
       Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '-----------------------
    
    If DbSave = False Then   
		Exit Function           
    End If                          '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												'��: Indicates that current mode is Crate mode
    
     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'��: This function lock the suitable field
    Call SetToolbar("11101000000011")
    
    frm1.txtPlantCd2.value = ""
    frm1.txtPlantNm2.value = ""
	
	frm1.txtValidFromDt.text	= StartDate
	frm1.txtValidToDt.text		= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")

	frm1.txtInvClsDt.text		= UNIFormatMonth(BaseDate)
	frm1.txtInvOpenDt.text		= UNIFormatMonth(BaseDate)
    
    frm1.txtPlantCd2.focus
    Set gActiveElement = document.activeElement 
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
	
    Err.Clear                                                               '��: Protect system from crashing
    
    LayerShowHide(1)
			
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd1=" & Trim(frm1.txtPlantCd1.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & "P"									'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables


    Err.Clear                                                               '��: Protect system from crashing
    
    LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd1=" & Trim(frm1.txtPlantCd1.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & "N"									'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)											'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")		'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG
    
    LayerShowHide(1)
		
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003						'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd2=" & Trim(frm1.txtPlantCd2.value)		'��: ���� ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call InitVariables()
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd1=" & Trim(frm1.txtPlantCd1.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = false
    
    frm1.txtPlantNm2.focus 
	Set gActiveElement = document.activeElement 
	
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call SetToolbar("11111000111111")
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave()
	Dim BlnRetCd
	Dim strVal
	Dim strYear1
	Dim strMonth1
	Dim strYear2
	Dim strMonth2
	Dim InvOpenDt
	Dim InvClsDt

	Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG
	
	strYear1 = frm1.txtInvClsDt.Year
	strMonth1 = frm1.txtInvClsDt.Month
	strYear2 = frm1.txtInvOpenDt.Year
	strMonth2 = frm1.txtInvOpenDt.Month

	InvOpenDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear2, strMonth2, "01")
	frm1.hInvOpenDt.value = UNIDateAdd("d", -1, UNIDateAdd("m", 1, InvOpenDt, parent.gDateFormat), parent.gDateFormat)
	InvClsDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear1, strMonth1, "01")
	frm1.hInvClsDt.value = UNIDateAdd("d", -1, UNIDateAdd("m", 1, InvClsDt, parent.gDateFormat), parent.gDateFormat)	
	
	BlnRetCd = ChkValidData

	If BlnRetCd = False Then
		Exit Function
	End if
	
	LayerShowHide(1)
		
	With frm1
		
		.txtMode.value = parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtCoCd.value = parent.gCompany 
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()															'��: ���� ������ ���� ���� 

    frm1.txtPlantCd1.value = frm1.txtPlantCd2.value 

    Call InitVariables
    
    Call MainQuery()

End Function

Sub InitComboBox()   
	   
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("C0004", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboSOFlag , lgF0, lgF1, Chr(11))
    
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�� ��</TD>
									<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="�� ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()"> <INPUT TYPE=TEXT ID="txtPlantNm1" SIZE=40 NAME="txtPlantNm1" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE CLASS="TB2" CELLSPACING=0>
								<TD WIDTH=50% valign=top>
									<FIELDSET>
										<LEGEND>�Ϲ�����</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>�� ��</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd2" SIZE=10 MAXLENGTH=4 tag="23XXXU"  ALT="�� ��">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm2" SIZE=30 MAXLENGTH=40 tag="22" ALT="�����"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�����</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=12 MAXLENGTH=10 tag="22XXXU"  ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=28 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>Į���� Ÿ��</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="22XXXU"  ALT="Į���� Ÿ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=30 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>��ȭ�ڵ�</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCurCd" SIZE=5 MAXLENGTH=3 tag="22X6XU"  ALT="��ȭ�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCur" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency()"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>Planning Horizon</TD>
												<TD CLASS="TD6" NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>																
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtPlngHrzn CLASSID=<%=gCLSIDFPDS%> tag="22X6Z" ALT="Planning Horizon"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>&nbsp;��
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>MPS Ȯ���Ⱓ(DTF)</TD>												
												<TD CLASS="TD6" NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtDtfForMps CLASSID=<%=gCLSIDFPDS%> tag="22X7Z" ALT="MPS Ȯ���Ⱓ(DTF)"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>&nbsp;��
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>MPS ��ȹ�Ⱓ(PTF)</TD>												
												<TD CLASS="TD6" NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtPtfForMps CLASSID=<%=gCLSIDFPDS%> ALT="MPS ��ȹ�Ⱓ(PTF)" tag="22X7Z"> </OBJECT>');</SCRIPT>
															</TD>
															<tD>&nbsp;��
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>MRP Ȯ���Ⱓ</TD>
												<TD CLASS="TD6" NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtPtfForMrp CLASSID=<%=gCLSIDFPDS%> ALT="MRP Ȯ���Ⱓ" tag="22X7Z"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>&nbsp;��
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET>
										<LEGEND>���������</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>���ʽ��۳��</TD>
												<TD CLASS="TD6" NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtInvOpenDt CLASSID=<%=gCLSIDFPDT%> tag="23X1" ALT="���ʽ��۳��"> </OBJECT>');</SCRIPT>
												</TD>										
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>����������� <INPUT TYPE=HIDDEN NAME="txtBomLastUpdatedDt" SIZE=10 MAXLENGTH=10  STYLE="display:none;" tag="2X" ></TD>
												<TD CLASS="TD6" NOWRAP>													
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtInvClsDt CLASSID=<%=gCLSIDFPDT%> tag="23X1" ALT="�����������"> </OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<LEGEND>��ȿ�Ⱓ</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS="TD5" NOWRAP>������ȿ�Ⱓ</TD>
													<TD CLASS="TD6" NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidFromDt CLASSID=<%=gCLSIDFPDT%> ALT="��ȿ�Ⱓ ������" tag="23X1"> </OBJECT>');</SCRIPT>
														&nbsp;~&nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidToDt CLASSID=<%=gCLSIDFPDT%> ALT="��ȿ�Ⱓ ������" tag="22X1"> </OBJECT>');</SCRIPT>
													</TD>													
												</TR>
											</TABLE>
									</FIELDSET>
									<FIELDSET>
										<LEGEND>��������</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCountryCd" SIZE=5 MAXLENGTH=2 tag="22XXXU"  ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountry" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCountry()">&nbsp;<INPUT TYPE=TEXT NAME="txtCountryNm" SIZE=30 tag="24"></TD>
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<LEGEND>��������</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS="TD5" NOWRAP>�����</TD>
												<TD CLASS="TD6" NOWRAP><SELECT NAME="cboSOFlag" CLASS=required ALT="�����" STYLE="Width: 140px;" tag="22"></SELECT></TD></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>	
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtCoCd" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="hInvOpenDt" tag="24"><INPUT TYPE=HIDDEN NAME="hInvClsDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

