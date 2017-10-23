<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : B2210MA1
'*  4. Program Name         : Company Register(�����������)
'*  5. Program Desc         : ����������� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/3/20
'*  8. Modified date(Last)  : 2000/8/29
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Kwon Yong Gyoun / Cho Ig Sung/kang eun kyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'***********************************************************************k*********************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '��: indicates that All variables must be declared in advance 


'********************************************  1.2 Global ����/��� ����  *********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================

Const BIZ_PGM_ID = "b2210mb1.asp"											 '��: �����Ͻ� ���� ASP�� 

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""

	frm1.txtCO_CD.value = parent.gCompany
	frm1.txtco_cd.focus  
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


 
'------------------------------------------  InitComboBox()  ----------------------------------------------
'	Name :InitComboBox()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox_One()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTaxPolicy ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Two()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCurPolicy ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Three()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = " & FilterVar("A1004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboxch_rate_fg ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Four()
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1020", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboOpenAcctFg ,lgF0  ,lgF1  ,Chr(11))  '�̰��������(�����ڵ��)
	Call SetCombo2(frm1.cboXchErrorUseFg ,lgF0  ,lgF1  ,Chr(11))  '�����ȯ����� 
End Sub

Sub InitComboBox_Five()
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("Z0015", "''", "S") & " ) ORDER BY MINOR_CD DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboQmdpalignopt ,lgF0  ,lgF1  ,Chr(11))  '��Ƽȭ��Ҽ�������(��ȸ��)
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("Z0016", "''", "S") & " ) ORDER BY MINOR_CD DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboImdpalignopt ,lgF0  ,lgF1  ,Chr(11))  '��Ƽȭ��Ҽ�������(�Է¿�)
End Sub

Sub InitComboBox_Six()
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("B9040", "''", "S") & " ) ORDER BY MINOR_CD DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboInvPostingFg ,lgF0  ,lgF1  ,Chr(11))  '��������ù�� 
End Sub

'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenCompanyInfo()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

Function OpenCompanyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"						' �˾� ��Ī 
	arrParam(1) = "B_COMPANY"						' TABLE ��Ī 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"

    arrField(0) = "Upper(CO_CD)"					' Field��(0)
    arrField(1) = "CO_FULL_NM"						' Field��(1)

    arrHeader(0) = "�����ڵ�"						' Header��(0)
    arrHeader(1) = "���θ�"						' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	Else
		Call SetCompanyInfo(arrRet,iWhere)
	End If	

End Function



'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
Function SetCompanyInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCO_CD.focus
			.txtCO_CD.value     = arrRet(0)
			.txtCO_FULLNM.value = arrRet(1)
		End If
'		lgBlnFlgChgValue = False
	End With

End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : OpenCountryInfo()
'	Description : �����ڵ� popup
'========================================================================================================= 

Function OpenCountryInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"							' �˾� ��Ī 
	arrParam(1) = "B_COUNTRY"							' TABLE ��Ī 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"

    arrField(0) = "COUNTRY_CD"							' Field��(0)
    arrField(1) = "COUNTRY_NM"							' Field��(1)

    arrHeader(0) = "�����ڵ�"							' Header��(0)
    arrHeader(1) = "������"							' Header��(1)
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCountryCd.focus
	    Exit Function
	Else
		Call SetCountryInfo(arrRet,iWhere)
	End If
End Function


'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetCountryInfo()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCountryInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCountryCd.focus
			.txtCountryCd.value = arrRet(0)
		End If
		lgBlnFlgChgValue = True
	End With

End Function


'========================================================================================================= 
Function OpenCurrencyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ڱ���ȭ �˾�"						' �˾� ��Ī 
	arrParam(1) = "B_CURRENCY"							' TABLE ��Ī 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "�ڱ���ȭ"

    arrField(0) = "CURRENCY"							' Field��(0)
    arrField(1) = "CURRENCY_DESC"						' Field��(1)

    arrHeader(0) = "�ڱ���ȭ�ڵ�"						' Header��(0)
    arrHeader(1) = "�ڱ���ȭ��"						' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtLOC_CUR.focus
	    Exit Function
	Else
		Call SetCurrencyInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function OpenIndclassInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"							' �˾� ��Ī 
	arrParam(1) = "B_MINOR"								' TABLE ��Ī 
	arrParam(2) =  strCode								' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("B9003", "''", "S") & "  "					' Where Condition
	arrParam(5) = "����"

    arrField(0) = "MINOR_CD"							' Field��(0)
    arrField(1) = "MINOR_NM"							' Field��(1)

    arrHeader(0) = "�����ڵ�"							' Header��(0)
    arrHeader(1) = "���¸�"							' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtInd_class.focus
	    Exit Function
	Else
		Call SetOpenIndclassInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function OpenIndTypeInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"							' �˾� ��Ī 
	arrParam(1) = "B_MINOR"								' TABLE ��Ī 
	arrParam(2) =  strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("B9002", "''", "S") & "  "					' Where Condition
	arrParam(5) = "����"

    arrField(0) = "MINOR_CD"							' Field��(0)
    arrField(1) = "MINOR_NM"							' Field��(1)

    arrHeader(0) = "�����ڵ�"							' Header��(0)
    arrHeader(1) = "������"							' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtInd_Type.focus
	    Exit Function
	Else
		Call SetOpenIndTypeInfo(arrRet,iWhere)
	End If
End Function


'========================================================================================================= 
Function OpenZipCode(ByVal strCode, ByVal iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If Trim(frm1.txtCountryCd.value) = "" Then
		MsgBox "������ ���� �Է��ϼ���", vbInformation, "uniERP(Information)"
		frm1.txtCountryCd.focus
		IsOpenPop = False
		Exit Function
	End IF
	iCalledAspName = AskPRAspName("ZipPopup")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.parent.VB_INFORMATION, "ZipPopup", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode
	arrParam(1) = ""
	arrParam(2) = Trim(frm1.txtCountryCd.value)

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
			frm1.txtzip_code.focus
	    Exit Function
	Else
		Call SetCurrencyInfo(arrRet,iWhere)
	End If
End Function


'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCurrency()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrencyInfo(Byval arrRet,byval iWhere)'
	With frm1
		If iWhere = 0 Then
			.txtLOC_CUR.focus
			.txtLOC_CUR.value = arrRet(0)
		ElseIf iWhere = 1 Then
			.txtzip_code.focus
			.txtzip_code.value = arrRet(0)
			.txtaddr.value     = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function SetOpenIndclassInfo(Byval arrRet,byval iWhere)'
	With frm1
		If iWhere = 0 Then

			.txtInd_class.focus
			.txtInd_class.value = arrRet(0)
			.txtInd_class_Nm.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function SetOpenIndTypeInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtInd_Type.focus
			.txtInd_Type.value = arrRet(0)
			.txtInd_Type_Nm.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function



'========================================================================================================= 
Sub Form_Load()
    Call InitVariables																'��: Initializes local global variables
    Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("1100100000001111")
    Call InitComboBox_One
    Call InitComboBox_Two
    Call InitComboBox_Three
	Call InitComboBox_Four
	Call InitComboBox_Five
	Call InitComboBox_Six

	Call ggoOper.FormatDate(frm1.txtFirstDeprYyyymm, parent.gDateFormat, 2)
    'Call ggoOper.FormatDate(frm1.txtLastDeprYyyymm, parent.gDateFormat, 2)

	frm1.txtco_cd.focus 

    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed

	FncQuery

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
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
    Call DbQuery
    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables

    Call SetToolbar("1100100000001111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strYear,strMonth,strDay
    Dim strYear1,strMonth1,strDay1

	FncSave = False
	Err.Clear

	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
	    Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '��: Check contents area
	   Exit Function
	End If

	If CompareDateByFormat(frm1.txtFISC_Start_DT.text,frm1.txtFISC_End_DT.text,frm1.txtFISC_Start_DT.Alt,frm1.txtFISC_End_DT.Alt, _
        	               "970024",frm1.txtFISC_Start_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFISC_Start_DT.focus
	   Exit Function
	End If
   
 	Call ExtractDateFrom(frm1.FDeprDateTime1.Text,frm1.FDeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    frm1.hFirstDeprYyyymm.value = strYear & strMonth

 	'Call ExtractDateFrom(frm1.LDeprDateTime1.Text,frm1.LDeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    'frm1.hLastDeprYyyymm.value = strYear1 & strMonth1


	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then
		Exit Function
	End If

	FncSave = True
End Function


'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode

     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
	lgBlnFlgChgValue = True

    frm1.txtCO_CD_Body.value = ""

    frm1.txtCO_CD_Body.focus
    
End Function


'========================================================================================
Function FncCancel()
     On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF

    response.write lgPrevNo

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtco_cd = " & lgPrevNo

	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						  '��: �����Ͻ� ó�� ASP�� ���°� 
    strVal = strVal & "&txtco_cd=" & lgNextNo

	Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtco_cd=" & Trim(frm1.txtco_cd.value)				'��: ���� ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function


'=======================================================================================================
'   Event Name : txtFISC_START_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFISC_START_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_START_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_START_DT.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFOUNDATION_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFOUNDATION_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFOUNDATION_DT.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFISC_END_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_END_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_END_DT.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFirstDeprYyyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtFirstDeprYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFirstDeprYyyymm.Focus
    End If
End Sub

'=======================================================================================================
Sub txtLastDeprYyyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtLastDeprYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLastDeprYyyymm.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFISC_START_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtFOUNDATION_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtFISC_END_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtTransStartDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub txtFirstDeprYyyymm_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub txtLastDeprYyyymm_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : cboQmdpalignopt_OnChange()
' Function Desc : 
'========================================================================================
Sub cboQmdpalignopt_OnChange()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================= 
Sub cboImdpalignopt_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub cboTaxPolicy_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub cboCurPolicy_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboXCH_RATE_FG_OnChange()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================
Sub txtFISC_CNT_Change() 
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboOpenAcctFg_OnChange() 
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboXchErrorUseFg_OnChange() 
	lgBlnFlgChgValue = True
End Sub

Sub cboInvPostingFg_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()
    Call SetToolbar("1100100000011111")
    lgBlnFlgChgValue = False
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	'20030916 jsk ���ʻ󰢳�� 
	Call settxtFirstDeprYyyymmMode '�����ڻ��������� ������ �� ���� 
    lgIntFlgMode = parent.OPMD_UMODE
End Function

Function SettxtFirstDeprYyyymmMode()

	call CommonQueryRs("TOP 1 ACQ_NO"," A_ASSET_ACQ "," ACQ_FG = " & FilterVar("03", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If lgF0 <> "" Then
		Call ggoOper.SetReqAttr(frm1.txtFirstDeprYyyymm, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtFirstDeprYyyymm, "N")
	End If	
End Function
'========================================================================================
Function DbSave() 

    Err.Clear
	DbSave = False

    Dim strVal

    Call LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
    frm1.txtCO_CD.value = frm1.txtCO_CD_Body.value 
    lgBlnFlgChgValue = False
    FncQuery
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtCO_CD" MAXLENGTH="10" SIZE=10 ALT ="�����ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCompanyInfo(frm1.txtco_cd.value,0)"> <INPUT NAME="txtCO_FULLNM" MAXLENGTH="30" SIZE=30 ALT ="���θ�" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_CD_Body" ALT="�����ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN:Left" tag = "23"></TD>
								<TD CLASS=TD5 NOWRAP>���ξ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_NM" ALT="���ξ��" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���θ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_FULL_NM_Body" ALT="���θ�" MAXLENGTH="50" SIZE=45 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
								<TD CLASS=TD5 NOWRAP>���ο�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtENG_NM" ALT="���ο�����" MAXLENGTH="50" SIZE=30 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ε�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOwn_Rgst_No" ALT="���ε�Ϲ�ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="22"></TD>
								<TD CLASS=TD5 NOWRAP>��ǥ�ڸ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtREPRE_NM" ALT="��ǥ�ڸ�" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǥ���ֹε�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_Rgst_No" ALT="��ǥ���ֹε�Ϲ�ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="22" ></TD>
								<TD CLASS=TD5 NOWRAP>FAX��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFAX_NO" ALT="FAX��ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="2" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>								
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_class" ALT="����" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenIndclassInfo(frm1.txtInd_class.value,0)">
								<INPUT NAME="txtInd_class_Nm" ALT="����" SIZE="20" tag = "24" ></TD>
								<TD CLASS=TD5 NOWRAP>��ȭ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTEL_NO" ALT="��ȭ��ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="2"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="����" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenIndTypeInfo(frm1.txtInd_Type.value,0)">
								<INPUT NAME="txtInd_Type_Nm" ALT="����" SIZE="20" tag = "24" ></TD>
								<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCountryCd" ALT="�����ڵ�" MAXLENGTH="2" SIZE="4" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCountryInfo(frm1.txtCountryCd.value,0)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ȸ��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE style="LEFT: 0px; WIDTH: 40px; TOP: 0px; HEIGHT: 20px" name=txtFISC_CNT CLASSID=<%=gCLSIDFPDS%> tag="22X6Z" ALT="ȸ��" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�ڱ���ȭ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLOC_CUR" ALT="�ڱ���ȭ" MAXLENGTH="3" SIZE="4" STYLE="TEXT-ALIGN:left" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCurrencyInfo(frm1.txtLOC_CUR.value,0)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_START_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="����������" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>���μ�������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFOUNDATION_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���μ�������" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_END_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�����������" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�̰��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOpenAcctFg" ALT="�̰��������" STYLE="WIDTH: 100px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ʰ����󰢳��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFirstDeprYyyymm CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���ʰ����󰢳��" tag="21X1" id=FDeprDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
<!--								<TD CLASS=TD5 NOWRAP>���������󰢳��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtLastDeprYyyymm CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���������󰢳��" tag="21X1" id=LDeprDateTime1></OBJECT>');</SCRIPT></TD>
-->
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtTransStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��������������" tag="24" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>��������ID</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCurOrgChangeID" ALT="��������ID" MAXLENGTH="5" Size = "5" STYLE="TEXT-ALIGN:Center" tag = "24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ݰ����å</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboTaxPolicy" ALT="���ݰ����å" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>ȯ������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboXCH_RATE_FG" ALT="ȯ������" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ȸ��Ҽ����ڸ���</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboQmdpalignopt" ALT="��ȸ��Ҽ����ڸ���" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>�Է¿�Ҽ����ڸ���</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboImdpalignopt" ALT="�Է¿�Ҽ����ڸ���" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ȯ������Ұ�</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboXchErrorUseFg" ALT="ȯ������Ұ�" STYLE="WIDTH: 100px" tag="22"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>��������ù��</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboInvPostingFg" ALT="��������ù��" STYLE="WIDTH: 100px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtzip_code" ALT="�����ȣ" MAXLENGTH="12" SIZE="11" STYLE="TEXT-ALIGN:left" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZipCode(frm1.txtZip_Code.value, 1)"></TD>
								<TD CLASS=TD5 NOWRAP>��ȯȯ����å</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCurPolicy" ALT="��ȯȯ����å" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ּ�</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtaddr" ALT="�ּ�" MAXLENGTH="128" SIZE="95" STYLE="TEXT-ALIGN:left"  tag="22" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ּ�</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txteng_addr" ALT="�����ּ�" MAXLENGTH="128" SIZE="95" STYLE="TEXT-ALIGN:left"  tag="2" ></TD>
							</TR>
<!--							<% Call SubFillRemBodyTd5656(2) %> -->
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hFirstDeprYyyymm" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hLastDeprYyyymm" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

