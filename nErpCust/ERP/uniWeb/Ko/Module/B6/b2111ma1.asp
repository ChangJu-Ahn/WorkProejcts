<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : B2111MA1
'*  4. Program Name         : Biz Area(������������)
'*  5. Program Desc         : ������������ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/17
'*  8. Modified date(Last)  : 2001/03/19
'*  9. Modifier (First)     : ahj
'* 10. Modifier (Last)      : hersheys / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit    


'============================================  1.2.1 Global ��� ����  ====================================

<%
StartDate = DateSerial(Year(Date),Month(Date),1)
StartDate = Year(StartDate) & "-" & Right("0" & Month(StartDate),2) & "-" & Right("0" & Day(StartDate),2)
EndDate = Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2)
%>

Const BIZ_PGM_ID = "b2111mb1.asp"											 '��: �����Ͻ� ���� ASP�� 

'============================================  1.2.2 Global ���� ����  ===================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2. Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

Dim lgBlnFlgChgValue				'��: Variable is for Dirty flag
Dim lgIntGrpCount				'��: Group View Size�� ������ ���� 
Dim lgIntFlgMode					'��: Variable is for Operation Status

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""

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
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenbizareaInfo()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
Function OpenbizareaInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����� �˾�"					' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"						' TABLE ��Ī 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name COndition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "�����"

    arrField(0) = "BIZ_AREA_CD"						' Field��(0)
    arrField(1) = "BIZ_AREA_NM"						' Field��(1)

    arrHeader(0) = "������ڵ�"					' Header��(0)
    arrHeader(1) = "������"						' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus

	    Exit Function
	Else
		Call SetbizareaInfo(arrRet,iWhere)
	End If

End Function

Function OpenZipCode(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("ZipPopup")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZipPopup", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode
	arrParam(1) = ""
	arrParam(2) = parent.gCountry

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtZipCode.focus
	    Exit Function
	Else
		Call SetBizAreaInfo(arrRet,iWhere)
	End If

End Function


'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
Function SetBizAreaInfo(ByVal arrRet, ByVal iWhere)

	With frm1
		If iWhere = 0 Then
			.txtBizAreaCd.focus
			.txtBizAreaCd.value = arrRet(0)
			.txtBizAreaNm.value = arrRet(1)
		ElseIf iWhere = 1 Then
			.txtZipCode.focus
			.txtZipCode.value = arrRet(0)
			.txtAddr1.value     = arrRet(1)

			lgBlnFlgChgValue = True
		End If
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

	arrParam(0) = "���� �˾�"						' �˾� ��Ī 
	arrParam(1) = "B_COUNTRY"							' TABLE ��Ī 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"

    arrField(0) = "COUNTRY_CD"							' Field��(0)
    arrField(1) = "COUNTRY_NM"							' Field��(1)


    arrHeader(0) = "�����ڵ�"						' Header��(0)
    arrHeader(1) = "������"						' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtconutry_cd.focus
	    Exit Function
	Else
		Call SetCountryInfo(arrRet,iWhere)
	End If	

End Function


'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetCountryInfo()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCountryInfo(Byval arrRet,byval iWhere)

	With frm1
		If iWhere = 0 Then
			.txtconutry_cd.focus
			.txtconutry_cd.value = arrRet(0)
		End If
		lgBlnFlgChgValue = True
	End With

End Function


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
Function OpenTaxOffice(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������ �˾�"						' �˾� ��Ī 
	arrParam(1) = "B_TAX_OFFICE"						' TABLE ��Ī 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "������"

    arrField(0) = "TAX_OFFICE_CD"						' Field��(0)
    arrField(1) = "TAX_OFFICE_NM"						' Field��(1)

    arrHeader(0) = "�������ڵ�"						' Header��(0)
    arrHeader(1) = "��������"						' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTaxOfficeCd.focus
	    Exit Function
	Else
		Call SetTaxOffice(arrRet,iWhere)
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
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetTaxOffice()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTaxOffice(Byval arrRet,byval iWhere)

	With frm1
		If iWhere = 1 Then
			.txtTaxOfficeCd.focus
			.txtTaxOfficeCd.value   = arrRet(0)
			.txtTaxOfficeNm.value = arrRet(1)
		End If

		lgBlnFlgChgValue = True
	End With

End Function


'========================================================================================================= 
Function OpenCommonPopupInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	select case iwhere
		case 0
			arrParam(0) = "���� �˾�"					' �˾� ��Ī 
			arrParam(1) = "B_MINOR"							' TABLE ��Ī 
			arrParam(2) =  strCode							 	' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("B9003", "''", "S") & "  "					' Where Condition
			arrParam(5) = "����"

			arrField(0) = "MINOR_CD"							' Field��(0)
			arrField(1) = "MINOR_NM"						' Field��(1)

			arrHeader(0) = "�����ڵ�"						' Header��(0)
			arrHeader(1) = "���¸�"					' Header��(1)
		case 1
			arrParam(0) = "���� �˾�"					' �˾� ��Ī 
			arrParam(1) = "B_MINOR"							' TABLE ��Ī 
			arrParam(2) =  strCode							 	' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("B9002", "''", "S") & "  "					' Where Condition
			arrParam(5) = "����"

			arrField(0) = "MINOR_CD"							' Field��(0)
			arrField(1) = "MINOR_NM"						' Field��(1)

			arrHeader(0) = "�����ڵ�"						' Header��(0)
			arrHeader(1) = "������"					' Header��(1)  

		case 2
			arrParam(0) = "�� ���ݽŰ����� �˾�"					' �˾� ��Ī 
			arrParam(1) = "B_TAX_BIZ_AREA"							' TABLE ��Ī 
			arrParam(2) =  strCode							 	' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""					' Where Condition
			arrParam(5) = "�� ���ݽŰ�����"

			arrField(0) = "TAX_BIZ_AREA_CD"							' Field��(0)
			arrField(1) = "TAX_BIZ_AREA_NM"						' Field��(1)

			arrHeader(0) = "�� ���ݽŰ�����"						' Header��(0)
			arrHeader(1) = "�� ���ݽŰ������"					' Header��(1)  
 	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	select case iwhere
		case 0
			frm1.txtInd_class.focus
		case 1
			frm1.txtInd_Type.focus
		case 2
			frm1.txtReportBizArea.focus
 	End Select	
	    Exit Function
	Else
		Call SetCommonPopupInfo(arrRet,iWhere)
	End If	

End Function

'========================================================================================================= 
Function SetCommonPopupInfo(Byval arrRet,byval iWhere)

	With frm1
		If iWhere = 0 Then
			.txtInd_class.focus
			.txtInd_class.value = arrRet(0)
			.txtInd_class_Nm.value = arrRet(1)
		Elseif iWhere = 1 Then
			.txtInd_Type.focus
			.txtInd_Type.value = arrRet(0)
			.txtInd_Type_Nm.value   = arrRet(1)
		Elseif iWhere = 2 Then
			.txtReportBizArea.focus
			.txtReportBizArea.value = arrRet(0)
			.txtReportBizAreaNm.value   = arrRet(1)
		End If

		lgBlnFlgChgValue = True
	End With

End Function


'==========================================================================================================
Sub Form_Load()

    Call InitVariables
    Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")

    Call SetToolBar("1110100000001111")
	frm1.txtBizAreaCd.focus	
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
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables

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
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call InitVariables

    Call SetToolBar("1110100000001111")

	frm1.txtBizAreaCd.focus

    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False

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
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete
    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 

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

  '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave

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

    lgIntFlgMode = parent.OPMD_CMODE

     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")

	lgBlnFlgChgValue = True

    Call SetToolBar("1110100000001111")
    frm1.txtBizAreaCd_Body.value = ""

    frm1.txtBizAreaCd_Body.focus
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

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
	End IF

    Call LayerShowHide(1)
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtBizAreaCd=" & Trim(frm1.txtBizAreaCd_Body.value)
    strVal = strVal & "&PrevNextFlg=" & "P"									'��: ��ȸ ���� ����Ÿ 
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext()
    Dim strVal
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

	Call LayerShowHide(1)
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtBizAreaCd=" & Trim(frm1.txtBizAreaCd_Body.value)
    strVal = strVal & "&PrevNextFlg=" & "N"									'��: ��ȸ ���� ����Ÿ 
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
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

    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtBizAreaCd=" & Trim(frm1.txtBizAreaCd.value)

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	lgBlnFlgChgValue = False
	Call FncNew()
End Function


'========================================================================================
' Function Name : cboXCH_RATE_FG_OnChange
' Function Desc : 
'========================================================================================
Sub cboXCH_RATE_FG_OnChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================
Function DbQuery()

    Err.Clear
    DbQuery = False

    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtBizAreaCd=" & Trim(frm1.txtBizAreaCd.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & ""
    call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True
End Function


'========================================================================================
Function DbQueryOk()
    Call SetToolBar("1111100011111111")
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    lgIntFlgMode = parent.OPMD_UMODE
End Function


'========================================================================================
Function DbSave() 

    Err.Clear
	DbSave = False

    Dim strVal
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	End With

    DbSave = True
End Function


'========================================================================================
Function DbSaveOk()
    frm1.txtBizAreaCd.value = frm1.txtBizAreaCd_Body.value 
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
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtBizAreaCd" MAXLENGTH="10" SIZE=10 ALT ="������ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenbizareaInfo(frm1.txtBizAreaCd.value,0)"> <INPUT NAME="txtBizAreaNm" MAXLENGTH="50" SIZE=30 ALT ="������" tag="14X"></TD>
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
								<TD CLASS=TD5 NOWRAP>������ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizAreaCd_Body" ALT="������ڵ�" MAXLENGTH="10" SIZE=10 tag = "23XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizAreaNm_Body" ALT="������" MAXLENGTH="50" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������幮��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizAreaFullNm" ALT="������幮��" MAXLENGTH="50" SIZE=30 tag ="22"></TD>
								<TD CLASS=TD5 NOWRAP>����念����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizAreaEngNm" ALT="����念����" MAXLENGTH="50" SIZE=30 tag ="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ڵ�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOwnRgstNo" ALT="����ڵ�Ϲ�ȣ" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
								<TD CLASS=TD5 NOWRAP>��ǥ�ڸ�</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepreNm" ALT="��ǥ�ڸ�" MAXLENGTH="50" STYLE="TEXT-ALIGN:left" tag  ="22"></TD>				    	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxOfficeCd"   ALT="�������ڵ�" Size = "12" MAXLENGTH="10" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenTaxOffice(frm1.txtTaxOfficeCd.value, 1)">
													 <INPUT NAME="txtTaxOfficeNm" MAXLENGTH="25" SIZE = "25" tag="24X"></TD>
							    <TD CLASS=TD5 NOWRAP>�Ű�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReportBizArea" ALT="�Ű�����" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCommonPopupInfo(frm1.txtReportBizArea.value,2)">
													 <INPUT NAME="txtReportBizAreaNm" ALT="�Ű������" SIZE="20" tag = "24" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_class" ALT="����" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCommonPopupInfo(frm1.txtInd_class.value,0)">
													 <INPUT NAME="txtInd_class_Nm" ALT="����" SIZE="20" tag = "24" ></TD>

								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="����" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCommonPopupInfo(frm1.txtInd_Type.value,1)">
													<INPUT NAME="txtInd_Type_Nm" ALT="����" SIZE="20" tag = "24" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>FAX��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFaxNo" ALT="FAX��ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="22" ></TD>	
 							    <TD CLASS=TD5 NOWRAP>��ȭ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo" ALT="��ȭ��ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="2"></TD>
 							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>�����ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZipCode" ALT="�����ȣ" MAXLENGTH="12" Size="11" STYLE="TEXT-ALIGN:left" tag  ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZipCode(frm1.txtZipCode.value, 1)"></TD>
							    <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ּ�</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr1"  ALT="�ּ�"     MAXLENGTH="100" SIZE="80" STYLE="TEXT-ALIGN:left" tag="22" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr2"  ALT="�ּ�"     MAXLENGTH="100" SIZE="80" STYLE="TEXT-ALIGN:left" tag="2" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ּ�</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtEng1Addr" ALT="�����ּ�" MAXLENGTH="50" Size="80" STYLE="TEXT-ALIGN: left" tag ="22"></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtEng2Addr" ALT="�����ּ�" MAXLENGTH="50" Size="80" STYLE="TEXT-ALIGN: left" tag ="2"></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtEng3Addr" ALT="�����ּ�" MAXLENGTH="50" Size="80" STYLE="TEXT-ALIGN: left" tag ="2"></TD>	
							</TR>
							<% Call SubFillRemBodyTd5656(2) %>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

