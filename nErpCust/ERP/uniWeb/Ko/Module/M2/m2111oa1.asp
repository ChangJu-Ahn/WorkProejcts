<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        :
'*  3. Program ID           : M2111OA1
'*  4. Program Name         : �̹��ֱ��ſ�û��� 
'*  5. Program Desc         : �̹��ֱ��ſ�û��� 
'*  6. Component List       :
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
Dim lblnWinEvent
Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim StartDate, EndDate

	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
    StartDate = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
    EndDate   = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

	frm1.txtFrDt.Text	= StartDate
	frm1.txtToDt.Text	= EndDate
	frm1.txtORGCd.Value = parent.gPurOrg

	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %>
End Sub

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'=========================================================================================================
Function OpenItem1()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	If  Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd1.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtitemcd1.focus
		Exit Function
	Else
		frm1.txtitemcd1.Value    = arrRet(0)
		frm1.txtitemNm1.Value    = arrRet(1)
		frm1.txtitemcd1.focus
	End If
End Function
Function OpenItem2()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	If  Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtitemcd2.Value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtitemcd2.focus
		Exit Function
	Else
		frm1.txtitemcd2.Value    = arrRet(0)
		frm1.txtitemNm2.Value    = arrRet(1)
		frm1.txtitemcd2.focus
	End If
End Function

Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "����"

    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"

    arrHeader(0) = "����"
    arrHeader(1) = "�����"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value=arrRet(0)
		frm1.txtPlantNm.value=arrret(1)
		frm1.txtPlantCd.focus
	End If
	frm1.txtitemcd1.value=""
	frm1.txtitemNm1.value=""
End Function

Function OpenBP()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_SpplCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	arrParam(0) = "����ó"
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	arrParam(4) = " In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "����ó"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "����ó"
    arrHeader(1) = "����ó��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBP(arrRet)
	End If
End Function

Function OpenGrp()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_grpCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	arrParam(0) = "���ִ��׷�"
	arrParam(1) = "B_pur_grp"
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""

	frm1.vspdData.Col=C_Org

	arrParam(4) = "Usage_flg=" & FilterVar("Y", "''", "S") & "  and PUR_ORG =  " & FilterVar(UCase(frm1.vspdData.Text), "''", "S") & " "
	arrParam(5) = "���ִ��׷�"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "���ִ��׷�"
    arrHeader(1) = "���ִ��׷��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGrp(arrRet)
	End If
End Function


Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"
	arrParam(1) = "B_Pur_Org"

	arrParam(2) = Trim(frm1.txtORGCd.Value)
'	arrParam(3) = Trim(frm1.txtORGNm.Value)

	arrParam(4) = ""
	arrParam(5) = "��������"

    arrField(0) = "PUR_ORG"
    arrField(1) = "PUR_ORG_NM"

    arrHeader(0) = "��������"
    arrHeader(1) = "����������"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtOrgCd.focus
		Exit Function
	Else
		frm1.txtOrgCd.Value = arrRet(0)
		frm1.txtOrgNm.Value = arrRet(1)
		frm1.txtOrgCd.focus
	End If
End Function

Function OpenPurGrpCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"
	arrParam(1) = "B_Pur_Grp"

	arrParam(2) = Trim(frm1.txtPurGrpCd1.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)

	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  And PUR_ORG =  " & FilterVar(frm1.txtORGCd.Value, "''", "S") & ""
	arrParam(5) = "���ű׷�"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "���ű׷�"
    arrHeader(1) = "���ű׷��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPurGrpCd1.focus
		Exit Function
	Else
		frm1.txtPurGrpCd1.Value = arrRet(0)
		frm1.txtPurGrpNm1.Value = arrRet(1)
		frm1.txtPurGrpCd1.focus
	End If
End Function

Function OpenPurGrpCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"
	arrParam(1) = "B_Pur_Grp"

	arrParam(2) = Trim(frm1.txtPurGrpCd2.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)

	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  And PUR_ORG =  " & FilterVar(frm1.txtORGCd.Value, "''", "S") & ""
	arrParam(5) = "���ű׷�"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "���ű׷�"
    arrHeader(1) = "���ű׷��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPurGrpCd2.focus
		Exit Function
	Else
		frm1.txtPurGrpCd2.Value = arrRet(0)
		frm1.txtPurGrpNm2.Value = arrRet(1)
		frm1.txtPurGrpCd2.focus
	End If
End Function

'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"
	arrParam(1) = "B_Biz_Partner"
	arrParam(2) = Trim(frm1.txtBpCd2.Value)
	'arrParam(3) = Trim(frm1.txtBpNm.Value)
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "����ó"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "����ó"
    arrHeader(1) = "����ó��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd2.focus
		Exit Function
	Else
		frm1.txtBpCd2.Value = arrRet(0)
		frm1.txtBpNm2.Value = arrRet(1)
		frm1.txtBpCd2.focus
	End If
End Function

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim IntRetCD
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
	arrParam(2) = Trim(frm1.txtPlantCd.value)	'���� 
	arrParam(3) = ""	'��ǰ�� 
	arrParam(4) = ""	'���ֹ�ȣ 
	arrParam(5) = ""	'�߰� Where�� 

'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 	iCalledAspName = AskPRAspName("S3135PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = Trim(arrRet)
		frm1.txtTrackNo.focus
	End If
End Function

'===========================================================================
' Function Name : OpenMrp
' Function Desc : OpenMrp Reference Popup
'===========================================================================
Function OpenMrp()
    Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

    If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "MRP Run��ȣ"				<%' �˾� ��Ī %>
	arrParam(1) = "(select distinct a.order_no A,a.confirm_dt B," & FilterVar("������������", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_EXPL_HISTORY a, m_pur_req b where a.order_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
    arrParam(1) = arrParam(1) & "union "
    arrParam(1) = arrParam(1) & "select distinct  a.run_no A, a.start_dt B ," & FilterVar("MRP����", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_MRP_HISTORY a, m_pur_req b where a.run_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " ) as g" <%' TABLE ��Ī %>


	arrParam(2) = Trim(frm1.txtMRP.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "MRP Run��ȣ"				<%' TextBox ��Ī %>

	arrField(0) = "A"
	arrField(1) = "B"
	arrField(2) = "D"

	arrHeader(0) = "MRP Run��ȣ"				<%' Header��(0)%>
	arrHeader(1) = "����"					<%' Header��(1)%>
	arrHeader(2) = "��������"				<%' Header��(2)%>

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtMRP.focus
		Exit Function
	Else
		frm1.txtMRP.value = arrRet(0)
		frm1.txtMRP.focus
	End If
End Function


Function SetGrp(byval arrRet)
	frm1.vspdData.Col = C_GrpCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col  = C_GrpNm
	frm1.vspdData.Text = arrret(1)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow
End Function

Function SetBP(byval arrRet)
	frm1.vspdData.Col = C_SpplCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col  = C_SpplNm
	frm1.vspdData.Text = arrret(1)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow

	Call SpplChange()
End Function

'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó�� 
'*********************************************************************************************************
 Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtFrDt.Focus
	End if
End Sub

 Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtToDt.Focus
	End if
End Sub
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")

    frm1.txtORGCd.focus
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
 Function FncPrint()
	Call parent.FncPrint()
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
 Function FncFind()
    Call parent.FncFind(parent.C_SINGLE , False)
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
'	Name : ChkKeyField()
'	Description :
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

    Err.Clear

	ChkKeyField = true

	strWhere = " PUR_ORG =  " & FilterVar(frm1.txtORGCd.value, "''", "S") & "  "

	Call CommonQueryRs(" PUR_ORG_NM "," B_PUR_ORG ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	IF Len(lgF0) < 1 Then
		Call DisplayMsgBox("17a003","X","��������","X")
		frm1.txtORGCd.focus
		frm1.txtORGNm.value = ""
		ChkKeyField = False
		Exit Function
	End If

	strDataNm = split(lgF0,chr(11))

	frm1.txtORGNm.value = strDataNm(0)
End Function


'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
 Function FncBtnPrint()
	Dim StrUrl
	Dim lngPos
	Dim intCnt
	dim var1,var2,var3,var4,var5,var6,var7,var8,var9, var10

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    IF ChkKeyField() = False Then
		frm1.txtORGCd.focus
		Exit Function
    End if

    with frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")
			Exit Function
		End if
	End with

	On Error Resume Next

	lngPos = 0

	var1 = UCase(frm1.txtORGCd.value)
	var2 = UCase(frm1.txtPurGrpCd1.value)
	var3 = UCase(frm1.txtPurGrpCd2.value)
	var4 = UCase(frm1.txtPlantCd.value)
	var5 = UCase(frm1.txtitemcd1.value)
	var6 = UCase(frm1.txtitemcd2.value)
	var7 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	var8 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
	var9 = UCase(frm1.txtMRP.value)
	var10 = UCase(frm1.txtTrackingNo.value)


	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	strUrl = strUrl & "pur_org|"	& var1
	strUrl = strUrl & "|pur_grp1|"	& var2
	strUrl = strUrl & "|pur_grp2|"	& var3
	strUrl = strUrl & "|plant|"		& var4
	strUrl = strUrl & "|item1|"		& var5
	strUrl = strUrl & "|item2|"		& var6
	strUrl = strUrl & "|fr_dt|"		& var7
	strUrl = strUrl & "|to_dt|"		& var8
	strUrl = strUrl & "|mrp|"		& var9
	strUrl = strUrl & "|TrackingNo|"		& var10


'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
'----------------------------------------------------------------
	'call FncEBRprint(EBAction, "m2111oa1.ebr", strUrl)
	'2002/12/10
	'ǥ�غ��� 
	'''''''''''''''''''''''''''''''''''''''''''''

	ObjName = AskEBDocumentName("m2111oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------

	Call BtnDisabled(0)
'	With frm1
'		.txtPurGrpNm1.value = ""
'		.txtPurGrpNm2.value = ""
'		.txtPlantNm.value = ""
'		.txtitemNm1.value = ""
'		.txtitemNm2.value = ""
'
'	End with

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()
	On Error Resume Next

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    IF ChkKeyField() = False Then
		frm1.txtORGCd.focus
		Exit Function
    End if

    With frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")
			Exit Function
		End if
	End With

	dim var1,var2,var3,var4,var5,var6,var7,var8,var9, var10
	dim strUrl
	dim arrParam, arrField, arrHeader

	var1 = UCase(frm1.txtORGCd.value)
	var2 = UCase(frm1.txtPurGrpCd1.value)
	var3 = UCase(frm1.txtPurGrpCd2.value)
	var4 = UCase(frm1.txtPlantCd.value)
	var5 = UCase(frm1.txtitemcd1.value)
	var6 = UCase(frm1.txtitemcd2.value)
	var7 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtFrDt.text)
	var8 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
	var9 = UCase(frm1.txtMRP.value)
	var10 = UCase(frm1.txtTrackingNo.value)

	strUrl = strUrl & "pur_org|"	& var1
	strUrl = strUrl & "|pur_grp1|"	& var2
	strUrl = strUrl & "|pur_grp2|"	& var3
	strUrl = strUrl & "|plant|"		& var4
	strUrl = strUrl & "|item1|"		& var5
	strUrl = strUrl & "|item2|"		& var6
	strUrl = strUrl & "|fr_dt|"		& var7
	strUrl = strUrl & "|to_dt|"		& var8
	strUrl = strUrl & "|mrp|"		& var9
	strUrl = strUrl & "|TrackingNo|"		& var10

	'2002/12/10
	'ǥ�غ��� 
	'''''''''''''''''''''''''''''''''''''''''''''
	'call FncEBRPreview("m2111oa1.ebr", strUrl)
	ObjName = AskEBDocumentName("m2111oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)

	Call BtnDisabled(0)
'	With frm1
'		.txtPurGrpNm1.value = ""
'		.txtPurGrpNm2.value = ""
'		.txtPlantNm.value = ""
'		.txtitemNm1.value = ""
'		.txtitemNm2.value = ""
'
'	End with
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
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
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../SChared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�̹��ֱ��ſ�û��Ȳ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtORGCd" ALT="��������" SIZE=10 MAXLENGTH=4  tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenORG()">
													   <INPUT TYPE=TEXT ID="txtORGNm" ALT="��������" NAME="arrCond" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrpCd1" SIZE=10 MAXLENGTH=10 ALT="���ű׷�" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd1()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm1" SIZE=20 MAXLENGTH=18 ALT="���ű׷�" tag="14"> ~ </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrpCd2" SIZE=10 MAXLENGTH=10 ALT="���ű׷�" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd2()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm2" SIZE=20 MAXLENGTH=18 ALT="���ű׷�" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X" ALT="����"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemcd1" SIZE=18 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem1()">
													   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemNm1" SIZE=20 tag="14X"> ~ </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemcd2" SIZE=18 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem2()">
													   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemNm2" SIZE=20 tag="14X"></TD>
							</TR>
							<TR><TD CLASS="TD5" NOWRAP>�ʿ���</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
											<td>~</td>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>MRP Run��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="MRP Run��ȣ" NAME="txtMRP" SIZE=32 MAXLENGTH=12 tag="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMrp"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
								<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No." TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
