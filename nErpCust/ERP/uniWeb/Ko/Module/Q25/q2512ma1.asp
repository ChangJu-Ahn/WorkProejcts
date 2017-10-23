<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2512MA1
'*  4. Program Name         : �˻��Ƿڵ�� 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID ="q2512mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "q2512mb2.asp"										 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID = "q2512mb3.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP1_ID = "Q2111MA1"
Const BIZ_PGM_JUMP2_ID = "Q2211MA1"
Const BIZ_PGM_JUMP3_ID = "Q2311MA1"
Const BIZ_PGM_JUMP4_ID = "Q2411MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop

Dim lgInspStatus	'�˻�������� �ڵ尪 ���� ���� 
Dim lgIFYesNo		'Ÿ ��⿡���� �˻��Ƿ� ���ΰ� ���� ���� 
Dim lgInspClass

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
Dim CompanyYMD	'#####
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)
'--------------- ������ coding part(�������,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	
	lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False                    'Indicates that no value changed
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'��: ����� ���� �ʱ�ȭ 
    lgInspStatus = ""
    lgIFYesNo = ""
    lgInspClass = ""
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
		
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If
		
	If ReadCookie("txtInspReqNo") <> "" Then
		frm1.txtInspReqNo1.Value = ReadCookie("txtInspReqNo")
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""	
	
	frm1.txtInspReqDt.Text = CompanyYMD
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd, lgF0, lgF1, Chr(11))
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant() 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)
		frm1.txtPlantNm.Value    = arrRet(1)
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true		
End Function

'------------------------------------------  OpenInspReqNo()  -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo()        
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtInspReqNo1.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo1.Value)
	Param4 = ""				'�˻�з� 
	Param5 = ""				'�˻�������Ȳ 
	
	iCalledAspName = AskPRAspName("q2512pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q2512pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3,  Param4, Param5), _
		"dialogWidth=820px; dialogHeight=500px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtInspReqNo1.value = arrRet(0)
	End If
	
	frm1.txtInspReqNo1.Focus	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item by Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value = arrRet(0)		
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemSpec.Value = arrRet(2)
		lgBlnFlgChgValue = True
	End If
	
	frm1.txtItemCd.Focus	
	Set gActiveElement = document.activeElement
	OpenItem = True
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & " "			
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtUnit.Value = arrRet(0)		
		lgBlnFlgChgValue = True
	End If
	
	frm1.txtUnit.Focus	
	Set gActiveElement = document.activeElement
End Function

'=============================================  2.5.1 LoadInspection()======================================
'=	Event Name : LoadInspection
'=	Event Desc :
'========================================================================================================
Function LoadInspection()
	Dim intRetCD
	Dim strInspClass
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		If lgInspClass <> frm1.cboInspClassCd.Value Then
			Call DisplayMsgBox("223710", "X", "X", "X") 		'�ڷ���Ʈ�� �˻�з��� ����Ǿ����ϴ�. Ȯ���Ͻʽÿ�.
			Exit Function
		End If 
		strInspClass = lgInspClass
	Else
		Call DisplayMsgBox("223719", "X", "X", "X") 		'��ȸ�� �����Ϳ� ���ؼ��� �˻��� ȭ������ �̵��� �� �ֽ��ϴ�.
		Exit Function
	End If
		
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If		

	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo2.value)
		If Trim(.hStatusFlag.value) = "N" Then
			WriteCookie "IsInspectionRequest", "True"
		End If
		
	End With	
		
	Select Case UCase(strInspClass)
		Case "R"
			PgmJump(BIZ_PGM_JUMP1_ID)
		Case "P"
			PgmJump(BIZ_PGM_JUMP2_ID)
		Case "F"
			PgmJump(BIZ_PGM_JUMP3_ID)
		Case "S"
			PgmJump(BIZ_PGM_JUMP4_ID)
	End Select 
End Function

'=============================================  2.6.1 ProtectByInspClass()======================================
'=	Event Name : ProtectByInspClass
'=	Event Desc :
'========================================================================================================
Sub ProtectByInspClass(Byval sInspClass)
	Select Case sInspClass
		Case "R"
			With ggoOper
				'���԰˻� ���� 
				Call .SetReqAttr(frm1.txtSupplierCd, "N")
				Call .SetReqAttr(frm1.txtSupplierNm, "D")
				Call .SetReqAttr(frm1.txtPRNo, "D")
				Call .SetReqAttr(frm1.txtPONo, "D")
				Call .SetReqAttr(frm1.txtPOSeq, "D")
				Call .SetReqAttr(frm1.txtSLCd1, "D")
				Call .SetReqAttr(frm1.txtSLNm1, "D")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtWcCd, "Q")
				Call .SetReqAttr(frm1.txtWcNm, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforP, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtProdtNo1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtProdtNo2, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforF, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtSLCd2, "Q")
				Call .SetReqAttr(frm1.txtSLNm2, "Q")
				
				'���ϰ˻� ���� 
				Call .SetReqAttr(frm1.txtBPCd, "Q")
				Call .SetReqAttr(frm1.txtDNNo, "Q")
				Call .SetReqAttr(frm1.txtDNSeq, "Q")
			End With
		Case "P"
			With ggoOper
				'���԰˻� ���� 
				Call .SetReqAttr(frm1.txtSupplierCd, "Q")
				Call .SetReqAttr(frm1.txtSupplierNm, "Q")
				Call .SetReqAttr(frm1.txtPRNo, "Q")
				Call .SetReqAttr(frm1.txtPONo, "Q")
				Call .SetReqAttr(frm1.txtPOSeq, "Q")
				Call .SetReqAttr(frm1.txtSLCd1, "Q")
				Call .SetReqAttr(frm1.txtSLNm1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtWcCd, "D")
				Call .SetReqAttr(frm1.txtWcNm, "D")
				Call .SetReqAttr(frm1.txtRoutNoforP, "N")
				Call .SetReqAttr(frm1.txtRoutNoDescforP, "D")
				Call .SetReqAttr(frm1.txtOprNoforP, "N")
				Call .SetReqAttr(frm1.txtOprNoDescforP, "D")
				Call .SetReqAttr(frm1.txtProdtNo1, "D")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtProdtNo2, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforF, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtSLCd2, "Q")
				Call .SetReqAttr(frm1.txtSLNm2, "Q")
				
				
				'���ϰ˻� ���� 
				Call .SetReqAttr(frm1.txtBPCd, "Q")
				Call .SetReqAttr(frm1.txtBPNm, "Q")
				Call .SetReqAttr(frm1.txtDNNo, "Q")
				Call .SetReqAttr(frm1.txtDNSeq, "Q")
			End With
		Case "F"
			With ggoOper
				'���԰˻� ���� 
				Call .SetReqAttr(frm1.txtSupplierCd, "Q")
				Call .SetReqAttr(frm1.txtSupplierNm, "Q")
				Call .SetReqAttr(frm1.txtPRNo, "Q")
				Call .SetReqAttr(frm1.txtPONo, "Q")
				Call .SetReqAttr(frm1.txtPOSeq, "Q")
				Call .SetReqAttr(frm1.txtSLCd1, "Q")
				Call .SetReqAttr(frm1.txtSLNm1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtWcCd, "Q")
				Call .SetReqAttr(frm1.txtWcNm, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforP, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtProdtNo1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtProdtNo2, "D")
				Call .SetReqAttr(frm1.txtRoutNoforF, "N")
				Call .SetReqAttr(frm1.txtRoutNoDescforF, "D")
				Call .SetReqAttr(frm1.txtOprNoforF, "N")
				Call .SetReqAttr(frm1.txtOprNoDescforF, "D")
				Call .SetReqAttr(frm1.txtSLCd2, "D")
				Call .SetReqAttr(frm1.txtSLNm2, "D")
				
				'���ϰ˻� ���� 
				Call .SetReqAttr(frm1.txtBPCd, "Q")
				Call .SetReqAttr(frm1.txtBPNm, "Q")
				Call .SetReqAttr(frm1.txtDNNo, "Q")
				Call .SetReqAttr(frm1.txtDNSeq, "Q")
			End With
		Case "S"
			With ggoOper
				'���԰˻� ���� 
				Call .SetReqAttr(frm1.txtSupplierCd, "Q")
				Call .SetReqAttr(frm1.txtSupplierNm, "Q")
				Call .SetReqAttr(frm1.txtPRNo, "Q")
				Call .SetReqAttr(frm1.txtPONo, "Q")
				Call .SetReqAttr(frm1.txtPOSeq, "Q")
				Call .SetReqAttr(frm1.txtSLCd1, "Q")
				Call .SetReqAttr(frm1.txtSLNm1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtWcCd, "Q")
				Call .SetReqAttr(frm1.txtWcNm, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforP, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtProdtNo1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtProdtNo2, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforF, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtSLCd2, "Q")
				Call .SetReqAttr(frm1.txtSLNm2, "Q")
				
				'���ϰ˻� ���� 
				Call .SetReqAttr(frm1.txtBPCd, "N")
				Call .SetReqAttr(frm1.txtBPNm, "D")
				Call .SetReqAttr(frm1.txtDNNo, "D")
				Call .SetReqAttr(frm1.txtDNSeq, "D")
			End With
		Case Else
			With ggoOper
				'���԰˻� ���� 
				Call .SetReqAttr(frm1.txtSupplierCd, "Q")
				Call .SetReqAttr(frm1.txtSupplierNm, "Q")
				Call .SetReqAttr(frm1.txtPRNo, "Q")
				Call .SetReqAttr(frm1.txtPONo, "Q")
				Call .SetReqAttr(frm1.txtPOSeq, "Q")
				Call .SetReqAttr(frm1.txtSLCd1, "Q")
				Call .SetReqAttr(frm1.txtSLNm1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtWcCd, "Q")
				Call .SetReqAttr(frm1.txtWcNm, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforP, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoforP, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforP, "Q")
				Call .SetReqAttr(frm1.txtProdtNo1, "Q")
				
				'�����˻� ���� 
				Call .SetReqAttr(frm1.txtProdtNo2, "Q")
				Call .SetReqAttr(frm1.txtRoutNoforF, "Q")
				Call .SetReqAttr(frm1.txtRoutNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoforF, "Q")
				Call .SetReqAttr(frm1.txtOprNoDescforF, "Q")
				Call .SetReqAttr(frm1.txtSLCd2, "Q")
				Call .SetReqAttr(frm1.txtSLNm2, "Q")
				
				'���ϰ˻� ���� 
				Call .SetReqAttr(frm1.txtBPCd, "Q")
				Call .SetReqAttr(frm1.txtBPNm, "Q")
				Call .SetReqAttr(frm1.txtDNNo, "Q")
				Call .SetReqAttr(frm1.txtDNSeq, "Q")
			End With
	End Select
End Sub

 '==========================================  3.1.1 Form_load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029	
	Call AppendNumberPlace("6", "3", "0")
	
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 							
	
	Call ggoOper.LockField(Document, "N")		
	Call InitComboBox
	Call SetDefaultVal
	Call ProtectByInspClass("")
	Call InitVariables						
	Call SetToolbar("11101000000011")
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus
	Else
		frm1.txtInspReqNo1.focus
    End If
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub cboInspClassCd_onChange()
	Call ProtectByInspClass(frm1.cboInspClassCd.value)
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtInspReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspReqDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspReqmtDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtInspReqmtDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspReqmtDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspSchdlDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtInspSchdlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspSchdlDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspReqDt_Change
'   Event Desc : 
'=======================================================================================================
Sub txtInspReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspReqmtDt_Change
'   Event Desc : 
'=======================================================================================================
Sub txtInspReqmtDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspSchdlDt_Change
'   Event Desc : 
'=======================================================================================================
Sub txtInspSchdlDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtLotSubNo_Change
'   Event Desc : 
'=======================================================================================================
Sub txtLotSubNo_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtLotSize_Change
'   Event Desc : 
'=======================================================================================================
Sub txtLotSize_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPOSeq_Change
'   Event Desc : 
'=======================================================================================================
Sub txtPOSeq_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDNSeq_Change
'   Event Desc : 
'=======================================================================================================
Sub txtDNSeq_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	FncQuery = False
	
	Dim IntRetCD 
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
    End If
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not ChkField(Document, "1") Then	Exit Function

	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call InitVariables
	
	
	Call ggoOper.LockField(Document, "N")								'��: This function lock the suitable field
   	
   	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function									'��: Query db data
	
	FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	FncNew = False
	
	Dim IntRetCD 
    	
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
	Call ProtectByInspClass("")
	Call SetDefaultVal
	Call InitVariables                                                      '��: Initializes local global variables
	Call SetToolbar("11101000000011")
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus
	Else
		frm1.txtInspReqNo1.focus
    End If
	Set gActiveElement = document.activeElement 
    
	FncNew = True 									'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	FncDelete = False
	
	Dim IntRetCD 
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If
	
	If lgIFYesNo <> "N" Then
		Select Case UCase(lgInspClass)
			Case "R"
				Call DisplayMsgBox("223706", "X", "X", "X")
						
			Case "P", "F"
				Call DisplayMsgBox("223707", "X", "X", "X")
						
			Case "S"
				Call DisplayMsgBox("223708", "X", "X", "X")
		End Select
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True                                                        '��: Processing is OK                   							'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	FncSave = False
	
	Dim IntRetCD 
	   
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		If lgIFYesNo <> "N" Then
			Select Case UCase(lgInspClass)
				Case "R"
					Call DisplayMsgBox("223706", "X", "X", "X")
						
				Case "P", "F"
					Call DisplayMsgBox("223707", "X", "X", "X")
						
				Case "S"
					Call DisplayMsgBox("223708", "X", "X", "X")
			End Select
			Exit Function
		End If
	End If
	    
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then Exit Function
    
    '-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then Exit Function                              '��: Save db data
    
	FncSave = True                                      	                    '��: Processing is OK
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
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim IntRetCD 
    
    FncPrev = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not ChkField(Document, "1") Then	Exit Function
	
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call InitVariables
	
	Call ggoOper.LockField(Document, "N")								'��: This function lock the suitable field
   	
   	'-----------------------
    'Query function call area
    '----------------------- 
    If DbPrev = False Then Exit Function           				'��: Query db data
    
	FncPrev = True
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim IntRetCD 
    
    FncNext = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not ChkField(Document, "1") Then Exit Function

	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call InitVariables
	
	Call ggoOper.LockField(Document, "N")								'��: This function lock the suitable field
   	
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbNext = False Then Exit Function           				'��: Query db data
    
	FncNext = False
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = False
	
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If
	
	Call InitVariables
	Call ggoOper.SetReqAttr(frm1.txtInspReqNo2, "D")
	
	frm1.txtInspReqNo1.value = ""
	frm1.txtInspReqNo2.value = ""
	frm1.txtInspStatus.value = ""
	
	lgIntFlgMode = Parent.OPMD_CMODE														'��: Indicates that current mode is Crate mode
	lgBlnFlgChgValue = True
	
	Call SetToolbar("11101000000011")
	
	frm1.txtInspReqNo2.focus
	Set gActiveElement = document.activeElement  
	
	FncCopy = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
    Call parent.FncExport(Parent.C_SINGLE)											'��: ȭ�� ���� 
    FncExcel = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE , False)                                   '��:ȭ�� ����, Tab ���� 
    FncFind = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = False
	
	Dim IntRetCD
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If

    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '��: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1)
       
    strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value) _
							& "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbPrev()
    DbPrev = False                                                         '��: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value)	_
							& "&PrevNextFlg=" & "P"									'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
	DbPrev = True
End Function

'========================================================================================
' Function Name : DbNext
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbNext()
    DbNext = False                                                         '��: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value) _
							& "&PrevNextFlg=" & "N"									'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
	DbNext = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	DbQueryOk = False
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    Call ProtectByInspClass(frm1.cboInspClassCd.value)
    
    If lgInspStatus = "N" Then				'�̰˻� �ϰ�� 
		Call SetToolbar("11111000111111")
	Else									'�˻����̳� Release�Ϸ��� ��� 
		Call SetToolbar("11100000111111")
	End If
    
    frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement 
    
    DbQueryOk = True
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 
	DbSave = False															'��: Processing is NG
	
	LayerShowHide(1)
		
	With frm1
		.txtFlgMode.value = lgIntFlgMode
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With
	
    DbSave = True                                                           '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()		
	DbSaveOk = False
	frm1.txtInspReqNo1.value = frm1.txtInspReqNo2.value 
	lgBlnFlgChgValue = False
    Call MainQuery()
    DbSaveOk = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = False
	
	Call LayerShowHide(1)
	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value) _
							& "&txtInspReqNo=" & Trim(frm1.txtInspReqNo1.value)
	
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True			                                                   			'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()
	DbDeleteOk = false
	lgBlnFlgChgValue = False												'��: ���� ������ ���� ���� 
	Call MainNew()
	DbDeleteOk = true
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- TAB, REFERENCE AREA START -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>�˻��Ƿ� ���</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- TAB, REFERENCE AREA END -->
	<!-- CONDITION/CONTENT AREA START -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<!-- CONDITION AREA START-->
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantPopup ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X">
									</TD>
									<TD CLASS="TD5" NOWRAP>�˻��Ƿڹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtInspReqNo1" SIZE="20" MAXLENGTH="18" ALT="�˻��Ƿڹ�ȣ" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnInspReqNoPopup ONCLICK=vbscript:OpenInspReqNo() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<!-- CONDITION AREA END-->
				<!-- CONTENT AREA START-->
				<TR>
					<TD HEIGHT=* WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_50%>>
							<!-- ����κ� START -->
							<TR>
								<TD Colspan=4>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>����</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�˻��Ƿڹ�ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo2" SIZE="20" MAXLENGTH="18" ALT="�˻��Ƿڹ�ȣ" TAG="25XXXU" ></TD>
												<td CLASS="TD5" NOWPAP>�˻�з�</TD>
												<td CLASS="TD6" NOWPAP><SELECT NAME="cboInspClassCd" ALT="�˻�з�" STYLE="WIDTH: 150px" TAG="22"><OPTION VALUE="" selected></OPTION></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>ǰ��</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE="20" MAXLENGTH="18" ALT="ǰ��" TAG="22XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnItemPopup ONCLICK=vbscript:OpenItem() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtItemNm" TAG="24"></TD>
												<TD CLASS="TD5" NOWPAP></TD>
												<TD CLASS="TD6" NOWPAP></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�԰�</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE="40" ALT="�԰�" TAG="24"></TD>
												<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE="25" MAXLENGTH="25" ALT="Tracking No." TAG="21XXXU" ></TD>
											</TR>
											<TR>
							                	<TD CLASS="TD5" NOWRAP>��Ʈ��ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE="20" MAXLENGTH="25" ALT="��Ʈ��ȣ" TAG="21XXXU">&nbsp;
							                		<script language =javascript src='./js/q2512ma1_txtLotSubNo_txtLotSubNo.js'></script>
												</TD>
							                	<TD CLASS="TD5" NOWRAP>��Ʈũ��</TD>        
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2512ma1_txtLotSize_txtLotSize.js'></script>&nbsp;<INPUT TYPE=TEXT NAME="txtUnit" SIZE="5" MAXLENGTH="3" TAG="22XXXU"  ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnitPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit()">
												</TD>
							                </TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�˻��Ƿ���</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2512ma1_txtInspReqDt_txtInspReqDt.js'></script>
												</TD>
												<TD CLASS="TD5" NOWRAP>�˻�䱸��</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2512ma1_txtInspReqmtDt_txtInspReqmtDt.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�˻��ȹ��</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q2512ma1_txtInspSchdlDt_txtInspSchdlDt.js'></script>
												</TD>
												<TD CLASS="TD5" NOWRAP>�˻��������</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspStatus" SIZE="20" MAXLENGTH="40" ALT="�˻��������" TAG="24"></TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<!-- ����κ� END -->
							<!-- ����/�����˻� START -->
							<TR>
								<!-- ���԰˻� START -->
								<TD WIDTH=50% VALIGN=TOP>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>���԰˻�</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>����ó</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE="10" MAXLENGTH="10" ALT="����ó" TAG="22XXXU">&nbsp;<INPUT NAME="txtSupplierNm" TAG="21">
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�԰��ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPRNo" SIZE="20" MAXLENGTH="18" ALT="�԰��ȣ" TAG="21XXXU" ></TD>
											</TR>
											<TR>
							                	<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPONo" SIZE="20" MAXLENGTH="18" ALT="���ֹ�ȣ" TAG="21XXXU">&nbsp;
							                		<script language =javascript src='./js/q2512ma1_txtPOSeq_txtPOSeq.js'></script>
												</TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWRAP>â��</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtSLCd1" SIZE="10" MAXLENGTH="7" ALT="â��" TAG="21XXXU">&nbsp;<INPUT NAME="txtSLNm1" TAG="21">
												</TD>
											</TR>
							                <TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<!-- ���԰˻� END -->
								<!-- �����˻� START -->
								<TD WIDTH=50% VALIGN=TOP>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>�����˻�</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�����</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNoforP" SIZE="20" MAXLENGTH="20" ALT="�����" TAG="22XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutNoDescforP" SIZE=20 MAXLENGTH=20 ALT="����ü���" tag="21"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>����</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNoforP" SIZE="5" MAXLENGTH="3" ALT="����" TAG="22XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOprNoDescforP" SIZE=20 MAXLENGTH=20 ALT="�����۾���" tag="21"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�۾���</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtWcCd" SIZE="10" MAXLENGTH="7" ALT="�۾���" TAG="21XXXU">&nbsp;<INPUT NAME="txtWcNm" TAG="21">
												</TD>
											</TR>
											<TR>
							                	<TD CLASS="TD5" NOWRAP>����������ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProdtNo1" SIZE="20" MAXLENGTH="16" ALT="����������ȣ" TAG="21XXXU"></TD>
							                </TR>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<!-- �����˻� END -->
							</TR>
							<!-- ����/�����˻� END -->
							<!-- ����/���ϰ˻� START -->
							<TR>
								<!-- �����˻� START -->
								<TD WIDTH=50% VALIGN=TOP>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>�����˻�</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
							                	<TD CLASS="TD5" NOWRAP>����������ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProdtNo2" SIZE="20" MAXLENGTH="16" ALT="����������ȣ" TAG="21XXXU"></TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWRAP>�����</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNoforF" SIZE="20" MAXLENGTH="20" ALT="�����" TAG="22XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutNoDescforF" SIZE=20 MAXLENGTH=20 ALT="����ü���" tag="21"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>����</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNoforF" SIZE="5" MAXLENGTH="3" ALT="����" TAG="22XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOprNoDescforF" SIZE=20 MAXLENGTH=20 ALT="�����۾���" tag="21"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>â��</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtSLCd2" SIZE="10" MAXLENGTH="7" ALT="â��" TAG="21XXXU">&nbsp;<INPUT NAME="txtSLNm2" TAG="21">
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<!-- �����˻� END -->
								<!-- ���ϰ˻� START -->
								<TD WIDTH=50% VALIGN=TOP>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>���ϰ˻�</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE=TEXT NAME="txtBPCd" SIZE="10" MAXLENGTH="10" ALT="�ŷ�ó" TAG="22XXXU">&nbsp;<INPUT NAME="txtBPNm" TAG="21">
												</TD>
											</TR>
											<TR>
							                	<TD CLASS="TD5" NOWRAP>���Ϲ�ȣ</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDNNo" SIZE="20" MAXLENGTH="18" ALT="���Ϲ�ȣ" TAG="21XXXU">&nbsp;
							                		<script language =javascript src='./js/q2512ma1_txtDNSeq_txtDNSeq.js'></script>
												</TD>
							                </TR>
							                <TR>
												<TD CLASS="TD5" NOWPAP HEIGHT=5></TD>
												<TD CLASS="TD6" NOWPAP HEIGHT=5></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<!-- ���ϰ˻� END -->
							</TR>
							<!-- ����/���ϰ˻� END -->
						</TABLE>
					</TD>
				</TR>
				<!-- CONTENT AREA END-->
			</TABLE>
		</TD>
	</TR>
	<!-- CONDITION/CONTENT AREA END -->
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
	      	<TD WIDTH="100%" >
	      		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	        		<TR>
	        			<TD WIDTH=10>&nbsp;</TD>
	        			<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspection">�˻���</A></TD>
	        		</TR>
	      		</TABLE>
	      	</TD>
         </TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="hStatusFlag" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
