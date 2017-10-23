<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : B1261MA1
'*  4. Program Name         : �ŷ�ó��� 
'*  5. Program Desc         : �ŷ�ó��� 
'*  6. Comproxy List        : PB5CS40.dll, PB5CS41.dll
'*  7. Modified date(First) : 2000/03/24
'*  8. Modified date(Last)  : 2000/08/22
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Sonbumyeol
'* 11. Comment              :	
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									
'*                            this mark(��) Means that "may  change"									
'*                            this mark(��) Means that "must change"									
'* 13. History              : 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate
iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "b1261mb1.asp"												'��: Head Query �����Ͻ� ���� ASP�� 

Const BIZ_PGM_JUMP_ID1 = "b1263ma1"	'������̷µ�� 
Const BIZ_PGM_JUMP_ID2 = "b1261ma8"	'�ŷ�ó��ȸ 
Const BIZ_PGM_JUMP_ID3 = "b1262ma8"	'�ŷ�ó������ȸ 

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2
Const TAB3 = 3																		'��: Tab�� ��ġ 
Const TAB4 = 4

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim IsOpenPop						' Popup
Dim gSelframeFlg 
Dim arrCollectVatType

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()

	With frm1

		.txtConBp_cd.focus
		<% '��뿩�� %>
		.rdoUsage_flag1.checked = True
		.txtRadioFlag.value = .rdoUsage_flag1.value 
		<% '�ΰ������Կ��� %>
		.rdoVATinc_1.checked = True
		.txtRadioVATinc.value = .rdoVATinc_1.value 
		<% '���Ű������� %>
		.rdoCredit_N.checked = True
		.txtRadioCredit.value = .rdoCredit_N.value 
		<% '�ΰ�������� %>
		.rdoVATcalc_Y.checked = True
		.txtRadioVATcalc.value = .rdoVATcalc_Y.value 
		
		<% '������������� %>
		.rdoReservePrice_N.checked = True
		.txtRadioDepositPrice.value = .rdoReservePrice_N.value 
	   
	    <% '������ ����ȭ�� %>
		.txtCash_Rate.Text = 0
		
		<% '���ݽŰ����� flag ���� %>
		.chkBpTypeT.disabled = True
		
		<% '�系�ܱ��� %>
		.rdoIn_out2.checked = True
		
		'��ǰ�ð˻��� 
		.rdoSoldInspectA.checked = True
		
	End With

End Sub

'========================================================================================================= 
<% '== ��� == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	<% '~~~ ù��° Tab %>
	gSelframeFlg = TAB1
	
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	<% '~~~ �ι�° Tab %>
	gSelframeFlg = TAB2

End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)	<% '~~~ ����° Tab %>
	gSelframeFlg = TAB3

End Function

Function ClickTab4()

	If gSelframeFlg = TAB4 Then Exit Function
	Call changeTabs(TAB4)	<% '~~~ �׹�° Tab %>
	gSelframeFlg = TAB4

End Function

'========================================================================================================= 
Function OpenConBp_cd()
	Dim arrRet
	Dim iCalledAspName

	On Error Resume Next

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("b1261pa1")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "b1261pa1", "x")
		IsOpenPop = False
		Exit Function
	End if

	frm1.txtConBp_cd.focus 
	
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, Trim(frm1.txtConBp_cd.value)),"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
		Call SetConBp_cd(arrRet)		
	End If	
End Function

'========================================================================================================= 
Function OpenMinor(ByVal iMinor)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iMinor
	Case 0												<%' ���� %>
		If lgIntFlgMode = parent.OPMD_UMODE Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtInd_Class.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9003", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "����"						<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "����"					<%' Header��(0)%>
	    arrHeader(1) = "���¸�"						<%' Header��(1)%>

		frm1.txtInd_Class.focus 
	Case 1												<%' ���� %>
		If lgIntFlgMode = parent.OPMD_UMODE Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtInd_Type.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9002", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "����"						<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "����"					<%' Header��(0)%>
	    arrHeader(1) = "������"						<%' Header��(1)%>

		frm1.txtInd_Type.focus 
	Case 2												<%' ��۹�� %>

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtTrans_Meth.value)	<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9009", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "��۹��"					<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "��۹��"				<%' Header��(0)%>
	    arrHeader(1) = "��۹����"					<%' Header��(1)%>

		frm1.txtTrans_Meth.focus 
	Case 3												<%' ��ü�򰡵�� %>

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtBp_Grade.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9010", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "��ü�򰡵��"				<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "��ü��"				<%' Header��(0)%>
	    arrHeader(1) = "��ü�򰡵�޸�"				<%' Header��(1)%>

		frm1.txtBp_Grade.focus 
	Case 4												<%' �ŷ����� %>

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtDeal_Type.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "�Ǹ�����"					<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "�Ǹ�����"				<%' Header��(0)%>
	    arrHeader(1) = "�Ǹ�������"					<%' Header��(1)%>

		frm1.txtDeal_Type.focus 
	Case 5												<%' �ΰ������� %>

		arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"	<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtvat_Type.value)				<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & "  And Config.MAJOR_CD = Minor.MAJOR_CD" _
						& " And Config.MINOR_CD = Minor.MINOR_CD" _
						& " And Config.SEQ_NO = 1"				<%' Where Condition%>
		arrParam(5) = "VAT����"							<%' TextBox ��Ī %>
		
	    arrField(0) = "Minor.MINOR_CD"							<%' Field��(0)%>
	    arrField(1) = "Minor.MINOR_NM"							<%' Field��(1)%>
	    arrField(2) = "Config.REFERENCE"						<%' Field��(2)%>
	    	    
	    arrHeader(0) = "VAT����"							<%' Header��(0)%>
	    arrHeader(1) = "VAT������"							<%' Header��(1)%>
		arrHeader(2) = "VAT��"							<%' Header��(2)%>

		frm1.txtvat_Type.focus 	    
	Case 6												<%' ������� %>

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtPay_meth.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "�������"					<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "�������"					<%' Header��(0)%>
	    arrHeader(1) = "���������"						<%' Header��(1)%>

		frm1.txtPay_meth.focus 
	Case 8												<%' �ŷ�ó�з� %>

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtBp_Group.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9014", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "�ŷ�ó�з�"					<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "�ŷ�ó�з�"					<%' Header��(0)%>
	    arrHeader(1) = "�ŷ�ó�з���"				<%' Header��(1)%>

		frm1.txtBp_Group.focus 
	Case 9												<%' �������(����)%>

		arrParam(1) = "B_MINOR"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtPay_meth_Pur.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "�������"					<%' TextBox ��Ī %>
		
	    arrField(0) = "MINOR_CD"						<%' Field��(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "�������"					<%' Header��(0)%>
	    arrHeader(1) = "���������"						<%' Header��(1)%>
	    
	    frm1.txtPay_meth_Pur.focus 
	End Select
    
	arrParam(0) = arrParam(5)							<%' �˾� ��Ī %>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinor(arrRet,iMinor)
	End If	
End Function

'========================================================================================================= 
Function OpenCardCO()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ī���"						<%' �˾� ��Ī %>
	arrParam(1) = "B_CARD_CO"		<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtCardCoCd.value)		<%' Code Condition%>

	arrParam(4) = "PAY_CARD_FG = " & FilterVar("Y", "''", "S") & " "			<%' Where Condition%>
	arrParam(5) = "ī���"					<%' TextBox ��Ī %>
		
	arrField(0) = "CARD_CO_CD"						<%' Field��(0)%>
	arrField(1) = "CARD_CO_NM"					<%' Field��(1)%>
    
	arrHeader(0) = "ī���"						<%' Header��(0)%>
	arrHeader(1) = "ī����"						<%' Header��(1)%>

	frm1.txtCardCoCd.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCardCO(arrRet)
	End If	

End Function

'========================================================================================================= 
Function SetCardCO(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtCardCoCd.value = arrRet(0)
	frm1.txtCardCoCdNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End If

End Function


'========================================================================================================= 
Function OpenBankCo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"						<%' �˾� ��Ī %>
	arrParam(1) = "B_BANK"						<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtBankCo.value)		<%' Code Condition%>

	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "����"					<%' TextBox ��Ī %>
		
	arrField(0) = "BANK_CD"						<%' Field��(0)%>
	arrField(1) = "BANK_NM"						<%' Field��(1)%>
    
    arrHeader(0) = "����"						<%' Header��(0)%>
    arrHeader(1) = "�����"						<%' Header��(1)%>

	frm1.txtBankCo.focus 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBankCO(arrRet)
	End If	

End Function

'========================================================================================================= 
Function SetBankCO(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtBankCO.value = arrRet(0)
	frm1.txtBankCONm.value = arrRet(1)
	lgBlnFlgChgValue = True
End If

End Function

'========================================================================================================= 
Function OpenBankAcctNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	
		arrParam(0) = "���¹�ȣ"						<%' �˾� ��Ī %>
		arrParam(1) = "B_BANK A,B_BANK_ACCT B"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtBankAcctNo.value)		<%' Code Condition%>

		arrParam(4) = "A.BANK_CD = B.BANK_CD AND B.BANK_CD like   " & FilterVar(Trim(frm1.txtBankCo.value), "'%'", "S") & ""							<%' Where Condition%>
		arrParam(5) = "���¹�ȣ"						<%' TextBox ��Ī %>
		
	    arrField(0) = "B.BANK_ACCT_NO" 						<%' Field��(0)%>
	    arrField(1) = "A.BANK_NM"						<%' Field��(0)%>
	    arrField(2) = "A.BANK_CD"						<%' Field��(1)%>
	
    
		arrHeader(0) = "���¹�ȣ"							<%' Header��(0)%>
		arrHeader(1) = "�����"						<%' Header��(0)%>
		arrHeader(2) = "����"				<%' Header��(1)%>
			
		frm1.txtBankAcctNo.focus 

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBankAcctNo(arrRet)
		End If	
	
	
	
End Function

'========================================================================================================= 
Function SetBankAcctNo(Byval arrRet)

	If arrRet(0) <> "" Then 
	    frm1.txtBankCO.value = arrRet(2)
	    frm1.txtBankCONm.value = arrRet(1)
		frm1.txtBankAcctNo.value = arrRet(0)
		lgBlnFlgChgValue = True
	End If

End Function

'========================================================================================================= 
Function OpenBiz_Grp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����׷�"						<%' �˾� ��Ī %>

	arrParam(1) = "B_SALES_GRP"						<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtBiz_Grp.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "				<%' Where Condition%>
	arrParam(5) = "�����׷�"					<%' TextBox ��Ī %>
		
	arrField(0) = "SALES_GRP"						<%' Field��(0)%>
	arrField(1) = "SALES_GRP_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "�����׷�"					<%' Header��(0)%>
    arrHeader(1) = "�����׷��"					<%' Header��(1)%>

	frm1.txtBiz_Grp.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBiz_Grp(arrRet)
	End If	

End Function						

'========================================================================================================= 
Function OpenTo_Grp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ݱ׷�"						<%' �˾� ��Ī %>
	arrParam(1) = "B_SALES_GRP"						<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtTo_Grp.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "				<%' Where Condition%>
	arrParam(5) = "���ݱ׷�"					<%' TextBox ��Ī %>
		
	arrField(0) = "SALES_GRP"						<%' Field��(0)%>
	arrField(1) = "SALES_GRP_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "���ݱ׷�"						<%' Header��(0)%>
    arrHeader(1) = "���ݱ׷��"						<%' Header��(1)%>

	frm1.txtTo_Grp.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTo_Grp(arrRet)
	End If	

End Function

'========================================================================================================= 
Function OpenPur_Grp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"						<%' �˾� ��Ī %>
	arrParam(1) = "B_PUR_GRP"						<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtPur_Grp.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "USAGE_FLG = " & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
	arrParam(5) = "���ű׷�"					<%' TextBox ��Ī %>
		
	arrField(0) = "PUR_GRP"							<%' Field��(0)%>
	arrField(1) = "PUR_GRP_NM"						<%' Field��(1)%>
    
    arrHeader(0) = "���ű׷�"						<%' Header��(0)%>
    arrHeader(1) = "���ű׷��"						<%' Header��(1)%>

	frm1.txtPur_Grp.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPur_Grp(arrRet)
	End If	

End Function

'========================================================================================================= 
Function OpenTax_Biz()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If frm1.txtTaxBizAreaCd.readOnly = True Then
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = "���ݽŰ�����"				' �˾� ��Ī 
	
	arrParam(1) = "B_TAX_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtTaxBizAreaCd.value)	' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "���ݽŰ�����"				' TextBox ��Ī 
					
	arrField(0) = "TAX_BIZ_AREA_CD"					' Field��(0)
	arrField(1) = "TAX_BIZ_AREA_NM"					' Field��(1)
				    
	arrHeader(0) = "���ݽŰ�����"				' Header��(0)
	arrHeader(1) = "���ݽŰ������"			' Header��(1)

	frm1.txtTaxBizAreaCd.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTax_Biz(arrRet)
	End If	

End Function

'========================================================================================================= 
Function OpenEtc(ByVal iMinor)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iMinor
	Case 0												<%' ���� %>

		If lgIntFlgMode = parent.OPMD_UMODE Then
			IsOpenPop = False
			Exit Function
		End If
	
		arrParam(1) = "B_COUNTRY"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtContry_cd.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = ""								<%' Where Condition%>
		arrParam(5) = "����"						<%' TextBox ��Ī %>
		
	    arrField(0) = "COUNTRY_CD"						<%' Field��(0)%>
	    arrField(1) = "COUNTRY_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "����"						<%' Header��(0)%>
	    arrHeader(1) = "������"						<%' Header��(1)%>

		frm1.txtContry_cd.focus 
	Case 1												<%' ���� %>

		If lgIntFlgMode = parent.OPMD_UMODE Then
			IsOpenPop = False
			Exit Function
		End If

		If Trim(frm1.txtContry_cd.value) = "" Then
			'Call parent.DisplayMsgBox("203150","X","X","X")
			MsgBox "������ ���� �Է��ϼ���", vbInformation, parent.gLogoName
			frm1.txtContry_cd.focus 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(1) = "B_PROVINCE"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtProvince_cd.value)	<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = "COUNTRY_CD =  " & FilterVar(frm1.txtContry_cd.value, "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = "����"						<%' TextBox ��Ī %>
		
	    arrField(0) = "PROVINCE_CD"						<%' Field��(0)%>
	    arrField(1) = "PROVINCE_NM"						<%' Field��(1)%>
	    
	    arrHeader(0) = "����"						<%' Header��(0)%>
	    arrHeader(1) = "�����"						<%' Header��(1)%>

		frm1.txtProvince_cd.focus 
	Case 2												<%' ȭ�� %>

		arrParam(1) = "B_CURRENCY"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtCurrency.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = ""								<%' Where Condition%>
		arrParam(5) = "ȭ��"						<%' TextBox ��Ī %>
		
	    arrField(0) = "CURRENCY"						<%' Field��(0)%>
	    arrField(1) = "CURRENCY_DESC"					<%' Field��(1)%>
	    
	    arrHeader(0) = "ȭ��"						<%' Header��(0)%>
	    arrHeader(1) = "ȭ���"						<%' Header��(1)%>

		frm1.txtCurrency.focus 
	End Select

	arrParam(0) = arrParam(5)							<%' �˾� ��Ī %>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetEtc(arrRet,iMinor)
	End If	
End Function

'========================================================================================================= 
Function OpenZip()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtContry_cd.value) = "" Then
		'Call parent.DisplayMsgBox("203150","X","X","X")
		MsgBox "������ ���� �Է��ϼ���", vbInformation, parent.gLogoName
		frm1.txtContry_cd.focus 
		IsOpenPop = False			
		Exit Function
	End IF

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtZIP_cd.value)
	arrParam(1) = ""
	arrParam(2) = Trim(frm1.txtContry_cd.value)

	frm1.txtZIP_cd.focus 
	
	arrRet = window.showModalDialog("../../comasp/ZipPopup.asp", Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetZip(arrRet)
	End If	
			
End Function

'========================================================================================================= 
Sub SetZip(arrRet)
	With frm1
		.txtZIP_cd.value = arrRet(0)
		.txtADDR1.value = arrRet(1)
		.txtADDR2.value = ""
		.txtProvince_cd.value = ""
		.txtProvince_nm.value = ""
		lgBlnFlgChgValue = True
	End With
End Sub

'========================================================================================================= 
Function OpenContentPopUp(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case "CreditGrp"	<%' ���Ű����׷� %>

		If frm1.txtCredit_grp.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "S_CREDIT_LIMIT"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtCredit_grp.Value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "���Ű����׷�"					<%' TextBox ��Ī %>
	
		arrField(0) = "CREDIT_GRP"							<%' Field��(0)%>
		arrField(1) = "CREDIT_GRP_NM"						<%' Field��(1)%>
    
		arrHeader(0) = "���Ű����׷�"					<%' Header��(0)%>
		arrHeader(1) = "���Ű����׷��"					<%' Header��(1)%>
		
		frm1.txtCredit_grp.focus 
		
	Case "BpGrade"	<%' ��ü�򰡵�� %>
		arrParam(1) = "B_MINOR"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtBp_Grade.Value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9010", "''", "S") & " "					<%' Where Condition%>
		arrParam(5) = "��ü�򰡵��"					<%' TextBox ��Ī %>
	
		arrField(0) = "MINOR_CD"							<%' Field��(0)%>
		arrField(1) = "MINOR_NM"							<%' Field��(1)%>
    
		arrHeader(0) = "��ü�򰡵��"					<%' Header��(0)%>
		arrHeader(1) = "��ü�򰡵�޸�"					<%' Header��(1)%>
		
		frm1.txtBp_Grade.focus 
		
	Case "PayTypeSales"		<%' ���������(����)%>

		If Trim(frm1.txtPay_meth.value) = "" Then
			Call DisplayMsgBox("205152","x",frm1.txtPay_meth.alt,"x")
			'MsgBox "��������� ���� �Է��ϼ���!"
			frm1.txtPay_meth.focus 
			IsOpenPop = False			
			Exit Function
		End If

		arrParam(1) = "B_CONFIGURATION A,B_MINOR B, B_CONFIGURATION C "			<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtPay_type.Value)			<%' Code Condition%>

		If Len(Trim(frm1.txtPay_meth.value)) Then

			CALL chkSaveValue()
		
			If Trim(frm1.txtBp_Type.value) = "CS" Then		'���Ը��� 
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("R", "''", "S") & " ) " 	<%' Where Condition%>
			
			Elseif Trim(frm1.txtBp_Type.value) = "C" Then	'����	
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("R", "''", "S") & " ) "	<%' Where Condition%>
			
			Elseif Trim(frm1.txtBp_Type.value) = "S" Then	'���� 
			
			Elseif Trim(frm1.txtBp_Type.value) = "*" Then	'���Ը��� ǥ��ȵȰ�			
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("R", "''", "S") & " )  " 	<%' Where Condition%>
		    
		    Elseif Trim(frm1.txtBp_Type.value) = "T" Then	'���ݽŰ�����			
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND C.REFERENCE = (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("R", "''", "S") & " )  " 	<%' Where Condition%>

			End If

		End If

		arrParam(5) = "���������"						<%' TextBox ��Ī %>
	
		arrField(0) = "A.REFERENCE"							<%' Field��(0)%>
		arrField(1) = "B.MINOR_NM"							<%' Field��(1)%>
    
		arrHeader(0) = "���������"						<%' Header��(0)%>
		arrHeader(1) = "�����������"					<%' Header��(1)%>

		frm1.txtPay_type.focus 
		
	Case "PayTypePur"		<%' ���������(����)%>

		If Trim(frm1.txtPay_meth_Pur.value) = "" Then
			Call DisplayMsgBox("205152","x",frm1.txtPay_meth_Pur.alt,"x")
			'MsgBox "��������� ���� �Է��ϼ���!"
			frm1.txtPay_meth_Pur.focus 
			IsOpenPop = False			
			Exit Function
		End If

		arrParam(1) = "B_CONFIGURATION A,B_MINOR B, B_CONFIGURATION C "			<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtPay_type_Pur.Value)			<%' Code Condition%>

		If Len(Trim(frm1.txtPay_meth_Pur.value)) Then

			CALL chkSaveValue()
		
			If Trim(frm1.txtBp_Type.value) = "CS" Then		'���Ը��� 
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth_Pur.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("P", "''", "S") & " )  " 	<%' Where Condition%>
			
			Elseif Trim(frm1.txtBp_Type.value) = "C" Then	'����	
			
			Elseif Trim(frm1.txtBp_Type.value) = "S" Then	'���� 
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth_Pur.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("P", "''", "S") & " ) "	<%' Where Condition%>
			
		    Elseif Trim(frm1.txtBp_Type.value) = "*" Then	'���Ը��� ǥ��ȵȰ�			
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth_Pur.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("P", "''", "S") & " )  " 	<%' Where Condition%>
		    Elseif Trim(frm1.txtBp_Type.value) = "T" Then	'���ݽŰ�����			
				arrParam(4) = "A.REFERENCE = B.MINOR_CD AND B.MAJOR_CD=" & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "  " _
			& "AND A.MINOR_CD =  " & FilterVar(frm1.txtPay_meth_Pur.value, "''", "S") & " AND A.SEQ_NO > 1 AND A.REFERENCE = C.MINOR_CD AND C.SEQ_NO = " & FilterVar("1", "''", "S") & "   AND (C.REFERENCE = " & FilterVar("RP", "''", "S") & "  OR C.REFERENCE = " & FilterVar("P", "''", "S") & " )  " 	<%' Where Condition%>

			End If

		End If

		arrParam(5) = "���������"						<%' TextBox ��Ī %>
	
		arrField(0) = "A.REFERENCE"							<%' Field��(0)%>
		arrField(1) = "B.MINOR_NM"							<%' Field��(1)%>
    
		arrHeader(0) = "���������"						<%' Header��(0)%>
		arrHeader(1) = "�����������"					<%' Header��(1)%>
		
		frm1.txtPay_type_Pur.focus 
		
	End Select

	arrParam(0) = arrParam(5)								<%' �˾� ��Ī %>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetContentPopUp(arrRet, iWhere)
	End If	
	
End Function

'========================================================================================================= 
Function SetConBp_cd(Byval arrRet)

	frm1.txtConBp_cd.value = arrRet(0)		
	frm1.txtConBp_nm.value = arrRet(1)
	Call MainQuery()
     
	lgBlnFlgChgValue = True

End Function

'========================================================================================================= 
Function SetMinor(Byval arrRet,ByVal iMinor)

If arrRet(0) <> "" Then 
	Select Case iMinor
	Case 0												<%' ���� %>
		frm1.txtInd_Class.value = arrRet(0)
		frm1.txtInd_ClassNm.value = arrRet(1)
	Case 1												<%' ���� %>
		frm1.txtInd_Type.value = arrRet(0)
		frm1.txtInd_TypeNm.value = arrRet(1)
	Case 2												<%' ��۹�� %>
		frm1.txtTrans_Meth.value = arrRet(0)
		frm1.txtTrans_Meth_nm.value = arrRet(1)
	Case 3												<%' ��ü�򰡵�� %>
		frm1.txtBp_Grade.value = arrRet(0)
		frm1.txtBp_Grade_nm.value = arrRet(1)
	Case 4												<%' �ŷ����� %>
		frm1.txtDeal_Type.value = arrRet(0)
		frm1.txtDeal_Type_nm.value = arrRet(1)
	Case 5												<%' �ΰ������� %>
		frm1.txtvat_Type.value = arrRet(0)
		frm1.txtvat_Type_nm.value = arrRet(1)
		frm1.txtvat_Rate.value = UNIFormatNumber(arrRet(2), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)		
	Case 6												<%' ������� %>
		frm1.txtPay_meth.value = arrRet(0)
		frm1.txtPay_meth_nm.value = arrRet(1)
	Case 7												<%' �ŷ�ó���� %>
		frm1.txtBp_Type.value = arrRet(0)
		frm1.txtBp_Type_Nm.value = arrRet(1)	
	Case 8												<%' �ŷ�ó���� %>
		frm1.txtBp_Group.value = arrRet(0)
		frm1.txtBp_Group_Nm.value = arrRet(1)	
	Case 9												<%' ������� %>
		frm1.txtPay_meth_Pur.value = arrRet(0)
		frm1.txtPay_meth_Pur_nm.value = arrRet(1)
	End Select
	lgBlnFlgChgValue = True
End If

End Function

'========================================================================================================= 
Function SetBiz_Grp(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtBiz_Grp.value = arrRet(0)
	frm1.txtBiz_Grp_nm.value = arrRet(1)
	lgBlnFlgChgValue = True
End If

End Function


'========================================================================================================= 
Function SetTo_Grp(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtTo_Grp.value = arrRet(0)
	frm1.txtTo_Grp_nm.value = arrRet(1)
	lgBlnFlgChgValue = True
End If

End Function


'========================================================================================================= 
Function SetPur_Grp(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtPur_Grp.value = arrRet(0)
	frm1.txtPur_Grp_nm.value = arrRet(1)
	lgBlnFlgChgValue = True
End If

End Function

'========================================================================================================= 
Function SetTax_BiZ(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtTaxBizAreaCd.value = arrRet(0)
	frm1.txtTaxBizAreaNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End If

End Function

'========================================================================================================= 
Function SetEtc(Byval arrRet,ByVal iMinor)

If arrRet(0) <> "" Then 
	Select Case iMinor
	Case 0												<%' ���� %>
		frm1.txtContry_cd.value = arrRet(0)
		frm1.txtCountry_nm.value = arrRet(1)
	Case 1												<%' ���� %>
		frm1.txtProvince_cd.value = arrRet(0)
		frm1.txtProvince_nm.value = arrRet(1)
	Case 2												<%' ȭ�� %>
		frm1.txtCurrency.value = arrRet(0)
	End Select
	lgBlnFlgChgValue = True
End If

End Function

'========================================================================================================= 
Function SetContentPopUp(Byval arrRet,ByVal iWhere)

If arrRet(0) <> "" Then 
	Select Case iWhere
	Case "CreditGrp"	<%' ���Ű����׷� %>
		frm1.txtCredit_grp.value = arrRet(0)
		frm1.txtCredit_grp_Nm.value = arrRet(1)
	Case "BpGrade"	<%' ��ü�򰡵�� %>
		frm1.txtBp_Grade.value = arrRet(0)
		frm1.txtBp_Grade_nm.value = arrRet(1)
	Case "PayTypeSales"		<%' ���������(����)%>
		frm1.txtPay_type.value = arrRet(0)
		frm1.txtPay_type_Nm.value = arrRet(1)
	Case "PayTypePur"		<%' ���������(����)%>
		frm1.txtPay_type_Pur.value = arrRet(0)
		frm1.txtPay_type_Pur_Nm.value = arrRet(1)
	End Select
	lgBlnFlgChgValue = True
End If

End Function


'========================================================================================================= 
Function CookiePage(ByVal Kubun)
	
	On Error Resume Next

	Const CookieSplit = 4877						

	Dim strTemp, arrVal

	If Kubun = 1 Then									
		WriteCookie CookieSplit , frm1.txtConBp_cd.value  & parent.gRowSep & frm1.txtConBp_nm.value

	ElseIf Kubun = 0 Then								

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtConBp_cd.value =  arrVal(0)
		frm1.txtConBp_nm.value =  arrVal(1)

		if Err.number <> 0 then
			Err.Clear
			WriteCookie CookieSplit , ""
			exit function
		end if
		
		Call MainQuery()		
			
		WriteCookie CookieSplit , ""

	End If
	
End Function


'===========================================================================
Function JumpChgCheck(strVal)

	Dim IntRetCD

	'************ �̱��� ��� **************
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)
	Call PgmJump(strVal)

End Function

<%
'========================================================================================
' Function Desc : This function is related to ID Check
'========================================================================================
%>
Function IDCheck(intIDFirst, intIDSecond)

<%
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
'    �ֹε�� üũ ��� 
'
'    Ex) 680312-1532520
'
'        6,  8,  0,  3,  1,  2,  1,  5,  3,  2,  5,  2
'    x)  2,  3,  4,  5,  6,  7,  8,  9,  2,  3,  4,  5
'    --------------------------------------------------
'    +) 12  24   0  15   6  14   8  45   6   6  20  10  = 166
'
'    11 - ( 166 / 11 ) = 11 - 1 = 10
'    ���� 680312-153252(0)
'    If [11-2=9] Then 680312-153252(9)


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
%>

    Dim arrID(1, 5)						'	�� �ֹε�Ϲ�ȣ �޴� �迭 
    Dim seqNum
    Dim logNum1, logNum2
    Dim TotalSum
    
    logNum1 = 1: logNum2 = 7			'	�ֹε�� ������ �ʿ��� �������� �ʱ�ȭ....
    
    For seqNum = 0 To 5
        
        logNum1 = logNum1 + 1			'	������� �� �ڸ��� �迭�� ���� 
        arrID(0, seqNum) = CInt(Mid(intIDFirst, seqNum + 1, 1)) * logNum1
        
        logNum2 = logNum2 + 1			'	�� 7�ڸ��� �� 6�ڸ��� �迭�� ���� 
        arrID(1, seqNum) = CInt(Mid(intIDSecond, seqNum + 1, 1)) * logNum2
        If logNum2 = 9 Then logNum2 = 1		'������ ����.... �ֹε�� �������� �ʿ�....	
    
    Next

    For seqNum = 0 To 5					'	�� �迭�� ���� �ڸ����� ���Ѵ�....

        TotalSum = TotalSum + arrID(0, seqNum) + arrID(1, seqNum)

    Next

    IDCheck = 11 - (TotalSum Mod 11)	'	�ֹε�� �ǵ��ڸ� ����....(���� �߿� ����)

End Function


'========================================================================================
Public Function CodeSect(ByVal strIndata) 
    
    Dim codehex , i
    Dim tmp1, tmp2

    CodeSect = "-1"
    
    If strIndata = "" Then
        Exit Function
    End If
    
    for i = 1 to len(strIndata)
		codehex = Right("0000" & Hex(Asc(Mid(strIndata,i,1))), 4)
    
		tmp1 = UCase(Left(codehex, 2))
		tmp2 = UCase(Right(codehex, 2))
    
		If (tmp2 >= "A1") And (tmp2 <= "F8") Then
			CodeSect = "0"
			Exit Function
		End If
    Next

End Function


<%
'========================================================================================
' Function Desc : ���ڸ� �Է¹޴� ���� üũ 
'========================================================================================
%>
Function NumericCheck()

	Dim objEl, KeyCode
	
	Set objEl = window.event.srcElement
	KeyCode = window.event.keycode

	Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
	Case Else
		window.event.keycode = 0
	End Select

End Function


<%
'==========================================================================================
'   Function Desc : ����� �ŷ�ó���п� ���� üũ�ڽ� Value Change
'==========================================================================================
%>
Function chkSaveValue()

  ' --> V ���� V ���� �̸� CS,  ���� V ���� �̸� S,
  ' --> V ����  ���� �̸� C, ���� ���� �̸� * �� ó���ٶ� 

	If frm1.chkBpTypeC.checked = True And frm1.chkBpTypeS.checked = True And frm1.chkBpTypeT.checked = false Then
		frm1.txtBp_Type.value = "CS"
		
		'## �ŷ�ó�� ����,����ó �Ѵ� �ϰ�� �ʵ带 ��� Ǭ��.


		'###1. ����ó 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_meth, "D") 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_type, "D")
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_dur, "D")
		frm1.btnPay_meth.disabled = False
		frm1.btnPay_type.disabled = False

		'###2. ����ó 

		
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_meth_Pur, "D") 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_type_Pur, "D")
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_dur_Pur, "D")
		frm1.btnPay_meth_Pur.disabled = False
		frm1.btnPay_type_Pur.disabled = False


		
	ElseIf frm1.chkBpTypeC.checked = True And frm1.chkBpTypeS.checked = False And frm1.chkBpTypeT.checked = false Then
		frm1.txtBp_Type.value = "C"
		
		'## �ŷ�ó�� ����ó�ϰ�� �����ʵ带 ���� 
		
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_meth_Pur, "Q")		 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_type_Pur, "Q")		
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_dur_Pur, "Q")
		frm1.btnPay_meth_Pur.disabled = true
		frm1.btnPay_type_Pur.disabled = true

		frm1.txtPay_meth_Pur.value = ""
		frm1.txtPay_meth_Pur_nm.value = ""
		frm1.txtPay_type_Pur.value = ""
		frm1.txtPay_type_Pur_nm.value = ""
		frm1.txtPay_dur_Pur.value = 0
		
		'## ����ó ��ư Ȱ��ȭ 
		frm1.btnPay_meth.disabled = False
		frm1.btnPay_type.disabled = False
		
	
		
	ElseIf frm1.chkBpTypeC.checked = False And frm1.chkBpTypeS.checked = True And frm1.chkBpTypeT.checked = false Then
		frm1.txtBp_Type.value = "S"
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_meth, "Q")	 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_type, "Q")	
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_dur, "Q")
		frm1.btnPay_meth.disabled = true
		frm1.btnPay_type.disabled = true

		frm1.txtPay_meth.value = ""
		frm1.txtPay_meth_nm.value = ""
		frm1.txtPay_type.value = ""
		frm1.txtPay_type_nm.value = ""
		frm1.txtPay_dur.value = 0
		
		'## ����ó ��ư Ȱ��ȭ 
		frm1.btnPay_meth_Pur.disabled = False
		frm1.btnPay_type_Pur.disabled = False

		
		
	ElseIf frm1.chkBpTypeC.checked = False And frm1.chkBpTypeS.checked = False And frm1.chkBpTypeT.checked = false Then
		frm1.txtBp_Type.value = "*"
		
		'## �ŷ�ó�� ����,����ó �Ѵ� �ƴҰ�� �ʵ带 ��� Ǭ��.
		
		'###1. ����ó 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_meth_Pur, "D") 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_type_Pur, "D")
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_dur_Pur, "D")
		frm1.btnPay_meth_Pur.disabled = False
		frm1.btnPay_type_Pur.disabled = False


		'###2. ����ó 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_meth, "D") 
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_type, "D")
		call ggoOper.SetReqAttr(window.document.frm1.txtPay_dur, "D")
		frm1.btnPay_meth.disabled = False
		frm1.btnPay_type.disabled = False
		
	End If

End Function

<% '==========================================================================================
'   Function Desc : ��ȸ�� �ŷ�ó���п� ���� üũ�ڽ� Value Change
'========================================================================================== %>
Function chkQueryValue()

	Select Case Trim(frm1.txtBp_Type.value)
	Case "CS"
		frm1.chkBpTypeC.checked = True
		frm1.chkBpTypeS.checked = True
		frm1.chkBpTypeT.checked = False
		
	Case "C"
		frm1.chkBpTypeC.checked = True
		frm1.chkBpTypeS.checked = False
		frm1.chkBpTypeT.checked = False
		
	Case "S"
		frm1.chkBpTypeC.checked = False
		frm1.chkBpTypeS.checked = True
		frm1.chkBpTypeT.checked = False
		
	Case "*"
		frm1.chkBpTypeC.checked = False
		frm1.chkBpTypeS.checked = False
		frm1.chkBpTypeT.checked = False
		
    Case "T"									'SON �ŷ�ó������ ���ݽŰ������ΰ�� 
		frm1.chkBpTypeC.checked = False
		frm1.chkBpTypeS.checked = False
		frm1.chkBpTypeT.checked = True		
		
	End Select

End Function

<% '==========================================================================================
'   Function Desc : ��ȸ�� �ŷ�ó������ üũ�ڽ��� üũ�� ���¼����Ұ� 
'========================================================================================== %>
Function chkQueryProtect()

	Select Case Trim(frm1.txtBp_Type.value)
	Case "CS"
		frm1.chkBpTypeC.disabled = True
		frm1.chkBpTypeS.disabled = True
		frm1.chkBpTypeT.disabled = True
	Case "C"
		frm1.chkBpTypeC.disabled = True
		frm1.chkBpTypeS.disabled = False
		frm1.chkBpTypeT.disabled = True
	Case "S"
		frm1.chkBpTypeC.disabled = False
		frm1.chkBpTypeS.disabled = True
		frm1.chkBpTypeT.disabled = True
	Case "*"
		frm1.chkBpTypeC.disabled = False
		frm1.chkBpTypeS.disabled = False
		frm1.chkBpTypeT.disabled = True
    Case "T"									
		frm1.chkBpTypeC.disabled = True
		frm1.chkBpTypeS.disabled = True
		frm1.chkBpTypeT.disabled = True
	End Select

End Function

<% '================================== =====================================================
' Function Desc : �Һ������ڵ�/��/�� �����ϱ� 
' ������� Ű���忡�� �Һ������ڵ带 ����� �Һ�������,�Һ���,���Աݾ�,NetAmount�� �����Ű�� �Լ� 
'======================================================================================== %>
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & "  And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation,Parent.gLogoName
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)  
		If arrCollectVatType(iCnt, 0) = UCase(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub

'========================================================================================
Sub SetVatType()
	Dim VatType, VatTypeNm, VatRate

	VatType = frm1.txtVat_Type.value
	
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

	frm1.txtvat_Type_nm.value = VatTypeNm	
	frm1.txtVat_rate.text = UNIFormatNumber(VatRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
End Sub

'========================================================================================
Function Check_ENTP_RGST(ByVal sNumber)
	Check_ENTP_RGST = False
<%	
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
'    ����ڵ�Ϲ�ȣüũ���� (2002-06-14 sonbumyeol - ���ο� ����ڵ�Ϲ�ȣüũ����)
'
'    �ŷ�ó���� ����� ��Ϲ�ȣüũ 
'    (�ش�ŷ�ó�� ���� �ڵ尡 �ѱ�(KR)�� �ƴҰ�� üũ�������� 
'
'    Ex) 603-81-13055
'
'    1. Ȯ�κ��� 0,3,7,0,3,7,0,3,0.5,0
'    2. Ȯ�κ����� '0'�ϰ���  ���ϰ�, '0'�̿��� ����� ���ڴ� ���� 
'    3. Ȯ�κ��� 0.5�� ���� ���Ͽ� ���¼��� �����ο� �Ҽ��� �� ���� 
'    4. ��������� �հ������ ���ڸ��� '0'�� �Ǹ� ��Ȯ�� ����� ��ȣ�� 
' 
'
'    <����� ��ȣ ������>
'    Ex) 603-81-13055
' 
'        Ȯ�κ���      
'
'    6  +  0        =  6 
'    0  *  3        =  3
'    3  *  7        =  21 
'    _________________ 
'    8  +  0        =  8
'    1  *  3        =  3
'    _________________
'    1  *  7        =  7
'    3  +  0        =  3
'    0  *  3        =  0
'    5  *  0.5      =  2.5 ( 2+5 =7)
'    5  +  0        =  5
'   _________________________________
'    �հ�              60     
'
'    --> �հ��� ���ڸ����� '0'�̹Ƿ� ��Ȯ�� ����� ��ȣ�� 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
%>	
		
	
	Dim sum,i  
	Dim li_chkvalue(9)  
	Dim NumCnt, Number,NumCnt0,NumCnt1,NumCnt2,NumCnt3,NumCnt4,NumCnt5,NumCnt6,NumCnt7,NumCnt8,NumCnt9



	Number = Replace(sNumber, "-", "")
	
	If isNumeric(Number) = False Then 
	   Exit Function
	End If
	
	NumCnt = Len(Number)
	
	Select Case NumCnt
	Case 13
		Exit Function
	
	Case 10
		
		sum = 0
				
		For i = 1 To 10
			li_chkvalue(i-1) = Mid(Number,i,1) 
		Next
										
				
		NumCnt0 = li_chkvalue(0) + 0
		NumCnt1 = li_chkvalue(1) * 3
		NumCnt2 = li_chkvalue(2) * 7
		NumCnt3 = li_chkvalue(3) + 0
		NumCnt4 = li_chkvalue(4) * 3
		NumCnt5 = li_chkvalue(5) * 7
		NumCnt6 = li_chkvalue(6) + 0
		NumCnt7 = li_chkvalue(7) * 3
		NumCnt8 = Int(li_chkvalue(8) * 0.5) + Int(((li_chkvalue(8) * 0.5) * 10) Mod 10)				
		NumCnt9 = li_chkvalue(9) + 0


		sum = (NumCnt0 + NumCnt1 + NumCnt2 + NumCnt3 + NumCnt4 + NumCnt5 + NumCnt6 + NumCnt7 + NumCnt8 + NumCnt9)
				
		if int(sum) MOD 10 <> 0 then Exit Function
				
	Case Else 
	
		Exit Function
	End Select 

	Check_ENTP_RGST = True
	
End Function

'========================================================================================
Function Check_INDI_RGST(ByVal sID) 

	Check_INDI_RGST = False

	Dim Weight 
	Dim Total 
	Dim Chk 
	Dim Rmn 
	Dim i 
	Dim dt 
	Dim wt 
	Dim Number, Numcnt

	Number = Replace(sID, "-", "")
	Numcnt = Len(Number)

	Select Case Numcnt
	Case 13
		Chk = CDbl(Right(Number, 1))

		Weight = "234567892345"
		Total = 0

		For i = 1 To 12
		dt = CDbl(Mid(Number, i, 1))
		wt = CDbl(Mid(Weight, i, 1))
		Total = Total + (dt * wt)
		Next 

		Rmn = 11 - (Total Mod 11)

		If Rmn > 9 Then Rmn = Rmn Mod 10

		If Rmn <> Chk Then Exit Function

 	Case 0
 		Check_INDI_RGST = True 
		Exit Function
		
 	Case Else 
		Exit Function
	End Select 

	Check_INDI_RGST = True
End Function

'========================================================================================
Sub FncSplitColumn()    
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call AppendNumberRange("0","0","100")
	Call AppendNumberRange("1","0","999")
	Call AppendNumberRange("2","0","99")
	Call AppendNumberRange("3","0","31")                                                 '��: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","3","2")
	Call AppendNumberPlace("8","15","0")
	Call AppendNumberPlace("9","2","0")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart) 

    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call SetDefaultVal   

    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("11101000000011")										'��: ��ư ���� ���� 
    Call InitVariables                                                      '��: Initializes local global variables
	Call CookiePage(0)
	Call ChangeTabs(TAB1)
	
	gIsTab = "Y"
	gTabMaxCnt = 4			

End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

<%
'==========================================================================================
'   Event Desc : Radio Button Click�� lgBlnFlgChgValue ó�� 
'==========================================================================================
%>
Sub rdoUsage_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoUsage_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoIn_out1_OnClick()
	lgBlnFlgChgValue = True
	<% '�����׷� %>
	Call ggoOper.SetReqAttr(frm1.txtBiz_Grp, "N")
End Sub

Sub rdoIn_out2_OnClick()
	lgBlnFlgChgValue = True
	<% '�����׷� %>
	Call ggoOper.SetReqAttr(frm1.txtBiz_Grp, "D")
End Sub

Sub rdoVATinc_1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoVATinc_2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoCredit_Y_OnClick()
	lgBlnFlgChgValue = True

	<% '���Ű����׷� %>
	Call ggoOper.SetReqAttr(frm1.txtCredit_grp, "N")
	'<% '����ȸ���� %>
	'Call ggoOper.SetReqAttr(frm1.txtCreditRotDt, "D")
End Sub

Sub rdoCredit_N_OnClick()
	lgBlnFlgChgValue = True

	<% '���Ű����׷� %>
	Call ggoOper.SetReqAttr(frm1.txtCredit_grp, "Q")
	'<% '����ȸ���� %>
	'Call ggoOper.SetReqAttr(frm1.txtCreditRotDt, "Q")
	
	frm1.txtCredit_grp.value = ""
	frm1.txtCredit_grp_Nm.value = ""
	'frm1.txtCreditRotDt.Text = 0
	
End Sub

Sub rdoVATcalc_Y_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoVATcalc_N_OnClick()
	lgBlnFlgChgValue = True
End Sub


Sub rdoReservePrice_Y_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoReservePrice_N_OnClick()
	lgBlnFlgChgValue = True
End Sub


Sub rdoSoldInspectA_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoSoldInspectB_OnClick()
	lgBlnFlgChgValue = True
End Sub

<%
'========================================================================================
' Sub Name : txtRepre_Rgst OnKeyPress Event
' Sub Desc : AutoTab Event
'========================================================================================
%>
Sub txtPay_type_OnKeyPress()
	If Trim(frm1.txtPay_meth.value) = "" Then
		Call DisplayMsgBox("205152","x",frm1.txtPay_meth.alt,"x")
		'MsgBox "��������� ���� �Է��ϼ���!"
		frm1.txtPay_meth.focus 
		IsOpenPop = False			
		Exit Sub
	End If
End Sub

Sub txtPay_type_pur_OnKeyPress()
	If Trim(frm1.txtPay_meth_pur.value) = "" Then
		Call DisplayMsgBox("205152","x",frm1.txtPay_meth_pur.alt,"x")
		'MsgBox "��������� ���� �Է��ϼ���!"
		frm1.txtPay_meth_pur.focus 
		IsOpenPop = False			
		Exit Sub
	End If
End Sub

'==========================================================================================
Sub txtContry_cd_OnChange()
	frm1.txtZIP_cd.value = ""
	frm1.txtProvince_cd.value = ""
	frm1.txtProvince_nm.value = ""
End Sub

'==========================================================================================
Function txtZIP_cd_OnFocus()
	If Trim(frm1.txtContry_cd.value) = "" Then
		MsgBox "������ ���� �Է��ϼ���", vbInformation, parent.gLogoName
		frm1.txtContry_cd.focus 
		Exit Function
	End IF
End Function

'==========================================================================================
Function txtZIP_cd_OnChange()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	If gLookUpEnable = False Then Exit Function

	frm1.txtADDR1.value = ""
	frm1.txtADDR2.value = ""
	frm1.txtProvince_cd.value = ""
	frm1.txtProvince_nm.value = ""

	If Trim(frm1.txtZIP_cd.value) = "" Then Exit Function
        
'--
    Call CommonQueryRs(" ADDRESS "," B_ZIP_CODE "," COUNTRY_CD =  " & FilterVar(frm1.txtContry_cd.value, "''", "S") & " AND ZIP_CD =  " & FilterVar(frm1.txtZIP_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if lgf0 = "" then 
		frm1.txtADDR1.value = ""
	else 
	    frm1.txtADDR1.value = Trim(Replace(lgF0,Chr(11),""))
    end if 
'--

End Function

<%
'==========================================================================================
'   Event Desc : OCX_Change()
'==========================================================================================
%>
<% 'â������� %>
Sub txtFnd_DT_Change()
	lgBlnFlgChgValue = True
End Sub

'*******************************************
' 2005.07.06 smj
' ����ڹ�ȣ ������ 
'*******************************************
Sub txtOwn_Rgst_DT_Change()
	lgBlnFlgChgValue = True
End Sub


<% '������ %>
Sub txtClose_day1_Change()
	lgBlnFlgChgValue = True
End Sub

<% '������ %>
Sub txtPay_Month_Change()
	lgBlnFlgChgValue = True
End Sub



<% '�������� %>
Sub txtEmp_Cnt_Change()
	lgBlnFlgChgValue = True
End Sub

<% '�Ⱓ����� %>
Sub txtSale_Amt_Change()
	lgBlnFlgChgValue = True
End Sub

<% '�ں��� %>
Sub txtCapital_Amt_Change()
	lgBlnFlgChgValue = True
End Sub

<% '���L/T %>
Sub txtTrans_LT_Change()
	lgBlnFlgChgValue = True
End Sub

<% '�������� %>
Sub txtComm_Rate_Change()
	lgBlnFlgChgValue = True
End Sub

<% 'vat�� %>
Sub txtvat_Rate_Change()
	lgBlnFlgChgValue = True
End Sub

<% '�����Ⱓ(����) %>
Sub txtPay_dur_Change()
	lgBlnFlgChgValue = True
End Sub

<% '�����Ⱓ(����) %>
Sub txtPay_dur_Pur_Change()
	lgBlnFlgChgValue = True
End Sub


<% '������(����) %>
Sub txtPay_day_Change()
	lgBlnFlgChgValue = True
End Sub


<% '����ȸ���� %>
Sub txtCreditRotDt_Change()
	lgBlnFlgChgValue = True
End Sub

<% '������ ����ȭ�� %>
Sub txtCash_Rate_Change()
	lgBlnFlgChgValue = True
End Sub

<%
'==========================================================================================
'   Event Desc : OCX_DbClick() �� Calendar Popup
'==========================================================================================
%>
Sub txtFnd_DT_DblClick(Button)
	If Button = 1 Then
		frm1.txtFnd_DT.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtFnd_DT.Focus
	End If
End Sub

<%
'==========================================================================================
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
%>
Sub txtFnd_DT_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'*******************************************
' 2005.07.06 SMJ
' ����ڹ�ȣ ������ 
'*******************************************

Sub txtOwn_Rgst_Dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOwn_Rgst_Dt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtOwn_Rgst_Dt.Focus
	End If
End Sub

<%
'==========================================================================================
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
%>
Sub txtOwn_Rgst_Dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


<%
'==========================================================================================
'   Event Name : chkBpType ...
'   Event Desc : �ŷ�ó���п� ���� üũ�ڽ� Value Change
'==========================================================================================
%>
Sub chkBpTypeS_OnPropertyChange()
End Sub

Sub chkBpTypeC_OnPropertyChange()
End Sub

Sub chkBpTypeC_OnClick()
	lgBlnFlgChgValue = True	
	Call chkSaveValue()
End Sub

Sub chkBpTypeS_OnClick()	
	lgBlnFlgChgValue = True
	Call chkSaveValue()
End Sub

'==========================================================================================
'   Event Desc : �������º��� �������� �ʼ��Է� ó�� 
'==========================================================================================
Sub txtVat_Type_OnChange()
	Call SetVatType()
End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x") '�� �ٲ�κ� 
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")										<%'��: Clear Contents  Field%>
    Call InitVariables															<%'��: Initializes local global variables%>
    
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If

	Call ggoOper.LockField(Document, "N")		                                      <%'��: Lock  Suitable  Field%>
	Call SetToolBar("11101000000011")
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    Call DbQuery																<%'��: Query db data%>
       
    FncQuery = True																
	       
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
<%  '-----------------------
    'Check previous data area
    '-----------------------%>
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x") 
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------%>
    Call ggoOper.ClearField(Document, "A")                                      <%'��: Clear Condition,Contents Field%>
    Call ggoOper.LockField(Document, "N")                                       <%'��: Lock  Suitable  Field%>
    Call SetDefaultVal
    Call SetToolBar("11101000000011")
    Call InitVariables															<%'��: Initializes local global variables%>

    FncNew = True																

End Function

'========================================================================================
Function FncDelete() 

    Dim IntRetCD
    
    FncDelete = False														
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        '��ȸ�� ���� �Ͻʽÿ�.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")
    '���� �Ͻðڽ��ϱ�?
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
    If frm1.chkBpTypeT.checked = True Then
        Call DisplayMsgBox("126143","x","x","x")
        '���ݽŰ������� �ŷ�ó�� �����Ҽ������ϴ�.
        Exit Function
    End If
    
    Call DbDelete															<%'��: Delete db data%>
    
    FncDelete = True                                                        
    
End Function

'========================================================================================
Function FncSave() 

	On Error Resume Next

    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")    
        Exit Function
    End If

    If Not chkField(Document, "2") Then
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
			Set gActiveElement = document.activeElement
        End If
        Exit Function
    End If         
    
    '����� ��Ϲ�ȣ üũ => �ߺ��Ǿ��ִ��� Ȯ��	(��Ͻø� üũ�ϰ� �����ô� üũ���� �ʴ´�.)
	If Trim(UCase(frm1.txtHConBp_cd.value)) = "" Then		
		IF Check_Double_ENTP_RGST = True Then
		  	Dim Check_Double_ENTP
		  	Check_Double_ENTP = DisplayMsgBox("126145", parent.VB_YES_NO, "X", "X")
		  	'%1 1���� �ŷ�ó�� ������ ����ڵ�Ϲ�ȣ�� ��ϵǾ��ֽ��ϴ�. ��� �����Ͻðڽ��ϱ�?
		  	If Check_Double_ENTP = vbNo Then                                                        
		  		Exit Function
		  		FncSave = True
		  	End If
		End If		  
	End If
	
	'����� ��Ϲ�ȣ üũ => ǥ�ػ���ڵ�Ϲ�ȣ���� Ȯ��		
	If Trim(UCase(frm1.txtContry_cd.value)) = "KR" AND Trim(UCase(frm1.txtHConBp_cd.value)) = "" Then				 	
		If Check_ENTP_RGST(Trim(frm1.txtOwn_Rgst_N.value)) = False AND UCase(parent.gCountry) = "KR"  Then
			Dim Check_ENTP
			Check_ENTP = DisplayMsgBox("126140", parent.VB_YES_NO, "X", "X")
			'�߸��� ����ڵ�Ϲ�ȣ�Դϴ�. �����Ͻðڽ��ϱ�?.
			If Check_ENTP = vbNo Then                                                        
				Exit Function
				FncSave = True
		    End If		
		End If
	End If
	
	'�ֹε�Ϲ�ȣ üũ    
    If Trim(UCase(frm1.txtContry_cd.value)) = "KR" AND Trim(UCase(frm1.txtHConBp_cd.value)) = "" Then	
		If Check_INDI_RGST(Trim(frm1.txtRepre_Rgst.value)) = False  AND UCase(parent.gCountry) = "KR"  Then
		    Dim Check_INDI
		    Check_INDI = DisplayMsgBox("126139", parent.VB_YES_NO, "X", "X")
		    '�߸��� �ֹε�Ϲ�ȣ�Դϴ�. �����Ͻðڽ��ϱ�?.
		    If Check_INDI = vbNo Then                                                       
		  		Exit Function
		  		FncSave = True 	 
		    End If
		End If
    End If
    
    '�ŷ�ó������ üũ   
	If CodeSect(frm1.txtBp_eng_nm.value ) = "0" Then
		Dim Check_CodeSect2
		Check_CodeSect2 = DisplayMsgBox("126144", parent.VB_YES_NO, "X", "X")
			'�ŷ�ó������ �ѱ��� �Էµƽ��ϴ�. �����Ͻðڽ��ϱ�?
			If Check_CodeSect2 = vbNo Then                                                        
				Exit Function
				FncSave = True
			End If			
	End If
	
    '�����ּ� üũ   
	If CodeSect(frm1.txtADDR1_Eng.value ) = "0" Or CodeSect(frm1.txtADDR2_Eng.value ) = "0" Or CodeSect(frm1.txtADDR3_Eng.value ) = "0" then
		Dim Check_CodeSect1
		Check_CodeSect1 = DisplayMsgBox("126314", parent.VB_YES_NO, "X", "X")
			'�����ּҿ� �ѱ��� �Էµƽ��ϴ�. �����Ͻðڽ��ϱ�?
			If Check_CodeSect1 = vbNo Then                                                        
				Exit Function
				FncSave = True
			End If			
	End If
	
	'-----------------------
    'Check RadioButton area
    '-----------------------
	With frm1
		<% '��뿩�� %>
		If .rdoUsage_flag1.checked = True Then
			.txtRadioFlag.value = .rdoUsage_flag1.value 
		ElseIf .rdoUsage_flag2.checked = True Then
			.txtRadioFlag.value = .rdoUsage_flag2.value 
		End IF
		
		<% '�系�ܱ��� %>
		If .rdoIn_out1.checked = True Then
			.txtRadioInOut.value = .rdoIn_out1.value 
		ElseIf .rdoIn_out2.checked = True Then
			.txtRadioInOut.value = .rdoIn_out2.value 
		End IF

		<% '�ΰ������Կ��� %>
		If .rdoVATinc_1.checked = True Then
			.txtRadioVATinc.value = .rdoVATinc_1.value 
		ElseIf .rdoVATinc_2.checked = True Then
			.txtRadioVATinc.value = .rdoVATinc_2.value 
		End IF

		<% '���Ű������� %>
		If .rdoCredit_N.checked = True Then
			.txtRadioCredit.value = .rdoCredit_N.value 
		ElseIf .rdoCredit_Y.checked = True Then
			.txtRadioCredit.value = .rdoCredit_Y.value 
		End IF

		<% '�ΰ�������� %>
		If .rdoVATcalc_Y.checked = True Then
			.txtRadioVATcalc.value = .rdoVATcalc_Y.value 
		ElseIf .rdoVATcalc_N.checked = True Then
			.txtRadioVATcalc.value = .rdoVATcalc_N.value 
		End IF

		<% '������������� %>
		If .rdoReservePrice_Y.checked = True Then
			.txtRadioDepositPrice.value = .rdoReservePrice_Y.value 
		ElseIf .rdoReservePrice_N.checked = True Then
			.txtRadioDepositPrice.value = .rdoReservePrice_N.value 
		End IF


		<% '��ǰ�ð˻��� %>
		If .rdoSoldInspectA.checked = True Then
			.txtRadioSoldInspect.value = .rdoSoldInspectA.value
		ElseIf .rdoSoldInspectB.checked = True Then
			.txtRadioSoldInspect.value = .rdoSoldInspectB.value 
		End IF

	End With

	If Len(Trim(frm1.txtZIP_cd.value)) Then
		If Trim(frm1.txtADDR1.value) = "" Then
			IntRetCD = DisplayMsgBox("970029", "x","�ּ�","x")			
			Exit Function
		ENd If
	End If
      
    Call chkSaveValue()
    
<%  '-----------------------
    'Save function call area
    '-----------------------%>
   
    Call DbSave				                                                <%'��: Save db data%>

    FncSave = True                                                          
    
End Function

'========================================================================================
Function FncCopy() 
	Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												<%'��: Indicates that current mode is Crate mode%>
    
    <% ' ���Ǻ� �ʵ带 �����Ѵ�. %>
    Call ggoOper.ClearField(Document, "1")                                      <%'��: Clear Condition Field%>
    Call ggoOper.LockField(Document, "N")									<%'��: This function lock the suitable field%>
    Call InitVariables															<%'��: Initializes local global variables%>
    Call SetToolBar("11101000000111")

	If frm1.rdoCredit_Y.checked = True Then
		Call rdoCredit_Y_OnClick()
	ElseIf frm1.rdoCredit_N.checked = True Then
		Call rdoCredit_N_OnClick()
	End If
    
    frm1.txtBp_cd.value = ""
    
    '���ݽŰ����� flag ���� 
	frm1.chkBpTypeT.disabled = True
	frm1.chkBpTypeT.checked = False
    
End Function

'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
Function FncPrev() 

	Dim strVal

	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

	frm1.txtPrevNext.value = "PREV"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtConBp_cd=" & Trim(frm1.txtBp_cd.value)		<%'��: ��ȸ ���� ����Ÿ %>
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'��: ��ȸ ���� ����Ÿ %>
         
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
Function FncNext() 
	
	Dim strVal
		    
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
	
	frm1.txtPrevNext.value = "NEXT"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtConBp_cd=" & Trim(frm1.txtBp_cd.value)		<%'��: ��ȸ ���� ����Ÿ %>
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'��: ��ȸ ���� ����Ÿ %>
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, TRUE)
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   '�� �ٲ�κ� 
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vb
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
    
        
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							<%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtBp_cd=" & Trim(frm1.txtBp_cd.value)				<%'��: ���� ���� ����Ÿ %>
    
	Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
	
    DbDelete = True                                                         

End Function

'========================================================================================
Function DbDeleteOk()														<%'��: ���� ������ ���� ���� %>
	Call FncNew()
End Function

'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         
    
        
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtConBp_cd=" & Trim(frm1.txtConBp_cd.value)		<%'��: ��ȸ ���� ����Ÿ %>
    
	Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
    DbQuery = True                                                          

End Function

'========================================================================================
Function DbQueryOk()														<%'��: ��ȸ ������ ������� %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												<%'��: Indicates that current mode is Update mode%>
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									<%'��: This function lock the suitable field%>
	Call SetToolBar("1111100011111111")

	If frm1.rdoIn_out1.checked = True Then	 <% '�系�ܿ��� %>
		<% '�����׷� %>
		Call ggoOper.SetReqAttr(frm1.txtBiz_Grp, "N")
	ElseIf frm1.rdoIn_out2.checked = True Then
		<% '�����׷� %>
		Call ggoOper.SetReqAttr(frm1.txtBiz_Grp, "D")
	End If

	If frm1.txtRadioCredit.value = "Y" Then	 <% '���Ű������� %>
		<% '���Ű����׷� %>
		Call ggoOper.SetReqAttr(frm1.txtCredit_grp, "N")
		'<% '����ȸ���� %>
		'Call ggoOper.SetReqAttr(frm1.txtCreditRotDt, "D")
	ElseIf frm1.txtRadioCredit.value = "N" Then
		<% '���Ű����׷� %>
		Call ggoOper.SetReqAttr(frm1.txtCredit_grp, "Q")
		'<% '����ȸ���� %>
		'Call ggoOper.SetReqAttr(frm1.txtCreditRotDt, "Q")
	End If

	Call chkQueryProtect()
	Call chkSaveValue()
	frm1.txtConBp_cd.focus


    frm1.txtHConBp_cd.value = frm1.txtBp_cd.value 
	
End Function

'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	    
	If LayerShowHide(1) = False Then
       Exit Function 
    End If

	
    Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002											<%'��: �����Ͻ� ó�� ASP �� ���� %>
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = parent.gUsrID 
		.txtUpdtUserId.value = parent.gUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           
    
End Function

'========================================================================================
Function DbSaveOk()															<%'��: ���� ������ ���� ���� %>

    frm1.txtConBp_cd.value = frm1.txtBp_cd.value 
    frm1.txtConBp_nm.value = frm1.txtBp_nm.value  
    
    Call InitVariables
    
    Call MainQuery()

End Function

'========================================================================================
Function Check_Double_ENTP_RGST()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" COUNT(BP_CD) ", " B_BIZ_PARTNER ", " BP_RGST_NO =  " & FilterVar(frm1.txtOwn_Rgst_N.value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    iCodeArr = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation,Parent.gLogoName
		Err.Clear 
		Exit Function
	End If

	If iCodeArr(0) = 0 Then 
		Check_Double_ENTP_RGST = False
	Else
		Check_Double_ENTP_RGST = True
	End If
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�Ϲ�����</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab4()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ȸ������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConBp_cd" ALT="�ŷ�ó" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConBp_cd()">&nbsp;<INPUT NAME="txtConBp_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TDT"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<!-- ù��° �� ���� -->
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó�ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_cd" ALT="�ŷ�ó�ڵ�" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU"></TD>								
								<TD CLASS=TD5 NOWRAP>�ŷ�ó����</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkBpTypeC" ID="chkBpTypeC" tag="21" Class="Check"><LABEL FOR="chkBpTypeC">����</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=CHECKBOX NAME="chkBpTypeS" ID="chkBpTypeS" tag="21" Class="Check"><LABEL FOR="chkBpTypeS">����</LABEL>&nbsp;&nbsp;
								    <INPUT TYPE=CHECKBOX NAME="chkBpTypeT" ID="chkBpTypeT" tag="21" Class="Check"><LABEL FOR="chkBpTypeT">���ݽŰ�����</LABEL>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ڵ�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOwn_Rgst_N" ALT="����ڵ�Ϲ�ȣ" TYPE="Text" MAXLENGTH="20" SIZE=40 tag="23XXX"></TD>
								<TD CLASS=TD5 NOWRAP>��뿩��</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag1" value="Y" tag = "21" checked>
										<label for="rdoUsage_flag1">��</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoUsage_flag" id="rdoUsage_flag2" value="N" tag = "21">
										<label for="rdoUsage_flag2">�ƴϿ�</label></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_full_nm" ALT="�ŷ�ó����" TYPE="Text" MAXLENGTH="120" SIZE=40 tag="23XXX"></TD>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó(MES)</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_alias_nm" TYPE="Text" ALT="�ŷ�ó��Ī" MAXLENGTH="80" SIZE=40 style="background:#FFE5CB"  tag="21XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó��Ī</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_nm" ALT="�ŷ�ó��Ī" TYPE="Text" MAXLENGTH="50" SIZE=40 tag="23XXX"></TD>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_eng_nm" TYPE="Text" ALT="�ŷ�ó������" MAXLENGTH="50" SIZE=40 tag="25XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǥ�ڸ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_nm" ALT="��ǥ�ڸ�" TYPE="Text" MAXLENGTH="50" SIZE=40 tag="23XXX"></TD>
								<TD CLASS=TD5 NOWRAP>��ǥ���ֹε�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_Rgst" TYPE="Text" ALT="��ǥ���ֹε�Ϲ�ȣ" MAXLENGTH="20" SIZE=40 tag="25XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Class" ALT="����" TYPE="Text" MAXLENGTH="5" SIZE=6 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinor 0">&nbsp;<INPUT NAME="txtInd_ClassNm" TYPE="Text" MAXLENGTH="30" SIZE=30 tag="24"></TD>							
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="����" TYPE="Text" MAXLENGTH="5" SIZE=6 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinor 1">&nbsp;<INPUT NAME="txtInd_TypeNm" TYPE="Text" MAXLENGTH="30" SIZE=30 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�系�ܱ���</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoIn_out" id="rdoIn_out1" value="I" tag = "21" >
										<label for="rdoIn_out1">�系</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoIn_out" id="rdoIn_out2" value="O" tag = "21" checked>
										<label for="rdoIn_out2">���</label></TD>
								<TD CLASS=TD5 NOWRAP>����ڹ�ȣ������</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtOwn_Rgst_DT" CLASS=FPDTYYYYMMDD tag="23X1" Title="FPDATETIME" ALT = "����ڹ�ȣ������"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<%Call SubFillRemBodyTD5656(9)%>
						</TABLE>
						</DIV>

						<!-- �ι�° �� ���� -->
						<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó�з�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_Group" ALT="�ŷ�ó�з�" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Group" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinor 8">&nbsp;<INPUT NAME="txtBp_Group_Nm" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtContry_cd" TYPE="Text" ALT="����" MAXLENGTH="2" SIZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnContry_Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEtc 0">&nbsp;<INPUT NAME="txtCountry_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIP_cd" TYPE="Text" ALT="�����ȣ" MAXLENGTH="12" SIZE=20 tag="25XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtProvince_cd" TYPE="Text" ALT="����" MAXLENGTH="5" SIZE=10 tag="25XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProvince_Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEtc 1">&nbsp;<INPUT NAME="txtProvince_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ּ�</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1" TYPE="Text" ALT="�ּ�" MAXLENGTH="100" SIZE=100 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2" TYPE="Text" ALT="�ּ�" MAXLENGTH="100" SIZE=100 tag="25XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ּ�</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1_Eng" TYPE="Text" ALT="�����ּ�" MAXLENGTH="50" SIZE=100 tag="25XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2_Eng" TYPE="Text" ALT="�����ּ�" MAXLENGTH="50" SIZE=100 tag="25XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR3_Eng" TYPE="Text" ALT="�����ּ�" MAXLENGTH="50" SIZE=100 tag="25XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ȭ��ȣ1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No1" TYPE="Text" ALT="��ȭ��ȣ1" MAXLENGTH="20" SIZE=34 tag="21"></TD>
								<TD CLASS=TD5 NOWRAP>��ȭ��ȣ2</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No2" TYPE="Text" ALT="��ȭ��ȣ2" MAXLENGTH="20" SIZE=34 tag="21"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ѽ���ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFax_No" TYPE="Text" ALT="�ѽ���ȣ" MAXLENGTH="20" SIZE=34 tag="21"></TD>
								<TD CLASS=TD5 NOWRAP>Ȩ�������ּ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHome_Url" TYPE="Text" ALT="Ȩ�������ּ�" MAXLENGTH="30" SIZE=34 tag="21"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>â�������</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFnd_DT" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtEmp_Cnt" CLASS=FPDS140 tag="21X8Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ⱓ�����</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtSale_Amt" CLASS=FPDS140 tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 NOWRAP>�ں���</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtCapital_Amt" CLASS=FPDS140 tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>							
								</TD>
							</TR>
							<%Call SubFillRemBodyTD5656(5)%>
						</TABLE>
						</DIV>

						<!-- ����° �� ���� -->					
						<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>��۹��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_Meth" TYPE="Text" ALT="��۹��" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrans_Meth" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinor 2">&nbsp;<INPUT NAME="txtTrans_Meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>���L/T</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtTrans_LT" CLASS=FPDS140 tag="21X6Z" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ǹ�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeal_Type" TYPE="Text" ALT="�Ǹ�����" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeal_Type" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinor 4">&nbsp;<INPUT NAME="txtDeal_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoVAT_inc_flag" id="rdoVATinc_1" value="1" tag = "21" checked>
										<label for="rdoVATinc_1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoVAT_inc_flag" id="rdoVATinc_2" value="2" tag = "21">
										<label for="rdoVATinc_2">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrade_Rgst" TYPE="Text" ALT="��������Ϲ�ȣ" MAXLENGTH="20" SIZE=34 tag="21XXX"></TD>
								<TD CLASS=TD5 NOWRAP>���������ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtClearance_ID" TYPE="Text" ALT="���������ȣ" MAXLENGTH="15" SIZE=34 tag="21XXX"></TD>
							</TR>		
							<TR>	
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle8 NAME="txtComm_Rate" CLASS=FPDS140 tag="21X7Z" Title="FPDOUBLESINGLE"> </OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL>
								</TD>
								<TD CLASS=TD5 NOWRAP>��ü�򰡵��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_Grade" TYPE="Text" ALT="��ü�򰡵��" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Grade" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenContentPopUp('BpGrade')">&nbsp;<INPUT NAME="txtBp_Grade_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>							
							</TR>
							
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó����ڸ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_prsn_Nm" TYPE="Text" ALT="�ŷ�ó����ڸ�" MAXLENGTH="50" SIZE=34 tag="21"></TD>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó����ڿ���ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_contact_Pt" TYPE="Text" ALT="�ŷ�ó����ڿ���ó" MAXLENGTH="30" SIZE=34 tag="21"></TD>
							</TR>
							<TR>							
								<TD CLASS=TD5 NOWRAP>��ǰ�ð˻���</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoSoldInspect" id="rdoSoldInspectA" value="A" tag = "21" checked>
										<label for="rdoSoldInspectA">�԰��İ˻�</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoSoldInspect" id="rdoSoldInspectB" value="B" tag = "21" >
										<label for="rdoSoldInspectB">�԰����˻�</label></TD>
								<TD CLASS=TD5 NOWRAP>���Ű�������</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS = "RADIO" name="rdoCredit_mgmt_flag" id="rdoCredit_N" value="N" tag = "21" checked>
										<label for="rdoCredit_N">�̰���</label>
									<input type=radio CLASS="RADIO" name="rdoCredit_mgmt_flag" id="rdoCredit_Y" value="Y" tag = "21" >
										<label for="rdoCredit_Y">����</label>&nbsp;&nbsp;&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���Ű����׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCredit_grp" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="24XXXU" ALT="���Ű����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrans_Meth" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenContentPopUp('CreditGrp')">&nbsp;<INPUT NAME="txtCredit_grp_Nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>����ȸ����</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtCreditRotDt" CLASS=FPDS140 tag="21X6Z" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL>
								</TD>
							</TR>
							<TR>							
								<TD CLASS=TD5 NOWRAP>�����׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBiz_Grp" TYPE="Text" ALT="�����׷�" MAXLENGTH="4" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBiz_Grp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBiz_Grp()">&nbsp;<INPUT NAME="txtBiz_Grp_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>���ݱ׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTo_Grp" TYPE="Text" ALT="���ݱ׷�" MAXLENGTH="4" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTo_Grp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTo_Grp()">&nbsp;<INPUT NAME="txtTo_Grp_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
							</TR>
							<TR>							
								<TD CLASS=TD5 NOWRAP>���ű׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPur_Grp" TYPE="Text" ALT="���ű׷�" MAXLENGTH="4" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPur_Grp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPur_Grp()">&nbsp;<INPUT NAME="txtPur_Grp_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>						
			
							</TR>
							<%Call SubFillRemBodyTD5656(7)%>
						</TABLE>
						</DIV>

						<!-- �׹�° �� ���� -->
						<DIV ID="TabDiv" SCROLL=no>					
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ȭ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCurrency" TYPE="Text" ALT="ȭ��" MAXLENGTH="3" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEtc 2"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>						
							</TR>
							<TR>				
								<TD CLASS=TD5 NOWRAP>�������������</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoReserve_Price_type" id="rdoReservePrice_Y" value="1" tag = "21">
										<label for="rdoReservePrice_Y">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoReserve_Price_type" id="rdoReservePrice_N" value="2" tag = "21" checked>
										<label for="rdoReservePrice_N">������</label></TD>
								<TD CLASS=TD5 NOWRAP>VAT�������</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoVAT_calc_type" id="rdoVATcalc_Y" value="1" tag = "21" checked>
										<label for="rdoVATcalc_Y">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoVAT_calc_type" id="rdoVATcalc_N" value="2" tag = "21">
										<label for="rdoVATcalc_N">����</label></TD>
							</TR>	
							<TR>	
								<TD CLASS=TD5 NOWRAP>VAT����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtvat_Type" TYPE="Text" ALT="VAT����" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnvat_Type" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinor 5">&nbsp;<INPUT NAME="txtvat_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>VAT��</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtvat_Rate" CLASS=FPDS140 tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�������(����)</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_meth" TYPE="Text" MAXLENGTH="4" SIZE=10 Alt="�������(����)" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPay_meth" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinor 6">&nbsp;<INPUT NAME="txtPay_meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>								
								<TD CLASS=TD5 NOWRAP>�������(����)</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_meth_Pur" TYPE="Text" MAXLENGTH="4" SIZE=10 Alt="�������(����)" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPay_meth_Pur" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinor 9">&nbsp;<INPUT NAME="txtPay_meth_Pur_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>														
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>���������(����)</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_type" TYPE="Text" ALT="���������(����)" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPay_type" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenContentPopUp('PayTypeSales')">&nbsp;<INPUT NAME="txtPay_type_Nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>							
								<TD CLASS=TD5 NOWRAP>���������(����)</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_type_Pur" TYPE="Text" ALT="���������(����)" MAXLENGTH="5" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPay_type_Pur" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenContentPopUp('PayTypePur')">&nbsp;<INPUT NAME="txtPay_type_Pur_Nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>�����Ⱓ(����)</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtPay_dur" CLASS=FPDS140 tag="21X6Z" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL>
								</TD>
								<TD CLASS=TD5 NOWRAP>�����Ⱓ(����)</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtPay_dur_Pur" CLASS=FPDS140 tag="21X6Z" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL>
								</TD>																
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtClose_day1" CLASS=FPDS40 tag="21XX3" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL>
								</TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtPay_Month" CLASS=FPDS40 tag="21XX2" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>�޵�</LABEL>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtPay_day" CLASS=FPDS40 tag="21XX3" Title="FPDOUBLESINGLE" align=absmiddle> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL>
								</TD>																							
							</TR>						
							<TR>
								<TD CLASS=TD5 NOWRAP>���ݽŰ�����</LABEL></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" ALT="���ݽŰ�����" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTax_Biz">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>����������ȭ��</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtCash_Rate" CLASS=FPDS140 tag="21XX0" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL>
								</TD>
							<TR> 
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtPay_terms_txt" TYPE="Text" ALT="��������" MAXLENGTH="120" SIZE=100 tag="21"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ī���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCardCoCd" TYPE="Text" MAXLENGTH="10" SIZE=10 Alt="ī���" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCardCoCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCardCO">&nbsp;<INPUT NAME="txtCardCoCdNm" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>��������ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCardMemNo" TYPE="Text" ALT="��������ȣ" MAXLENGTH="20" SIZE=34 tag="21XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankCo" TYPE="Text" MAXLENGTH="10" SIZE=10 Alt="����" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBankCo">&nbsp;<INPUT NAME="txtBankCoNm" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankAcctNo" TYPE="Text" ALT="���¹�ȣ" MAXLENGTH="30" SIZE=34 tag="21XXX"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBankAcctNo"></TD>
							</TR>
							
							<%Call SubFillRemBodyTD5656(5)%>
						</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID1)">������̷µ��</a>&nbsp;|&nbsp;<a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID2)">�ŷ�ó��ȸ</a>&nbsp;|&nbsp;<a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID3)">�ŷ�ó������ȸ</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioVATinc" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioCredit" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioVATcalc" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioDepositPrice" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioSoldInspect" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioInOut" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBp_Type" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHConBp_cd" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
