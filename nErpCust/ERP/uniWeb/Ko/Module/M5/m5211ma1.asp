
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5211ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : ����B/L��� ASP															*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2003/05/27																			*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must cccccchange"								*
'* 13. History              : 1. 2000/04/18 : ȭ�� design												*
'*							  2. 2000/04/18 : Coding Start												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc ����  ********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	
<!--
'============================================  1.1.2 ���� Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
 Option Explicit
	
	Dim interface_Account

	Const BIZ_PGM_QRY_ID = "m5211mb1.asp"			'dbquery mode
	Const BIZ_PGM_SAVE_ID = "m5211mb2.asp"			'dbsave mode
	Const BIZ_PGM_DEL_ID = "m5211mb3.asp"			'dbdelete mode
	Const BIZ_PGM_POQRY_ID = "m5211mb4.asp"			'�������� 
	Const BIZ_PGM_POSTQRY_ID = "m5211mb5.asp"		'Ȯ����ư Ŭ���� 
	Const BIZ_PGM_LCQRY_ID = "m5211mb6.asp"		    'L/C ���� 
	Const IMBL_DETAIL_ENTRY_ID = "m5212ma1"         'B/C ������� ���� 
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"		    '��� ��� 
	Const IV_Payment_ID = "m5113ma1"                '���޳������ 

	Const TAB1 = 1
	Const TAB2 = 2
    '�� setting �� ���а����� ��� 
	Const gstrTransportMajor = "B9009"				'��۹�� 
	Const gstrIncotermsMajor = "B9006"				'��������	
	Const gstrPayTypeMajor = "A1006"				'�������� 
	Const gstrPayMethodMajor = "B9004"				'�������	
	Const gstrVatTypeMajor = "B9001"				'VAT���� 
	Const gstrFreightMajor = "S9007"				'�������ҹ�� 
	Const gstrPackingTypeMajor = "B9007"			'�������� 
	Const gstrOriginMajor      = "B9094"			'������ 
	Const gstrDeliveryPlceMajor = "B9095"			'�ε���� 
    Const gstrLoadingPortMajor = "B9092"			'������ 
	Const gstrDischgePortMajor = "B9092"			'������ 
	
	Const CID_POST  = 5211					'Ȯ�� 
	
	Dim  lgBlnFlgChgValue			
	Dim  lgIntGrpCount				
	Dim  lgIntFlgMode				

	Dim gSelframeFlg			                    'tab1,tab2 ����	
	Dim gblnWinEvent				

	Dim lsPosting
	Dim serverDate
	serverDate = "<%=GetSvrDate%>"
	serverDate = UniConvDateAToB(serverDate, Parent.gServerDateFormat, parent.gDateFormat) 
<!--
'==========================================  2.1.1 InitVariables()  =====================================
-->
Function InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE	
	lgBlnFlgChgValue = False	
	lgIntGrpCount = 0			
		
	gblnWinEvent = False
End Function
<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
Sub SetDefaultVal()
		
	Call SetToolbar("1110000000001111")
		
	if frm1.hdnDefaultFlg.value = "" then
	    frm1.txtBLIssueDt.text = serverDate            'B/L������ 
	end if
		
	frm1.txtLoadingDt.text = serverDate            '������ 
	frm1.hdnEnddate.value  = serverDate
	frm1.rdoPostingFlg(1).Checked = True           'Ȯ������ 
	frm1.btnPosting.disabled = true                'Ȯ�� ��ư 
	frm1.btnPosting.value = "Ȯ��"
	frm1.btnGlSel.value = "��ǥ��ȸ"
	frm1.btnGlSel.disabled = true
	frm1.txtDischgeDt.text = serverDate            '������ 
	frm1.chkPoNoCnt.Checked = True                 '���ֹ�ȣ ���� check box
	frm1.chkLcNoCnt.Checked = True                 'L/C��ȣ ���� check box
	frm1.ChkPrepay.Checked =   false                 '���ޱݿ��� ���� check box
	Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
		
	Call ClickTab1()                               '���� B/L ���� 
	frm1.txtBLNo.focus                             '��ȸ�� B/L ������ȣ 
	Set gActiveElement = document.activeElement
	interface_Account = GetSetupMod(parent.gSetupMod, "a") 'Y or N return

End Sub
	
<!--
'==========================================  2.2.2 LoadInfTB19029()  ====================================
-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>

End Sub
<!--
'===========================================  2.3.1 Tab Click ó��  =====================================
-->
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
		
	Call changeTabs(TAB1)
		
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
		
	gSelframeFlg = TAB2
End Function
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBlNoPop()  +++++++++++++++++++++++++++++++++++++++
'+	Name : OpenBlNoPop()																				+
'+	Description : B/L Reference Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBlNoPop()
	Dim strRet,IntRetCD
	Dim iCalledAspName 	
			
	If gblnWinEvent = True Or UCase(frm1.txtBLNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M5211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5211PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
			
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
			
	If strRet = "" Then
		frm1.txtBLNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBLNo.value = strRet
		frm1.txtBLNo.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenPoRef()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenPoRef()
	Dim strRet,IntRetCD
	Dim iCalledAspName  
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")  '%1 ������忡���� ������ �� �����ϴ�.
		Exit function
	End If	
			
	If frm1.rdoPostingflg(0).Checked = True then    'ȸ��ó�������̹Ƿ� ���� �Ҽ� �����ϴ�(Ȯ�� Y�̸� �����Ұ� 
		Call DisplayMsgBox("17a009","X","X","X")
		Exit Function
	End if
		
	iCalledAspName = AskPRAspName("M3111RA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111RA2", "X")
		gblnWinEvent = False
		Exit Function
	End If

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
		
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	If strRet = "" Then
		frm1.txtBLNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPORef(strRet)
	End If
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCRef()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCRef()
	
	Dim strRet,IntRetCD
	Dim arrParam(1)
	Dim iCalledAspName  
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If	
		
	if frm1.rdoPostingflg(0).Checked = True then
		Call DisplayMsgBox("17a009","X","X","X")
		Exit Function
	End if
		
	arrParam(0) = "BL"
		
	iCalledAspName = AskPRAspName("M3211RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211RA1", "X")
		gblnWinEvent = False
		Exit Function
	End If

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        		
	gblnWinEvent = False

	If strRet = "" Then
		frm1.txtBLNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetLCRef(strRet)
	End If
		
End Function

<!--
'------------------------------------------  OpenIvType()  -------------------------------------------------
-->
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 

    If frm1.txtIvType.readOnly = True  then 
        Exit Function
    end if

	arrHeader(0) = "��������"					
    arrHeader(1) = "�������¸�"					
    
    arrField(0) = "IV_TYPE_CD"						
    arrField(1) = "IV_TYPE_NM"						
    
	arrParam(0) = "��������"						
	arrParam(1) = "M_IV_TYPE"						
	arrParam(2) = Trim(frm1.txtIvType.Value)		
	arrParam(4) = "import_flg=" & FilterVar("Y", "''", "S") & " "							
	arrParam(5) = "��������"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    If arrRet(0) <> "" then
    	lgBlnFlgChgValue = True
		frm1.txtIvType.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
    end if
    frm1.txtIvType.focus
    Set gActiveElement = document.activeElement
End Function

Function OpenBpMpa()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 

	If gblnWinEvent = True Or UCase(frm1.txtPayeeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	gblnWinEvent = True

	arrHeader(0) = "����ó"
    arrHeader(1) = "����ó��"
    
    arrField(0) = "bpftn.partner_bp_cd"
    arrField(1) = "bp.bp_nm"
    
	arrParam(0) = "����ó"
	arrParam(1) = "b_biz_partner_ftn bpftn,b_biz_partner bp"
	arrParam(2) = Trim(frm1.txtPayeeCd.Value) 	 '//Trim(frm1.txtPayeeCd.Value)
	
	arrParam(4) = "bpftn.partner_bp_cd=bp.bp_cd And bpftn.partner_ftn=" & FilterVar("MPA", "''", "S") & " and bpftn.bp_cd= " & FilterVar(Trim(frm1.txtBeneficiary.Value), " " , "S") & " " 	'//Trim(frm1.txtBeneficiary.value) 
	arrParam(5) = "����ó"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	If arrRet(0) <> "" then
    	lgBlnFlgChgValue = True
		frm1.txtPayeeCd.Value = arrRet(0)
		frm1.txtPayeeNm.Value = arrRet(1)
    End If
    frm1.txtPayeeCd.focus	
	Set gActiveElement = document.activeElement
	
	gblnWinEvent = False
    
    Call txtPayeeCd_Change()
 
End Function

Function OpenBpMgs()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 

	If gblnWinEvent = True Or UCase(frm1.txtBuildCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	gblnWinEvent = True

	arrHeader(0) = "���ݰ�꼭����ó"
    arrHeader(1) = "���ݰ�꼭����ó��"
    
    arrField(0) = "bpftn.partner_bp_cd"
    arrField(1) = "bp.bp_nm"
    
	arrParam(0) = "���ݰ�꼭����ó"
	arrParam(1) = "b_biz_partner_ftn bpftn,b_biz_partner bp"
	arrParam(2) = Trim(frm1.txtBuildCd.Value)
	arrParam(4) = "bpftn.partner_bp_cd=bp.bp_cd And bpftn.partner_ftn=" & FilterVar("MBI", "''", "S") & " and bpftn.bp_cd= " & FilterVar(Trim(frm1.txtBeneficiary.Value), " " , "S") & " "
	arrParam(5) = "���ݰ�꼭����ó"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False

    If arrRet(0) <> "" then
    	lgBlnFlgChgValue = True
		frm1.txtBuildCd.Value = arrRet(0)
		frm1.txtBuildNm.Value = arrRet(1)
    End If
    frm1.txtBuildCd.focus	
	Set gActiveElement = document.activeElement
    
End Function

<!--
'------------------------------------------  OpenPpNo()  -------------------------------------------------
-->
Function OpenPpNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If frm1.ChkPrepay.Checked = false then 'or frm1.rdoPostingflg(0).Checked = True then
        Exit Function
    end if

	If gblnWinEvent = True  Then Exit Function

	if Trim(frm1.txtCurrency.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "ȭ��","X")
		Exit Function
	elseif Trim(frm1.txtBeneficiary.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "������","X")
		Exit Function
	end if
	
	gblnWinEvent = True

	arrParam(0) = "���ޱݹ�ȣ"	
	arrParam(1) = "F_PRPAYM"
	
	arrParam(2) = ""
	
	arrParam(4) = "DOC_CUR =  " & FilterVar(Trim(frm1.txtCurrency.Value), " " , "S") & "  And BP_CD =  " & FilterVar(Trim(frm1.txtBeneficiary.Value), " " , "S") & "  AND BAL_AMT > 0"
	arrParam(5) = "���ޱݹ�ȣ"			
	
    arrField(0) = "PRPAYM_NO"
    arrField(1) = "F2" & parent.gColSep & "PRPAYM_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "BAL_AMT"
    
    arrHeader(0) = "���ޱݹ�ȣ"		
    arrHeader(1) = "���ޱ�"		
    arrHeader(2) = "���ޱ�ȭ��"
    arrHeader(3) = "���ޱ��ܾ�"
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	'If arrRet(0) = "" Then
		Exit Function
	'Else
	'	frm1.txtPrePayNo.Value 		= arrRet(0)
	'	frm1.txtPrePayDocAmt.Text 	= arrRet(3)
	'	lgBlnFlgChgValue = True
	'	Call changePpNo()
	'End If	

End Function

<!-- '------------------------------------------  OpenBizArea()  -------------------------------------------------
-->
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtTaxBizArea.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "���ݽŰ�����"	
	arrParam(1) = "B_Tax_Biz_Area"
	arrParam(2) = Trim(frm1.txtTaxBizArea.Value)
	arrParam(4) = ""
	arrParam(5) = "���ݽŰ�����"			
	
    arrField(0) = "Tax_Biz_Area_Cd"
    arrField(1) = "Tax_Biz_Area_Nm"
    
    arrHeader(0) = "���ݽŰ�����"
    arrHeader(1) = "���ݽŰ������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtTaxBizArea.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtTaxBizArea.Value = arrRet(0)
		frm1.txtTaxBizAreaNm.Value = arrRet(1)
		frm1.txtTaxBizArea.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If	

End Function
<!--
'------------------------------------------  OpenLoanNo()  -------------------------------------------------
-->
Function OpenLoanNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	if interface_Account = "N" then
		Exit Function
	End if

	If gblnWinEvent = True Or UCase(frm1.txtLoanNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	if Trim(frm1.txtCurrency.Value) = "" then'or Trim(frm1.txtBeneficiary.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "ȭ��","X")
		Exit Function
	end if
	
	gblnWinEvent = True

	arrParam(0) = "���Աݹ�ȣ"	
	arrParam(1) = "F_LOAN"
	
	arrParam(2) = Trim(frm1.txtLoanNo.Value)
	
	'arrParam(4) = "DOC_CUR = '" & Trim(frm1.txtCurrency.Value) & "'"
	arrParam(4) = "DOC_CUR =  " & FilterVar(Trim(frm1.txtCurrency.Value), " " , "S") & "  AND LOAN_BAL_AMT > 0"
	arrParam(5) = "���Աݹ�ȣ"			
	
    arrField(0) = "LOAN_NO"
    arrField(1) = "F2" & parent.gColSep & "LOAN_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "LOAN_BAL_AMT"
    
    arrHeader(0) = "���Աݹ�ȣ"		
    arrHeader(1) = "���Ա�"		
    arrHeader(2) = "���Ա�ȭ��"
    arrHeader(3) = "���Ա��ܾ�"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtLoanNo.Value 	= arrRet(0)
		frm1.txtLoanAmt.Text 	= arrRet(3)
		frm1.txtLoanNo.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue 		= True
		'Call changeLoanNo()
	End If	

End Function
<!--
'+++++++++++++++++++++++++++++++++++++++++++++  OpenPurGrp()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "���Դ��"						
	arrParam(1) = "B_PURCHASE_GROUP"				
	arrParam(2) = Trim(frm1.txtPurGrp.value)		
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "���Դ��"						

	arrField(0) = "PUR_GRP"							
	arrField(1) = "PUR_GRP_NM"						

	arrHeader(0) = "���Դ��"						
	arrHeader(1) = "���Դ���"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtPurGrp.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrp.value = arrRet(0)
		frm1.txtPurGrpNm.value = arrRet(1)
		frm1.txtPurGrp.focus	
		Set gActiveElement = document.activeElement

		lgBlnFlgChgValue = True
	End If
End Function

<!--
'------------------------------------------  OpenPayType()  -------------------------------------------------
-->
Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	if Trim(frm1.txtPayMethod.Value) = "" then
		Call DisplayMsgBox("17a002", parent.VB_YES_NO,"�������", "X")
		Exit Function
	End if

	If gblnWinEvent = True Or UCase(frm1.txtPayType.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "��������"
	arrParam(1) = "B_MINOR,B_CONFIGURATION," _
	& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""_
		& "And MINOR_CD= " & FilterVar(Trim(frm1.txtPayMethod.Value), " " , "S") & "  And SEQ_NO>=2)C"
	
	arrParam(2) = Trim(frm1.txtPayType.Value)

	if Trim(frm1.txtPayMethod.Value) <> "" then
		
		arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
					& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("P", "''", "S") & " )"
	else
	
		arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
					& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("P", "''", "S") & " )"
	end if

	arrParam(5) = "��������"
    
	arrField(0) = "B_MINOR.MINOR_CD"						
	arrField(1) = "left(B_MINOR.MINOR_NM,100)"	

    arrHeader(0) = "��������"
    arrHeader(1) = "����������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtPayType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPayType.Value = arrRet(0)
		frm1.txtPayTypeNm.Value = arrRet(1)
		frm1.txtPayType.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue 		= True
	End If	
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenMinorCd()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenMinorCd()																				+
'+	Description : Minor Code PopUp Window Call                                                          +	
'+  (�ڵ尪(aa.value),�̸���(bb.value),"��۹��(display�÷���)","��۹����(display�÷���)",code��)    +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
	Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strPopPosNm, strMajorCd)
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = strPopPos								
		arrParam(1) = "B_Minor"								
		arrParam(2) = Trim(strMinorCD)						
'		arrParam(3) = Trim(strMinorNM)						
		arrParam(4) = "MAJOR_CD= " & FilterVar(Trim(strMajorCd), " " , "S") & " "		
		arrParam(5) = strPopPos								

		arrField(0) = "Minor_CD"							
		arrField(1) = "Minor_NM"							

		arrHeader(0) = strPopPos							
		arrHeader(1) = strPopPosNm							

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetMinorCd(strMajorCd, arrRet)
		End If
	End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenCountry()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenCountry(strCntryCD, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "����"							
	arrParam(1) = "B_COUNTRY"						
	arrParam(2) = Trim(strCntryCD)					
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "����"							

	arrField(0) = "COUNTRY_CD"						
	arrField(1) = "COUNTRY_NM"						

	arrHeader(0) = "����"						
	arrHeader(1) = "������"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCountry(strPopPos, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++  OpenBizPartner()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBizPartner(strBizPartnerCD, strBizPartnerNM, strPopPos, strPopPosNm)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos				
	arrParam(1) = "B_BIZ_PARTNER"					
	arrParam(2) = Trim(strBizPartnerCD)				
'		arrParam(3) = Trim(strBizPartnerNM)				
	arrParam(4) = ""								
	arrParam(5) = strPopPos				

	arrField(0) = "BP_CD"				
	arrField(1) = "BP_NM"				

	arrHeader(0) = strPopPos			
	arrHeader(1) = strPopPosNm			

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(strPopPos, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++++  OpenUnit()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenUnit(strUnitCD, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
	if strPopPos = "WT" then
		arrParam(0) = "�߷�����"				
		arrParam(1) = "B_UNIT_OF_MEASURE"		
		arrParam(2) = Trim(strUnitCD)			
		arrParam(3) = ""						
		arrParam(4) = "DIMENSION= " & FilterVar(Trim(strPopPos), " " , "S") & " "	
		arrParam(5) = "�߷�����"						
	
		arrField(0) = "UNIT"							
		arrField(1) = "UNIT_NM"							
	
		arrHeader(0) = "�߷�����"					
		arrHeader(1) = "�߷�������"					
	else
		arrParam(0) = "��������"						
		arrParam(1) = "B_UNIT_OF_MEASURE"				
		arrParam(2) = Trim(strUnitCD)					
		arrParam(3) = ""								
		arrParam(4) = "DIMENSION= " & FilterVar(Trim(strPopPos), " " , "S") & " "	
		arrParam(5) = "��������"						
	
		arrField(0) = "UNIT"							
		arrField(1) = "UNIT_NM"							
	
		arrHeader(0) = "��������"					
		arrHeader(1) = "����������"					
	End if
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUnit(strPopPos, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenDischgePort()  +++++++++++++++++++++++++++++++++++++
-->
Function OpenDischgePort()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "������"				
	arrParam(1) = "B_Minor"					
	arrParam(2) = Trim(frm1.txtDischgePort.Value)			
'		arrParam(3) = Trim(frm1.txtDischgePortNm.Value)			
	arrParam(4) = "MAJOR_CD= " & FilterVar(Trim(gstrDisChgePortMajor), " " , "S") & " "	
	arrParam(5) = "������"				

	arrField(0) = "Minor_CD"				
	arrField(1) = "Minor_NM"				

	arrHeader(0) = "������"				
	arrHeader(1) = "�����׸�"			

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtdischgePort.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtdischgePort.Value = arrRet(0)
		frm1.txtDischgePortNm.Value = arrRet(1)
		frm1.txtdischgePort.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenTaxOffice()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenTaxOffice()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "�Ű�����"					
	arrParam(1) = "B_TAX_OFFICE"					
	arrParam(2) = Trim(frm1.txtTaxOfficeCd.value)	
'		arrParam(3) = Trim(frm1.txtTaxOfficeCdNm.value)	
	arrParam(4) = ""								
	arrParam(5) = "�Ű�����"					

	arrField(0) = "TAX_OFFICE_CD"					
	arrField(1) = "TAX_OFFICE_NM"					

	arrHeader(0) = "�Ű�����"					
	arrHeader(1) = "�Ű�������"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtTaxOfficeCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtTaxOfficeCd.value = arrRet(0)
		frm1.txtTaxOfficeCdNm.value = arrRet(1)
		frm1.txtTaxOfficeCd.focus
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If
End Function

Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtGlNo.value)          '��ǥ��ȣ 
	arrParam(1) = Trim(frm1.txtIvNo.value)          '���Թ�ȣ 
	'arrParam(2) = Trim(frm1.txtGrpCd.value)
	'arrParam(3) = Trim(frm1.txtGrpNm.value)
	
   If frm1.hdnGlType.Value = "A" Then               'ȸ����ǥ�˾� 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "T" Then          '������ǥ�˾� 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '���� ��ǥ�� �������� �ʾҽ��ϴ�. 
    End if

	gblnWinEvent = False
	
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++  OpenVat()  *++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenVat()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "VAT����"						
	arrParam(1) = "B_MINOR,b_configuration"			
	
	'arrParam(2) = Trim(frm1.txtVattype.Value)	<%' Code Condition%>
		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_configuration.SEQ_NO=1"							<%' Where Condition%>
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd"
	arrParam(5) = "VAT����"						
	
    arrField(0) = "b_minor.MINOR_CD"				
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"		
    
    arrHeader(0) = "VAT����"						
    arrHeader(1) = "VAT���¸�"					
    arrHeader(2) = "Reference"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		'frm1.txtVatType.Value 	= arrRet(0)
		'frm1.txtVatRate.Text 	= arrRet(2)
		lgBlnFlgChgValue 		= True
	End If	
	
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetPORef()  ++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetPORef(strRet)
	Dim strVal
	    
    Call ggoOper.ClearField(Document, "A")			

	frm1.txtPONo.value = strRet

	strVal = BIZ_PGM_POQRY_ID & "?txtPONo=" & Trim(frm1.txtPONo.value)	
		
	if Trim(frm1.txtBlIssueDt.text) <> "" then
	    strVal = strVal & "&hdnBlIssueDt=" & Trim(frm1.txtBlIssueDt.text)	
	else
	    strVal = strVal & "&hdnBlIssueDt=" & serverDate
	end if
	    strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
			
	frm1.hdnDefaultFlg.value = "Y"
	Call SetDefaultVal
	Call ChangeTag(False)
			
    If  LayerShowHide(1) = False Then
       	Exit Function
    End If

	Call RunMyBizASP(MyBizASP, strVal)									
End Function
	
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetLCRef()  ++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetLCRef(strRet)
	Dim strVal
        
    Call ggoOper.ClearField(Document, "A")								

	frm1.txtLCNo.value = strRet

	strVal = BIZ_PGM_LCQRY_ID & "?txtLCNo=" & Trim(frm1.txtLCNo.value)	
		
	if Trim(frm1.txtBlIssueDt.text) <> "" then
	    strVal = strVal & "&hdnBlIssueDt=" & Trim(frm1.txtBlIssueDt.text)	
	else
	    strVal = strVal & "&hdnBlIssueDt=" & serverDate
	end if
		
	frm1.hdnDefaultFlg.value = "Y"
	Call SetDefaultVal
	Call ChangeTag(False)

    If  LayerShowHide(1) = False Then
       	Exit Function
    End If
		
	Call RunMyBizASP(MyBizASP, strVal)									
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
-->
Function SetMinorCd(strMajorCd, arrRet)
	Select Case strMajorCd
				
		Case gstrTransportMajor
			frm1.txtTransport.Value = arrRet(0)
			frm1.txtTransportNm.Value = arrRet(1)
			frm1.txtTransport.focus	
		

		Case gstrPayTypeMajor
			frm1.txtPayType.Value = arrRet(0)
			frm1.txtPayTypeNm.Value = arrRet(1)
			frm1.txtPayType.focus	

		Case gstrPayMethodMajor
			frm1.txtPayMethod.Value = arrRet(0)
			frm1.txtPayMethodNm.Value = arrRet(1)
			frm1.txtPayMethod.focus	
				
		Case gstrVatTypeMajor
			'frm1.txtVatType.Value = arrRet(0)

		Case gstrFreightMajor
			frm1.txtFreight.Value = arrRet(0)
			frm1.txtFreightNm.Value = arrRet(1)
			frm1.txtFreight.focus	
				
		Case gstrPackingTypeMajor
			frm1.txtPackingType.Value = arrRet(0)
			frm1.txtPackingTypeNm.Value = arrRet(1)
			frm1.txtPackingType.focus

		Case gstrDeliveryPlceMajor
			frm1.txtDeliveryPlce.Value = arrRet(0)
			frm1.txtDeliveryPlceNm.Value = arrRet(1)
			frm1.txtDeliveryPlce.focus
				
		Case gstrLoadingPortMajor
			frm1.txtLoadingPort.Value = arrRet(0)
			frm1.txtLoadingPortNm.Value = arrRet(1)
			frm1.txtLoadingPort.focus
				
		Case gstrOriginMajor
			frm1.txtOrigin.Value = arrRet(0)
			frm1.txtOriginNm.Value = arrRet(1)
			frm1.txtOrigin.focus
		Case Else
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
		
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++  SetCountry()  +++++++++++++++++++++++++++++++++++++++++++
-->
Function SetCountry(strPopPos, arrRet)
	Select Case UCase(strPopPos)
		Case "VESSEL"
			frm1.txtVesselCntry.Value = arrRet(0)
			frm1.txtVesselCntry.focus

		Case "TRANSHIP"
			frm1.txtTranshipCntry.Value = arrRet(0)
			frm1.txtTranshipCntry.focus

		Case "ORIGIN"
			frm1.txtOriginCntry.Value = arrRet(0)
			frm1.txtOriginCntry.focus

		Case Else
	End Select
	Set gActiveElement = document.activeElement

	lgBlnFlgChgValue = True
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++  SetBizPartner()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function SetBizPartner(strPopPos, arrRet)
	Select Case (strPopPos)
		Case "������"
			frm1.txtAgent.Value = arrRet(0)
			frm1.txtAgentNm.Value = arrRet(1)
			frm1.txtAgent.focus
				
		Case "������"
			frm1.txtManufacturer.Value = arrRet(0)
			frm1.txtManufacturerNm.Value = arrRet(1)
			frm1.txtManufacturer.focus
				
		Case "����ȸ��"
			frm1.txtForwarder.Value = arrRet(0)
			frm1.txtForwarderNm.Value = arrRet(1)
			frm1.txtForwarder.focus
				
		Case Else
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  SetUnit()  ++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetUnit(strPopPos, arrRet)
	Select Case UCase(strPopPos)
		Case "WT"
			frm1.txtWeightUnit.Value = arrRet(0)
			frm1.txtWeightUnit.focus

		Case "WD"
			frm1.txtVolumnUnit.Value = arrRet(0)
			frm1.txtVolumnUnit.focus

		Case Else
	End Select
	lgBlnFlgChgValue = True
End Function

<%'======================================   GetPayDt()  =====================================
'	Name : GetPayDt()
'	Description : ���ҿ������� �����´�.
'==================================================================================================== %>
Sub GetPayDt()
   	Dim strSelectList, strFromList, strWhereList
	Dim strPayeeCd, strBlIssueDt,temp
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp

    	strPayeeCd  = frm1.txtPayeeCd.value                   '����ó	
    	temp    = UNIConvDate(frm1.txtBlIssueDt.text)         'B/L������ 
		strBlIssueDt = mid(temp,1,4)
		strBlIssueDt = strBlIssueDt & mid(temp,6,2)
		strBlIssueDt = strBlIssueDt & mid(temp,9,2) 
		<%'����ó�� ���� �׷��� ��� ��ϵǾ� �ִ� ��� �����ڵ忡 ������ rule�� ������ %>
    
	
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetPayDt( " & FilterVar(Trim(strPayeeCd), " " , "S") & " ,  " & FilterVar(Trim(strBlIssueDt), " " , "S") & " ) "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
	    frm1.txtSetlmnt.text = UNIFormatDate(arrTemp(1))
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Sub
		End If

		frm1.txtSetlmnt.text = ""

	End if
End Sub
'======================================   GetTaxBizArea()  =====================================
'	Name : GetTaxBizArea()
'	Description : ���ݽŰ������� �����´� 
'==================================================================================================== %>

Sub GetTaxBizArea(Byval strFlag)
   	Dim strSelectList, strFromList, strWhereList
	Dim strBilltoParty, strSalesGrp, strTaxBizArea
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp
	
     
	If strFlag = "NM" Then                              '���ݽŰ����� ����� �̸����� �����´� 
		strTaxBizArea = frm1.txtTaxBizArea.value
	Else
		strBilltoParty = frm1.txtBuildCd.value          '��꼭 ����ó 
		strSalesGrp    = frm1.txtPurGrp.value            '���ű׷� 
		<%'����ó�� ���� �׷��� ��� ��ϵǾ� �ִ� ��� �����ڵ忡 ������ rule�� ������ %>
		If Len(strBillToParty) > 0 And Len(strSalesGrp) > 0	Then strFlag = "*"
	End if
	
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetTaxBizArea ( " & FilterVar(Trim(strBilltoParty), " " , "S") & " ,  " & FilterVar(Trim(strSalesGrp), " " , "S") & " ,  " & FilterVar(Trim(strTaxBizArea), " " , "S") & " ,  " & FilterVar(Trim(strFlag), " " , "S") & " ) "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		frm1.txtTaxBizArea.value = arrTemp(1)
		frm1.txtTaxBizAreaNm.value = arrTemp(2)
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Sub
		End If

		' ���� �Ű� ������� Editing�� ��� 
		'If strFlag = "NM" Then
		'	If Not OpenBillHdr(3) Then
			frm1.txtTaxBizArea.value = ""
			frm1.txtTaxBizAreaNm.value = ""
		'	End if
		'End if
	End if
End Sub


'======================================   CheckPrePayedAmtYN()  =====================================
'	Name : CheckPrePayedAmtYN()
'	Description : ���ޱݿ��θ� üũ�Ѵ�.
'============================================================================================
Sub CheckPrePayedAmtYN()
	Dim strSelectList,strFromList,strWhereList
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim iCount

	strSelectList = " count(*) "
	strFromList = " F_PRPAYM " 
	strWhereList = " BP_CD= " & FilterVar(Trim(frm1.txtPayeeCd.Value), " " , "S") & "  and DOC_CUR= " & FilterVar(Trim(frm1.txtCurrency.value), " " , "S") & " and BAL_AMT > 0 AND CONF_FG = " & FilterVar("C", "''", "S") & " "

	Call CommonQueryRs(strSelectList,strFromList,strWhereList, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Err.number <> 0 Then
		'MsgBox Err.description 
		Exit Sub
	End If

	 iCount = Split(lgF0, Chr(11))
	    
	if UNICDbl(Trim(iCount(0))) > 0 then
		frm1.ChkPrepay.checked = true
	else
		frm1.ChkPrepay.checked = false
	End if

End Sub

<!--
'==========================================================================================
'   Event Name : ChangeCurOrDt()
'   Event Desc : B/L�����Ϻ����� ȯ�������� 
'==========================================================================================
-->
Function ChangeCurOrDt()
   
    Dim strVal
    Err.Clear                                                               '��: Protect system from crashing

    frm1.hdnDefaultFlg.value = "Y"

    if Trim(frm1.hdnchangeflg.value) = "Y" then
        frm1.hdnchangeflg.value = "N"
        Exit Function
    end if
	
	if Trim(frm1.txtCurrency.value) = parent.gCurrency  then
	    frm1.txtXchRate.text = UNIFormatNumber(1,ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	    Exit Function
	end if

    With frm1
		
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & "LookupDailyExRt"			
		strVal = strVal & "&txtBlIssueDt=" & Trim(.txtBlIssueDt.text)	
        strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)	
				
    End With
	
    If  LayerShowHide(1) = False Then
      	Exit Function
    End If

	Call RunMyBizASP(MyBizASP, strVal)
        
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.Value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
<!--
'++++++++++++++++++++++++++++++++++++++++++++  changePpNo()  *++++++++++++++++++++++++++++++++++++++++++
-->
Function changePpNo()

	If Trim(frm1.txtPrePayNo.Value) <> "" Then
		Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"N")
	Else
		frm1.txtPrePayDocAmt.Text = "0"
		Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"Q")
	End if
	
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++  changeLoanNo()  *++++++++++++++++++++++++++++++++++++++++++
-->
Function changeLoanNo()

	If Trim(frm1.txtLoanNo.Value) = "" Then
		frm1.txtLoanAmt.Text = 0		
	End If
	
End Function

Function XchRateOk()

End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++  dbRefQueryOK()  *++++++++++++++++++++++++++++++++++++++++++
-->
Function dbRefQueryOK()

	Call SetToolbar("1110100000001111")
	Call ChangeTag(False)	
	'Call setLoan()
	
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++  setLoan()  *++++++++++++++++++++++++++++++++++++++++++
-->
Function setLoan()

	'if frm1.hdnLoanflg.value = "Y" then
		if interface_Account = "N" then		
			Call ggoOper.SetReqAttr(frm1.txtLoanNo,"D")	
			Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"D")		
		else
			Call ggoOper.SetReqAttr(frm1.txtLoanNo,"D")	
			Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"Q")		
		end if
	'else
	'	Call ggoOper.SetReqAttr(frm1.txtLoanNo,"Q")	
	'	Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"Q")		
	'end if

end Function
<!--
'--------------------------------------------------------------------
'		Field�� Tag�Ӽ��� Protect�� ��ȯ,���� ��Ű�� �Լ� 
'--------------------------------------------------------------------
-->

 Function ChangeTag(Byval Changeflg)

	if Changeflg = true then  'Ȯ���̸� ��ü lock
		''''ù��° �� 
		Call ggoOper.SetReqAttr(frm1.txtBLDocNo,"Q")
		Call ggoOper.SetReqAttr(frm1.chkPoNoCnt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtLoadingDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBlIssueDt,"Q")
		Call ggoOper.SetReqAttr(frm1.chkLcNoCnt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDischgeDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtSetlmnt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTransport,"Q")
		Call ggoOper.SetReqAttr(frm1.txtForwarder,"Q")
		Call ggoOper.SetReqAttr(frm1.txtVesselNm,"Q")
		Call ggoOper.SetReqAttr(frm1.txtVoyageNo,"Q")
		Call ggoOper.SetReqAttr(frm1.txtVesselCntry,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTotPackingCnt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDischgePort,"Q")
		Call ggoOper.SetReqAttr(frm1.txtXchRate,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayType,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayeeCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBuildCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayTermsTxt,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtVatType,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtVatRate,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtPrePayNo,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtLoanNo,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtCashAmt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtIvType,"Q")
		''''�ι�° �� 
		Call ggoOper.SetReqAttr(frm1.txtPackingType,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPackingTxt,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtGrossWeight,"Q")
		Call ggoOper.SetReqAttr(frm1.txtNetWeight,"Q")
		Call ggoOper.SetReqAttr(frm1.txtWeightUnit,"Q")
		Call ggoOper.SetReqAttr(frm1.txtContainerCnt,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtGrossVolumn,"Q")
		Call ggoOper.SetReqAttr(frm1.txtVolumnUnit,"Q")
		Call ggoOper.SetReqAttr(frm1.txtFreight,"Q")
		Call ggoOper.SetReqAttr(frm1.txtFreightPlce,"Q")
		Call ggoOper.SetReqAttr(frm1.txtFinalDest,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDeliveryPlce,"Q")
		Call ggoOper.SetReqAttr(frm1.txtReceiptPlce,"Q")
		Call ggoOper.SetReqAttr(frm1.txtLoadingPort,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTranshipCntry,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTranshipDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtOrigin,"Q")
		Call ggoOper.SetReqAttr(frm1.txtOriginCntry,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBLIssuePlce,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBLIssueCnt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtAgent,"Q")
		Call ggoOper.SetReqAttr(frm1.txtManufacturer,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTaxBizArea,"Q")
		Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
		Call ggoOper.SetReqAttr(frm1.txtRemark,"Q")
		
		if frm1.ChkPrepay.checked = false then
			Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
		else 		
			Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"N")
		end if
	else
		''''ù��° �� 
		Call ggoOper.SetReqAttr(frm1.txtBLDocNo,"N")         'B/L��ȣ 
		Call ggoOper.SetReqAttr(frm1.chkPoNoCnt,"D")         '���ֹ�ȣ���� 
		Call ggoOper.SetReqAttr(frm1.txtLoadingDt,"N")       '������ 
		Call ggoOper.SetReqAttr(frm1.txtBlIssueDt,"N")       'B/L������ 
		Call ggoOper.SetReqAttr(frm1.chkLcNoCnt,"D")         'L/C��ȣ���� 
		Call ggoOper.SetReqAttr(frm1.txtDischgeDt,"D")       '������ 
		Call ggoOper.SetReqAttr(frm1.txtSetlmnt,"D")         '���ҿ����� 
		Call ggoOper.SetReqAttr(frm1.txtTransport,"D")       '��۹�� 
		Call ggoOper.SetReqAttr(frm1.txtForwarder,"D")       '����ȸ�� 
		Call ggoOper.SetReqAttr(frm1.txtVesselNm,"D")        'Vessel�� 
		Call ggoOper.SetReqAttr(frm1.txtVoyageNo,"D")        '������ȣ 
		Call ggoOper.SetReqAttr(frm1.txtVesselCntry,"D")     '���ڱ��� 
		Call ggoOper.SetReqAttr(frm1.txtTotPackingCnt,"D")   '�����尹�� 
		Call ggoOper.SetReqAttr(frm1.txtDischgePort,"D")     '������ 
		
		if parent.gCurrency <> Trim(frm1.txtCurrency.value) then
		    Call ggoOper.SetReqAttr(frm1.txtXchRate,"N")         'ȯ�� 
		else
		   Call ggoOper.SetReqAttr(frm1.txtXchRate,"Q")         'ȯ�� 
		end if
		Call ggoOper.SetReqAttr(frm1.txtPayType,"D")         '�������� 
		Call ggoOper.SetReqAttr(frm1.txtPayeeCd,"N")         '����ó 
		Call ggoOper.SetReqAttr(frm1.txtBuildCd,"N")         '���ݰ�꼭����ó 
		Call ggoOper.SetReqAttr(frm1.txtPayTermsTxt,"D")     '��ݰ������� 
		'Call ggoOper.SetReqAttr(frm1.txtVatType,"N")
		'Call ggoOper.SetReqAttr(frm1.txtVatRate,"N")
		'Call ggoOper.SetReqAttr(frm1.txtPrePayNo,"D")
		'if Trim(frm1.txtPrePayNo.value) <> "" then
			'Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"N")
		'else
			'Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"Q")
		'End if
		'Call ggoOper.SetReqAttr(frm1.txtLoanNo,"D")          '���Աݹ�ȣ 
		'if interface_Account = "N" then
		'	Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"D")     '���Ա� 
		'else
		'	Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"Q")
		'End if

		'Call ggoOper.SetReqAttr(frm1.txtCashAmt,"D")        
		Call ggoOper.SetReqAttr(frm1.txtIvType,"N")          '�������� 
		''''�ι�° �� 
		Call ggoOper.SetReqAttr(frm1.txtPackingType,"D")
		Call ggoOper.SetReqAttr(frm1.txtPackingTxt,"D")
		'Call ggoOper.SetReqAttr(frm1.txtGrossWeight,"D")
		Call ggoOper.SetReqAttr(frm1.txtNetWeight,"D")
		Call ggoOper.SetReqAttr(frm1.txtWeightUnit,"D")
		Call ggoOper.SetReqAttr(frm1.txtContainerCnt,"D")
		'Call ggoOper.SetReqAttr(frm1.txtGrossVolumn,"D")
		Call ggoOper.SetReqAttr(frm1.txtVolumnUnit,"D")
		Call ggoOper.SetReqAttr(frm1.txtFreight,"D")
		Call ggoOper.SetReqAttr(frm1.txtFreightPlce,"D")
		Call ggoOper.SetReqAttr(frm1.txtFinalDest,"D")
		Call ggoOper.SetReqAttr(frm1.txtDeliveryPlce,"D")
		Call ggoOper.SetReqAttr(frm1.txtReceiptPlce,"D")
		Call ggoOper.SetReqAttr(frm1.txtLoadingPort,"D")
		Call ggoOper.SetReqAttr(frm1.txtTranshipCntry,"D")
		Call ggoOper.SetReqAttr(frm1.txtTranshipDt,"D")
		Call ggoOper.SetReqAttr(frm1.txtOrigin,"D")
		Call ggoOper.SetReqAttr(frm1.txtOriginCntry,"D")
		Call ggoOper.SetReqAttr(frm1.txtBLIssuePlce,"D")
		Call ggoOper.SetReqAttr(frm1.txtBLIssueCnt,"D")
		Call ggoOper.SetReqAttr(frm1.txtAgent,"D")
		Call ggoOper.SetReqAttr(frm1.txtManufacturer,"D")
		Call ggoOper.SetReqAttr(frm1.txtTaxBizArea,"D")
		Call ggoOper.SetReqAttr(frm1.txtRemark,"D")
	End if 
	
End Function

<!--
'=============================================  2.5.1 LoadBlDtl()  ======================================
-->
Function LoadBlDtl()
	Dim strDtlOpenParam
	Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then          
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End if
	    	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	WriteCookie "BlNo", UCase(Trim(frm1.txtBLNo.value))
	WriteCookie "PoNo", UCase(Trim(frm1.txtPONo.value))
		
	PgmJump(IMBL_DETAIL_ENTRY_ID)

End Function

<!--
'=============================================  2.5.1 LoadChargeHdr()  ======================================
-->
Function LoadChargeHdr()
	Dim strHdrOpenParam
	Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then          
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End if
	    	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    WriteCookie "Process_Step" , "VB"
	WriteCookie "Po_No" , Trim(frm1.txtBLNo.value)
	WriteCookie "Pur_Grp", Trim(frm1.txtPurGrp.Value)
	WriteCookie "Po_Cur", Trim(frm1.txtCurrency.Value)
	WriteCookie "Po_Xch", Trim(frm1.txtXchRate.Value)
		
	PgmJump(CHARGE_HDR_ENTRY_ID)

End Function
	
<!--
'=============================================  LoadIvPayment()  ======================================
-->
Function LoadIvPayment()
	Dim strHdrOpenParam
	Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then          
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End if
	    	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    WriteCookie "txtIvNo" , Trim(frm1.txtIvNo.value)
		
	PgmJump(IV_Payment_ID)

End Function
	
<!--
'============================================  2.5.2 OpenCookie()  ======================================
-->
Function OpenCookie()
	frm1.txtBLNo.value = ReadCookie("BlNo")
	WriteCookie "BlNo", ""
End Function
<!--
'============================================ ValidDateCheckLocal()  ======================================
-->
 Function ValidDateCheckLocal(pObjFromDt, pObjToDt)

	ValidDateCheckLocal = False

	If Len(Trim(pObjToDt.Text)) And Len(Trim(pObjFromDt.Text)) Then

		If UniConvDateToYYYYMMDD(pObjFromDt.Text,parent.gDateFormat,"") > UniConvDateToYYYYMMDD(pObjToDt.Text,parent.gDateFormat,"") Then
			ClickTab1()
			Call DisplayMsgBox("970023","X", pObjToDt.Alt, pObjFromDt.Alt)
			pObjToDt.Focus
            Set gActiveElement = document.activeElement                            
			
			Exit Function
		End If

	End If

	ValidDateCheckLocal = True

End Function

<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	Call LoadInfTB19029()							
	'Call AppendNumberRange("0","0","9999999999")	
	Call AppendNumberRange("1","0","999")			'tag 5��°�� 1�ΰ�� �ִ밪 (�����̳ʼ�)
	Call AppendNumberRange("2","0","99")			'tag 5��°�� 2�ΰ�� �ִ밪 (B/L����μ�)
	Call AppendNumberPlace("7","2","0")
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart) '��ȸ ���� ���� 
	Call ggoOper.LockField(Document, "N")			
	
    frm1.hdnDefaultFlg.value = ""
	Call SetDefaultVal
	Call InitVariables
	Call OpenCookie()		
	If UCase(Trim(frm1.txtBLNo.value)) <> "" Then   '��Ŷ���� ������ �ٷ� �����Ѵ� 
		frm1.hdnchangeflg.value = "Y"
		Call MainQuery
	End If
        
    frm1.hdnchangeflg.value = "Y"
	gSelframeFlg = TAB1
	gIsTab     = "Y" 
    gTabMaxCnt = 2                                   'tab ���� 
		
End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	 
End Sub
<!--
'==========================================================================================
'   Event Name : OCX Event
'==========================================================================================
-->
Sub txtLoadingDt_DblClick(Button)
	if Button = 1 then
		frm1.txtLoadingDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoadingDt.focus
	End if
End Sub

Sub txtLoadingDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtSetlmnt_DblClick(Button)
	if Button = 1 then
		frm1.txtSetlmnt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtSetlmnt.focus
	End if
End Sub

Sub txtSetlmnt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtVatRate_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtPrePayDocAmt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtTotPackingCnt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtXchRate_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtLoanAmt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtCashAmt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtPayeeCd_Change()

	call CheckPrePayedAmtYN()
	
    if Trim(frm1.txtBlIssueDt.Text) = ""  then
	    Call DisplayMsgBox("970021","X","B/L������","X")
	    Exit Sub
	End if	
	
	if Trim(frm1.txtPayeeCd.value) = ""  then
	    Call DisplayMsgBox("970021","X","����ó","X")
	    Exit Sub
    End if
	
	Call GetPayDt()
	
	lgBlnFlgChgValue = true	
End Sub

Sub ChangePayeeCd()
	
	Call CheckPrePayedAmtYN() '����ó�� �ٲ�� ���ޱݿ��θ� �ٽ� üũ�ϱ� ���� 

    if Trim(frm1.txtBlIssueDt.Text) = ""  then
	    Call DisplayMsgBox("970021","X","B/L������","X")
	    Exit Sub
	End if	
	
	if Trim(frm1.txtPayeeCd.value) = ""  then
	    Call DisplayMsgBox("970021","X","����ó","X")
	    Exit Sub
    End if
	
	Call GetPayDt()
	
	lgBlnFlgChgValue = true	
End Sub

Sub ChangeCurrency()
	Call CheckPrePayedAmtYN() 'ȭ�� �ٲ�� ���ޱݿ��θ� �ٽ� üũ�ϱ� ���� 
End Sub

Sub txtTranshipDt_DblClick(Button)
	if Button = 1 then
		frm1.txtTranshipDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTranshipDt.focus
	End if
End Sub

Sub txtTranshipDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtDischgeDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDischgeDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDischgeDt.focus
	End if
End Sub

Sub txtDischgeDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtGrossWeight_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtNetWeight_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtContainerCnt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtGrossVolumn_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtBLIssueCnt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtBlIssueDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBlIssueDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBlIssueDt.focus
	End if
End Sub


'Sub txtBlIssueDt_onBlur()
'	lgBlnFlgChgValue = true
'End Sub 

Sub txtBlIssueDt_Change()

    lgBlnFlgChgValue = true
       
    if Trim(frm1.txtBlIssueDt.Text) = ""  then
	    Exit Sub
	End if	
	
	Call GetPayDt()
		
	if Trim(frm1.txtCurrency.value) <> "" and Trim(frm1.txtCurrency.value) <> parent.gCurrency  then
	    Call ChangeCurOrDt()
	elseif Trim(frm1.txtCurrency.value) = parent.gCurrency  then
	    frm1.txtXchRate.text = UNIFormatNumber(1,ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	end if
	
	lgBlnFlgChgValue = true	
		
End Sub

<!--
'=======================================  3.2.1 chkPoNoCnt_onchange()  ======================================
'=	Event Name : chkPoNoCnt_onpropertychange																		=
'=	Event Desc : check box ������ �������� true 															=
'========================================================================================================
-->
Sub chkPoNoCnt_onpropertychange()
	lgBlnFlgChgValue = true	
End Sub

<!--
'=======================================  3.2.1 chkLcNoCnt_onchange()  ======================================
'=	Event Name : chkLcNoCnt_onpropertychange																		=
'=	Event Desc :																						=
'========================================================================================================
-->
Sub chkLcNoCnt_onpropertychange()
	lgBlnFlgChgValue = true	
End Sub

<!--
'=======================================  3.2.1 btnBLNoOnClick()  ======================================
'=	Event Name : btnBLNoOnClick																		=
'=	Event Desc : ��ȸ ������ȣ Ŭ���� ȣ��															=
'========================================================================================================
-->
Sub btnBLNoOnClick()
	Call OpenBlNoPop()
End Sub

<!--
'====================================  3.2.2 btnPurGrpOnClick()  ===================================
'=	Event Name : btnPurGrpOnClick																	=
'=	Event Desc : ���ű׷�(�˾��� ��������)															=
'========================================================================================================
-->
Sub btnPurGrpOnClick()
	If frm1.txtPurGrp.readOnly <> True Then
		Call OpenPurGrp()
	End If
End Sub

<!--
'====================================  3.2.2 btnPrePayNoOnClick()  ===================================
'=	Event Name : btnPrePayNo																	=
'=	Event Desc : ���ޱݾ�(��������)																			=
'========================================================================================================
-->
Sub btnPrePayNoOnClick()
	If frm1.ChkPrepay.Checked = True Then
		Call OpenPpNo()
	End If
End Sub

<!--
'====================================  3.2.2 btnLoanNoOnClick()  ===================================
'=	Event Name : btnLoanNo																	=
'=	Event Desc :���Աݹ�ȣ ȣ��(tab1)   																=
'========================================================================================================
-->
Sub btnLoanNoOnClick()
	If frm1.txtLoanNo.readOnly <> True Then
		Call OpenLoanNo()
	End If
End Sub

<!--
'=====================================  3.2.3 btnTransportOnClick()  ===================================
-->
Sub btnTransportOnClick()
	If frm1.txtTransport.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "��۹��", "��۹����", gstrTransportMajor)
	End If
End Sub

<!--
'======================================  3.2.4 btnFreightOnClick()  ====================================
-->
Sub btnFreightOnClick()
	If frm1.txtFreight.readOnly <> True Then
		Call OpenMinorCd(frm1.txtFreight.value, frm1.txtFreightNm.value, "�������ҹ��", "�������ҹ����", gstrFreightMajor)
	End If
End Sub

<!--
'====================================  3.2.5 btnPackingTypeOnClick()  ==================================
-->
Sub btnPackingTypeOnClick()
	If frm1.txtPackingType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtPackingType.value, frm1.txtPackingTypeNm.value, "��������", "�������¸�", gstrPackingTypeMajor)
	End If
End Sub

<!--
'======================================  3.2.4 btnIncotermsOnClick()  ====================================
-->
Sub btnIncotermsOnClick()
	If frm1.txtIncoterms.readOnly <> True Then
		Call OpenMinorCd(frm1.txtIncoterms.value, frm1.txtIncotermsNm.value, "��������", "�������Ǹ�", gstrIncotermsMajor)
	End If
End Sub

<!--
'====================================  3.2.5 btnPayTypeOnClick()  ==================================
-->
Sub btnPayTypeOnClick()
	If frm1.txtPayType.readOnly <> True Then
		Call OpenPayType()
	End If
End Sub
	
<!--
'====================================  3.2.5 btnPayMethodOnClick()  ==================================
-->
Sub btnPayMethodOnClick()
	If frm1.txtPayMethod.readOnly <> True Then
		Call OpenMinorCd(frm1.txtPayMethod.value, frm1.txtPayMethodNm.value, "�������", "���������", gstrPayMethodMajor)
	End If
End Sub
	
<!--
'====================================  3.2.5 btnVatTypeOnClick()  ==================================
-->
Sub btnVatTypeOnClick()
	'If frm1.txtVatType.readOnly <> True Then
	'	Call OpenVat()
	'End If
End Sub	
	
<!--
'====================================  3.2.5 btnDeliveryPlceOnClick()  ==================================
-->
Sub btnDeliveryPlceOnClick()
	If frm1.txtDeliveryPlce.readOnly <> True Then
		Call OpenMinorCd(frm1.txtDeliveryPlce.value, frm1.txtDeliveryPlceNm.value, "�ε����", "�ε���Ҹ�", gstrDeliveryPlceMajor)
	End If
End Sub	

<!--
'====================================  3.2.5 btnLoadingPortOnClick()  ==================================
-->
Sub btnLoadingPortOnClick()
	If frm1.txtLoadingPort.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "������", "�����׸�", gstrLoadingPortMajor)
	End If
End Sub		

<!--
'====================================  3.2.5 btnDischgePortOnClick()  ==================================
-->
Sub btnDischgePortOnClick()
	If frm1.txtDischgePort.readOnly <> True Then
		Call OpenDischgePort()
	End If
End Sub		
<!--
'=====================================  3.2.6 btnVesselCntryOnClick()  =================================
-->
Sub btnVesselCntryOnClick()
	If frm1.txtVesselCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtVesselCntry.value, "VESSEL")
	End If
End Sub
<!--
'====================================  3.2.7 btnTranshipCntryOnClick()  ================================
-->
Sub btnTranshipCntryOnClick()
	If frm1.txtTranshipCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtTranshipCntry.value, "TRANSHIP")
	End If
End Sub

<!--
'====================================  3.2.8 btnOriginCntryOnClick()  ==================================
-->
Sub btnOriginCntryOnClick()
	If frm1.txtOriginCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtOriginCntry.value, "ORIGIN")
	End If
End Sub

<!--
'====================================  3.2.8 btnOriginOnClick()  ==================================
-->
Sub btnOriginOnClick()
	If frm1.txtOrigin.readOnly <> True Then
		Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "������", "��������", gstrOriginMajor)
	End If
End Sub
	
	
<!--
'======================================  3.2.9 btnAgentOnClick()  ======================================
-->
Sub btnAgentOnClick()
	If frm1.txtAgent.readOnly <> True Then
		Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "������", "�����ڸ�")
	End If
End Sub

<!--
'=================================  3.2.10 btnManufacturerOnClick()  ===================================
-->
Sub btnManufacturerOnClick()
	If frm1.txtManufacturer.readOnly <> True Then
		Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "������", "�����ڸ�")
	End If
End Sub

<!--
'=================================  3.2.10 btnTaxBizAreaOnClick()  ===================================
-->
Sub btnTaxBizAreaOnClick()
	If frm1.txtTaxBizArea.readOnly <> True Then
		Call OpenBizArea()
	End If
End Sub

<!--
'=================================  3.2.11 btnForwarderOnClick()  ======================================
-->
Sub btnForwarderOnClick()
	If frm1.txtForwarder.readOnly <> True Then
		Call OpenBizPartner(frm1.txtForwarder.value, frm1.txtForwarderNm.value, "����ȸ��", "����ȸ���")
	End If
End Sub

<!--
'===================================  3.2.12 btnWeightUnitOnClick()  ===================================
-->
Sub btnWeightUnitOnClick()
	If frm1.txtWeightUnit.readOnly <> True Then
		Call OpenUnit(frm1.txtWeightUnit.value, "WT")
	End If
End Sub

<!--
'===================================  3.2.13 btnVolumnUnitOnClick()  ===================================
-->
Sub btnVolumnUnitOnClick()
	If frm1.txtVolumnUnit.readOnly <> True Then
		Call OpenUnit(frm1.txtVolumnUnit.value, "WD")
	End If
End Sub

<!--
'==================================  3.2.14 btnTaxOfficeCdOnClick()  ===================================
-->
Sub btnTaxOfficeCdOnClick()
	If frm1.txtTaxOfficeCd.readOnly <> True Then
		Call OpenTaxOffice()
	End If
End Sub
<!--
'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
-->
Sub btnPosting_OnClick()
	Dim Answer
	Dim strVal
	Dim strBLNo
	
    Err.Clear               

	strBLNo = frm1.txtBLNo.value
		
	If strBLNo = "" Then	
		Call DisplayMsgBox("900002","X","X","X")   '��ȸ�� ���� �Ͻʽÿ�.
		Exit Sub
	Else
		
		if frm1.rdoPostingFlg2.Checked = True then 
			Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")   '�۾��� ���� �Ͻðڽ��ϱ�?
			If Answer = vbNo Then
				frm1.btnPosting.disabled = False	'20040315          
				Exit Sub
			Else 
				frm1.btnPosting.disabled = True		'20040315   
			End If	
		else
			Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
			If Answer = vbNo Then
				frm1.btnPosting.disabled = False	'200308    
				Exit Sub
			Else 		
				frm1.btnPosting.disabled = True		'200308 	
			End If
		End if
			
		If Answer = VBYES Then
			
    	    If  LayerShowHide(1) = False Then
            	Exit Sub
            End If
				
			frm1.txtUpdtUserId.value = parent.gUsrID
			frm1.txtInsrtUserId.value = parent.gUsrID

			strVal = BIZ_PGM_POSTQRY_ID & "?txtMode=" & CID_POST		
			strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)          'B/L ������ȣ 
			strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo.value)          '���Թ�ȣ 
			strVal = strVal & "&txtBlIssueDt=" & Trim(frm1.txtBlIssueDt.Text) 'B/L ������ 
			if frm1.rdoPostingFlg2.Checked = True then                        'Ȯ������ 
				strVal = strVal & "&txtPost=" & "C"
			else
				strVal = strVal & "&txtPost=" & "D"
			End if
			strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)

			Call RunMyBizASP(MyBizASP, strVal)							

		End IF
	End IF
		
End Sub
<!--
'=========================================  5.1.1 FncQuery()  ===========================================
-->
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													

	Err.Clear															

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")							
	Call InitVariables												

	If Not chkField(Document, "1") Then			
		Exit Function
	End If

	If DbQuery = False Then Exit Function

	FncQuery = True			
End Function
	
<!--
'===========================================  5.1.2 FncNew()  ===========================================
-->
Function FncNew()
		
	Dim IntRetCD 

	FncNew = False  

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    frm1.hdnchangeflg.value = "Y"
    frm1.hdnDefaultFlg.value = ""
		
	Call ClickTab1()
	Call ggoOper.ClearField(Document, "A")								
	Call ggoOper.LockField(Document, "N")								
	Call SetDefaultVal
	Call InitVariables													
		
	Call ChangeTag(False)
		
	frm1.txtBLNo.focus
	Set gActiveElement = document.activeElement
		
	FncNew = True														
		
End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
Function FncDelete()
		
	Dim IntRetCD
	
	FncDelete = False				
		
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then							
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	If DbDelete = False Then Exit Function

	FncDelete = True										
End Function

<!--
'===========================================  5.1.4 FncSave()  ==========================================
-->
Function FncSave()
	Dim IntRetCD
		
	FncSave = False											

	Err.Clear												
		
	If CheckRunningBizProcess = True Then
    	Exit Function
	End If	

	If lgBlnFlgChgValue = False Then					
	    IntRetCD = DisplayMsgBox("900001","X","X","X")	
	    Exit Function
	End If
		
    If Not chkField(Document, "2") Then                      
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
	        
        Exit Function
    End If
	    
    If Trim(frm1.txtTaxBizArea.value) = "" then 
		Call GetTaxBizArea("*")
	end if	    
    
    if Trim(UNICDbl(frm1.txtXchRate.text)) = "" Or Trim(UNICDbl(frm1.txtXchRate.text)) = "0" then
		Call DisplayMsgBox("970021", "X","ȯ��", "X")
		Call ClickTab1()
		frm1.txtXchRate.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	    
    'if Trim(frm1.txtPrePayNo.value) <> "" And (Trim(UNICDbl(frm1.txtPrePayDocAmt.text)) = "" Or Trim(UNICDbl(frm1.txtPrePayDocAmt.text)) = "0") then
	'	Call DisplayMsgBox("970021", "X","���ޱ�", "X")
	'	Call ClickTab1()
	'	frm1.txtPrePayDocAmt.focus
	'	Set gActiveElement = document.activeElement
	'	Exit Function
	'End if
				
	If ValidDateCheckLocal(frm1.txtLoadingDt, frm1.txtBlIssueDt) = False Then Exit Function
	'���ҿ����� üũ���� ����(2003.09.22)
	'If ValidDateCheckLocal(frm1.txtBlIssueDt, frm1.txtSetlmnt) = False Then Exit Function			
	If ValidDateCheckLocal(frm1.txtLoadingDt, frm1.txtDischgeDt) = False Then Exit Function	
		
	'if UCase(frm1.hdnLoanflg.value) = "Y" then
	'	if Trim(frm1.txtLoanNo.value) <> "" And (Trim(UNICDbl(frm1.txtLoanAmt.text)) = "" Or Trim(UNICDbl(frm1.txtLoanAmt.text)) = "0") then
	'		Call DisplayMsgBox("970021", "X","���Ա�", "X")
	'		Call ClickTab1()
	'		frm1.txtLoanAmt.focus
	'		Set gActiveElement = document.activeElement
	'		Exit Function
	'	End if
	'	if UNICDbl(frm1.txtDocAmt.Text) > 0 then
	'		if UNICDbl(frm1.txtLoanAmt.text) > UNICDbl(frm1.txtDocAmt.text) then 
	'			Call DisplayMsgBox("970023", "X","B/L�ݾ�","���Ա�")
	'			Exit Function
	'		end if
	'	end if
	'end if
				
	If DbSave = False Then Exit Function
		
	frm1.txtBLNo.focus
    Set gActiveElement = document.ActiveElement   
		
		
		
	FncSave = True	
End Function

<!--
'===========================================  5.1.5 FncCopy()  ==========================================
-->
Function FncCopy()
		
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = parent.OPMD_CMODE												

		
	Call ggoOper.ClearField(Document, "1")									
	Call ggoOper.LockField(Document, "N")									
		
	Call SetToolbar("11111000001111")
	Call ChangeTag(False)
			
End Function

<!--
'===========================================  5.1.6 FncCancel()  ========================================
-->
Function FncCancel() 
	On Error Resume Next
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow()
	On Error Resume Next
End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
Function FncDeleteRow()
	On Error Resume Next
End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
Function FncPrint()
   Call parent.FncPrint()
End Function

<!--
'============================================  5.1.10 FncPrev()  ========================================
-->
Function FncPrev() 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then	
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgPrevNo = "" Then			
		Call DisplayMsgBox("900011","X","X","X")
	End If
End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
Function FncNext()
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then						
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgNextNo = "" Then								
		Call DisplayMsgBox("900012","X","X","X")
	End If
End Function

<!--
'===========================================  5.1.12 FncExcel()  ========================================
-->
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, True)
End Function
	
<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function
<!--
'=============================================  5.2.1 DbQuery()  ========================================
-->
Function DbQuery()
	Dim strVal

	Err.Clear													

	DbQuery = False												

    If  LayerShowHide(1) = False Then
       	Exit Function
    End If

	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001			
	strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)	

    frm1.hdnchangeflg.value = "Y"
	Call RunMyBizASP(MyBizASP, strVal)							
	
	DbQuery = True												
End Function

<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
Function DbSave()
	Dim strVal

	Err.Clear													

	DbSave = False												

    If  LayerShowHide(1) = False Then
       	Exit Function
    End If
		
	With frm1
		
		if .rdoPostingFlg1.Checked = True then
			.txtPost.Value = "Y"
		Else
			.txtPost.Value = "N"
		End if
			
		if .chkPoNoCnt.checked = true then
			.hdnChkPoNo.Value = "1"
		else
			.hdnChkPoNo.Value = "0"
		End if
		if .chkLcNoCnt.checked = true then
			.hdnChkLcDocNo.Value = "1"
		else
			.hdnChkLcDocNo.Value = "0"
		End if
		
		.txtMode.value = parent.UID_M0002								
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.hdninterface_Account.value = interface_Account

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)

	End With

	DbSave = True												
End Function
	
<!--
'=============================================  5.2.3 DbDelete()  =======================================
-->
Function DbDelete()
	Dim strVal

	Err.Clear													

	DbDelete = False											
		
    If  LayerShowHide(1) = False Then
       	Exit Function
    End If
		
	strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003			
	strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)	

	Call RunMyBizASP(MyBizASP, strVal)							

	DbDelete = True												
End Function

<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function DbQueryOk()											
	lgIntFlgMode = parent.OPMD_UMODE									

	lgBlnFlgChgValue = False

	Call ggoOper.LockField(Document, "Q")				'������ �����Ұ�		
		
	'**����(2003.03.26)-ȸ������ ��� Ȯ��,Ȯ����� �����ϵ��� ������.
	if frm1.rdoPostingflg(0).Checked = True then        'Ȯ���� �Ȱ�� ����,���� �Ұ� 
		Call SetToolbar("11100000000111")			
		Call ChangeTag(True)
		frm1.btnPosting.value = "Ȯ�����"
		frm1.btnPosting.disabled = false
		if interface_Account <> "N" then
			frm1.btnGlSel.disabled = false
		Else
			frm1.btnGlSel.disabled = True
		End If
	else
		Call SetToolbar("11111000000111")
		Call ChangeTag(False)
		frm1.btnPosting.value = "Ȯ��"
		frm1.btnPosting.disabled = false
		frm1.btnGlSel.disabled = true
		'Call setLoan()
	end if
		
    if frm1.hdnGlType.Value = "A" Then
       frm1.btnGlSel.value = "ȸ����ǥ��ȸ"
    elseif frm1.hdnGlType.Value = "T" Then
       frm1.btnGlSel.value = "������ǥ��ȸ"
    elseif frm1.hdnGlType.Value = "B" Then
       frm1.btnGlSel.value = "��ǥ��ȸ"
    end if	
    '�߰�(Detail�� �������� ������ Ȯ����ư Disable��Ŵ)
    if UNICDbl(Trim(frm1.txtDocAmt.Text)) <> 0 then	
		frm1.btnPosting.Disabled = False
	else
		frm1.btnPosting.Disabled = True
	End if
		
		
	if frm1.ChkPrepay.checked = true then
	    Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
	end if
		
	lgBlnFlgChgValue = False
		
	Call ClickTab1()
	frm1.txtBLNo.focus 
	Set gActiveElement = document.activeElement
		
	Call CheckPrePayedAmtYN()
		
End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()												
	Call InitVariables
	frm1.hdnchangeflg.value = "Y"
	Call MainQuery()
End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()											
	Call MainNew()
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����B/L����</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����B/L��Ÿ</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenPoRef">��������</A> | <A href="vbscript:OpenLCRef">L/C����</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
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
												<TD CLASS=TD5 NOWRAP>B/L ������ȣ</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=32 MAXLENGTH=18 TAG="12XXXU" ALT="B/L ������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnBLNoOnClick()"></TD>
												<TD CLASS=TD6>&nbsp;</TD>
												<TD CLASS=TD6>&nbsp;</TD>
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
									<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L ������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo1" SIZE=34 MAXLENGTH=18 TAG="25XXXU" ALT="B/L ������ȣ"></TD>
											<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX CHECKED ID="ChkPrepay" tag="24" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">&nbsp;���ޱݿ���
											                     <INPUT TYPE=TEXT NAME="ChkPrepay1"  style="HEIGHT: 19px; WIDTH: 1px" MAXLENGTH=0 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrepayNo" align=top TYPE="BUTTON" onclick="vbscript:OpenPpNo()"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT ALT=B/L��ȣ NAME="txtBLDocNo" MAXLENGTH=18 TYPE=TEXT SIZE=34  TAG="22XXXU">
											<TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
 											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPONo" TYPE=TEXT SIZE=20  TAG="24XXXU">
 														         <INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="21" CHECKED ID="chkPoNoCnt">&nbsp;���ֹ�ȣ����</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
  			        						<TD CLASS=TD6 NOWRAP>
  			        							<Table Cellspacing=0 Cellpadding=0>	
  			        								<TR>
  			        									<TD NOWRAP>
  			        										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=������ NAME="txtLoadingDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;
  			        									</TD>
  			        									<TD NOWRAP>
														 	 &nbsp;&nbsp;B/L������
														</TD>
														<TD NOWRAP>
														 	<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L������ NAME="txtBlIssueDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=20  TAG="24XXXU">
																 <INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="21" CHECKED ID="chkLcNoCnt">&nbsp;L/C��ȣ����</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
  			        						<TD CLASS=TD6 NOWRAP>
  			        							<Table Cellpadding=0 Cellspacing=0>
  			        								<TR>
  			        									<TD NOWRAP>
  			        										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=������ NAME="txtDischgeDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;
  			        									</TD>
  			        									<TD NOWRAP>
  			        										 &nbsp;���ҿ�����
  			        									</TD>
  			        									<TD NOWRAP>
  			        										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���ҿ����� NAME="txtSetlmnt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
  			        									</TD>
  			        								</TR>
  			        							</Table>
  			        						</TD>
											<TD CLASS=TD5 NOWRAP>��۹��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="��۹��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" onclick="vbscript:btnTransportOnClick()">
																 <INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>����ȸ��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtForwarder" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="����ȸ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnForwarder" align=top TYPE="BUTTON" onclick="vbscript:btnForwarderOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtForwarderNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>VESSEL��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL��" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X">
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVoyageNo" TYPE=TEXT SIZE=34 MAXLENGTH=20  TAG="21XXXU">											
											<TD CLASS=TD5 NOWRAP>���ڱ���</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVesselCntry" SIZE=10  MAXLENGTH=3 TAG="21XXXU" ALT="���ڱ���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVesselCntry" align=top TYPE="BUTTON" onclick="vbscript:btnVesselCntryOnClick()"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����尹��</TD>
 	        								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�����尹�� NAME="txtTotPackingCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDischgePort" SIZE=10 MAXLENGTH=5  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" onclick="vbscript:btnDischgePortOnClick()">
																 <INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȭ��</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table CellPadding=0 cellspacing=0>
													<TR>
														<TD>
															<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��" OnChange="VBScript:ChangeCurrency()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
														</TD>
													</TR>
												</Table>
											</TD>
	        								<TD CLASS=TD5 NOWRAP>ȯ��</TD>
			        						<TD CLASS=TD6 NOWRAP>
				        						<Table Cellpadding=0 Cellspacing=0>
				        							<TR>
				        								<TD NOWRAP>
				        									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=ȯ�� NAME="txtXchRate" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="22X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
				        								</TD>
													</TR>
												</Table>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L�ݾ�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L�ݾ� NAME="txtDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>B/L�ڱ��ݾ�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L�ڱ��ݾ� NAME="txtLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayType" align=top TYPE="BUTTON" onclick="vbscript:btnPayTypeOnClick()">
																 <INPUT TYPE=TEXT NAME="txtPayTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>��ݰ�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsTxt" ALT="��ݰ�������" TYPE=TEXT MAXLENGTH=120 SIZE=34 TAG="21X"></TD>
										</TR>																													
										<TR>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayMethod" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="�������">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtPayMethodNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellspacing=0 Cellpadding=0>
													<TR>
														<TD NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�����Ⱓ NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="24X7" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
														<TD NOWRAP>
											                &nbsp;DAYS&nbsp;��������&nbsp;
											            </TD>
											            <TD NOWRAP>
											                <INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10  MAXLENGTH=5 TAG="24XXXU" ALT="��������"></TD>
											            </TD>
											        </TR>
											    </Table>
											</TD>
										</TR>
										<TR><!--
											<TD CLASS=TD5 NOWRAP>���ޱݹ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
															<INPUT TYPE=TEXT NAME="txtPrePayNo" SIZE=32 STYLE="Text-Transform: uppercase" MAXLENGTH=18 TAG="21X" ALT="���ޱݹ�ȣ" OnChange="changePpNo()"><IMG SRC="../../image/btnPopup.gif" NAME="btnPrePayNo" align=top TYPE="BUTTON">
														</TD>
													</TR>
												</Table>
											</TD>-->
													<TD CLASS=TD5 NOWRAP>����ó</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayeeCd" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="����ó" OnChange="VBScript:ChangePayeeCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" OnClick="OpenBpMpa()">
																 <INPUT TYPE=TEXT NAME="txtPayeeNm" SIZE=20 TAG="24"></TD>
											<!--TD CLASS=TD5 NOWRAP>���Աݹ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNo" SIZE=32  MAXLENGTH=18 TAG="21XXXU" ALT="���Աݹ�ȣ" OnChange="VBScript:ChangeLoanNo()"><IMG SRC="../../image/btnPopup.gif" NAME="btnLoanNo" align=top TYPE="BUTTON" onclick="vbscript:btnLoanNoOnClick()"></TD-->
											<TD CLASS=TD5 NOWRAP>���Թ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=34 TAG="24XXXU" MAXLENGTH=18></TD> <!-- TAG="25XNXU" -->
										</TR>
										<TR>
											<!--<TD CLASS=TD5 NOWRAP>���ޱ�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���ޱ� NAME="txtPrePayDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											-->
											<TD CLASS=TD5 NOWRAP>���ݰ�꼭����ó</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBuildCd" SIZE=10  MAXLENGTH=10 TAG="22XXXU" ALT="���ݰ�꼭����ó" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpMgs" align=top TYPE="BUTTON" OnClick="OpenBpMgs()">
																 <INPUT TYPE=TEXT NAME="txtBuildNm" SIZE=20 TAG="24"></TD>
											<!--TD CLASS=TD5 NOWRAP>���Ա�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���Ա� NAME="txtLoanAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" Title="FPDOUBLESINGLE" ></OBJECT>');</SCRIPT></TD-->
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR><!--
											<TD CLASS=TD5 NOWRAP>������ݾ�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=������ݾ� NAME="txtCashAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											-->
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvType" SIZE=10 MAXLENGTH=5 TAG="22XNXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON" OnClick="OpenIvType()">
																 <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
     	         							<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingFlg" TAG="24X" VALUE="Y" ID="rdoPostingFlg1"><LABEL FOR="rdoPostingFlg1">&nbsp;Y&nbsp;</LABEL> 
     	         							                     <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingFlg" TAG="24X" VALUE="N" CHECKED ID="rdoPostingFlg2"><LABEL FOR="rdoPostingFlg2">&nbsp;N&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</LABEL>
     	         							                     <INPUT TYPE=TEXT NAME="txtGlNo" ALT="��ǥ��ȣ" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD></TD>
										</TR>	
										<%Call SubFillRemBodyTD5656(2)%>
									</TABLE>
									</DIV>
									<!-- �ι�° �� ���� 
									<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>-->
									<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPackingType" align=top TYPE="BUTTON" onclick="vbscript:btnPackingTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�����������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackingTxt" ALT="�����������" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���߷�</TD>
 	        								<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���߷� NAME="txtGrossWeight" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP>�����̳ʼ�</TD>
 	        								<TD CLASS=TD6 NOWRAP> 	        								
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�����̳ʼ� NAME="txtContainerCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="21X31Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</Table>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���߷�</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
														    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���߷� NAME="txtNetWeight" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="21X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;
														</TD>
														<TD>
														    <INPUT TYPE=TEXT NAME="txtWeightUnit" SIZE=10 MAXLENGTH=3 SIZE=20 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON" onclick="vbscript:btnWeightUnitOnClick()">
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP>�ѿ���</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
														    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�ѿ��� NAME="txtGrossVolumn" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;
														</TD>
														<TD NOWRAP>
														    <INPUT TYPE=TEXT NAME="txtVolumnUnit" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVolumnUnit" align=top TYPE="BUTTON" onclick="vbscript:btnVolumnUnitOnClick()"></TD>
														</TD>
													</TR>
												</Table>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������ҹ��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="�������ҹ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFreight" align=top TYPE="BUTTON" onclick="vbscript:btnFreightOnClick()">
																 <INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�����������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFreightPlce" ALT="�����������" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>����������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="����������" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>�ε����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeliveryPlce" SIZE=10 MAXLENGTH=5  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeliveryPlce" align=top TYPE="BUTTON" onclick="vbscript:btnDeliveryPlceOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeliveryPlceNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiptPlce" ALT="�������" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoadingPort" SIZE=10 MAXLENGTH=5  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" onclick="vbscript:btnLoadingPortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȯ������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTranshipCntry" SIZE=10 MAXLENGTH=3 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTranshipCntry" align=top TYPE="BUTTON" onclick="vbscript:btnTranshipCntryOnClick()"></TD>
											<TD CLASS=TD5 NOWRAP>ȯ����</TD>
  			        						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=ȯ���� NAME="txtTranshipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
    										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrigin" SIZE=10 MAXLENGTH=5 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" onclick="vbscript:btnOriginOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>����������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="����������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" onclick="vbscript:btnOriginCntryOnClick()"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLIssuePlce" ALT="B/L�������" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>B/L����μ�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L����μ� NAME="txtBLIssueCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X72" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											
										</TR>
    									<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" onclick="vbscript:btnAgentOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" onclick="vbscript:btnManufacturerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
										</TR>
    									<TR>
											<TD CLASS=TD5 NOWRAP>���ű׷�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="24XXXU" ALT="���ű׷�">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTaxBizArea" SIZE=10 MAXLENGTH=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" onclick="vbscript:btnTaxBizAreaOnClick()">
																 <INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="24"></TD>
										</TR>										
    									<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="��������">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5>���</TD>
											<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT TYPE=TEXT  NAME="txtRemark" ALT="���" tag = "21" SIZE=90 MAXLENGTH=70></TD>
										</TR>
										<%Call SubFillRemBodyTD5656(6)%>
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
							<TD><BUTTON NAME="btnPosting" CLASS="CLSMBTN">Ȯ��</BUTTON>&nbsp;
							<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">��ǥ��ȸ</BUTTON>&nbsp;
							</TD>
							<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadBLDtl()">B/L �������</A> | <A href="vbscript:LoadChargeHdr()">�����</A> | <A href="vbscript:LoadIvPayment()">���޳������</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtLCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHBLNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtPost" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnChkPoNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnChkLcDocNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnLoanflg" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdninterface_Account" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnEnddate" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnchangeFlg" tag="24">
		<INPUT TYPE=HIDDEN NAME="hdnDefaultFlg" tag="24">
		
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>	
</BODY>
</HTML>
