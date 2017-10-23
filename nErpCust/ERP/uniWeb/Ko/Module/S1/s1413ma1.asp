<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1413MA1
'*  4. Program Name         : �㺸��� 
'*  5. Program Desc         : �㺸��� 
'*  6. Comproxy List        : PS1G114.dll, PS1G115.dll, PS1G116.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2002/12/09 : INCLUDE �ٽ� ���� ����, Kang Jun Gu
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<Script Language="VBS">
Option Explicit					<% '��: indicates that All variables must be declared in advance %>
	
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)


	Const BIZ_PGM_ID = "s1413mb1.asp"						<% '��: �����Ͻ� ���� ASP�� %>
	Const BIZ_PGM_JUMP_ID = "s1413qa1"

	<% '------ Minor Code PopUp�� ���� Major Code���� ------ %>
	Const gstrWarrantTypeMajor = "S0002"
	Const gstrDelTypeMajor = "S0003"


'============================================  1.2.2 Global ���� ����  ==================================
	Dim lgBlnFlgChgValue					<% '��: Variable is for Dirty flag %>
	Dim lgIntGrpCount					<% '��: Group View Size�� ������ ���� %>
	Dim lgIntFlgMode						<% '��: Variable is for Operation Status %>
	
	
	Dim gSelframeFlg					<% '���� TAB�� ��ġ�� ��Ÿ���� Flag %>
	Dim gblnWinEvent					<% '~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
												'	PopUp Window�� ��������� ���θ� ��Ÿ���� variable %>
	Dim lgBlnFlawChgFlg	
	Dim gtxtChargeType

'========================================================================================================
	Function InitVariables()
		lgIntFlgMode = Parent.OPMD_CMODE						<%'��: Indicates that current mode is Create mode%>
		lgBlnFlgChgValue = False								<%'��: Indicates that no value changed%>
		lgBlnFlawChgFlg = False
		lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
		
		<% '------ Coding part ------ %>
		gblnWinEvent = False
	End Function

'========================================================================================================
	Sub SetDefaultVal()
		frm1.txtWarrentNo.focus
		frm1.txtAsignDt.text = EndDate
		frm1.txtEstimateAmt.text = UNIFormatNumber(0,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		frm1.txtWarrentAbleAmt.text = UNIFormatNumber(0,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		frm1.txtWarrentAsignAmt.text = UNIFormatNumber(0,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		lgBlnFlgChgValue = False
	End Sub

'========================================================================================================
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
		<% Call LoadBNumericFormatA( "I", "*", "NOCOOKIE", "MA") %>
	End Sub

'========================================================================================================
	Function OpenCollateralNoPop()
		Dim iCalledAspName
		Dim strRet
		If gblnWinEvent = True Then Exit Function
		gblnWinEvent = True

		iCalledAspName = AskPRAspName("S1413PA1")
		
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S1413PA1", "x")
			gblnWinEvent = False
			exit Function
		end if
		
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False
		
		If strRet = "" Then
			Exit Function
		Else
			Call SetCollateralNo(strRet)
		End If	
	End Function

'========================================================================================================
	Function OpenBizPartner()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "��"							<%' �˾� ��Ī %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtCustomer.value)		    <%' Code Condition%>
		arrParam(3) = ""				                    <%' Name Cindition%>
		arrParam(4) = "CREDIT_MGMT_FLAG = " & FilterVar("Y", "''", "S") & " "				<%' Where Condition%>
		arrParam(5) = "��"							<%' TextBox ��Ī %>

		arrField(0) = "BP_CD"								<%' Field��(0)%>
		arrField(1) = "BP_NM"								<%' Field��(1)%>

		arrHeader(0) = "��"							<%' Header��(0)%>
		arrHeader(1) = "����"							<%' Header��(1)%>

		frm1.txtCustomer.focus
		 
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBizPartner(arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = strPopPos								<%' �˾� ��Ī %>
		arrParam(1) = "B_Minor"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(strMinorCD)						<%' Code Condition%>
		arrParam(3) = ""						            <%' Name Cindition%>
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = strPopPos								<%' TextBox ��Ī %>

		arrField(0) = "Minor_CD"							<%' Field��(0)%>
		arrField(1) = "Minor_NM"							<%' Field��(1)%>

		arrHeader(0) = strPopPos							<%' Header��(0)%>
		arrHeader(1) = strPopPos & "��"					<%' Header��(1)%>


		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetMinorCd(strMajorCd, arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenSalesGroup()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "�����׷�"						<%' �˾� ��Ī %>
		arrParam(1) = "B_SALES_GRP"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtSalesGrp.value)			<%' Code Condition%>
		arrParam(3) = ""                        			<%' Name Cindition%>
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		arrParam(5) = "�����׷�"						<%' TextBox ��Ī %>

		arrField(0) = "SALES_GRP"							<%' Field��(0)%>
		arrField(1) = "SALES_GRP_NM"						<%' Field��(1)%>

		arrHeader(0) = "�����׷�"						<%' Header��(0)%>
		arrHeader(1) = "�����׷��"						<%' Header��(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetSalesGroup(arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenCurrency()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		If lgIntFlgMode = parent.OPMD_UMODE Then
			gblnWinEvent = False
			Exit Function
		End If
		
		arrParam(0) = "ȭ��"						<%' �˾� ��Ī %>
		arrParam(1) = "B_CURRENCY"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtCurrency.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = ""								<%' Where Condition%>
		arrParam(5) = "ȭ��"						<%' TextBox ��Ī %>
		
	    arrField(0) = "CURRENCY"						<%' Field��(0)%>
	    arrField(1) = "CURRENCY_DESC"					<%' Field��(1)%>
	    
	    arrHeader(0) = "ȭ��"						<%' Header��(0)%>
	    arrHeader(1) = "ȭ���"						<%' Header��(1)%>
	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetCurrency(arrRet)
		End If
	End Function

'========================================================================================================
	Function SetCollateralNo(arrRet)
		frm1.txtWarrentNo.Value = arrRet
	End Function

'========================================================================================================
	Function SetBizPartner(arrRet)
		frm1.txtCustomer.Value = arrRet(0)
		frm1.txtCustomerNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetMinorCd(strMajorCd, arrRet)
		Select Case strMajorCd
			Case gstrWarrantTypeMajor
				frm1.txtWarrentType.Value = arrRet(0)
				frm1.txtWarrentTypeNm.Value = arrRet(1)
			Case gstrDelTypeMajor
				frm1.txtDelType.Value = arrRet(0)
				frm1.txtDelTypeNm.Value = arrRet(1)
			Case Else
		End Select

		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetSalesGroup(arrRet)
		frm1.txtSalesGrp.value = arrRet(0)
		frm1.txtSalesGrpNm.value = arrRet(1)

		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetCurrency(arrRet)
		frm1.txtCurrency.Value = arrRet(0)

		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function CookiePage(ByVal Kubun)

		On Error Resume Next

		Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
		Dim strTemp, arrVal

		Select Case Kubun
		
		Case 1
			WriteCookie CookieSplit , frm1.txtSalesGrp.value & Parent.gRowSep & frm1.txtSalesGrpNm.value & Parent.gRowSep & _
			frm1.txtWarrentType.value & Parent.gRowSep & frm1.txtWarrentTypeNm.value & Parent.gRowSep & _
			frm1.txtAsignDt.text & Parent.gRowSep & frm1.txtDelType.value & Parent.gRowSep & _
			frm1.txtDelTypeNm.value & Parent.gRowSep & frm1.txtCustomer.value & Parent.gRowSep & frm1.txtCustomerNm.value & Parent.gRowSep

		Case 0
			strTemp = ReadCookie(CookieSplit)

			If strTemp = "" then Exit Function
				
			frm1.txtWarrentNo.value =  strTemp
			
			If Err.number <> 0 Then
				Err.Clear
				WriteCookie CookieSplit , ""
				Exit Function 
			End If
			
			Call MainQuery()
						
			WriteCookie CookieSplit , ""
		Case Else
			Exit Function
		End Select 		
	End Function

'========================================================================================================
Function JumpChgCheck()

	Dim IntRetCD

	'************ �̱��� ��� **************
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function
	
'========================================================================================================
	Sub Form_Load()
		Call LoadInfTB19029																<% '��: Load table , B_numeric_format %>
		Call AppendNumberPlace("6", "4", "0")
		Call AppendNumberPlace("7", "3", "0")
		Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
 		Call ggoOper.LockField(Document, "N")											<% '��: Lock  Suitable  Field %>
		Call SetDefaultVal()
		<% '----------  Coding part  ------------------------------------------------------------- %>
		Call chkDeleteFlg_OnPropertyChange()
		Call SetToolBar("1110100000001111")												<% '��: ��ư ���� ���� %>
		Call InitVariables
		Call CookiePage (0)
			
	End Sub
	
'========================================================================================================
	Sub btnWarrentNoOnClick()
		frm1.txtWarrentNo.focus 
		Call OpenCollateralNoPop()
	End Sub

'========================================================================================================
	Sub btnCustomerOnClick()
		If Not frm1.txtCustomer.readOnly Then
			Call OpenBizPartner()
		End If
	End Sub

'========================================================================================================
	Sub btnWarrentTypeOnClick()
		frm1.txtWarrentType.focus 
		Call OpenMinorCd(frm1.txtWarrentType.value, frm1.txtWarrentTypeNm.value, "�㺸����", gstrWarrantTypeMajor)
	End Sub

'========================================================================================================
	Sub btnCurrencyOnClick()
		frm1.txtCurrency.focus 
		Call OpenCurrency()
	End Sub

'========================================================================================================
	Sub btnDelTypeOnClick()
		If frm1.txtDelType.readOnly = False Then
			frm1.txtDelType.focus 
			Call OpenMinorCd(frm1.txtDelType.value, frm1.txtDelTypeNm.value, "��������", gstrDelTypeMajor)
		End If
	End Sub

'========================================================================================================
	Sub btnSalesGrpOnClick()
		frm1.txtSalesGrp.focus 
		Call OpenSalesGroup()
	End Sub

'========================================================================================================
	Sub txtEstimateDt_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtAsignDt_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtExpiryDt_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtDelDt_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtEstimateAmt_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtWarrentAbleAmt_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtWarrentAsignAmt_Change()
		lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtAsgnSeq_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtFloorSpace_Change()
		lgBlnFlgChgValue = True
	End Sub

	Sub txtGroundSpace_Change()
		lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtCreditCheckDt_Change()
		lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtEstimateDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtEstimateDt.Action = 7 
			Call SetFocusToDocument("M")
			Frm1.txtEstimateDt.Focus
	    End If
	    lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtAsignDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtAsignDt.Action = 7
			Call SetFocusToDocument("M")
			Frm1.txtAsignDt.Focus
	    End If
	    lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtExpiryDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtExpiryDt.Action = 7
			Call SetFocusToDocument("M")
			Frm1.txtExpiryDt.Focus
	    End If
	    lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtDelDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtDelDt.Action = 7
			Call SetFocusToDocument("M")
			Frm1.txtDelDt.Focus
	    End If
	    lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub chkDeleteFlg_OnPropertyChange()
	    With frm1
			If .chkDeleteFlg.checked Then
				Call ggoOper.SetReqAttr(.txtDelType, "N")
				Call ggoOper.SetReqAttr(.txtDelDt, "N")
			Else
				Call ggoOper.SetReqAttr(.txtDelType, "Q")
				Call ggoOper.SetReqAttr(.txtDelDt, "Q")
				frm1.txtDelType.value = ""
				frm1.txtDelTypeNm.value = "" 
				frm1.txtDelDt.text = "" 
			End If
		End With						
	End Sub

	Sub chkDeleteFlg_OnClick()
	    lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Function FncQuery()
		Dim IntRetCD
		
		FncQuery = False													<% '��: Processing is NG %>

		Err.Clear															<% '��: Protect system from crashing %>
	
		<% '------ Check previous data area ------ %>
		
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			<% '��: "Will you destory previous data" %>
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '��: Clear Contents  Field %>
		Call InitVariables
		<% '------ Check condition area ------ %>

		If Not chkField(Document, "1") Then							<% '��: This function check indispensable field %>
			Exit Function
		End If

		<% '------ Query function call area ------ %>
		Call ggoOper.LockField(Document, "N")								<% '��: This function lock the suitable field %>

		Call DbQuery()														<% '��: Query db data %>

		FncQuery = True														<% '��: Processing is OK %>
	End Function
	
'========================================================================================================
	Function FncNew()
		Dim IntRetCD 

		FncNew = False                                                          <%'��: Processing is NG%>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase condition area ------ %>
		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "A")									<%'��: Clear Condition Field%>
		Call ggoOper.LockField(Document, "N")									<%'��: Lock  Suitable  Field%>
		Call SetDefaultVal
		Call InitVariables														<%'��: Initializes local global variables%>
		Call SetToolBar("1110100000001111")												<% '��: ��ư ���� ���� %>
		Call chkDeleteFlg_OnPropertyChange()
		lgBlnFlgChgValue = False
		FncNew = True															<%'��: Processing is OK%>
	End Function
	
'========================================================================================================
	Function FncDelete()
		Dim IntRetCD

		FncDelete = False												<% '��: Processing is NG %>
		
		<% '------ Precheck area ------ %>
		If lgIntFlgMode <> Parent.OPMD_UMODE Then								<% 'Check if there is retrived data %>
			Call DisplayMsgBox("900002","x","x","x")
			Exit Function
		End If

		IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")

		If IntRetCD = vbNo Then
			Exit Function
		End If

		<% '------ Delete function call area ------ %>
		Call CreditCheck("DELETE")
'		Call DbDelete													<% '��: Delete db data %>

		FncDelete = True												<% '��: Processing is OK %>
	End Function

'========================================================================================================
	Function FncSave()
		Dim IntRetCD
		
		FncSave = False													<% '��: Processing is NG %>
		
		Err.Clear														<% '��: Protect system from crashing %>
		
		<% '------ Precheck area ------ %>
		If lgBlnFlgChgValue = False Then								<% 'Check if there is retrived data %>
		    IntRetCD = DisplayMsgBox("900001",Parent.VB_YES_NO,"x","x")					<% '��: No data changed!! %>
		    Exit Function
		End If
		
		<% '------ Check contents area ------ %>
		If Not chkField(Document, "2") Then						<% '��: Check contents area %>
			Exit Function
		End If
	
		<% '------ Save function call area ------ %>
		Call CreditCheck("SAVE") 
		
		'Call DbSave														<% '��: Save db data %>
		
		FncSave = True													<% '��: Processing is OK %>
	End Function

'========================================================================================================
	Function FncCopy()
		Dim IntRetCD

		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")			<%'��: "Will you destory previous data"%>

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		lgIntFlgMode = Parent.OPMD_CMODE													<%'��: Indicates that current mode is Crate mode%>

		<% '------ ���Ǻ� �ʵ带 �����Ѵ�. ------ %>
		Call ggoOper.ClearField(Document, "1")										<%'��: Clear Condition Field%>
		Call ggoOper.LockField(Document, "N")	
		frm1.txtWarrentNo1.value = "" 
		
		If frm1.chkDeleteFlg.checked = False Then
			frm1.txtDelType.value = ""
			frm1.txtDelTypeNm.value = ""
			frm1.txtDelDt.text = ""
		End If
		Call chkDeleteFlg_OnPropertyChange()		
	End Function

'========================================================================================================
	Function FncPrint()
		Call Parent.FncPrint()														<%'��: Protect system from crashing%>
	End Function

'========================================================================================
	Function FncPrev() 
	    Dim strVal
	    
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
	        Call DisplayMsgBox("900002","x","x","x")  '�� �ٲ�κ� 
	        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
	        Exit Function
	    End If

		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		frm1.txtPrevNext.value = "PREV"

	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
	    strVal = strVal & "&txtWarrentNo=" & Trim(frm1.txtWarrentNo.value)				<%'��: ��ȸ ���� ����Ÿ %>
	    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'��: ��ȸ ���� ����Ÿ %>
	         
		Call RunMyBizASP(MyBizASP, strVal)
	End Function

'========================================================================================
	Function FncNext() 
	    Dim strVal
	    
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
	        Call DisplayMsgBox("900002","x","x","x")  '�� �ٲ�κ� 
	        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
	        Exit Function
	    End If

		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		frm1.txtPrevNext.value = "NEXT"

	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
	    strVal = strVal & "&txtWarrentNo=" & Trim(frm1.txtWarrentNo.value)				<%'��: ��ȸ ���� ����Ÿ %>
	    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'��: ��ȸ ���� ����Ÿ %>
	         
		Call RunMyBizASP(MyBizASP, strVal)
	End Function

'========================================================================================================
	Function FncExcel() 
		Call Parent.FncExport(Parent.C_SINGLE)
	End Function

'========================================================================================================
	Function FncFind() 
		Call Parent.FncFind(Parent.C_SINGLE, True)
	End Function

'========================================================================================================
	Function FncExit()
		Dim IntRetCD

		FncExit = False
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			<%'��: "Will you destory previous data"%>

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		FncExit = True
	End Function


'========================================================================================================
	Function DbQuery()
		Err.Clear															<%'��: Protect system from crashing%>

		DbQuery = False														<%'��: Processing is NG%>

		Dim strVal

		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtWarrentNo=" & Trim(frm1.txtWarrentNo.value)		<%'��: ��ȸ ���� ����Ÿ %>

		Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
		
		DbQuery = True															<%'��: Processing is NG%>
	End Function

<%
'========================================================================================================
'=	Event Name : CreditCheck																			
'=	Event Desc : �㺸�ݾ� ������ ���� �����ѵ� üũ 
'========================================================================================================
%>
	Function CreditCheck(ModeFlag)
		Err.Clear															<%'��: Protect system from crashing%>
		
		CreditCheck = False														<%'��: Processing is NG%>
		
		Dim strVal

		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If
		
		With frm1
			.txtMode.value = "CHECK"										<%'��: �����Ͻ� ó�� ASP �� ���� %>
			.txtCHKMode.value = ModeFlag 
			.txtFlgMode.value = lgIntFlgMode
			.txtInsrtUserId.value = Parent.gUsrID
				
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		End With
			
		CreditCheck = True	
	End Function


'========================================================================================================
	Function DbSave()
		Err.Clear															<%'��: Protect system from crashing%>
		
		DbSave = False														<%'��: Processing is NG%>
		
		Dim strVal
		
		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If
		
		With frm1
			.txtMode.value = Parent.UID_M0002										<%'��: �����Ͻ� ó�� ASP �� ���� %>
			.txtFlgMode.value = lgIntFlgMode
			.txtInsrtUserId.value = Parent.gUsrID
				
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		End With
			
		DbSave = True														<%'��: Processing is NG%>
	End Function

'========================================================================================================
	Function DbDelete()
		Err.Clear																<%'��: Protect system from crashing%>
		
		DbDelete = False														<%'��: Processing is NG%>
		
		Dim strVal
		
		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If
		
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003					<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtWarrentNo=" & Trim(frm1.txtWarrentNo1.value)		<%'��: ���� ���� ����Ÿ %>
		
		Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
		
		DbDelete = True															<%'��: Processing is NG%>
	End Function


'========================================================================================================
	Function DbQueryOk()														<% '��: ��ȸ ������ ������� %>
		<% '------ Reset variables area ------ %>
		lgIntFlgMode = Parent.OPMD_UMODE										<% '��: Indicates that current mode is Update mode %>
		lgBlnFlgChgValue = False
		
		Call ggoOper.LockField(Document, "Q")									<% '��: This function lock the suitable field %>
		Call SetToolBar("1111100000111111")
		
		Call chkDeleteFlg_OnPropertyChange()
		
		lgBlnFlgChgValue = False
			
	End Function
	

'========================================================================================================
	Function DbSaveOk()														<%'��: ���� ������ ���� ���� %>
		Call InitVariables
		Call MainQuery()
	End Function
	

'========================================================================================================
	Function DbDeleteOk()													<%'��: ���� ������ ���� ���� %>
		lgBlnFlgChgValue = False
		Call MainNew()
	End Function
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�㺸</font></td>
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
									<TD CLASS=TD5 NOWRAP>�㺸������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWarrentNo" SIZE="20" MAXLENGTH="18" TAG="12XXXU" ALT="�㺸������ȣ" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWarrentNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnWarrentNoOnClick()"></TD>
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�㺸������ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWarrentNo1" TYPE=TEXT SIZE="20" MAXLENGTH="18"   TAG="25XXXU" ALT="�㺸������ȣ"></TD>
								<TD CLASS=TD5 NOWRAP>��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCustomer" SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCustomer" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnCustomerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtCustomerNm" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�㺸����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWarrentType" TYPE=TEXT SIZE=10  MAXLENGTH="5" TAG="22XXXU" ALT="�㺸����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWarrentType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnWarrentTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtWarrentTypeNm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>�����׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE="10"  MAXLENGTH="4" TAG="22XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGrpOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>ȭ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE="10"  MAXLENGTH="3"  MAXLENGTH=3 TAG="23XXXU" ALT="ȭ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnCurrencyOnClick()"></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtWarrentAsignAmt.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s1413ma1_fpDateTime2_txtAsignDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s1413ma1_fpDateTime2_txtExpiryDt.js'></script></TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s1413ma1_fpDateTime2_txtEstimateDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>����ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtEstimatePlace" MAXLENGTH="35" SIZE=35 TAG="21X" ALT="����ó"></TD>
							</TR>
							<TR>				
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtEstimateAmt.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
								<TD CLASS=TD5 NOWRAP>���ɴ㺸��</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtWarrentAbleAmt.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProffer" MAXLENGTH="50" SIZE="35" TAG="21XXX" ALT="������"></TD>
								<TD CLASS=TD5 NOWRAP>������ �ֹε�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProfferRgstNo" MAXLENGTH="20" SIZE="35" TAG="21XXX" ALT="������ �ֹε�Ϲ�ȣ"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������ ����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRelationShip" MAXLENGTH="50" SIZE="35" TAG="21XXX" ALT="������ ����"></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtAsgnSeq.js'></script></TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>�������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWarrentOrgNm" MAXLENGTH="50" SIZE="35" TAG="21XXX" ALT="�������"></TD>
								<TD CLASS=TD5 NOWRAP>������� ����ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrgTelNo" MAXLENGTH="20" SIZE="35" TAG="21XXX" ALT="������� ����ó"></TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>���ǹ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStockNo" MAXLENGTH="35" SIZE="35" TAG="21XXX" ALT="���ǹ�ȣ"></TD>
								<TD CLASS=TD5 NOWRAP>�Ӵ���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLenderNm" MAXLENGTH="50" SIZE="35" TAG="21XXX" ALT="�Ӵ���"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 COLSPAN=3><INPUT TYPE=TEXT NAME="txtRocation" MAXLENGTH="120" SIZE="85" TAG="21XXX" ALT="������"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtFloorSpace.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtGroundSpace.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>	
							<TR>	
								<TD CLASS=TD5 NOWRAP>�ſ����Ȯ�αⰣ</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s1413ma1_fpDoubleSingle1_txtCreditCheckDt.js'></script>&nbsp;��</TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21XXX" VALUE="Y" NAME="chkDeleteFlg" ID="chkDeleteFlg">
									<LABEL FOR="chkDeleteFlg">�㺸����</LABEL>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDelType" SIZE="10"  MAXLENGTH=5 TAG="22XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDelType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnDelTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDelTypeNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s1413ma1_fpDateTime2_txtDelDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 COLSPAN=3><INPUT TYPE=TEXT NAME="txtRemark" MAXLENGTH="120" SIZE="85" TAG="21XXX" ALT="���"></TD>
							</TR>			
							<%Call SubFillRemBodyTD5656(3)%>
						</TABLE>
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
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck()">�㺸��Ȳ��ȸ</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCHKMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24"> 
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

