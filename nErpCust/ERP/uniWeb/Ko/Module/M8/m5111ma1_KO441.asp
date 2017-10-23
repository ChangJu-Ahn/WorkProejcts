<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procuremen
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/05/12
'*  8. Modified date(Last)  : 2003/09/22
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                           
'**********************************************************************************************-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc ����   ********************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->							
<!-- #Include file="../../inc/IncSvrHTML.inc" -->								
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--'==========================================  1.1.2 ���� Include   ======================================-->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                          '��: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

Const BIZ_PGM_ID 		= "m5111mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID 	= "m5112ma1_KO441"
Const BIZ_PGM_JUMP_ID2 	= "m5113ma1"

Const ivType = "ST"

Dim lgBlnFlgChgValue				'��: Variable is for Dirty flag
Dim lgIntFlgMode					'��: Variable is for Operation Status

Dim lgMpsFirmDate, lgLlcGivenDt											
Dim gSelframeFlg			                    'tab1,tab2 ����	
Dim lblnWinEvent
Dim interface_Account
Dim arrCollectVatType

'==========================================   ChangeSupplier()  ======================================
Sub ChangeSupplier(BpType)
	lgBlnFlgChgValue = true	
	if CheckRunningBizProcess = true then
		exit sub
	end if
	Call SpplRef(BpType)
End Sub

'==========================================   SpplRef()  ======================================
'	Name : SpplRef()
'	Description : It is Call at txtSupplier Change Event
'=========================================================================================================
Sub SpplRef(BpType)
	If gLookUpEnable = False Then
		Exit Sub
	End If

    Err.Clear                                                      '��: Protect system from crashing
    
    Dim strVal, StrvalBpCd
	Select Case BpType
		Case "1"                                                   '����ó�ΰ�� ȭ�� ���� 
		    if Trim(frm1.txtSpplCd.Value) = "" then
    			Exit Sub
    		End if
    		
    		StrvalBpCd = FilterVar(Trim(frm1.txtSpplCd.value), "", "SNM")
    	    
    	    if Trim(frm1.txtIvDt.Text) = ""  then
	            Call DisplayMsgBox("970021","X","���Ե����","X")
	            Exit Sub
	        End if
    	   
    	    if Trim(frm1.txtSpplCd.value) = ""  then
	            Call DisplayMsgBox("970021","X","����ó","X")
	            Exit Sub
	        End if
    	    
    	    Call GetPayDt()                                        '���ҿ����� setting
    	Case "2"                                                   '����ó�ΰ�� �����Ⱓ,��ݰ�������,������������ 
    		if Trim(frm1.txtPayeeCd.Value) = "" then               '���ֹ�ȣ no checked��� ����������� 
    			Exit Sub
    		End if
			StrvalBpCd = FilterVar(Trim(frm1.txtPayeeCd.value), "", "SNM")
    	Case "3"                                                  '���ݰ�꼭����ó�� ��� VAT,VAT�̸�,����ڵ�Ϲ�ȣ 
    		if Trim(frm1.txtBuildCd.Value) = "" then
    			Exit Sub
    		End if   	
    		StrvalBpCd = FilterVar(Trim(frm1.txtBuildCd.value), "", "SNM")
    		
	        Call GetTaxBizArea("BP")
	        '2003.1�� ������ġ(S) : ���ޱ� �˾���ư ����.(KJH : 03-01-06)
	        Call CheckPrePayedAmtYN()
	End Select
 
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"			'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strval & "&txtBpType=" & BpType
    strVal = strVal & "&txtBpCd=" & StrvalBpCd		'��: ��ȸ ���� ����Ÿ 
    
    if LayerShowHide(1) = False then
		Exit sub
	end if

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����	
	
End Sub

'--------------------------------------------------------------------
'		Field�� Tag�Ӽ��� Protect�� ��ȯ,���� ��Ű�� �Լ� 
'--------------------------------------------------------------------
Function ChangeTag(Byval Changeflg)

	if Changeflg = true then
		'Call ggoOper.SetReqAttr(frm1.txtIvTypeCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtIvDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPostDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtCur,"Q")
		Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtSpplCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayeeCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBuildCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtSpplIvNo,"Q")
		Call ggoOper.SetReqAttr(frm1.txtVatCd,"Q")
		'Call ggoOper.SetReqAttr(frm1.txtVatRt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtGrpCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayMethCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayDur,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayTypeCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPayTermstxt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtMemo,"Q")
		Call ggoOper.SetReqAttr(frm1.ChkPoNo, "Q")
        Call ggoOper.SetReqAttr(frm1.rdoVatFlg1,"Q")
        Call ggoOper.SetReqAttr(frm1.rdoVatFlg2,"Q")
        Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
        Call ggoOper.SetReqAttr(frm1.rdoIssueDTFg1,"Q")
        Call ggoOper.SetReqAttr(frm1.rdoIssueDTFg2,"Q")

	else
		'Call ggoOper.SetReqAttr(frm1.txtIvTypeCd,"N")
		Call ggoOper.SetReqAttr(frm1.txtIvDt,"N")                '���Ե���� 
		Call ggoOper.SetReqAttr(frm1.txtPayDt,"N")               '���ҿ����� 
		Call ggoOper.SetReqAttr(frm1.txtPostDt,"D")              '������ 
		'Call ggoOper.SetReqAttr(frm1.txtCur,"N")
		'Call ggoOper.SetReqAttr(frm1.txtXchRt,"D")

		Call ggoOper.SetReqAttr(frm1.txtPayeeCd,"N")             '����ó 
		Call ggoOper.SetReqAttr(frm1.txtBuildCd,"N")             '���ݰ�꼭����ó 
		Call ggoOper.SetReqAttr(frm1.txtSpplIvNo,"D")            '����ó 
		Call ggoOper.SetReqAttr(frm1.txtVatCd,"N")             'VAT
		Call ggoOper.SetReqAttr(frm1.txtGrpCd,"N")               '���ű׷� 

		Call ggoOper.SetReqAttr(frm1.txtPayDur,"D")              '�����Ⱓ 
		Call ggoOper.SetReqAttr(frm1.txtPayTypeCd,"D")           '�������� 
		Call ggoOper.SetReqAttr(frm1.txtPayTermstxt,"D")         '��ݰ������� 
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCd,"D")           '���ݽŰ����� 
		Call ggoOper.SetReqAttr(frm1.txtMemo,"D")                '��� 
		
 		if (UNICDbl(Trim(frm1.txtIvAmt.Value)) <> 0 and Trim(frm1.txtIvAmt.Value) <> "") or Trim(frm1.txtPoNo.value) = "" then	'iv detail�� �����ϸ� PO NO�� �����Ͽ� ����� �� ����.
			ggoOper.SetReqAttr	frm1.ChkPoNo, "Q"   'N: REQUIRED, D: UNREQUIRED ,Q:PROTECTED
		else                                        '���ֹ�ȣ ���� 
			ggoOper.SetReqAttr	frm1.ChkPoNo, "D"
		End if
	
		if Trim(frm1.txtPoNo.value) <> "" then                  '���������� ������� ȭ�� protect
			ggoOper.SetReqAttr	frm1.txtCur, "Q"                'ȭ�� 
			Call ggoOper.SetReqAttr(frm1.txtSpplCd, "Q" )       '����ó 
		else
			ggoOper.SetReqAttr	frm1.txtCur, "N"
			Call ggoOper.SetReqAttr(frm1.txtSpplCd,"N")         'txtSpplCd
		End if
		'��������� �׻� Required�̵��� ������.(2003.03.18)-Lee,Eun Hee
		Call ggoOper.SetReqAttr(frm1.txtPayMethCd,"N")      '������� 
		
	    if UCase(Trim(frm1.txtCur.value)) = UCase(parent.gCurrency) _
	      Or UCase(Trim(frm1.hdnRetflg.Value)) = "Y"   then
		   Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")            'ȯ�� 
		else       
		   Call ggoOper.SetReqAttr(frm1.txtXchRt,"D")
	    End if
	End if
	
End Function
'==========================================   Posting()  ======================================
'	Name : Posting()
'	Description : Ȯ����ư,Ȯ����ҹ�ư�� Event �ռ� 
'========================================================================================================= 
Sub Posting()
    Dim IntRetCD 
    
    Err.Clear                                                         '��: Protect system from crashing
    
    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217","X","X","X")                      '����Ÿ�� ����Ǿ����ϴ�. ������ �����Ͻʽÿ�.
		Exit sub
	End if
	
    if frm1.rdoApFlg(0).checked = True then                           'Ȯ������ 
					
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")       '�۾��� ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			frm1.btnPosting.disabled = False	'20040315  
			Exit Sub
		Else 
				frm1.btnPosting.disabled = True		'20040315 
		End If
		
		
			
		Call Changeflg()                                               'hidden�� setting �Լ� 
		Call DbSave("Posting")				             

	Elseif frm1.rdoApFlg(1).checked = True then
		
		if Trim(frm1.txtPostDt.text) = "" then
			Call DisplayMsgBox("17A002","X" , "������","X")        '%1�� �Է��ϼ���.
			Exit sub
		End if
		
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnPosting.disabled = False	'20040315  
			Exit Sub
		Else 
				frm1.btnPosting.disabled = True		'20040315 
		End If
		
		Call Changeflg()                                               'hidden�� setting �Լ� 
		Call DbSave("UnPosting")
		
	End if
	
End Sub
'--------------------------------------------------------------------
'		Cookie ����Լ� 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD
	If Kubun = 1 Then                                                  '���Գ������ ���� 

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then                                '����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
		
		WriteCookie "txtIvNo" , UCase(Trim(frm1.txtIvNo.value))
		WriteCookie "txtPoNo" , UCase(Trim(frm1.txtPoNo.value))
		Call PgmJump(BIZ_PGM_JUMP_ID)				  
		
	ElseIf Kubun = 0 Then

		strTemp = ReadCookie("txtIvNo")
		
		If strTemp = "" then Exit Function
		
		frm1.txtIvNo.value = strTemp
		
		WriteCookie "txtIvNo" , ""
		
		Call MainQuery()
	ElseIf Kubun = 2 Then                                               '���޳������ ���� 

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                              'Check if there is retrived data
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
		
		WriteCookie "txtIvNo" , UCase(Trim(frm1.txtIvNo.value))
		
		Call PgmJump(BIZ_PGM_JUMP_ID2)			
	End IF
	
End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lblnWinEvent = False

End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	With frm1
		.rdoApflg(0).Checked = False
		.rdoApflg(1).Checked = True
		.hdnApFlg.value= "N"
    	
		.txtIvDt.Text = EndDate
    	.txtPostDt.Text =EndDate
		Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
		'.PrepayNo.disabled = true
		.ChkPrepay.Checked =   false                 '���ޱݿ��� ���� check box		
		.btnPosting.disabled = true             'Ȯ��������ư 
    	.btnGlSel.disabled = true               '��ǥ��ȸ��ư 
    	.hdnLocCur.value = parent.gCurrency
    	.txtGrpCd.Value = parent.gPurGrp
    	.hdnUsrId.value = parent.gUsrID
    	.btnPosting.value = "Ȯ��"
    	.txtIvNo.focus 
    	Set gActiveElement = document.activeElement
    	
    	frm1.chkPoNo.checked = False
    	
    	Call ClickTab1()   
    	
    	Call SetToolBar("1110100000001111")
    	interface_Account = GetSetupMod(parent.gSetupMod, "a")
		'**����(2003.03.26)-ȸ������ ��� Ȯ��,��� �����ϵ��� ���� 
		'if interface_Account = "N" then
		'	'btnintAcc.style.display = "none"
		'	frm1.btnPosting.disabled = true
		'End if
		Call ggoOper.SetReqAttr(frm1.txtIvTypeCd,"N")        '�������� 
	End With
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
	
'------------------------------------------  OpenPoRef()  -------------------------------------------------
Function OpenPoRef()

	Dim strRet
	Dim arrParam(7)
	Dim iCalledAspName
	
	If lgIntFlgMode = parent.OPMD_UMODE Then 
			Call DisplayMsgBox("200005", "X", "X", "X")
			Exit function
	End If	
	if frm1.rdoApFlg(0).checked = true then
		Call DisplayMsgBox("17a009","X","X","X")
		Exit Function
	End if
	
	if Trim(frm1.txtIvDt.Text) = ""  then
	    Call DisplayMsgBox("970021","X","���Ե����","X")
	    Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
		
	iCalledAspName = AskPRAspName("m3111ra4_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m3111ra4_KO441", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		Call ClickTab1()
		frm1.txtIvNo1.focus
		Exit Function
	Else
		Call SetPoRef(strRet)
	End If	
		
End Function

 '------------------------------------------  OpenGLRef()  -------------------------------------------------
'	Name : OpenGLRef()
'	Description : ��ǥ��ȸ 
'--------------------------------------------------------------------------------------------------------- 
Function OpenGLRef()
	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtGlNo.value)

	If frm1.hdnGlType.Value = "A" Then
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "T" Then
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")
    End if
        
	lblnWinEvent = False
	
End Function

'------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : ���������� setting
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)
    Dim strVal
    
	Call ggoOper.ClearField(Document, "A")
    Call SetDefaultVal
	Call ggoOper.SetReqAttr(frm1.txtCur, "Q" )
	Call ggoOper.SetReqAttr(frm1.txtSpplCd, "Q" )
	'����(2003.03.18)-Lee,Eun Hee
	'��������� ���氡���ϵ��� ������.
	'Call ggoOper.SetReqAttr(frm1.txtPayMethCd, "Q" )

    Call InitVariables
    
	frm1.hdnPoNo.Value = strRet(0)
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpPo"							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPoNo=" & Trim(frm1.hdnPoNo.Value)				'��: ��ȸ ���� ����Ÿ 
   	If Trim(frm1.txtPoNo.value) <> "" Then frm1.chkPoNo.checked = True
 
    if LayerShowHide(1) = false then
		exit sub
	end if
    
    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
End Sub
'------------------------------------------  OpenIvNo()  -------------------------------------------------
Function OpenIvNo()
	Dim strRet
	Dim arrParam(0)
	Dim iCalledAspName
		
	If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	arrParam(0) = ivType
		
	iCalledAspName = AskPRAspName("m5111pa1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m5111pa1_KO441", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtIvNo.focus
		Exit Function
	Else
		frm1.txtIvNo.value = strRet(0)
		frm1.txtIvNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'------------------------------------------  OpenCommPopup()  -------------------------------------------------
Function OpenCommPopup(arrHeader, arrField, arrParam, arrRet)


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	If arrRet(0) = "" Then
		OpenCommPopup = False
	Else
		OpenCommPopup = True
		lgBlnFlgChgValue = True
	End If
	
End Function

'------------------------------------------  OpenCur()  -------------------------------------------------
Function OpenCur()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 

	If lblnWinEvent = True Or UCase(frm1.txtCur.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	lblnWinEvent = True
	arrHeader(0) = "ȭ��"						' Header��(0)
    arrHeader(1) = "ȭ���"						' Header��(1)
    
    arrField(0) = "Currency"						' Field��(0)
    arrField(1) = "Currency_Desc"					' Field��(1)
    
	arrParam(0) = "ȭ��"						' �˾� ��Ī 
	arrParam(1) = "B_Currency"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtCur.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtCur.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "ȭ��"						' TextBox ��Ī 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) Then
		frm1.txtCur.Value 	= arrRet(0)
		frm1.txtCurNm.Value = arrRet(1)
		Call ChangeCurr()
    End If
	lblnWinEvent = False
	frm1.txtCur.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenSppl()  -------------------------------------------------
Function OpenSppl(Byval BpType)
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	arrHeader(2) = "����ڵ�Ϲ�ȣ"									' Header��(2)
    arrField(0) = "B_BIZ_PARTNER.BP_Cd"									' Field��(0)
    arrField(1) = "B_BIZ_PARTNER.BP_Nm"								    ' Field��(1)
    arrField(2) = "B_BIZ_PARTNER.BP_RGST_NO"							' Field��(2)
    
	Select Case BpType
		Case "1"  '����ó 
			If lblnWinEvent = True Or UCase(frm1.txtSpplCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True
			arrHeader(0) = "����ó"											' Header��(0)
			arrHeader(1) = "����ó��"										' Header��(1)

		    arrParam(0) = "����ó"											' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER "					                    ' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtSpplCd.Value)		
			'arrParam(2) = Trim(frm1.txtSpplCd.Value)							' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "		' Where Condition
			arrParam(5) = "����ó"											' TextBox ��Ī 
		Case "2"   '����ó 
			If lblnWinEvent = True Or UCase(frm1.txtPayeeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True

			arrHeader(0) = "����ó"											' Header��(0)
			arrHeader(1) = "����ó��"										' Header��(1)

			arrParam(0) = "����ó"											' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
			'arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MPA", "''", "S") & " "
			arrParam(5) = "����ó"											' TextBox ��Ī 
		Case "3"   '���ݰ�꼭����ó 
			If lblnWinEvent = True Or UCase(frm1.txtBuildCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True

			arrHeader(0) = "���ݰ�꼭����ó"											' Header��(0)
			arrHeader(1) = "���ݰ�꼭����ó��" 										' Header��(1)

			arrParam(0) = "���ݰ�꼭����ó"											' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"           					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			'arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MBI", "''", "S") & " "
			arrParam(5) = "���ݰ�꼭����ó"											' TextBox ��Ī 
	End Select
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) Then
		Select Case BpType
			Case "1"   '����ó 
				frm1.txtSpplCd.Value = arrRet(0) : frm1.txtSpplNm.Value = arrRet(1)
				frm1.txtSpplCd.focus
			Case "2"   '����ó 
				frm1.txtPayeeCd.Value = arrRet(0) : frm1.txtPayeeNm.Value = arrRet(1)
				frm1.txtPayeeCd.focus
			Case "3"   '���ݰ�꼭����ó 
				frm1.txtBuildCd.Value = arrRet(0) : frm1.txtBuildNm.Value = arrRet(1) ': frm1.txtSpplRegNo.Value = arrRet(2)				
		        Call GetTaxBizArea("BP")
		        frm1.txtBuildCd.focus
		End Select 
		Call ChangeSupplier(BpType)
    End If
    lblnWinEvent = False
    Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenVat()  -------------------------------------------------
Function OpenVat()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Or UCase(frm1.txtVatCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
	lblnWinEvent = True
	
	arrHeader(0) = "VAT����"									' Header��(0)
    arrHeader(1) = "VAT���¸�"									' Header��(1)
    arrHeader(2) = "VAT��"									    ' Header��(2)
    
    arrField(0) = "b_minor.MINOR_CD"					            ' Field��(0)
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"					    ' Field��(1)
    
	arrParam(0) = "VAT"	            							' �˾� ��Ī 
	arrParam(1) = "B_MINOR,b_configuration"
	arrParam(2) = Trim(frm1.txtVatCd.Value)						    ' Code Condition
	'arrParam(2) = Trim(frm1.txtVatCd.Value)						    ' Code Condition
	'arrParam(3) = Trim(frm1.txtVatNm.Value)						' Name Cindition
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_configuration.seq_no=1 and b_minor.major_cd=b_configuration.major_cd"
	arrParam(5) = "VAT"										    ' TextBox ��Ī 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtVatCd.Value = arrRet(0)
		frm1.txtVatNm.Value = arrRet(1)
		frm1.txtVatRt.text  = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		'frm1.txtVatRt.text = arrRet(2)
    end if
    frm1.txtVatCd.focus
    Set gActiveElement = document.activeElement
    lblnWinEvent = False
End Function

'------------------------------------------  OpenGrp()  -------------------------------------------------
Function OpenGrp()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Or UCase(frm1.txtGrpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
	lblnWinEvent = True	
	arrHeader(0) = "���ű׷�"									' Header��(0)
    arrHeader(1) = "���ű׷��"									' Header��(1)
    
    arrField(0) = "PUR_GRP"											' Field��(0)
    arrField(1) = "PUR_GRP_NM"										' Field��(1)
    
	arrParam(0) = "���ű׷�"									' �˾� ��Ī 
	arrParam(1) = "B_PUR_GRP"										' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtGrpCd.Value)							' Code Condition
	'arrParam(2) = Trim(frm1.txtGrpCd.Value)							' Code Condition
																	' Where Condition
	arrParam(4) = "USAGE_FLG = " & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "���ű׷�"									' TextBox ��Ī 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtGrpCd.Value = arrRet(0)
		frm1.txtGrpNm.Value = arrRet(1)  
    end if
    Call GetTaxBizArea("*")
    lblnWinEvent = False
    frm1.txtGrpCd.focus
    Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenPayType()  -------------------------------------------------
Function OpenPayType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	if Trim(frm1.txtPayMethCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "�������","X")
		Exit Function
	End if

	If lblnWinEvent = True Or UCase(frm1.txtPaytypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	lblnWinEvent = True
	arrHeader(0) = "��������"						' Header��(0)
    arrHeader(1) = "����������"						' Header��(1)
    
    arrField(0) = "b_configuration.REFERENCE"			' Field��(0)
    arrField(1) = "B_Minor.Minor_Nm"					' Field��(1)
    
	arrParam(0) = "��������"						' �˾� ��Ī 
	arrParam(1) = "B_Minor,b_configuration"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPaytypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtPaytypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtPayTypeNm.Value)		' Name Cindition
	arrParam(4) = "b_configuration.minor_cd= " & FilterVar(Trim(frm1.txtPayMethCd.Value), "''", "S") & _
				  " And b_configuration.Major_Cd= " & FilterVar("B9004", "''", "S") & " and " & _
				  "b_minor.minor_cd=*b_configuration.REFERENCE and b_configuration.SEQ_NO>" & FilterVar("1", "''", "S") & "  And " & _
				  "b_minor.Major_Cd=" & FilterVar("A1006", "''", "S") & " and substring(b_configuration.REFERENCE,2,1) <> " & FilterVar("R", "''", "S") & "  "
				  
	arrParam(5) = "��������"						' TextBox ��Ī 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtPaytypeCd.Value = arrRet(0) : frm1.txtPayTypeNm.Value = arrRet(1)
    end if
    lblnWinEvent = False
    frm1.txtPaytypeCd.focus
    Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenPayMeth()  -------------------------------------------------
Function OpenPayMeth()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	If lblnWinEvent = True Or UCase(frm1.txtPayMethCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	lblnWinEvent = True
	arrHeader(0) = "�������"						        ' Header��(0)
    arrHeader(1) = "���������"						        ' Header��(1)
    arrHeader(2) = "Reference"
    
    arrField(0) = "B_Minor.MINOR_CD"							' Field��(0)
    arrField(1) = "B_Minor.MINOR_NM"							' Field��(1)
    arrField(2) = "b_configuration.REFERENCE"
    
	arrParam(0) = "�������"						        ' �˾� ��Ī 
	arrParam(1) = "B_Minor,b_configuration"				        ' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPayMethCd.Value)			        ' Code Condition
	'arrParam(2) = Trim(frm1.txtPayMethCd.Value)			        ' Code Condition
	'arrParam(3) = Trim(frm1.txtPayMethNM.Value)			    ' Name Cindition
	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & " and B_Minor.minor_cd =b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"	 
	arrParam(5) = "�������"						        ' TextBox ��Ī 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtPayMethCd.Value = arrRet(0) : frm1.txtPayMethNm.Value = arrRet(1)
		Call changePayMeth()
    end if
    lblnWinEvent = False
    frm1.txtPayMethCd.focus
    Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenBizArea()  -------------------------------------------------
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lblnWinEvent = True Or UCase(frm1.txtBizAreaCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True

	arrParam(0) = "���ݽŰ�����"	
	arrParam(1) = "B_TAX_BIZ_AREA"
	
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	'arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	
	'arrParam(4) = "Tax_Flag = 'Y'"
	arrParam(4) = ""
	arrParam(5) = "���ݽŰ�����"			
	
    arrField(0) = "TAX_BIZ_AREA_CD"
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "���ݽŰ�����"
    arrHeader(1) = "���ݽŰ������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd.Value = arrRet(0)
		frm1.txtBizAreaNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
	End If	
	frm1.txtBizAreaCd.focus
	Set gActiveElement = document.activeElement
End Function


'------------------------------------------  OpenIvType()  -------------------------------------------------
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Or UCase(frm1.txtIvTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True
	
	arrHeader(0) = "��������"						' Header��(0)
    arrHeader(1) = "�������¸�"						' Header��(1)
    
    arrField(0) = "IV_TYPE_CD"							' Field��(0)
    arrField(1) = "IV_TYPE_NM"							' Field��(1)
    
	arrParam(0) = "��������"						' �˾� ��Ī 
	arrParam(1) = "M_IV_TYPE"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			' Name Cindition
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and import_flg=" & FilterVar("N", "''", "S") & " "						' Where Condition
	arrParam(5) = "��������"						' TextBox ��Ī 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtIvTypeCd.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
    end if
    lblnWinEvent = False
    frm1.txtIvTypeCd.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenLoanNo()  -------------------------------------------------
Function OpenLoanNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lblnWinEvent = True Or UCase(frm1.txtLoanNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	if Trim(frm1.txtSpplCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "����ó","X")
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	elseif Trim(frm1.txtCur.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "ȭ��","X")
		frm1.txtCur.focus
		Set gActiveElement = document.activeElement
		Exit Function
	end if
	
	lblnWinEvent = True

	arrParam(0) = "���Աݹ�ȣ"	
	arrParam(1) = "F_LOAN"
	arrParam(2) = Trim(frm1.txtLoanNo.Value)
	'arrParam(2) = Trim(frm1.txtLoanNo.Value)
	
	arrParam(4) = "DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  And BP_CD =  " & FilterVar(frm1.txtSpplCd.Value, "''", "S") & " "
	'arrParam(4) = "DOC_CUR = '" & Trim(frm1.txtCur.Value) & "' And BP_CD = '" & Trim(frm1.txtSpplCd.Value) & "'"
	arrParam(5) = "���Աݹ�ȣ"			
	
    arrField(0) = "LOAN_NO"
    arrField(1) = "F2" & parent.gColSep & "LOAN_AMT"
    arrHeader(0) = "���Աݹ�ȣ"		
    arrHeader(1) = "���Ա�"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanNo.focus
		Exit Function
	Else
		frm1.txtLoanNo.Value 	= arrRet(0)
		frm1.txtLoanAmt.Text 	= arrRet(1)
	End If	
	frm1.txtLoanNo.focus
	Set gActiveElement = document.activeElement

End Function

<!--
'------------------------------------------  OpenPpNo()  -------------------------------------------------
-->
Function OpenPpNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lblnWinEvent = True or frm1.ChkPrepay.Checked = false Then 
	'or frm1.rdoApFlg(0).checked = true or Trim(UCase(frm1.hdnImportflg.Value)) = "Y" Then 
	    Exit Function
	end if

	if Trim(frm1.txtCur.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "ȭ��","X")
		Exit Function
	elseif Trim(frm1.txtSpplCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "������","X")
		Exit Function
	end if
	
	lblnWinEvent = True

	arrParam(0) = "���ޱݹ�ȣ"	
	arrParam(1) = "F_PRPAYM"
	
	arrParam(2) = ""
	'====================== 1�� ������ġ(S) : ���ޱ� �˾���ư ����.(KJH : 03-01-06)============================
	arrParam(4) = "DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  And BP_CD =  " & FilterVar(frm1.txtPayeeCd.Value, "''", "S") & "  AND BAL_AMT > 0 AND CONF_FG = " & FilterVar("C", "''", "S") & " "
	'====================== 1�� ������ġ(E) : ���ޱ� �˾���ư ����.(KJH : 03-01-06)============================
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
	
	lblnWinEvent = False
	

End Function

'=====================================  Changeflg()  ============================================
Sub Changeflg()
	if frm1.rdoApFlg(0).checked = true then
		frm1.hdnApFlg.value= "Y"
	else
		frm1.hdnApFlg.value= "N"
	end if 
End Sub

'=====================================  ChangeCurr()  ============================================
Sub ChangeCurr()
	if UCase(Trim(frm1.txtCur.value)) = UCase(parent.gCurrency) then
		frm1.txtXchRt.Text = 1
		Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")
	else
		frm1.txtXchRt.Text = ""
		Call ggoOper.SetReqAttr(frm1.txtXchRt,"D")
	end if 
	Call CurFormatNumericOCX()
End Sub

'=====================================  InitCollectType()  =========================================
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
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
'=====================================  GetCollectTypeRef()  =========================================
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
'=====================================  SetVatType()  =========================================
Sub SetVatType()

	Dim VatType, VatTypeNm, VatRate

	VatType = Trim(frm1.txtVatCd.value)
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
    
	frm1.txtVatNm.value = VatTypeNm
	frm1.txtVatRt.text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

End Sub

'======================================   GetTaxBizArea()  =====================================
Sub GetTaxBizArea(Byval strFlag)
   	Dim strSelectList, strFromList, strWhereList
	Dim strBilltoParty, strSalesGrp, strTaxBizArea
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp
	
    
	If strFlag = "NM" Then                              '���ݽŰ����� ����� �̸����� �����´� 
		strTaxBizArea = frm1.txtBizAreaCd.value
	Else
		strBilltoParty = frm1.txtBuildCd.value          '��꼭 ����ó 
		strSalesGrp    = frm1.txtGrpCd.value            '���ű׷� 
		
		<%'����ó�� ���� �׷��� ��� ��ϵǾ� �ִ� ��� �����ڵ忡 ������ rule�� ������ %>
		If Len(strBillToParty) > 0 And Len(strSalesGrp) > 0	Then strFlag = "*"
	End if
		
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetTaxBizArea ( " & FilterVar(strBilltoParty, "''", "S") & " ,  " & FilterVar(strSalesGrp, "''", "S") & " ,  " & FilterVar(strTaxBizArea, "''", "S") & " ,  " & FilterVar(strFlag, "''", "S") & " ) "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		frm1.txtBizAreaCd.value = arrTemp(1)
		frm1.txtBizAreaNm.value = arrTemp(2)
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Sub
		End If

		' ���� �Ű� ������� Editing�� ��� 
		'If strFlag = "NM" Then
		'	If Not OpenBillHdr(3) Then
				frm1.txtBizAreaCd.value = ""
				frm1.txtBizAreaNm.value = ""
		'	End if
		'End if
	End if
End Sub

<%'======================================   GetPayDt()  =====================================
'	Name : GetPayDt()
'	Description : ���ҿ������� �����´�.
'==================================================================================================== %>
Sub GetPayDt()
   	Dim strSelectList, strFromList, strWhereList
	Dim strSpplCd, strIvDt,temp
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp

    	strSpplCd  = frm1.txtSpplCd.value                       '����ó	
    	temp    = UNIConvDate(frm1.txtIvDt.text)            '���Ե���� 
		strIvDt = mid(temp,1,4)
		strIvDt = strIvDt & mid(temp,6,2)
		strIvDt = strIvDt & mid(temp,9,2) 
		<%'����ó�� ���� �׷��� ��� ��ϵǾ� �ִ� ��� �����ڵ忡 ������ rule�� ������ %>
    
	
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetPayDt( " & FilterVar(strSpplCd, "''", "S") & " ,  " & FilterVar(strIvDt, "''", "S") & " ) "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		frm1.txtPayDt.text = UNIFormatDate(arrTemp(1))
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Sub
		End If

		frm1.txtPayDt.text = ""

	End if
End Sub

'========================================================================================
' Function Name : DbPoQueryOK()
'========================================================================================
Function DbPoQueryOK()
	'Call SetToolBar("1110100000001111")
	Call ggoOper.SetReqAttr(frm1.txtIvTypeCd,"Q")
	'2003.1�� ������ġ(S) : ���ޱ� �˾���ư ����.(KJH : 03-01-06)
	Call CheckPrePayedAmtYN()
End Function

'========================================================================================
' Function Name : changePayMeth
'========================================================================================
Sub changePayMeth()
	
	frm1.txtPayTypeCd.Value = ""
	frm1.txtPayTypeNm.Value = ""
	frm1.txtPayDur.Text = 0	

End Sub

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	With frm1   
	
		'���Աݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtIvAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'���Լ��ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtnetDocAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'VAT�ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    '�����ڱ��ݾ� 
	    ggoOper.FormatFieldByObjectOfCur .txtnetLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    '�Ѹ����ڱ��ݾ� 
	    ggoOper.FormatFieldByObjectOfCur .txtIvLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    'vat�ڱ��ݾ� 
	    ggoOper.FormatFieldByObjectOfCur .txtVatLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    'ȯ�� 
	    ggoOper.FormatFieldByObjectOfCur .txtXchRt, .txtCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	    
	    
	    
	    
	End With

End Sub
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

   
    Call LoadInfTB19029	    														'��: Load table , B_numeric_format
    
    Call AppendNumberRange("0","0","999")					'�Ⱓ	
    Call AppendNumberPlace("7","2","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field    
    Call GetValue_ko441()
    Call SetDefaultVal   
    Call InitVariables
    Call cookiepage(0)
    Call changeTabs(TAB1)
    gSelframeFlg = TAB1
	gIsTab     = "Y" 
    gTabMaxCnt = 2                                   'tab ���� 
End Sub

'========================================  Form_QueryUnload()  ======================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'========================================  OCX_EVENT  ====================================
Sub txtIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIvDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIvDt.focus
	End if
End Sub

Sub txtIvDt_Change()
    
    lgBlnFlgChgValue = true	
    
    if Trim(frm1.txtIvDt.Text) = ""  then
  	    Exit Sub
  	End if	    
    
    Call GetPayDt()                       '���ҿ�����  
End Sub

Sub txtPayDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPayDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPayDt.focus
	End if
End Sub

Sub txtPayDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtPostDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPostDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPostDt.focus
	End if
End Sub

Sub txtPostDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'���ֹ�ȣ���� Ŭ���� 
'========================================  chkPoNo_onpropertychange() ====================================
Sub chkPoNo_onpropertychange()

	if frm1.ChkPoNo.checked = True and Trim(frm1.txtPoNo.Value) <> "" then     '���������� ������� ȭ�� protect
		Call ggoOper.SetReqAttr(frm1.txtCur, "Q" )
		Call ggoOper.SetReqAttr(frm1.txtSpplCd, "Q" )
		'����(2003.03.18)-Lee,Eun Hee
		'Call ggoOper.SetReqAttr(frm1.txtPayMethCd, "Q" )

		frm1.txtCur.value = frm1.hdnCur.value
		frm1.txtSpplCd.value =frm1.hdnSpplCd.value
		frm1.txtPayMethCd.value =frm1.hdnPayMethCd.value
	else
		Call ggoOper.SetReqAttr(frm1.txtCur, "N" )
		Call ggoOper.SetReqAttr(frm1.txtSpplCd, "N" )
		Call ggoOper.SetReqAttr(frm1.txtPayMethCd, "N" )

	End if
End Sub

Sub chkPoNo_OnClick()
	lgBlnFlgChgValue = true
End Sub

'===================================  Change_Event  ============================================
Sub txtIvAmt_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtVatRt_Change()
	'lgBlnFlgChgValue = true	
End Sub
Sub txtVatAmt_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtPayDur_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtXchRt_Change()
	lgBlnFlgChgValue = true	
End Sub

<%'==========================================================================================
'   Event Name : txtBillToPartyCd_OnChange
'   Event Desc : ����ó ������ ����Ǿ����� ���� �׸� LookUp
'==========================================================================================%>
Sub txtBuildCd_OnChange()
		If Trim(frm1.txtBuildCd.value) = "" Then
			'frm1.txtBillToPartyNm.value = ""
		Else
			Call GetTaxBizArea("BP")
		End if
End Sub

Sub txtGrpCd_OnChange()
		If Trim(frm1.txtGrpCd.value) = "" Then
			'frm1.txtBillToPartyNm.value = ""
		Else
			Call GetTaxBizArea("*")
		End if
End Sub

Sub txtBizAreaCd_OnChange()
		If Trim(frm1.txtBizAreaCd.value) = "" Then
			frm1.txtBizAreaNm.value = ""
		Else
			Call GetTaxBizArea("NM")
		End if
End Sub

Sub rdoVatFlg1_OnClick()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoVatFlg2_OnClick()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoIssueDTFg1_OnClick()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoIssueDTFg2_OnClick()
	lgBlnFlgChgValue = true	
End Sub

'====================== 2003.1�� ������ġ(S) : ���ޱ� �˾���ư ����.(KJH : 03-01-06)==========
'======================================   CheckPrePayedAmtYN()  =============================
'	Name : CheckPrePayedAmtYN()
'	Description : ���ޱݿ��θ� üũ�Ѵ�.
'============================================================================================
Sub CheckPrePayedAmtYN()
	Dim strSelectList,strFromList,strWhereList
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim iCount
 
	strSelectList	= " COUNT(*) "
	strFromList		= " F_PRPAYM " 
	strWhereList	= " BP_CD= " & FilterVar(frm1.txtPayeeCd.Value, "''", "S") & " "
    strWhereList	= strWhereList & " AND DOC_CUR =  " & FilterVar(frm1.txtCur.value, "''", "S") & "  AND BAL_AMT > 0 AND CONF_FG = " & FilterVar("C", "''", "S") & "  "
    
	Call CommonQueryRs(strSelectList,strFromList,strWhereList, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Err.number <> 0 Then
		Exit Sub
	End If

	 iCount = Split(lgF0, Chr(11))
	    
	if UNICDbl(Trim(iCount(0))) > 0 then
		frm1.ChkPrepay.checked = true
	else
		frm1.ChkPrepay.checked = false
	End if

End Sub
'==============================  FncQuery()  ================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                             '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
       
    If Not chkField(Document, "1") Then											'��: This function check indispensable field
       Exit Function
    End If
    
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables 
    
    If DbQuery = False Then Exit Function										'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    Set gActiveElement = document.activeElement    
End Function
'==============================  FncNew()  ================================================
Function FncNew() 

	Dim IntRetCD
    FncNew = False                                                          '��: Processing is NG
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ClickTab1()
    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call ChangeTag(False)
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
    
    FncNew = True																'��: Processing is OK
	Set gActiveElement = document.activeElement
End Function

'==============================  FncDelete()  ================================================
Function FncDelete() 
    
Dim IntRetCD

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function

    FncDelete = False														'��: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function									'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'==============================  FncSave()  ================================================
Function FncSave() 
Dim IntRetCD 

    FncSave = False                                                         '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing
    
	if CheckRunningBizProcess = true then
		exit function
	end if

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
	
    If Not chkField(Document, "2") Then                             '��: Check contents area
       If gPageNo > 0 Then
	       gSelframeFlg = gPageNo
	   End If
       Exit Function
    End If
    
    '**_**_*_*_*_*_*_*_*__*_*_*_***__*_*_*_*_*
    ' 2002-08-21
    ' ���ݽŰ������� Null�� �� �� �ҷ����� 
    '**_**_*_*_*_*_*_*_*__*_*_*_***__*_*_*_*_*
    If Trim(frm1.txtBizAreaCd.value) = "" then 
		Call GetTaxBizArea("BP")
	end if 

    if frm1.rdoApFlg(0).checked = true then
    	frm1.hdnApFlg.Value = "Y"
    else
    	frm1.hdnApFlg.Value = "N"
    End if
    'vat ���Կ��� 
    if frm1.rdoVatFlg1.checked = true then
    	frm1.hdvatFlg.Value = "1"
    else
    	frm1.hdvatFlg.Value = "2"
    End if

    '���ҿ����� "" �� ���
    If frm1.txtPayDt.Text = "" then
	Call DisplayMsgBox("17A002","X" , "���ҿ�����","X")
	Exit Function
    End IF

    ' ���ڼ��ݰ�꼭���� 
    if frm1.rdoIssueDTFg1.checked = true then
    	frm1.hdIssueDTFg.Value = "Y"
    else
    	frm1.hdIssueDTFg.Value = "N"
    End If
  
    If DbSave("toolbar") = False Then Exit Function                         '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function
'==============================  FncCopy()  ================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE											'��: Indicates that current mode is Crate mode
    
     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                              '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'��: This function lock the suitable field
    Call SetToolBar("11101000000011")
    frm1.rdoApFlg(1).checked = true
    frm1.btnPosting.disabled = true
    frm1.txtIvNo1.value = ""
    frm1.txtPoNo.Value = ""
    frm1.chkPoNo.checked = False
    Call ChangeTag(False)
    lgBlnFlgChgValue = True
    Set gActiveElement = document.activeElement 
End Function

'==============================  FncCancel()  ================================================
Function FncCancel() 
    On Error Resume Next                                                 '��: Protect system from crashing
End Function
'==============================  FncInsertRow()  ================================================
Function FncInsertRow() 
     On Error Resume Next                                                 '��: Protect system from crashing
End Function
'==============================  FncDeleteRow()  ================================================
Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function
'==============================  FncPrint()  ================================================
Function FncPrint() 
    Call parent.FncPrint()                                                '��: Protect system from crashing
    Set gActiveElement = document.activeElement
End Function
'==============================  FncPrev()  ================================================
Function FncPrev() 
    Dim strVal
End Function
'==============================  FncNext()  ================================================
Function FncNext() 
    Dim strVal
End Function
'==============================  FncExcel()  ================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)             '��: Protect system from crashing
End Function
'==============================  FncFind()  ================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                  '��:ȭ�� ����, Tab ���� 
    Set gActiveElement = document.activeElement
End Function
'==============================  FncExit()  ================================================
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
    Set gActiveElement = document.activeElement
End Function
'==============================  DbDelete()  ================================================
Function DbDelete() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtIvNo=" & FilterVar(Trim(frm1.txtIvNo.value), "", "SNM")				'��: ���� ���� ����Ÿ 
    
    if LayerShowHide(1) = false then
		exit function
	end if

    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG

End Function
'==============================  DbDeleteOk()  ================================================
Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
'==============================  DbQuery()  ================================================
Function DbQuery() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo.value)				'��: ��ȸ ���� ����Ÿ %>
    
    if LayerShowHide(1) = false then
		exit function
	end if
    
    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG

End Function
'==============================  DbQueryOk()  ================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	

    Call InitVariables
    
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
 
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field


    if frm1.rdoApFlg(0).checked = true or Trim(UCase(frm1.hdnImportflg.Value)) = "Y"  then  'Ȯ������ 
		Call ChangeTag(true)
 
		if Trim(UCase(frm1.hdnImportflg.Value)) = "Y"  then
			Call SetToolBar("11100000000111")
			frm1.btnPosting.disabled = true
			frm1.rdoVatFlg1.disabled = true
			frm1.rdoVatFlg2.disabled = true
		else
			Call SetToolBar("11100000001111")
			if UNICDbl(frm1.txtIvAmt.Value) = 0 or Trim(frm1.txtIvAmt.Value) = "" then       '���Գ����� ������ 
				
					frm1.btnPosting.disabled = true
					frm1.rdoVatFlg1.disabled = false
					frm1.rdoVatFlg2.disabled = false
			else
					frm1.btnPosting.disabled = false
					frm1.rdoVatFlg1.disabled = true
					frm1.rdoVatFlg2.disabled = true
			end if
		end if
        
	        frm1.rdoIssueDTFg1.disabled = true
	        frm1.rdoIssueDTFg2.disabled = true

	Else
		Call ChangeTag(False)
		Call SetToolBar("11111000001111")
		if UNICDbl(frm1.txtIvAmt.Value) = 0 or Trim(frm1.txtIvAmt.Value) = "" then         '���Գ����� ������ 
				
				frm1.btnPosting.disabled = true                                            '���Աݾ��� ������ Ȯ���Ұ� 
				frm1.rdoVatFlg1.disabled = false
				frm1.rdoVatFlg2.disabled = false
		else
				
				frm1.btnPosting.disabled = false
				frm1.rdoVatFlg1.disabled = true
				frm1.rdoVatFlg2.disabled = true
		end if
	end if
  	
	if frm1.rdoApFlg(0).checked = true then                                                'Ȯ���̵Ǿ� ��ǥ��ȸ�� ���� 
		frm1.btnPosting.value = "Ȯ�����"
		
		if interface_Account <> "N" then
			frm1.btnGlSel.disabled = false
		Else
			frm1.btnGlSel.disabled = true
		end if
		
		if frm1.hdnGlType.Value = "A" Then
		   frm1.btnGlSel.value = "ȸ����ǥ��ȸ"
		elseif frm1.hdnGlType.Value = "T" Then
		   frm1.btnGlSel.value = "������ǥ��ȸ"
		end if
	else
		frm1.btnPosting.value = "Ȯ��"
		frm1.btnGlSel.disabled = true
	end if
	  
  if frm1.ChkPrepay.checked = false then
    Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
  end if
  
'2009-09-02 ������ ���� ��û���� ȭ���ʵ� ����
  'if UNICDbl(Trim(frm1.txtnetLocAmt.Value)) <> 0  then
  '	Call ggoOper.SetReqAttr(frm1.txtCur,"Q") 
  '	Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")
  'End if 
  Call ClickTab1()
End Function
'==============================  DbSave()  ================================================
Function DbSave(byval btnflg) 

    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG

    Dim strVal

	With frm1
		.hdnUsrId.value = parent.gUsrID
		.txtMode.value = parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		
		if btnflg = "Posting" then
			.txtMode.value = "Release" 				                        '��: Y=Ȯ�� ��ư 
		elseif btnflg = "UnPosting" then
			.txtMode.value = "UnRelease" 				                    '��: Y=Ȯ����� ��ư 
		end if
		      	
		if LayerShowHide(1) = false then
			exit function
		end if

		If .chkPoNo.checked = True Then
			.txtChkPoNo.value = "Y"                                         'hidden
		Else
			.txtChkPoNo.value = "N"
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function
'==============================  DbSaveOk()  ================================================
Function DbSaveOk()														'��: ���� ������ ���� ���� 
	lgBlnFlgChgValue = False
	Call MainQuery()
End Function

'============================================================================================================
' Name : SubGetGlNo
' Desc : Get Gl_no : 2003.03 KJH ��ǥ��ȣ �������� ���� ���� 
'============================================================================================================
Sub SubGetGlNo()
	Dim lgStrFrom
	Dim strTempGlNo, strGlNo
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	On Error Resume Next
	Err.Clear 
	
	lgStrFrom =  " ufn_a_GetGlNo( " & FilterVar(frm1.hdnIvNo.Value, "''", "S") & " )"
	
	Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", lgStrFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 <> "" then 
		strTempGlNo = Split(lgF0, Chr(11))
		strGlNo		= Split(lgF1, Chr(11))
					
		If strGlNo(0) = "" and strTempGlNo(0) = "" then 
			frm1.txtGlNo.Value		=	""
			frm1.hdnGlType.value	=	"B"
		Elseif strGlNo(0) = "" and strTempGlNo(0) <> "" then
			frm1.txtGlNo.Value		=	strTempGlNo(0) 
			frm1.hdnGlType.value	=	"T"
		Elseif strGlNo(0) <> "" then 
			frm1.txtGlNo.Value		=	strGlNo(0) 
			frm1.hdnGlType.value	=	"A"
		End If
	Else
		frm1.txtGlNo.Value		=	""
		frm1.hdnGlType.value	=	"B"
	End if
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���Լ��ݰ�꼭</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���Լ��ݰ�꼭��Ÿ</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPORef()">��������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT="*">
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" nowrap>���Թ�ȣ</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvNo" style="HEIGHT: 20px; WIDTH: 250px" MAXLENGTH=18 ALT="���Թ�ȣ" STYLE="TEXT-ALIGN:left; TEXT-TRANSFORM:UPPERCASE" tag="12N"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvNo" style="HEIGHT: 21px; WIDTH: 16px" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvNo()"></TD>
									<TD CLASS="TD6" nowrap></TD>
									<TD CLASS="TD6" nowrap></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
					<!-- ù��° �� ���� -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">	
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" nowrap>���Թ�ȣ</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvNo1" ALT="���Թ�ȣ" style="HEIGHT: 20px; WIDTH: 250px" MAXLENGTH=18 STYLE="TEXT-ALIGN:left; TEXT-TRANSFORM:UPPERCASE" tag="25X"></TD>
								<TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtPoNo" ALT="���ֹ�ȣ" TYPE=TEXT MAXLENGTH=35 SIZE=25 TAG="24XXXU">
								    <INPUT TYPE=CHECKBOX NAME="chkPoNo" tag="25" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkPoNo">���ֹ�ȣ����</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>��������</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvTypeCd" ALT="��������" MAXLENGTH=5 style="HEIGHT: 20px; WIDTH: 80px" tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px"  align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
													   <INPUT TYPE=TEXT NAME="txtIvTypeNm" ALT="��������" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<TD CLASS="TD5" nowrap>���Ե����</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���Ե���� NAME="txtIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22N1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>Ȯ������</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=radio NAME="rdoApFlg" ALT="Ȯ������" CLASS="RADIO" tag="24X"><label for="rdoApFlg"> Yes </label>
													   <INPUT TYPE=radio NAME="rdoApFlg" ALT="Ȯ������" CLASS="RADIO" checked tag="24X"><label for="rdoApFlg">  No&nbsp;&nbsp;</label>
													   <INPUT TYPE=TEXT NAME="txtGlNo" ALT="��ǥ��ȣ" style="HEIGHT: 20px; WIDTH: 148px" tag="24X"></TD>
							
								<TD CLASS="TD5" nowrap>���ҿ�����</TD>
								<TD CLASS="TD6" nowrap>
								    <Table Cellspacing=0 Cellpadding=0>
								        <TR>
										    <TD NOWRAP>
								            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���ҿ����� NAME="txtPayDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								            <TD NOWRAP>&nbsp;������</TD>
								            <TD NOWRAP>
								            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=������ NAME="txtPostDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							            </TR>
							        </Table>
							    </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>����ó</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtSpplCd" ALT="����ó" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(1)"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnSupplier" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(1)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT NAME="txtSpplNm" ALT="����ó" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<TD CLASS="TD5" nowrap>����ڵ�Ϲ�ȣ</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtSpplRegNo" ALT="����ڵ�Ϲ�ȣ" MAXLENGTH=10 style="HEIGHT: 20px; WIDTH: 250px"  tag="24X"></TD>
                            </TR>
							
							<TR>
	                            <TD CLASS="TD5" nowrap>���ݰ�꼭����ó</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBuildCd" ALT="���ݰ�꼭����ó" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(3)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSupplier" style="HEIGHT: 21px; WIDTH: 16px"  align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(3)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT NAME="txtBuildNm" ALT="���ݰ�꼭����ó" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<TD CLASS="TD5" nowrap>����ó</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayeeCd" ALT="����ó" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(2)"><IMG SRC="../../../CShared/image/btnPopup.gif"  style="HEIGHT: 21px; WIDTH: 16px" NAME="btnSupplier" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(2)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT NAME="txtPayeeNm" ALT="����ó" tag="24X"></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>ȭ��</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtCur" ALT="ȭ��" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=3 tag="22NXXU" onChange="ChangeCurr()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCur() ">&nbsp; 
										   <INPUT TYPE=HIDDEN NAME="txtCurNm" ALT="ȭ��" style="HEIGHT: 20px; WIDTH: 46px" tag="24X">
											
								</TD>
								<TD CLASS="TD5" nowrap>ȯ��</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=ȯ�� NAME="txtXchRt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>�Ѹ��Աݾ�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�Ѹ��Աݾ� NAME="txtIvAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" nowrap>�Ѹ����ڱ��ݾ�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�Ѹ����ڱ��ݾ� NAME="txtIvLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>���Աݾ�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���Աݾ� NAME="txtnetDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" nowrap>�����ڱ��ݾ�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�����ڱ��ݾ� NAME="txtnetLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>VAT</TD>
								<TD CLASS="TD6" NOWRAP>
									<Table cellpadding=0 cellspacing=0>
										<TR>
											<TD NOWRAP><INPUT TYPE=TEXT NAME="txtVatCd" ALT="VAT" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU"
											ONChange="vbscript:SetVatType()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnVat" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenVat()">
													   <INPUT TYPE=TEXT NAME="txtVatNm" ALT="VAT" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" >&nbsp;
											</TD>
				
							
										</TR>
									</Table>
								<TD CLASS="TD5" nowrap>VAT���Կ���</TD>
								<TD CLASS="TD6" nowrap>
								     <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT���Ա���" CLASS="RADIO" checked id="rdoVatFlg1" tag="21X"><label for="rdoVatFlg"> ���� </label>
									 <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT���Ա���" CLASS="RADIO" id="rdoVatFlg2"  tag="21X"><label for="rdoVatFlg">  ����&nbsp;</label></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>VAT��</TD>
								<TD CLASS="TD6" nowrap>								
									<Table cellpadding=0 cellspacing=0>
										<TR>
											<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT NAME="txtVatRt" MAXLENGTH=10 CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 style="HEIGHT: 20px; WIDTH: 96px" tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
											</TD>
											<TD WIDTH="*" NOWRAP>%
											</TD>
										</TR>
									</Table>
								</TD>
								<TD CLASS="TD5" nowrap></TD>
								<TD CLASS="TD6" nowrap></td>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>VAT�ݾ�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT�ݾ� NAME="txtVatAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
							    <TD CLASS="TD5" nowrap>VAT�ڱ��ݾ�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT�ڱ��ݾ� NAME="txtVatLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>�������</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayMethCd" ALT="�������" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU" OnChange="VBScript:changePayMeth()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayMeth()">
													   <INPUT TYPE=TEXT NAME="txtPayMethNm" ALT="�������" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								<TD CLASS="TD5" nowrap>�����Ⱓ</TD>
								<TD CLASS="TD6" NOWRAP>
									<Table cellpadding=0 cellspacing=0>
										<TR>
											<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=�����Ⱓ NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												&nbsp;��
											</TD>
										</TR>
									</Table>
								</TD>
							</TR>
							
							
							
							<TR>
								<TD CLASS="TD5" nowrap>��������</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTypeCd" ALT="��������" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayType()">
													   <INPUT TYPE=TEXT NAME="txtPayTypeNm" ALT="��������" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=CHECKBOX CHECKED ID="ChkPrepay" tag="24" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">&nbsp;���ޱݿ���
											           <INPUT TYPE=TEXT NAME="ChkPrepay1"  style="HEIGHT: 19px; WIDTH: 1px" MAXLENGTH=0 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" id="PrepayNo" NAME="PrepayNo" align=top TYPE="BUTTON" onclick="vbscript:OpenPpNo()"></TD>
											           
		                                               
							</TR>

							<TR>
								<TD CLASS="TD5" nowrap>���ݽŰ�����</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBizAreaCd" ALT="���ݽŰ�����" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
													   <INPUT TYPE=TEXT NAME="txtBizAreaNm" ALT="���ݽŰ�����" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								<!--<TD CLASS="TD5" nowrap>���Աݹ�ȣ</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtLoanNo" ALT="���Աݹ�ȣ" MAXLENGTH=18 style="HEIGHT: 20px; WIDTH: 250px" tag="24N"><!--<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenLoanNo() "></TD>-->
								<TD CLASS="TD5" nowrap>B/L������ȣ</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBlDocNo" ALT="B/L������ȣ"  style="HEIGHT: 20px; WIDTH: 250px" tag="24X"></TD>
							</TR>
							
								<!--TD CLASS="TD5" nowrap>������ݾ�</TD-->
								<!--TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=������ݾ� NAME="txtCashAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="21N2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>-->
								
													
							
							<TR>
								 <TD CLASS="TD5" nowrap>���ű׷�</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtGrpCd" ALT="���ű׷�" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=4 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrp()" >
													   <INPUT TYPE=TEXT NAME="txtGrpNm" ALT="���ű׷�" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<!--<TD CLASS="TD5" nowrap>���Ա�</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���Ա� NAME="txtLoanAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24N2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>-->
								<TD CLASS="TD5" nowrap>B/L��ȣ</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBlNo" ALT="B/L��ȣ" style="HEIGHT: 20px; WIDTH: 250px" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>���ڼ��ݰ�꼭����</TD>
								<TD CLASS="TD6" nowrap>
								     <INPUT TYPE=radio NAME="rdoIssueDTFg" ALT="���ڼ��ݰ�꼭����" CLASS="RADIO" id="rdoIssueDTFg1" tag="21X"><label for="rdoVatFlg"> YES </label>
									 <INPUT TYPE=radio NAME="rdoIssueDTFg" ALT="���ڼ��ݰ�꼭����" CLASS="RADIO" checked id="rdoIssueDTFg2"  tag="21X"><label for="rdoVatFlg"> NO </label>
								</TD>
								<TD CLASS="TD5" nowrap>&nbsp;</TD>
								<TD CLASS="TD6" nowrap>&nbsp;</TD>
							</TR>
														
							<%Call SubFillRemBodyTD5656(2)%>
							</TABLE>
							</DIV>
							<!-- �ι�° �� ���� -->
							<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>   
							    <TD CLASS="TD5" NOWRAP>����ó INVOICE NO.</TD>
								<TD CLASS="TD6" colspan=3 width=100% nowrap><INPUT TYPE=TEXT NAME="txtSpplIvNo" ALT="����ó INVOICE NO."  style="HEIGHT: 20px; WIDTH:250px" MAXLENGTH=50 tag="21"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>��ݰ�������</TD>
								<TD CLASS="TD6" colspan=3 width=100% nowrap><INPUT TYPE=TEXT NAME="txtPayTermsTxt" ALT="��ݰ�������"  style="HEIGHT: 20px; WIDTH: 624px" MAXLENGTH=120 tag="21N"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" nowrap>���</TD>
								<TD CLASS="TD6" colspan=3 width=100% nowrap><INPUT TYPE=TEXT  NAME="txtMemo" ALT="���" tag = "21"  style="HEIGHT: 20px; WIDTH: 624px" MAXLENGTH=70></TD>
							</TR>
                            <% Call SubFillRemBodyTD5656(12) %>
							
						   </TABLE>
					       </DIV>
					</TD>	
				</TR>
			</table>
		</TD>
	</TR>
    <tr>
      <td <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
      <td WIDTH="100%">
		<table <%=LR_SPACE_TYPE_30%>>
		  <tr>
			<TD WIDTH=10>&nbsp;</TD>
<!--			<td align="Left"><Div ID="btnintAcc"><a><button name="btnPostingSel" id="btnPosting" class="clsmbtn" ONCLICK="Posting()">Ȯ��</button></a><Div></td> -->
            <td> 
			   <BUTTON NAME="btnPosting" CLASS="CLSSBTN"  ONCLICK="Posting()">Ȯ��ó��</BUTTON>&nbsp;
			   <BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">��ǥ��ȸ</BUTTON>&nbsp;
			</td>
		    <td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">���Գ������</a>|<a href="VBSCRIPT:CookiePage(2)">���޳������</a></td>
		    <TD WIDTH=10>&nbsp;</TD>
		  </tr>
		</table>
      </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnApFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdvatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLocCur" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnUsrId" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtChkPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCur" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSpplCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdIssueDTFg" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
