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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                           
'**********************************************************************************************-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc 선언   ********************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->							
<!-- #Include file="../../inc/IncSvrHTML.inc" -->								
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--'==========================================  1.1.2 공통 Include   ======================================-->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                          '☜: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Const BIZ_PGM_ID 		= "m5111mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID 	= "m5112ma1_KO441"
Const BIZ_PGM_JUMP_ID2 	= "m5113ma1"

Const ivType = "ST"

Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgIntFlgMode					'☜: Variable is for Operation Status

Dim lgMpsFirmDate, lgLlcGivenDt											
Dim gSelframeFlg			                    'tab1,tab2 구분	
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

    Err.Clear                                                      '☜: Protect system from crashing
    
    Dim strVal, StrvalBpCd
	Select Case BpType
		Case "1"                                                   '공급처인경우 화폐 변동 
		    if Trim(frm1.txtSpplCd.Value) = "" then
    			Exit Sub
    		End if
    		
    		StrvalBpCd = FilterVar(Trim(frm1.txtSpplCd.value), "", "SNM")
    	    
    	    if Trim(frm1.txtIvDt.Text) = ""  then
	            Call DisplayMsgBox("970021","X","매입등록일","X")
	            Exit Sub
	        End if
    	   
    	    if Trim(frm1.txtSpplCd.value) = ""  then
	            Call DisplayMsgBox("970021","X","공급처","X")
	            Exit Sub
	        End if
    	    
    	    Call GetPayDt()                                        '지불예정일 setting
    	Case "2"                                                   '지급처인경우 결제기간,대금결제참조,지급유형변동 
    		if Trim(frm1.txtPayeeCd.Value) = "" then               '발주번호 no checked경우 결제방법변동 
    			Exit Sub
    		End if
			StrvalBpCd = FilterVar(Trim(frm1.txtPayeeCd.value), "", "SNM")
    	Case "3"                                                  '세금계산서발행처인 경우 VAT,VAT이름,사업자등록번호 
    		if Trim(frm1.txtBuildCd.Value) = "" then
    			Exit Sub
    		End if   	
    		StrvalBpCd = FilterVar(Trim(frm1.txtBuildCd.value), "", "SNM")
    		
	        Call GetTaxBizArea("BP")
	        '2003.1월 정기패치(S) : 선급금 팝업버튼 관련.(KJH : 03-01-06)
	        Call CheckPrePayedAmtYN()
	End Select
 
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"			'☜: 비지니스 처리 ASP의 상태 
    strVal = strval & "&txtBpType=" & BpType
    strVal = strVal & "&txtBpCd=" & StrvalBpCd		'☆: 조회 조건 데이타 
    
    if LayerShowHide(1) = False then
		Exit sub
	end if

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동	
	
End Sub

'--------------------------------------------------------------------
'		Field의 Tag속성을 Protect로 전환,복구 시키는 함수 
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
		Call ggoOper.SetReqAttr(frm1.txtIvDt,"N")                '매입등록일 
		Call ggoOper.SetReqAttr(frm1.txtPayDt,"N")               '지불예정일 
		Call ggoOper.SetReqAttr(frm1.txtPostDt,"D")              '매입일 
		'Call ggoOper.SetReqAttr(frm1.txtCur,"N")
		'Call ggoOper.SetReqAttr(frm1.txtXchRt,"D")

		Call ggoOper.SetReqAttr(frm1.txtPayeeCd,"N")             '지급처 
		Call ggoOper.SetReqAttr(frm1.txtBuildCd,"N")             '세금계산서발행처 
		Call ggoOper.SetReqAttr(frm1.txtSpplIvNo,"D")            '공급처 
		Call ggoOper.SetReqAttr(frm1.txtVatCd,"N")             'VAT
		Call ggoOper.SetReqAttr(frm1.txtGrpCd,"N")               '구매그룹 

		Call ggoOper.SetReqAttr(frm1.txtPayDur,"D")              '결제기간 
		Call ggoOper.SetReqAttr(frm1.txtPayTypeCd,"D")           '지급유형 
		Call ggoOper.SetReqAttr(frm1.txtPayTermstxt,"D")         '대금결제참조 
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCd,"D")           '세금신고사업장 
		Call ggoOper.SetReqAttr(frm1.txtMemo,"D")                '비고 
		
 		if (UNICDbl(Trim(frm1.txtIvAmt.Value)) <> 0 and Trim(frm1.txtIvAmt.Value) <> "") or Trim(frm1.txtPoNo.value) = "" then	'iv detail이 존재하면 PO NO를 제한하여 등록할 수 없음.
			ggoOper.SetReqAttr	frm1.ChkPoNo, "Q"   'N: REQUIRED, D: UNREQUIRED ,Q:PROTECTED
		else                                        '발주번호 지정 
			ggoOper.SetReqAttr	frm1.ChkPoNo, "D"
		End if
	
		if Trim(frm1.txtPoNo.value) <> "" then                  '발주참조를 했을경우 화폐를 protect
			ggoOper.SetReqAttr	frm1.txtCur, "Q"                '화폐 
			Call ggoOper.SetReqAttr(frm1.txtSpplCd, "Q" )       '공급처 
		else
			ggoOper.SetReqAttr	frm1.txtCur, "N"
			Call ggoOper.SetReqAttr(frm1.txtSpplCd,"N")         'txtSpplCd
		End if
		'결제방법을 항상 Required이도록 수정함.(2003.03.18)-Lee,Eun Hee
		Call ggoOper.SetReqAttr(frm1.txtPayMethCd,"N")      '결제방법 
		
	    if UCase(Trim(frm1.txtCur.value)) = UCase(parent.gCurrency) _
	      Or UCase(Trim(frm1.hdnRetflg.Value)) = "Y"   then
		   Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")            '환율 
		else       
		   Call ggoOper.SetReqAttr(frm1.txtXchRt,"D")
	    End if
	End if
	
End Function
'==========================================   Posting()  ======================================
'	Name : Posting()
'	Description : 확정버튼,확정취소버튼의 Event 합수 
'========================================================================================================= 
Sub Posting()
    Dim IntRetCD 
    
    Err.Clear                                                         '☜: Protect system from crashing
    
    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217","X","X","X")                      '데이타가 변경되었습니다. 저장후 진행하십시오.
		Exit sub
	End if
	
    if frm1.rdoApFlg(0).checked = True then                           '확정여부 
					
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")       '작업을 수행 하시겠습니까?
		If IntRetCD = vbNo Then
			frm1.btnPosting.disabled = False	'20040315  
			Exit Sub
		Else 
				frm1.btnPosting.disabled = True		'20040315 
		End If
		
		
			
		Call Changeflg()                                               'hidden값 setting 함수 
		Call DbSave("Posting")				             

	Elseif frm1.rdoApFlg(1).checked = True then
		
		if Trim(frm1.txtPostDt.text) = "" then
			Call DisplayMsgBox("17A002","X" , "매입일","X")        '%1을 입력하세요.
			Exit sub
		End if
		
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnPosting.disabled = False	'20040315  
			Exit Sub
		Else 
				frm1.btnPosting.disabled = True		'20040315 
		End If
		
		Call Changeflg()                                               'hidden값 setting 함수 
		Call DbSave("UnPosting")
		
	End if
	
End Sub
'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD
	If Kubun = 1 Then                                                  '매입내역등록 점프 

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then                                '데이타가 변경되었습니다. 계속 하시겠습니까?
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
	ElseIf Kubun = 2 Then                                               '지급내역등록 점프 

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

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
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
		.ChkPrepay.Checked =   false                 '선급금여부 지정 check box		
		.btnPosting.disabled = true             '확정유무버튼 
    	.btnGlSel.disabled = true               '전표조회버튼 
    	.hdnLocCur.value = parent.gCurrency
    	.txtGrpCd.Value = parent.gPurGrp
    	.hdnUsrId.value = parent.gUsrID
    	.btnPosting.value = "확정"
    	.txtIvNo.focus 
    	Set gActiveElement = document.activeElement
    	
    	frm1.chkPoNo.checked = False
    	
    	Call ClickTab1()   
    	
    	Call SetToolBar("1110100000001111")
    	interface_Account = GetSetupMod(parent.gSetupMod, "a")
		'**수정(2003.03.26)-회계모듈이 없어도 확정,취소 가능하도록 수정 
		'if interface_Account = "N" then
		'	'btnintAcc.style.display = "none"
		'	frm1.btnPosting.disabled = true
		'End if
		Call ggoOper.SetReqAttr(frm1.txtIvTypeCd,"N")        '매입형태 
	End With
End Sub

<!--
'===========================================  2.3.1 Tab Click 처리  =====================================
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
	    Call DisplayMsgBox("970021","X","매입등록일","X")
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
'	Description : 전표조회 
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
'	Description : 발주참조값 setting
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)
    Dim strVal
    
	Call ggoOper.ClearField(Document, "A")
    Call SetDefaultVal
	Call ggoOper.SetReqAttr(frm1.txtCur, "Q" )
	Call ggoOper.SetReqAttr(frm1.txtSpplCd, "Q" )
	'수정(2003.03.18)-Lee,Eun Hee
	'결제방법을 변경가능하도록 수정함.
	'Call ggoOper.SetReqAttr(frm1.txtPayMethCd, "Q" )

    Call InitVariables
    
	frm1.hdnPoNo.Value = strRet(0)
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpPo"							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPoNo=" & Trim(frm1.hdnPoNo.Value)				'☆: 조회 조건 데이타 
   	If Trim(frm1.txtPoNo.value) <> "" Then frm1.chkPoNo.checked = True
 
    if LayerShowHide(1) = false then
		exit sub
	end if
    
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
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
	arrHeader(0) = "화폐"						' Header명(0)
    arrHeader(1) = "화폐명"						' Header명(1)
    
    arrField(0) = "Currency"						' Field명(0)
    arrField(1) = "Currency_Desc"					' Field명(1)
    
	arrParam(0) = "화폐"						' 팝업 명칭 
	arrParam(1) = "B_Currency"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCur.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtCur.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "화폐"						' TextBox 명칭 
	
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
	
	arrHeader(2) = "사업자등록번호"									' Header명(2)
    arrField(0) = "B_BIZ_PARTNER.BP_Cd"									' Field명(0)
    arrField(1) = "B_BIZ_PARTNER.BP_Nm"								    ' Field명(1)
    arrField(2) = "B_BIZ_PARTNER.BP_RGST_NO"							' Field명(2)
    
	Select Case BpType
		Case "1"  '공급처 
			If lblnWinEvent = True Or UCase(frm1.txtSpplCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True
			arrHeader(0) = "공급처"											' Header명(0)
			arrHeader(1) = "공급처명"										' Header명(1)

		    arrParam(0) = "공급처"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER "					                    ' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtSpplCd.Value)		
			'arrParam(2) = Trim(frm1.txtSpplCd.Value)							' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "		' Where Condition
			arrParam(5) = "공급처"											' TextBox 명칭 
		Case "2"   '지급처 
			If lblnWinEvent = True Or UCase(frm1.txtPayeeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True

			arrHeader(0) = "지급처"											' Header명(0)
			arrHeader(1) = "지급처명"										' Header명(1)

			arrParam(0) = "지급처"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
			'arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MPA", "''", "S") & " "
			arrParam(5) = "지급처"											' TextBox 명칭 
		Case "3"   '세금계산서발행처 
			If lblnWinEvent = True Or UCase(frm1.txtBuildCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True

			arrHeader(0) = "세금계산서발행처"											' Header명(0)
			arrHeader(1) = "세금계산서발행처명" 										' Header명(1)

			arrParam(0) = "세금계산서발행처"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"           					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			'arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MBI", "''", "S") & " "
			arrParam(5) = "세금계산서발행처"											' TextBox 명칭 
	End Select
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) Then
		Select Case BpType
			Case "1"   '공급처 
				frm1.txtSpplCd.Value = arrRet(0) : frm1.txtSpplNm.Value = arrRet(1)
				frm1.txtSpplCd.focus
			Case "2"   '지급처 
				frm1.txtPayeeCd.Value = arrRet(0) : frm1.txtPayeeNm.Value = arrRet(1)
				frm1.txtPayeeCd.focus
			Case "3"   '세금계산서발행처 
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
	
	arrHeader(0) = "VAT형태"									' Header명(0)
    arrHeader(1) = "VAT형태명"									' Header명(1)
    arrHeader(2) = "VAT율"									    ' Header명(2)
    
    arrField(0) = "b_minor.MINOR_CD"					            ' Field명(0)
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"					    ' Field명(1)
    
	arrParam(0) = "VAT"	            							' 팝업 명칭 
	arrParam(1) = "B_MINOR,b_configuration"
	arrParam(2) = Trim(frm1.txtVatCd.Value)						    ' Code Condition
	'arrParam(2) = Trim(frm1.txtVatCd.Value)						    ' Code Condition
	'arrParam(3) = Trim(frm1.txtVatNm.Value)						' Name Cindition
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_configuration.seq_no=1 and b_minor.major_cd=b_configuration.major_cd"
	arrParam(5) = "VAT"										    ' TextBox 명칭 
	
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
	arrHeader(0) = "구매그룹"									' Header명(0)
    arrHeader(1) = "구매그룹명"									' Header명(1)
    
    arrField(0) = "PUR_GRP"											' Field명(0)
    arrField(1) = "PUR_GRP_NM"										' Field명(1)
    
	arrParam(0) = "구매그룹"									' 팝업 명칭 
	arrParam(1) = "B_PUR_GRP"										' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtGrpCd.Value)							' Code Condition
	'arrParam(2) = Trim(frm1.txtGrpCd.Value)							' Code Condition
																	' Where Condition
	arrParam(4) = "USAGE_FLG = " & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "구매그룹"									' TextBox 명칭 
	
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
		Call DisplayMsgBox("17A002","X" , "결제방법","X")
		Exit Function
	End if

	If lblnWinEvent = True Or UCase(frm1.txtPaytypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	lblnWinEvent = True
	arrHeader(0) = "지급유형"						' Header명(0)
    arrHeader(1) = "지급유형명"						' Header명(1)
    
    arrField(0) = "b_configuration.REFERENCE"			' Field명(0)
    arrField(1) = "B_Minor.Minor_Nm"					' Field명(1)
    
	arrParam(0) = "지급유형"						' 팝업 명칭 
	arrParam(1) = "B_Minor,b_configuration"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPaytypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtPaytypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtPayTypeNm.Value)		' Name Cindition
	arrParam(4) = "b_configuration.minor_cd= " & FilterVar(Trim(frm1.txtPayMethCd.Value), "''", "S") & _
				  " And b_configuration.Major_Cd= " & FilterVar("B9004", "''", "S") & " and " & _
				  "b_minor.minor_cd=*b_configuration.REFERENCE and b_configuration.SEQ_NO>" & FilterVar("1", "''", "S") & "  And " & _
				  "b_minor.Major_Cd=" & FilterVar("A1006", "''", "S") & " and substring(b_configuration.REFERENCE,2,1) <> " & FilterVar("R", "''", "S") & "  "
				  
	arrParam(5) = "지급유형"						' TextBox 명칭 
	
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
	arrHeader(0) = "결제방법"						        ' Header명(0)
    arrHeader(1) = "결제방법명"						        ' Header명(1)
    arrHeader(2) = "Reference"
    
    arrField(0) = "B_Minor.MINOR_CD"							' Field명(0)
    arrField(1) = "B_Minor.MINOR_NM"							' Field명(1)
    arrField(2) = "b_configuration.REFERENCE"
    
	arrParam(0) = "결제방법"						        ' 팝업 명칭 
	arrParam(1) = "B_Minor,b_configuration"				        ' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPayMethCd.Value)			        ' Code Condition
	'arrParam(2) = Trim(frm1.txtPayMethCd.Value)			        ' Code Condition
	'arrParam(3) = Trim(frm1.txtPayMethNM.Value)			    ' Name Cindition
	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & " and B_Minor.minor_cd =b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"	 
	arrParam(5) = "결제방법"						        ' TextBox 명칭 
	
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

	arrParam(0) = "세금신고사업장"	
	arrParam(1) = "B_TAX_BIZ_AREA"
	
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	'arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	
	'arrParam(4) = "Tax_Flag = 'Y'"
	arrParam(4) = ""
	arrParam(5) = "세금신고사업장"			
	
    arrField(0) = "TAX_BIZ_AREA_CD"
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "세금신고사업장"
    arrHeader(1) = "세금신고사업장명"
    
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
	
	arrHeader(0) = "매입형태"						' Header명(0)
    arrHeader(1) = "매입형태명"						' Header명(1)
    
    arrField(0) = "IV_TYPE_CD"							' Field명(0)
    arrField(1) = "IV_TYPE_NM"							' Field명(1)
    
	arrParam(0) = "매입형태"						' 팝업 명칭 
	arrParam(1) = "M_IV_TYPE"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			' Name Cindition
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and import_flg=" & FilterVar("N", "''", "S") & " "						' Where Condition
	arrParam(5) = "매입형태"						' TextBox 명칭 
	
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
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	elseif Trim(frm1.txtCur.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "화폐","X")
		frm1.txtCur.focus
		Set gActiveElement = document.activeElement
		Exit Function
	end if
	
	lblnWinEvent = True

	arrParam(0) = "차입금번호"	
	arrParam(1) = "F_LOAN"
	arrParam(2) = Trim(frm1.txtLoanNo.Value)
	'arrParam(2) = Trim(frm1.txtLoanNo.Value)
	
	arrParam(4) = "DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  And BP_CD =  " & FilterVar(frm1.txtSpplCd.Value, "''", "S") & " "
	'arrParam(4) = "DOC_CUR = '" & Trim(frm1.txtCur.Value) & "' And BP_CD = '" & Trim(frm1.txtSpplCd.Value) & "'"
	arrParam(5) = "차입금번호"			
	
    arrField(0) = "LOAN_NO"
    arrField(1) = "F2" & parent.gColSep & "LOAN_AMT"
    arrHeader(0) = "차입금번호"		
    arrHeader(1) = "차입금"	
    
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
		Call DisplayMsgBox("17A002","X" , "화폐","X")
		Exit Function
	elseif Trim(frm1.txtSpplCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "수출자","X")
		Exit Function
	end if
	
	lblnWinEvent = True

	arrParam(0) = "선급금번호"	
	arrParam(1) = "F_PRPAYM"
	
	arrParam(2) = ""
	'====================== 1월 정기패치(S) : 선급금 팝업버튼 관련.(KJH : 03-01-06)============================
	arrParam(4) = "DOC_CUR =  " & FilterVar(frm1.txtCur.Value, "''", "S") & "  And BP_CD =  " & FilterVar(frm1.txtPayeeCd.Value, "''", "S") & "  AND BAL_AMT > 0 AND CONF_FG = " & FilterVar("C", "''", "S") & " "
	'====================== 1월 정기패치(E) : 선급금 팝업버튼 관련.(KJH : 03-01-06)============================
	arrParam(5) = "선급금번호"			
	
    arrField(0) = "PRPAYM_NO"
    arrField(1) = "F2" & parent.gColSep & "PRPAYM_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "BAL_AMT"
    
    arrHeader(0) = "선급금번호"		
    arrHeader(1) = "선급금"		
    arrHeader(2) = "선급금화폐"
    arrHeader(3) = "선급금잔액"
        
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
	
    
	If strFlag = "NM" Then                              '세금신고사업장 변경시 이름값만 가져온다 
		strTaxBizArea = frm1.txtBizAreaCd.value
	Else
		strBilltoParty = frm1.txtBuildCd.value          '계산서 발행처 
		strSalesGrp    = frm1.txtGrpCd.value            '구매그룹 
		
		<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
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

		' 세금 신고 사업장을 Editing한 경우 
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
'	Description : 지불예정일을 가져온다.
'==================================================================================================== %>
Sub GetPayDt()
   	Dim strSelectList, strFromList, strWhereList
	Dim strSpplCd, strIvDt,temp
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp

    	strSpplCd  = frm1.txtSpplCd.value                       '공급처	
    	temp    = UNIConvDate(frm1.txtIvDt.text)            '매입등록일 
		strIvDt = mid(temp,1,4)
		strIvDt = strIvDt & mid(temp,6,2)
		strIvDt = strIvDt & mid(temp,9,2) 
		<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
    
	
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
	'2003.1월 정기패치(S) : 선급금 팝업버튼 관련.(KJH : 03-01-06)
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
	
		'매입금액 
		ggoOper.FormatFieldByObjectOfCur .txtIvAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'매입순금액 
		ggoOper.FormatFieldByObjectOfCur .txtnetDocAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'VAT금액 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    '매입자국금액 
	    ggoOper.FormatFieldByObjectOfCur .txtnetLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    '총매입자국금액 
	    ggoOper.FormatFieldByObjectOfCur .txtIvLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    'vat자국금액 
	    ggoOper.FormatFieldByObjectOfCur .txtVatLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	    '환율 
	    ggoOper.FormatFieldByObjectOfCur .txtXchRt, .txtCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	    
	    
	    
	    
	End With

End Sub
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

   
    Call LoadInfTB19029	    														'⊙: Load table , B_numeric_format
    
    Call AppendNumberRange("0","0","999")					'기간	
    Call AppendNumberPlace("7","2","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field    
    Call GetValue_ko441()
    Call SetDefaultVal   
    Call InitVariables
    Call cookiepage(0)
    Call changeTabs(TAB1)
    gSelframeFlg = TAB1
	gIsTab     = "Y" 
    gTabMaxCnt = 2                                   'tab 갯수 
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
    
    Call GetPayDt()                       '지불예정일  
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
'발주번호지정 클릭시 
'========================================  chkPoNo_onpropertychange() ====================================
Sub chkPoNo_onpropertychange()

	if frm1.ChkPoNo.checked = True and Trim(frm1.txtPoNo.Value) <> "" then     '발주참조를 했을경우 화폐를 protect
		Call ggoOper.SetReqAttr(frm1.txtCur, "Q" )
		Call ggoOper.SetReqAttr(frm1.txtSpplCd, "Q" )
		'수정(2003.03.18)-Lee,Eun Hee
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
'   Event Desc : 발행처 내용이 변경되었을때 관련 항목 LookUp
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

'====================== 2003.1월 정기패치(S) : 선급금 팝업버튼 관련.(KJH : 03-01-06)==========
'======================================   CheckPrePayedAmtYN()  =============================
'	Name : CheckPrePayedAmtYN()
'	Description : 선급금여부를 체크한다.
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

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                             '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
       
    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
       Exit Function
    End If
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables 
    
    If DbQuery = False Then Exit Function										'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    Set gActiveElement = document.activeElement    
End Function
'==============================  FncNew()  ================================================
Function FncNew() 

	Dim IntRetCD
    FncNew = False                                                          '⊙: Processing is NG
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ClickTab1()
    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call ChangeTag(False)
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
    FncNew = True																'⊙: Processing is OK
	Set gActiveElement = document.activeElement
End Function

'==============================  FncDelete()  ================================================
Function FncDelete() 
    
Dim IntRetCD

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function

    FncDelete = False														'⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function									'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'==============================  FncSave()  ================================================
Function FncSave() 
Dim IntRetCD 

    FncSave = False                                                         '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing
    
	if CheckRunningBizProcess = true then
		exit function
	end if

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
	
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       If gPageNo > 0 Then
	       gSelframeFlg = gPageNo
	   End If
       Exit Function
    End If
    
    '**_**_*_*_*_*_*_*_*__*_*_*_***__*_*_*_*_*
    ' 2002-08-21
    ' 세금신고사업장이 Null일 때 값 불러오기 
    '**_**_*_*_*_*_*_*_*__*_*_*_***__*_*_*_*_*
    If Trim(frm1.txtBizAreaCd.value) = "" then 
		Call GetTaxBizArea("BP")
	end if 

    if frm1.rdoApFlg(0).checked = true then
    	frm1.hdnApFlg.Value = "Y"
    else
    	frm1.hdnApFlg.Value = "N"
    End if
    'vat 포함여부 
    if frm1.rdoVatFlg1.checked = true then
    	frm1.hdvatFlg.Value = "1"
    else
    	frm1.hdvatFlg.Value = "2"
    End if

    '지불예정일 "" 일 경우
    If frm1.txtPayDt.Text = "" then
	Call DisplayMsgBox("17A002","X" , "지불예정일","X")
	Exit Function
    End IF

    ' 전자세금계산서여부 
    if frm1.rdoIssueDTFg1.checked = true then
    	frm1.hdIssueDTFg.Value = "Y"
    else
    	frm1.hdIssueDTFg.Value = "N"
    End If
  
    If DbSave("toolbar") = False Then Exit Function                         '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
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
    
    lgIntFlgMode = parent.OPMD_CMODE											'⊙: Indicates that current mode is Crate mode
    
     ' 조건부 필드를 삭제한다. 
    Call ggoOper.ClearField(Document, "1")                              '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
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
    On Error Resume Next                                                 '☜: Protect system from crashing
End Function
'==============================  FncInsertRow()  ================================================
Function FncInsertRow() 
     On Error Resume Next                                                 '☜: Protect system from crashing
End Function
'==============================  FncDeleteRow()  ================================================
Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function
'==============================  FncPrint()  ================================================
Function FncPrint() 
    Call parent.FncPrint()                                                '☜: Protect system from crashing
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
    Call parent.FncExport(parent.C_SINGLE)             '☜: Protect system from crashing
End Function
'==============================  FncFind()  ================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                  '☜:화면 유형, Tab 유무 
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
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtIvNo=" & FilterVar(Trim(frm1.txtIvNo.value), "", "SNM")				'☜: 삭제 조건 데이타 
    
    if LayerShowHide(1) = false then
		exit function
	end if

    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG

End Function
'==============================  DbDeleteOk()  ================================================
Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
'==============================  DbQuery()  ================================================
Function DbQuery() 
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbQuery = False                                                         '⊙: Processing is NG
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo.value)				'☆: 조회 조건 데이타 %>
    
    if LayerShowHide(1) = false then
		exit function
	end if
    
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG

End Function
'==============================  DbQueryOk()  ================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	

    Call InitVariables
    
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
 
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field


    if frm1.rdoApFlg(0).checked = true or Trim(UCase(frm1.hdnImportflg.Value)) = "Y"  then  '확정여부 
		Call ChangeTag(true)
 
		if Trim(UCase(frm1.hdnImportflg.Value)) = "Y"  then
			Call SetToolBar("11100000000111")
			frm1.btnPosting.disabled = true
			frm1.rdoVatFlg1.disabled = true
			frm1.rdoVatFlg2.disabled = true
		else
			Call SetToolBar("11100000001111")
			if UNICDbl(frm1.txtIvAmt.Value) = 0 or Trim(frm1.txtIvAmt.Value) = "" then       '매입내역이 없으면 
				
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
		if UNICDbl(frm1.txtIvAmt.Value) = 0 or Trim(frm1.txtIvAmt.Value) = "" then         '매입내역이 없으면 
				
				frm1.btnPosting.disabled = true                                            '매입금액이 없으면 확정불가 
				frm1.rdoVatFlg1.disabled = false
				frm1.rdoVatFlg2.disabled = false
		else
				
				frm1.btnPosting.disabled = false
				frm1.rdoVatFlg1.disabled = true
				frm1.rdoVatFlg2.disabled = true
		end if
	end if
  	
	if frm1.rdoApFlg(0).checked = true then                                                '확정이되야 전표조회가 가능 
		frm1.btnPosting.value = "확정취소"
		
		if interface_Account <> "N" then
			frm1.btnGlSel.disabled = false
		Else
			frm1.btnGlSel.disabled = true
		end if
		
		if frm1.hdnGlType.Value = "A" Then
		   frm1.btnGlSel.value = "회계전표조회"
		elseif frm1.hdnGlType.Value = "T" Then
		   frm1.btnGlSel.value = "결의전표조회"
		end if
	else
		frm1.btnPosting.value = "확정"
		frm1.btnGlSel.disabled = true
	end if
	  
  if frm1.ChkPrepay.checked = false then
    Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
  end if
  
'2009-09-02 김지한 과정 요청으로 화폐필드 수정
  'if UNICDbl(Trim(frm1.txtnetLocAmt.Value)) <> 0  then
  '	Call ggoOper.SetReqAttr(frm1.txtCur,"Q") 
  '	Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")
  'End if 
  Call ClickTab1()
End Function
'==============================  DbSave()  ================================================
Function DbSave(byval btnflg) 

    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG

    Dim strVal

	With frm1
		.hdnUsrId.value = parent.gUsrID
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		
		if btnflg = "Posting" then
			.txtMode.value = "Release" 				                        '☜: Y=확정 버튼 
		elseif btnflg = "UnPosting" then
			.txtMode.value = "UnRelease" 				                    '☜: Y=확정취소 버튼 
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
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function
'==============================  DbSaveOk()  ================================================
Function DbSaveOk()														'☆: 저장 성공후 실행 로직 
	lgBlnFlgChgValue = False
	Call MainQuery()
End Function

'============================================================================================================
' Name : SubGetGlNo
' Desc : Get Gl_no : 2003.03 KJH 전표번호 가져오는 로직 수정 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입세금계산서</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입세금계산서기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPORef()">발주참조</A></TD>
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
									<TD CLASS="TD5" nowrap>매입번호</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvNo" style="HEIGHT: 20px; WIDTH: 250px" MAXLENGTH=18 ALT="매입번호" STYLE="TEXT-ALIGN:left; TEXT-TRANSFORM:UPPERCASE" tag="12N"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvNo" style="HEIGHT: 21px; WIDTH: 16px" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvNo()"></TD>
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
					<!-- 첫번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">	
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" nowrap>매입번호</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvNo1" ALT="매입번호" style="HEIGHT: 20px; WIDTH: 250px" MAXLENGTH=18 STYLE="TEXT-ALIGN:left; TEXT-TRANSFORM:UPPERCASE" tag="25X"></TD>
								<TD CLASS=TD5 NOWRAP>발주번호</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtPoNo" ALT="발주번호" TYPE=TEXT MAXLENGTH=35 SIZE=25 TAG="24XXXU">
								    <INPUT TYPE=CHECKBOX NAME="chkPoNo" tag="25" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkPoNo">발주번호지정</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>매입형태</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvTypeCd" ALT="매입형태" MAXLENGTH=5 style="HEIGHT: 20px; WIDTH: 80px" tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px"  align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
													   <INPUT TYPE=TEXT NAME="txtIvTypeNm" ALT="매입형태" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<TD CLASS="TD5" nowrap>매입등록일</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=매입등록일 NAME="txtIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22N1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>확정여부</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=radio NAME="rdoApFlg" ALT="확정여부" CLASS="RADIO" tag="24X"><label for="rdoApFlg"> Yes </label>
													   <INPUT TYPE=radio NAME="rdoApFlg" ALT="확정여부" CLASS="RADIO" checked tag="24X"><label for="rdoApFlg">  No&nbsp;&nbsp;</label>
													   <INPUT TYPE=TEXT NAME="txtGlNo" ALT="전표번호" style="HEIGHT: 20px; WIDTH: 148px" tag="24X"></TD>
							
								<TD CLASS="TD5" nowrap>지불예정일</TD>
								<TD CLASS="TD6" nowrap>
								    <Table Cellspacing=0 Cellpadding=0>
								        <TR>
										    <TD NOWRAP>
								            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=지불예정일 NAME="txtPayDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								            <TD NOWRAP>&nbsp;매입일</TD>
								            <TD NOWRAP>
								            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=매입일 NAME="txtPostDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							            </TR>
							        </Table>
							    </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>공급처</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtSpplCd" ALT="공급처" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(1)"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnSupplier" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(1)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT NAME="txtSpplNm" ALT="공급처" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<TD CLASS="TD5" nowrap>사업자등록번호</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtSpplRegNo" ALT="사업자등록번호" MAXLENGTH=10 style="HEIGHT: 20px; WIDTH: 250px"  tag="24X"></TD>
                            </TR>
							
							<TR>
	                            <TD CLASS="TD5" nowrap>세금계산서발행처</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBuildCd" ALT="세금계산서발행처" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(3)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSupplier" style="HEIGHT: 21px; WIDTH: 16px"  align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(3)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT NAME="txtBuildNm" ALT="세금계산서발행처" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<TD CLASS="TD5" nowrap>지급처</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayeeCd" ALT="지급처" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(2)"><IMG SRC="../../../CShared/image/btnPopup.gif"  style="HEIGHT: 21px; WIDTH: 16px" NAME="btnSupplier" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(2)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT NAME="txtPayeeNm" ALT="지급처" tag="24X"></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>화폐</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtCur" ALT="화폐" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=3 tag="22NXXU" onChange="ChangeCurr()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCur() ">&nbsp; 
										   <INPUT TYPE=HIDDEN NAME="txtCurNm" ALT="화폐" style="HEIGHT: 20px; WIDTH: 46px" tag="24X">
											
								</TD>
								<TD CLASS="TD5" nowrap>환율</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환율 NAME="txtXchRt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>총매입금액</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=총매입금액 NAME="txtIvAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" nowrap>총매입자국금액</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=총매입자국금액 NAME="txtIvLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>매입금액</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=매입금액 NAME="txtnetDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" nowrap>매입자국금액</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=매입자국금액 NAME="txtnetLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
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
								<TD CLASS="TD5" nowrap>VAT포함여부</TD>
								<TD CLASS="TD6" nowrap>
								     <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" checked id="rdoVatFlg1" tag="21X"><label for="rdoVatFlg"> 별도 </label>
									 <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" id="rdoVatFlg2"  tag="21X"><label for="rdoVatFlg">  포함&nbsp;</label></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" nowrap>VAT율</TD>
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
								<TD CLASS="TD5" nowrap>VAT금액</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT금액 NAME="txtVatAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
							    <TD CLASS="TD5" nowrap>VAT자국금액</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT자국금액 NAME="txtVatLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>결제방법</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayMethCd" ALT="결제방법" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU" OnChange="VBScript:changePayMeth()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayMeth()">
													   <INPUT TYPE=TEXT NAME="txtPayMethNm" ALT="결제방법" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								<TD CLASS="TD5" nowrap>결제기간</TD>
								<TD CLASS="TD6" NOWRAP>
									<Table cellpadding=0 cellspacing=0>
										<TR>
											<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=결제기간 NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												&nbsp;일
											</TD>
										</TR>
									</Table>
								</TD>
							</TR>
							
							
							
							<TR>
								<TD CLASS="TD5" nowrap>지급유형</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTypeCd" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayType()">
													   <INPUT TYPE=TEXT NAME="txtPayTypeNm" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=CHECKBOX CHECKED ID="ChkPrepay" tag="24" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">&nbsp;선급금여부
											           <INPUT TYPE=TEXT NAME="ChkPrepay1"  style="HEIGHT: 19px; WIDTH: 1px" MAXLENGTH=0 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" id="PrepayNo" NAME="PrepayNo" align=top TYPE="BUTTON" onclick="vbscript:OpenPpNo()"></TD>
											           
		                                               
							</TR>

							<TR>
								<TD CLASS="TD5" nowrap>세금신고사업장</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBizAreaCd" ALT="세금신고사업장" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=10 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
													   <INPUT TYPE=TEXT NAME="txtBizAreaNm" ALT="세금신고사업장" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								<!--<TD CLASS="TD5" nowrap>차입금번호</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtLoanNo" ALT="차입금번호" MAXLENGTH=18 style="HEIGHT: 20px; WIDTH: 250px" tag="24N"><!--<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenLoanNo() "></TD>-->
								<TD CLASS="TD5" nowrap>B/L관리번호</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBlDocNo" ALT="B/L관리번호"  style="HEIGHT: 20px; WIDTH: 250px" tag="24X"></TD>
							</TR>
							
								<!--TD CLASS="TD5" nowrap>현금출금액</TD-->
								<!--TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=현금출금액 NAME="txtCashAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="21N2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>-->
								
													
							
							<TR>
								 <TD CLASS="TD5" nowrap>구매그룹</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtGrpCd" ALT="구매그룹" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=4 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrp()" >
													   <INPUT TYPE=TEXT NAME="txtGrpNm" ALT="구매그룹" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
								<!--<TD CLASS="TD5" nowrap>차입금</TD>
								<TD CLASS="TD6" nowrap><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=차입금 NAME="txtLoanAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24N2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>-->
								<TD CLASS="TD5" nowrap>B/L번호</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBlNo" ALT="B/L번호" style="HEIGHT: 20px; WIDTH: 250px" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>전자세금계산서여부</TD>
								<TD CLASS="TD6" nowrap>
								     <INPUT TYPE=radio NAME="rdoIssueDTFg" ALT="전자세금계산서여부" CLASS="RADIO" id="rdoIssueDTFg1" tag="21X"><label for="rdoVatFlg"> YES </label>
									 <INPUT TYPE=radio NAME="rdoIssueDTFg" ALT="전자세금계산서여부" CLASS="RADIO" checked id="rdoIssueDTFg2"  tag="21X"><label for="rdoVatFlg"> NO </label>
								</TD>
								<TD CLASS="TD5" nowrap>&nbsp;</TD>
								<TD CLASS="TD6" nowrap>&nbsp;</TD>
							</TR>
														
							<%Call SubFillRemBodyTD5656(2)%>
							</TABLE>
							</DIV>
							<!-- 두번째 탭 내용 -->
							<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>   
							    <TD CLASS="TD5" NOWRAP>공급처 INVOICE NO.</TD>
								<TD CLASS="TD6" colspan=3 width=100% nowrap><INPUT TYPE=TEXT NAME="txtSpplIvNo" ALT="공급처 INVOICE NO."  style="HEIGHT: 20px; WIDTH:250px" MAXLENGTH=50 tag="21"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" nowrap>대금결제참조</TD>
								<TD CLASS="TD6" colspan=3 width=100% nowrap><INPUT TYPE=TEXT NAME="txtPayTermsTxt" ALT="대금결제참조"  style="HEIGHT: 20px; WIDTH: 624px" MAXLENGTH=120 tag="21N"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" nowrap>비고</TD>
								<TD CLASS="TD6" colspan=3 width=100% nowrap><INPUT TYPE=TEXT  NAME="txtMemo" ALT="비고" tag = "21"  style="HEIGHT: 20px; WIDTH: 624px" MAXLENGTH=70></TD>
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
<!--			<td align="Left"><Div ID="btnintAcc"><a><button name="btnPostingSel" id="btnPosting" class="clsmbtn" ONCLICK="Posting()">확정</button></a><Div></td> -->
            <td> 
			   <BUTTON NAME="btnPosting" CLASS="CLSSBTN"  ONCLICK="Posting()">확정처리</BUTTON>&nbsp;
			   <BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
			</td>
		    <td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">매입내역등록</a>|<a href="VBSCRIPT:CookiePage(2)">지급내역등록</a></td>
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
