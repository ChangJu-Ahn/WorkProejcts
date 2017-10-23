
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5211ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입B/L등록 ASP															*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2003/05/27																			*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must cccccchange"								*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	
<!--
'============================================  1.1.2 공통 Include  ======================================
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
	Const BIZ_PGM_POQRY_ID = "m5211mb4.asp"			'발주참조 
	Const BIZ_PGM_POSTQRY_ID = "m5211mb5.asp"		'확정버튼 클릭시 
	Const BIZ_PGM_LCQRY_ID = "m5211mb6.asp"		    'L/C 참조 
	Const IMBL_DETAIL_ENTRY_ID = "m5212ma1"         'B/C 내역등록 점프 
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"		    '경비 등록 
	Const IV_Payment_ID = "m5113ma1"                '지급내역등록 

	Const TAB1 = 1
	Const TAB2 = 2
    '값 setting 시 구분값으로 사용 
	Const gstrTransportMajor = "B9009"				'운송방법 
	Const gstrIncotermsMajor = "B9006"				'가격조건	
	Const gstrPayTypeMajor = "A1006"				'지급유형 
	Const gstrPayMethodMajor = "B9004"				'결제방법	
	Const gstrVatTypeMajor = "B9001"				'VAT유형 
	Const gstrFreightMajor = "S9007"				'운임지불방법 
	Const gstrPackingTypeMajor = "B9007"			'포장형태 
	Const gstrOriginMajor      = "B9094"			'원산지 
	Const gstrDeliveryPlceMajor = "B9095"			'인도장소 
    Const gstrLoadingPortMajor = "B9092"			'선적항 
	Const gstrDischgePortMajor = "B9092"			'도착항 
	
	Const CID_POST  = 5211					'확정 
	
	Dim  lgBlnFlgChgValue			
	Dim  lgIntGrpCount				
	Dim  lgIntFlgMode				

	Dim gSelframeFlg			                    'tab1,tab2 구분	
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
	    frm1.txtBLIssueDt.text = serverDate            'B/L접수일 
	end if
		
	frm1.txtLoadingDt.text = serverDate            '선적일 
	frm1.hdnEnddate.value  = serverDate
	frm1.rdoPostingFlg(1).Checked = True           '확정여부 
	frm1.btnPosting.disabled = true                '확정 버튼 
	frm1.btnPosting.value = "확정"
	frm1.btnGlSel.value = "전표조회"
	frm1.btnGlSel.disabled = true
	frm1.txtDischgeDt.text = serverDate            '도착일 
	frm1.chkPoNoCnt.Checked = True                 '발주번호 지정 check box
	frm1.chkLcNoCnt.Checked = True                 'L/C번호 지정 check box
	frm1.ChkPrepay.Checked =   false                 '선급금여부 지정 check box
	Call ggoOper.SetReqAttr(frm1.ChkPrepay1,"Q")
		
	Call ClickTab1()                               '수입 B/L 정보 
	frm1.txtBLNo.focus                             '조회부 B/L 관리번호 
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
		Call DisplayMsgBox("200005", "X", "X", "X")  '%1 수정모드에서는 참조할 수 없습니다.
		Exit function
	End If	
			
	If frm1.rdoPostingflg(0).Checked = True then    '회계처리상태이므로 참조 할수 없습니다(확정 Y이면 참조불가 
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

	arrHeader(0) = "매입형태"					
    arrHeader(1) = "매입형태명"					
    
    arrField(0) = "IV_TYPE_CD"						
    arrField(1) = "IV_TYPE_NM"						
    
	arrParam(0) = "매입형태"						
	arrParam(1) = "M_IV_TYPE"						
	arrParam(2) = Trim(frm1.txtIvType.Value)		
	arrParam(4) = "import_flg=" & FilterVar("Y", "''", "S") & " "							
	arrParam(5) = "매입형태"						

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

	arrHeader(0) = "지급처"
    arrHeader(1) = "지급처명"
    
    arrField(0) = "bpftn.partner_bp_cd"
    arrField(1) = "bp.bp_nm"
    
	arrParam(0) = "지급처"
	arrParam(1) = "b_biz_partner_ftn bpftn,b_biz_partner bp"
	arrParam(2) = Trim(frm1.txtPayeeCd.Value) 	 '//Trim(frm1.txtPayeeCd.Value)
	
	arrParam(4) = "bpftn.partner_bp_cd=bp.bp_cd And bpftn.partner_ftn=" & FilterVar("MPA", "''", "S") & " and bpftn.bp_cd= " & FilterVar(Trim(frm1.txtBeneficiary.Value), " " , "S") & " " 	'//Trim(frm1.txtBeneficiary.value) 
	arrParam(5) = "지급처"

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

	arrHeader(0) = "세금계산서발행처"
    arrHeader(1) = "세금계산서발행처명"
    
    arrField(0) = "bpftn.partner_bp_cd"
    arrField(1) = "bp.bp_nm"
    
	arrParam(0) = "세금계산서발행처"
	arrParam(1) = "b_biz_partner_ftn bpftn,b_biz_partner bp"
	arrParam(2) = Trim(frm1.txtBuildCd.Value)
	arrParam(4) = "bpftn.partner_bp_cd=bp.bp_cd And bpftn.partner_ftn=" & FilterVar("MBI", "''", "S") & " and bpftn.bp_cd= " & FilterVar(Trim(frm1.txtBeneficiary.Value), " " , "S") & " "
	arrParam(5) = "세금계산서발행처"

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
		Call DisplayMsgBox("17A002","X" , "화폐","X")
		Exit Function
	elseif Trim(frm1.txtBeneficiary.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "수출자","X")
		Exit Function
	end if
	
	gblnWinEvent = True

	arrParam(0) = "선급금번호"	
	arrParam(1) = "F_PRPAYM"
	
	arrParam(2) = ""
	
	arrParam(4) = "DOC_CUR =  " & FilterVar(Trim(frm1.txtCurrency.Value), " " , "S") & "  And BP_CD =  " & FilterVar(Trim(frm1.txtBeneficiary.Value), " " , "S") & "  AND BAL_AMT > 0"
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

	arrParam(0) = "세금신고사업장"	
	arrParam(1) = "B_Tax_Biz_Area"
	arrParam(2) = Trim(frm1.txtTaxBizArea.Value)
	arrParam(4) = ""
	arrParam(5) = "세금신고사업장"			
	
    arrField(0) = "Tax_Biz_Area_Cd"
    arrField(1) = "Tax_Biz_Area_Nm"
    
    arrHeader(0) = "세금신고사업장"
    arrHeader(1) = "세금신고사업장명"
    
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
		Call DisplayMsgBox("17A002","X" , "화폐","X")
		Exit Function
	end if
	
	gblnWinEvent = True

	arrParam(0) = "차입금번호"	
	arrParam(1) = "F_LOAN"
	
	arrParam(2) = Trim(frm1.txtLoanNo.Value)
	
	'arrParam(4) = "DOC_CUR = '" & Trim(frm1.txtCurrency.Value) & "'"
	arrParam(4) = "DOC_CUR =  " & FilterVar(Trim(frm1.txtCurrency.Value), " " , "S") & "  AND LOAN_BAL_AMT > 0"
	arrParam(5) = "차입금번호"			
	
    arrField(0) = "LOAN_NO"
    arrField(1) = "F2" & parent.gColSep & "LOAN_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "LOAN_BAL_AMT"
    
    arrHeader(0) = "차입금번호"		
    arrHeader(1) = "차입금"		
    arrHeader(2) = "차입금화폐"
    arrHeader(3) = "차입금잔액"
    
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

	arrParam(0) = "수입담당"						
	arrParam(1) = "B_PURCHASE_GROUP"				
	arrParam(2) = Trim(frm1.txtPurGrp.value)		
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "수입담당"						

	arrField(0) = "PUR_GRP"							
	arrField(1) = "PUR_GRP_NM"						

	arrHeader(0) = "수입담당"						
	arrHeader(1) = "수입담당명"						

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
		Call DisplayMsgBox("17a002", parent.VB_YES_NO,"결제방법", "X")
		Exit Function
	End if

	If gblnWinEvent = True Or UCase(frm1.txtPayType.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "지급유형"
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

	arrParam(5) = "지급유형"
    
	arrField(0) = "B_MINOR.MINOR_CD"						
	arrField(1) = "left(B_MINOR.MINOR_NM,100)"	

    arrHeader(0) = "지급유형"
    arrHeader(1) = "지급유형명"
    
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
'+  (코드값(aa.value),이름값(bb.value),"운송방법(display컬럼명)","운송방법명(display컬럼명)",code값)    +
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

	arrParam(0) = "국가"							
	arrParam(1) = "B_COUNTRY"						
	arrParam(2) = Trim(strCntryCD)					
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "국가"							

	arrField(0) = "COUNTRY_CD"						
	arrField(1) = "COUNTRY_NM"						

	arrHeader(0) = "국가"						
	arrHeader(1) = "국가명"						

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
		arrParam(0) = "중량단위"				
		arrParam(1) = "B_UNIT_OF_MEASURE"		
		arrParam(2) = Trim(strUnitCD)			
		arrParam(3) = ""						
		arrParam(4) = "DIMENSION= " & FilterVar(Trim(strPopPos), " " , "S") & " "	
		arrParam(5) = "중량단위"						
	
		arrField(0) = "UNIT"							
		arrField(1) = "UNIT_NM"							
	
		arrHeader(0) = "중량단위"					
		arrHeader(1) = "중량단위명"					
	else
		arrParam(0) = "용적단위"						
		arrParam(1) = "B_UNIT_OF_MEASURE"				
		arrParam(2) = Trim(strUnitCD)					
		arrParam(3) = ""								
		arrParam(4) = "DIMENSION= " & FilterVar(Trim(strPopPos), " " , "S") & " "	
		arrParam(5) = "용적단위"						
	
		arrField(0) = "UNIT"							
		arrField(1) = "UNIT_NM"							
	
		arrHeader(0) = "용적단위"					
		arrHeader(1) = "용적단위명"					
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

	arrParam(0) = "도착항"				
	arrParam(1) = "B_Minor"					
	arrParam(2) = Trim(frm1.txtDischgePort.Value)			
'		arrParam(3) = Trim(frm1.txtDischgePortNm.Value)			
	arrParam(4) = "MAJOR_CD= " & FilterVar(Trim(gstrDisChgePortMajor), " " , "S") & " "	
	arrParam(5) = "도착항"				

	arrField(0) = "Minor_CD"				
	arrField(1) = "Minor_NM"				

	arrHeader(0) = "도착항"				
	arrHeader(1) = "도착항명"			

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

	arrParam(0) = "신고세무서"					
	arrParam(1) = "B_TAX_OFFICE"					
	arrParam(2) = Trim(frm1.txtTaxOfficeCd.value)	
'		arrParam(3) = Trim(frm1.txtTaxOfficeCdNm.value)	
	arrParam(4) = ""								
	arrParam(5) = "신고세무서"					

	arrField(0) = "TAX_OFFICE_CD"					
	arrField(1) = "TAX_OFFICE_NM"					

	arrHeader(0) = "신고세무서"					
	arrHeader(1) = "신고세무서명"				

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
	
	arrParam(0) = Trim(frm1.txtGlNo.value)          '전표번호 
	arrParam(1) = Trim(frm1.txtIvNo.value)          '매입번호 
	'arrParam(2) = Trim(frm1.txtGrpCd.value)
	'arrParam(3) = Trim(frm1.txtGrpNm.value)
	
   If frm1.hdnGlType.Value = "A" Then               '회계전표팝업 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
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

	arrParam(0) = "VAT형태"						
	arrParam(1) = "B_MINOR,b_configuration"			
	
	'arrParam(2) = Trim(frm1.txtVattype.Value)	<%' Code Condition%>
		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_configuration.SEQ_NO=1"							<%' Where Condition%>
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd"
	arrParam(5) = "VAT형태"						
	
    arrField(0) = "b_minor.MINOR_CD"				
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"		
    
    arrHeader(0) = "VAT형태"						
    arrHeader(1) = "VAT형태명"					
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
		Case "대행자"
			frm1.txtAgent.Value = arrRet(0)
			frm1.txtAgentNm.Value = arrRet(1)
			frm1.txtAgent.focus
				
		Case "제조자"
			frm1.txtManufacturer.Value = arrRet(0)
			frm1.txtManufacturerNm.Value = arrRet(1)
			frm1.txtManufacturer.focus
				
		Case "선박회사"
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
'	Description : 지불예정일을 가져온다.
'==================================================================================================== %>
Sub GetPayDt()
   	Dim strSelectList, strFromList, strWhereList
	Dim strPayeeCd, strBlIssueDt,temp
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp

    	strPayeeCd  = frm1.txtPayeeCd.value                   '지급처	
    	temp    = UNIConvDate(frm1.txtBlIssueDt.text)         'B/L접수일 
		strBlIssueDt = mid(temp,1,4)
		strBlIssueDt = strBlIssueDt & mid(temp,6,2)
		strBlIssueDt = strBlIssueDt & mid(temp,9,2) 
		<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
    
	
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
'	Description : 세금신고사업장을 가져온다 
'==================================================================================================== %>

Sub GetTaxBizArea(Byval strFlag)
   	Dim strSelectList, strFromList, strWhereList
	Dim strBilltoParty, strSalesGrp, strTaxBizArea
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp
	
     
	If strFlag = "NM" Then                              '세금신고사업장 변경시 이름값만 가져온다 
		strTaxBizArea = frm1.txtTaxBizArea.value
	Else
		strBilltoParty = frm1.txtBuildCd.value          '계산서 발행처 
		strSalesGrp    = frm1.txtPurGrp.value            '구매그룹 
		<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
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

		' 세금 신고 사업장을 Editing한 경우 
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
'	Description : 선급금여부를 체크한다.
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
'   Event Desc : B/L접수일변동시 환율값변동 
'==========================================================================================
-->
Function ChangeCurOrDt()
   
    Dim strVal
    Err.Clear                                                               '☜: Protect system from crashing

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
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
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
'		Field의 Tag속성을 Protect로 전환,복구 시키는 함수 
'--------------------------------------------------------------------
-->

 Function ChangeTag(Byval Changeflg)

	if Changeflg = true then  '확정이면 전체 lock
		''''첫번째 탭 
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
		''''두번째 탭 
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
		''''첫번째 탭 
		Call ggoOper.SetReqAttr(frm1.txtBLDocNo,"N")         'B/L번호 
		Call ggoOper.SetReqAttr(frm1.chkPoNoCnt,"D")         '발주번호지정 
		Call ggoOper.SetReqAttr(frm1.txtLoadingDt,"N")       '선적일 
		Call ggoOper.SetReqAttr(frm1.txtBlIssueDt,"N")       'B/L접수일 
		Call ggoOper.SetReqAttr(frm1.chkLcNoCnt,"D")         'L/C번호지정 
		Call ggoOper.SetReqAttr(frm1.txtDischgeDt,"D")       '도착일 
		Call ggoOper.SetReqAttr(frm1.txtSetlmnt,"D")         '지불예정일 
		Call ggoOper.SetReqAttr(frm1.txtTransport,"D")       '운송방법 
		Call ggoOper.SetReqAttr(frm1.txtForwarder,"D")       '선박회사 
		Call ggoOper.SetReqAttr(frm1.txtVesselNm,"D")        'Vessel명 
		Call ggoOper.SetReqAttr(frm1.txtVoyageNo,"D")        '항차번호 
		Call ggoOper.SetReqAttr(frm1.txtVesselCntry,"D")     '선박국적 
		Call ggoOper.SetReqAttr(frm1.txtTotPackingCnt,"D")   '총포장갯수 
		Call ggoOper.SetReqAttr(frm1.txtDischgePort,"D")     '도착항 
		
		if parent.gCurrency <> Trim(frm1.txtCurrency.value) then
		    Call ggoOper.SetReqAttr(frm1.txtXchRate,"N")         '환율 
		else
		   Call ggoOper.SetReqAttr(frm1.txtXchRate,"Q")         '환율 
		end if
		Call ggoOper.SetReqAttr(frm1.txtPayType,"D")         '지급유형 
		Call ggoOper.SetReqAttr(frm1.txtPayeeCd,"N")         '지급처 
		Call ggoOper.SetReqAttr(frm1.txtBuildCd,"N")         '세금계산서발행처 
		Call ggoOper.SetReqAttr(frm1.txtPayTermsTxt,"D")     '대금결제참조 
		'Call ggoOper.SetReqAttr(frm1.txtVatType,"N")
		'Call ggoOper.SetReqAttr(frm1.txtVatRate,"N")
		'Call ggoOper.SetReqAttr(frm1.txtPrePayNo,"D")
		'if Trim(frm1.txtPrePayNo.value) <> "" then
			'Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"N")
		'else
			'Call ggoOper.SetReqAttr(frm1.txtPrePayDocAmt,"Q")
		'End if
		'Call ggoOper.SetReqAttr(frm1.txtLoanNo,"D")          '차입금번호 
		'if interface_Account = "N" then
		'	Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"D")     '차입금 
		'else
		'	Call ggoOper.SetReqAttr(frm1.txtLoanAmt,"Q")
		'End if

		'Call ggoOper.SetReqAttr(frm1.txtCashAmt,"D")        
		Call ggoOper.SetReqAttr(frm1.txtIvType,"N")          '매입형태 
		''''두번째 탭 
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
	Call AppendNumberRange("1","0","999")			'tag 5번째가 1인경우 최대값 (컨데이너수)
	Call AppendNumberRange("2","0","99")			'tag 5번째가 2인경우 최대값 (B/L발행부수)
	Call AppendNumberPlace("7","2","0")
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart) '조회 영역 포멧 
	Call ggoOper.LockField(Document, "N")			
	
    frm1.hdnDefaultFlg.value = ""
	Call SetDefaultVal
	Call InitVariables
	Call OpenCookie()		
	If UCase(Trim(frm1.txtBLNo.value)) <> "" Then   '쿠킷값이 있으면 바로 쿼리한다 
		frm1.hdnchangeflg.value = "Y"
		Call MainQuery
	End If
        
    frm1.hdnchangeflg.value = "Y"
	gSelframeFlg = TAB1
	gIsTab     = "Y" 
    gTabMaxCnt = 2                                   'tab 갯수 
		
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
	    Call DisplayMsgBox("970021","X","B/L접수일","X")
	    Exit Sub
	End if	
	
	if Trim(frm1.txtPayeeCd.value) = ""  then
	    Call DisplayMsgBox("970021","X","지급처","X")
	    Exit Sub
    End if
	
	Call GetPayDt()
	
	lgBlnFlgChgValue = true	
End Sub

Sub ChangePayeeCd()
	
	Call CheckPrePayedAmtYN() '지급처가 바뀌면 선급금여부를 다시 체크하기 위해 

    if Trim(frm1.txtBlIssueDt.Text) = ""  then
	    Call DisplayMsgBox("970021","X","B/L접수일","X")
	    Exit Sub
	End if	
	
	if Trim(frm1.txtPayeeCd.value) = ""  then
	    Call DisplayMsgBox("970021","X","지급처","X")
	    Exit Sub
    End if
	
	Call GetPayDt()
	
	lgBlnFlgChgValue = true	
End Sub

Sub ChangeCurrency()
	Call CheckPrePayedAmtYN() '화폐가 바뀌면 선급금여부를 다시 체크하기 위해 
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
'=	Event Desc : check box 수정후 수정여부 true 															=
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
'=	Event Desc : 조회 관리번호 클릭시 호출															=
'========================================================================================================
-->
Sub btnBLNoOnClick()
	Call OpenBlNoPop()
End Sub

<!--
'====================================  3.2.2 btnPurGrpOnClick()  ===================================
'=	Event Name : btnPurGrpOnClick																	=
'=	Event Desc : 구매그룹(팝업이 없어졌음)															=
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
'=	Event Desc : 선급금액(없어졌음)																			=
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
'=	Event Desc :차입금번호 호출(tab1)   																=
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
		Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "운송방법", "운송방법명", gstrTransportMajor)
	End If
End Sub

<!--
'======================================  3.2.4 btnFreightOnClick()  ====================================
-->
Sub btnFreightOnClick()
	If frm1.txtFreight.readOnly <> True Then
		Call OpenMinorCd(frm1.txtFreight.value, frm1.txtFreightNm.value, "운임지불방법", "운임지불방법명", gstrFreightMajor)
	End If
End Sub

<!--
'====================================  3.2.5 btnPackingTypeOnClick()  ==================================
-->
Sub btnPackingTypeOnClick()
	If frm1.txtPackingType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtPackingType.value, frm1.txtPackingTypeNm.value, "포장형태", "포장형태명", gstrPackingTypeMajor)
	End If
End Sub

<!--
'======================================  3.2.4 btnIncotermsOnClick()  ====================================
-->
Sub btnIncotermsOnClick()
	If frm1.txtIncoterms.readOnly <> True Then
		Call OpenMinorCd(frm1.txtIncoterms.value, frm1.txtIncotermsNm.value, "가격조건", "가격조건명", gstrIncotermsMajor)
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
		Call OpenMinorCd(frm1.txtPayMethod.value, frm1.txtPayMethodNm.value, "결제방법", "결제방법명", gstrPayMethodMajor)
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
		Call OpenMinorCd(frm1.txtDeliveryPlce.value, frm1.txtDeliveryPlceNm.value, "인도장소", "인도장소명", gstrDeliveryPlceMajor)
	End If
End Sub	

<!--
'====================================  3.2.5 btnLoadingPortOnClick()  ==================================
-->
Sub btnLoadingPortOnClick()
	If frm1.txtLoadingPort.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "선적항", "선적항명", gstrLoadingPortMajor)
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
		Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "원산지", "원산지명", gstrOriginMajor)
	End If
End Sub
	
	
<!--
'======================================  3.2.9 btnAgentOnClick()  ======================================
-->
Sub btnAgentOnClick()
	If frm1.txtAgent.readOnly <> True Then
		Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "대행자", "대행자명")
	End If
End Sub

<!--
'=================================  3.2.10 btnManufacturerOnClick()  ===================================
-->
Sub btnManufacturerOnClick()
	If frm1.txtManufacturer.readOnly <> True Then
		Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "제조자", "제조자명")
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
		Call OpenBizPartner(frm1.txtForwarder.value, frm1.txtForwarderNm.value, "선박회사", "선박회사명")
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
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
-->
Sub btnPosting_OnClick()
	Dim Answer
	Dim strVal
	Dim strBLNo
	
    Err.Clear               

	strBLNo = frm1.txtBLNo.value
		
	If strBLNo = "" Then	
		Call DisplayMsgBox("900002","X","X","X")   '조회를 먼저 하십시오.
		Exit Sub
	Else
		
		if frm1.rdoPostingFlg2.Checked = True then 
			Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")   '작업을 수행 하시겠습니까?
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
			strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)          'B/L 관리번호 
			strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo.value)          '매입번호 
			strVal = strVal & "&txtBlIssueDt=" & Trim(frm1.txtBlIssueDt.Text) 'B/L 접수일 
			if frm1.rdoPostingFlg2.Checked = True then                        '확정여부 
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
		Call DisplayMsgBox("970021", "X","환율", "X")
		Call ClickTab1()
		frm1.txtXchRate.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	    
    'if Trim(frm1.txtPrePayNo.value) <> "" And (Trim(UNICDbl(frm1.txtPrePayDocAmt.text)) = "" Or Trim(UNICDbl(frm1.txtPrePayDocAmt.text)) = "0") then
	'	Call DisplayMsgBox("970021", "X","선급금", "X")
	'	Call ClickTab1()
	'	frm1.txtPrePayDocAmt.focus
	'	Set gActiveElement = document.activeElement
	'	Exit Function
	'End if
				
	If ValidDateCheckLocal(frm1.txtLoadingDt, frm1.txtBlIssueDt) = False Then Exit Function
	'지불예정일 체크하지 않음(2003.09.22)
	'If ValidDateCheckLocal(frm1.txtBlIssueDt, frm1.txtSetlmnt) = False Then Exit Function			
	If ValidDateCheckLocal(frm1.txtLoadingDt, frm1.txtDischgeDt) = False Then Exit Function	
		
	'if UCase(frm1.hdnLoanflg.value) = "Y" then
	'	if Trim(frm1.txtLoanNo.value) <> "" And (Trim(UNICDbl(frm1.txtLoanAmt.text)) = "" Or Trim(UNICDbl(frm1.txtLoanAmt.text)) = "0") then
	'		Call DisplayMsgBox("970021", "X","차입금", "X")
	'		Call ClickTab1()
	'		frm1.txtLoanAmt.focus
	'		Set gActiveElement = document.activeElement
	'		Exit Function
	'	End if
	'	if UNICDbl(frm1.txtDocAmt.Text) > 0 then
	'		if UNICDbl(frm1.txtLoanAmt.text) > UNICDbl(frm1.txtDocAmt.text) then 
	'			Call DisplayMsgBox("970023", "X","B/L금액","차입금")
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

	Call ggoOper.LockField(Document, "Q")				'쿼리후 수정불가		
		
	'**수정(2003.03.26)-회계모듈이 없어도 확정,확정취소 가능하도록 수정함.
	if frm1.rdoPostingflg(0).Checked = True then        '확정이 된경우 수정,삭제 불가 
		Call SetToolbar("11100000000111")			
		Call ChangeTag(True)
		frm1.btnPosting.value = "확정취소"
		frm1.btnPosting.disabled = false
		if interface_Account <> "N" then
			frm1.btnGlSel.disabled = false
		Else
			frm1.btnGlSel.disabled = True
		End If
	else
		Call SetToolbar("11111000000111")
		Call ChangeTag(False)
		frm1.btnPosting.value = "확정"
		frm1.btnPosting.disabled = false
		frm1.btnGlSel.disabled = true
		'Call setLoan()
	end if
		
    if frm1.hdnGlType.Value = "A" Then
       frm1.btnGlSel.value = "회계전표조회"
    elseif frm1.hdnGlType.Value = "T" Then
       frm1.btnGlSel.value = "결의전표조회"
    elseif frm1.hdnGlType.Value = "B" Then
       frm1.btnGlSel.value = "전표조회"
    end if	
    '추가(Detail이 존재하지 않으면 확정버튼 Disable시킴)
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수입B/L정보</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수입B/L기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenPoRef">발주참조</A> | <A href="vbscript:OpenLCRef">L/C참조</A></TD>
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
												<TD CLASS=TD5 NOWRAP>B/L 관리번호</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=32 MAXLENGTH=18 TAG="12XXXU" ALT="B/L 관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnBLNoOnClick()"></TD>
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
									<!-- 첫번째 탭 내용 -->
									<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L 관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo1" SIZE=34 MAXLENGTH=18 TAG="25XXXU" ALT="B/L 관리번호"></TD>
											<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX CHECKED ID="ChkPrepay" tag="24" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">&nbsp;선급금여부
											                     <INPUT TYPE=TEXT NAME="ChkPrepay1"  style="HEIGHT: 19px; WIDTH: 1px" MAXLENGTH=0 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrepayNo" align=top TYPE="BUTTON" onclick="vbscript:OpenPpNo()"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT ALT=B/L번호 NAME="txtBLDocNo" MAXLENGTH=18 TYPE=TEXT SIZE=34  TAG="22XXXU">
											<TD CLASS=TD5 NOWRAP>발주번호</TD>
 											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPONo" TYPE=TEXT SIZE=20  TAG="24XXXU">
 														         <INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="21" CHECKED ID="chkPoNoCnt">&nbsp;발주번호지정</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>선적일</TD>
  			        						<TD CLASS=TD6 NOWRAP>
  			        							<Table Cellspacing=0 Cellpadding=0>	
  			        								<TR>
  			        									<TD NOWRAP>
  			        										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=선적일 NAME="txtLoadingDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;
  			        									</TD>
  			        									<TD NOWRAP>
														 	 &nbsp;&nbsp;B/L접수일
														</TD>
														<TD NOWRAP>
														 	<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L접수일 NAME="txtBlIssueDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP>L/C번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=20  TAG="24XXXU">
																 <INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="21" CHECKED ID="chkLcNoCnt">&nbsp;L/C번호지정</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>도착일</TD>
  			        						<TD CLASS=TD6 NOWRAP>
  			        							<Table Cellpadding=0 Cellspacing=0>
  			        								<TR>
  			        									<TD NOWRAP>
  			        										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=도착일 NAME="txtDischgeDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;
  			        									</TD>
  			        									<TD NOWRAP>
  			        										 &nbsp;지불예정일
  			        									</TD>
  			        									<TD NOWRAP>
  			        										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=지불예정일 NAME="txtSetlmnt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
  			        									</TD>
  			        								</TR>
  			        							</Table>
  			        						</TD>
											<TD CLASS=TD5 NOWRAP>운송방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="운송방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" onclick="vbscript:btnTransportOnClick()">
																 <INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>선박회사</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtForwarder" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="선박회사"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnForwarder" align=top TYPE="BUTTON" onclick="vbscript:btnForwarderOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtForwarderNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>VESSEL명</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL명" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X">
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>항차번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVoyageNo" TYPE=TEXT SIZE=34 MAXLENGTH=20  TAG="21XXXU">											
											<TD CLASS=TD5 NOWRAP>선박국적</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVesselCntry" SIZE=10  MAXLENGTH=3 TAG="21XXXU" ALT="선박국적"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVesselCntry" align=top TYPE="BUTTON" onclick="vbscript:btnVesselCntryOnClick()"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>총포장갯수</TD>
 	        								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=총포장갯수 NAME="txtTotPackingCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>도착항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDischgePort" SIZE=10 MAXLENGTH=5  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" onclick="vbscript:btnDischgePortOnClick()">
																 <INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>화폐</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table CellPadding=0 cellspacing=0>
													<TR>
														<TD>
															<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐" OnChange="VBScript:ChangeCurrency()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
														</TD>
													</TR>
												</Table>
											</TD>
	        								<TD CLASS=TD5 NOWRAP>환율</TD>
			        						<TD CLASS=TD6 NOWRAP>
				        						<Table Cellpadding=0 Cellspacing=0>
				        							<TR>
				        								<TD NOWRAP>
				        									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환율 NAME="txtXchRate" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="22X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
				        								</TD>
													</TR>
												</Table>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L금액</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L금액 NAME="txtDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>B/L자국금액</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L자국금액 NAME="txtLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>지급유형</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="지급유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayType" align=top TYPE="BUTTON" onclick="vbscript:btnPayTypeOnClick()">
																 <INPUT TYPE=TEXT NAME="txtPayTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>대금결제참조</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsTxt" ALT="대금결제참조" TYPE=TEXT MAXLENGTH=120 SIZE=34 TAG="21X"></TD>
										</TR>																													
										<TR>
											<TD CLASS=TD5 NOWRAP>결제방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayMethod" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="결제방법">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtPayMethodNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>결제기간</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellspacing=0 Cellpadding=0>
													<TR>
														<TD NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=결제기간 NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="24X7" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
														<TD NOWRAP>
											                &nbsp;DAYS&nbsp;가격조건&nbsp;
											            </TD>
											            <TD NOWRAP>
											                <INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10  MAXLENGTH=5 TAG="24XXXU" ALT="가격조건"></TD>
											            </TD>
											        </TR>
											    </Table>
											</TD>
										</TR>
										<TR><!--
											<TD CLASS=TD5 NOWRAP>선급금번호</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
															<INPUT TYPE=TEXT NAME="txtPrePayNo" SIZE=32 STYLE="Text-Transform: uppercase" MAXLENGTH=18 TAG="21X" ALT="선급금번호" OnChange="changePpNo()"><IMG SRC="../../image/btnPopup.gif" NAME="btnPrePayNo" align=top TYPE="BUTTON">
														</TD>
													</TR>
												</Table>
											</TD>-->
													<TD CLASS=TD5 NOWRAP>지급처</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayeeCd" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="지급처" OnChange="VBScript:ChangePayeeCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" OnClick="OpenBpMpa()">
																 <INPUT TYPE=TEXT NAME="txtPayeeNm" SIZE=20 TAG="24"></TD>
											<!--TD CLASS=TD5 NOWRAP>차입금번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNo" SIZE=32  MAXLENGTH=18 TAG="21XXXU" ALT="차입금번호" OnChange="VBScript:ChangeLoanNo()"><IMG SRC="../../image/btnPopup.gif" NAME="btnLoanNo" align=top TYPE="BUTTON" onclick="vbscript:btnLoanNoOnClick()"></TD-->
											<TD CLASS=TD5 NOWRAP>매입번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=34 TAG="24XXXU" MAXLENGTH=18></TD> <!-- TAG="25XNXU" -->
										</TR>
										<TR>
											<!--<TD CLASS=TD5 NOWRAP>선급금</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=선급금 NAME="txtPrePayDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											-->
											<TD CLASS=TD5 NOWRAP>세금계산서발행처</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBuildCd" SIZE=10  MAXLENGTH=10 TAG="22XXXU" ALT="세금계산서발행처" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpMgs" align=top TYPE="BUTTON" OnClick="OpenBpMgs()">
																 <INPUT TYPE=TEXT NAME="txtBuildNm" SIZE=20 TAG="24"></TD>
											<!--TD CLASS=TD5 NOWRAP>차입금</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=차입금 NAME="txtLoanAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" Title="FPDOUBLESINGLE" ></OBJECT>');</SCRIPT></TD-->
											<TD CLASS=TD5 NOWRAP>수출자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR><!--
											<TD CLASS=TD5 NOWRAP>현금출금액</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=현금출금액 NAME="txtCashAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											-->
											<TD CLASS=TD5 NOWRAP>매입형태</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvType" SIZE=10 MAXLENGTH=5 TAG="22XNXU" ALT="매입형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON" OnClick="OpenIvType()">
																 <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>확정여부</TD>
     	         							<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingFlg" TAG="24X" VALUE="Y" ID="rdoPostingFlg1"><LABEL FOR="rdoPostingFlg1">&nbsp;Y&nbsp;</LABEL> 
     	         							                     <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingFlg" TAG="24X" VALUE="N" CHECKED ID="rdoPostingFlg2"><LABEL FOR="rdoPostingFlg2">&nbsp;N&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</LABEL>
     	         							                     <INPUT TYPE=TEXT NAME="txtGlNo" ALT="전표번호" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD></TD>
										</TR>	
										<%Call SubFillRemBodyTD5656(2)%>
									</TABLE>
									</DIV>
									<!-- 두번째 탭 내용 
									<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>-->
									<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>포장조건</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="포장형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPackingType" align=top TYPE="BUTTON" onclick="vbscript:btnPackingTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>포장참고사항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackingTxt" ALT="포장참고사항" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>총중량</TD>
 	        								<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=총중량 NAME="txtGrossWeight" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP>컨테이너수</TD>
 	        								<TD CLASS=TD6 NOWRAP> 	        								
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=컨테이너수 NAME="txtContainerCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="21X31Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</Table>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>순중량</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
														    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=순중량 NAME="txtNetWeight" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="21X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;
														</TD>
														<TD>
														    <INPUT TYPE=TEXT NAME="txtWeightUnit" SIZE=10 MAXLENGTH=3 SIZE=20 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON" onclick="vbscript:btnWeightUnitOnClick()">
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP>총용적</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table Cellpadding=0 Cellspacing=0>
													<TR>
														<TD NOWRAP>
														    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=총용적 NAME="txtGrossVolumn" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;
														</TD>
														<TD NOWRAP>
														    <INPUT TYPE=TEXT NAME="txtVolumnUnit" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVolumnUnit" align=top TYPE="BUTTON" onclick="vbscript:btnVolumnUnitOnClick()"></TD>
														</TD>
													</TR>
												</Table>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>운임지불방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="운임지불방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFreight" align=top TYPE="BUTTON" onclick="vbscript:btnFreightOnClick()">
																 <INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>운임지불장소</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFreightPlce" ALT="운임지불장소" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>최종목적지</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="최종목적지" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>인도장소</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeliveryPlce" SIZE=10 MAXLENGTH=5  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeliveryPlce" align=top TYPE="BUTTON" onclick="vbscript:btnDeliveryPlceOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeliveryPlceNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>수취장소</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiptPlce" ALT="수취장소" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>선적항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoadingPort" SIZE=10 MAXLENGTH=5  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" onclick="vbscript:btnLoadingPortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>환적국가</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTranshipCntry" SIZE=10 MAXLENGTH=3 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTranshipCntry" align=top TYPE="BUTTON" onclick="vbscript:btnTranshipCntryOnClick()"></TD>
											<TD CLASS=TD5 NOWRAP>환적일</TD>
  			        						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환적일 NAME="txtTranshipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>원산지</TD>
    										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrigin" SIZE=10 MAXLENGTH=5 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" onclick="vbscript:btnOriginOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>원산지국가</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="원산지국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" onclick="vbscript:btnOriginCntryOnClick()"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L발행장소</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLIssuePlce" ALT="B/L발행장소" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>B/L발행부수</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L발행부수 NAME="txtBLIssueCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X72" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											
										</TR>
    									<TR>
											<TD CLASS=TD5 NOWRAP>대행자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="대행자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" onclick="vbscript:btnAgentOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>제조자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="제조자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" onclick="vbscript:btnManufacturerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
										</TR>
    									<TR>
											<TD CLASS=TD5 NOWRAP>구매그룹</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="24XXXU" ALT="구매그룹">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTaxBizArea" SIZE=10 MAXLENGTH=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" onclick="vbscript:btnTaxBizAreaOnClick()">
																 <INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="24"></TD>
										</TR>										
    									<TR>
											<TD CLASS=TD5 NOWRAP>구매조직</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="구매조직">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>수입자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수입자">&nbsp;&nbsp;&nbsp;&nbsp;
																 <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5>비고</TD>
											<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT TYPE=TEXT  NAME="txtRemark" ALT="비고" tag = "21" SIZE=90 MAXLENGTH=70></TD>
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
							<TD><BUTTON NAME="btnPosting" CLASS="CLSMBTN">확정</BUTTON>&nbsp;
							<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
							</TD>
							<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadBLDtl()">B/L 내역등록</A> | <A href="vbscript:LoadChargeHdr()">경비등록</A> | <A href="vbscript:LoadIvPayment()">지급내역등록</A></TD>
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
