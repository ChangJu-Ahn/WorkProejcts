Option Explicit          

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const TAB1 = 1									
Const TAB2 = 2

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID 					= "m3111mb1_ko441.asp"											
Const BIZ_OnLine_ID 				= "m3111ab1.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL 		= "M3112MA1_KO441"
Const BIZ_PGM_JUMP_ID_PUR_CHARGE	= "M6111MA2_KO441"

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgIntFlgMode					'☜: Variable is for Operation Status
Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim gSelframeFlg
Dim IsOpenPop          
Dim lblnWinEvent
Dim lgOpenFlag  
Dim lgTabClickFlag  
Dim arrCollectVatType
Dim StartDate, EndDate
Dim iDBSYSDate


'========================================================================================
' Function Name : OnLineQueryOK
' Function Desc : 
'========================================================================================
Function OnLineQueryOK() 
	If Trim(frm1.txtSupplierCd.value) <> "" Then Call SupplierLookUp()    
	'======================== 추후에 수정=======================
	if Trim(frm1.txtPotypeCd.Value) <> "" then Call ChangePotype()
	'======================== 추후에 수정=======================
End Function

'==========================================================================================
'   Event Name : SupplierLookUp
'   Event Desc : 
'==========================================================================================
Function SupplierLookup()
    Err.Clear                                          
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
    
    Dim strVal
    
    if Trim(frm1.txtSupplierCd.Value) = "" then
    	Exit Function
    End if
    
	lgBlnFlgChgValue = true
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "SupplierLookupAfterOnline"			
    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.txtSupplierCd.value)	
    
    If LayerShowHide(1) = False Then Exit Function
    
	Call RunMyBizASP(MyBizASP, strVal)										
	
End Function
'==========================================================================================
'   Event Name : ChangePotype
'   Event Desc : txtPotypeCd Chagne Event
'==========================================================================================
Sub ChangePotype()

	If gLookUpEnable = False Then
		Exit Sub
	End If
	
	Call PotypeRef()
	
End Sub

'==========================================================================================
'   Event Name : ChangeSupplier
'   Event Desc : txtSupplierCd Chagne Event
'==========================================================================================
Sub ChangeSupplier()
	If gLookUpEnable = False Then
		Exit Sub
	End If

	Call SpplRef()
End Sub

'==========================================   PotypeRef()  ======================================
'	Name : PotypeRef()
'	Description : 
'=========================================================================================================
Sub PotypeRef()
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
    Dim strVal
    
    if Trim(frm1.txtPotypeCd.Value) = "" then
		Call DisplayMsgBox("205152", "X", "발주형태", "X")
		frm1.txtPotypeCd.focus
    	Exit Sub
    End if
    
	if lgIntFlgMode <> Parent.OPMD_UMODE Then 
		lgBlnFlgChgValue = true
	end if		
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpPoType"			
    strVal = strVal & "&txtPoTypeCd=" & Trim(frm1.txtPoTypeCd.value)
    strVal = strVal & "&txtTabClickFlag=" & lgTabClickFlag

    If LayerShowHide(1) = False Then Exit Sub
    
	Call RunMyBizASP(MyBizASP, strVal)								
End Sub

'==========================================   SpplRef()  ======================================
'	Name : SpplRef()
'	Description : It is Call at txtSupplier Change Event
'=========================================================================================================
Sub SpplRef()
    Err.Clear
    
    Dim strVal
    
    if Trim(frm1.txtSupplierCd.Value) = "" then
    	Exit Sub
    End if
    
	lgBlnFlgChgValue = true
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"			
    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.txtSupplierCd.value)	
    strVal = strVal & "&lgPGCd=" & lgPGCd

    If LayerShowHide(1) = False Then Exit Sub
    
	Call RunMyBizASP(MyBizASP, strVal)										
End Sub

'==========================================   Cfm()  ======================================
'	Name : Cfm()
'	Description : 확정버튼,확정취소버튼의 Event 합수 
'=========================================================================================================
Sub Cfm()
    Dim IntRetCD 
    
    Err.Clear                                                               
    
    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if
	
	if frm1.rdoRelease(0).checked = True then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnCfm.disabled = False	'20040315          
			Exit Sub
		Else 
			frm1.btnCfm.disabled = True		'20040315   
		End If
		Call DbSave("Cfm")				                                    
					                                                 
	elseif frm1.rdoRelease(1).checked = True then
			
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnCfm.disabled = False	'200308    
			Exit Sub
		Else 		
			frm1.btnCfm.disabled = True		'200308 	
		End If
		Call DbSave("UnCfm")
	End if
	
End Sub

'-------------------------------------------------------------------
'		확정여부에 따라 Field의 속성을 Protect로 전환,복구 시키는 함수 
'--------------------------------------------------------------------
Function ChangeTag(Byval Changeflg)
	with frm1
	
		If Changeflg = true Then
			'첫번째 Tab
			'ggoOper.SetReqAttr	.txtPoTypeCd, "Q"
			ggoOper.SetReqAttr	.txtPoDt, "Q"
			ggoOper.SetReqAttr	.txtGroupCd, "Q"
			ggoOper.SetReqAttr	.txtSupplierCd, "Q"
			ggoOper.SetReqAttr	.txtCurr, "Q"
			ggoOper.SetReqAttr	.txtXch, "Q"
			ggoOper.SetReqAttr	.txtVatType, "Q"
			ggoOper.SetReqAttr	.txtPayTermCd, "Q"
			ggoOper.SetReqAttr	.txtPayDur, "Q"
			ggoOper.SetReqAttr	.txtPayTermstxt, "Q"
			ggoOper.SetReqAttr	.txtPayTypeCd, "Q"
			ggoOper.SetReqAttr	.txtSuppSalePrsn, "Q"
			ggoOper.SetReqAttr	.txtTel, "Q"
			ggoOper.SetReqAttr	.txtRemark, "Q"
			ggoOper.SetReqAttr	.rdoMergPurFlg1, "Q"
			ggoOper.SetReqAttr	.rdoMergPurFlg2, "Q"
			ggoOper.SetReqAttr  .rdoVatFlg1,"Q"
            ggoOper.SetReqAttr  .rdoVatFlg2,"Q"
			
			'두번째 Tab
			ggoOper.SetReqAttr	.txtOffDt, "Q"
			ggoOper.SetReqAttr	.txtDvryDt, "Q"
			ggoOper.SetReqAttr	.txtExpiryDt, "Q"
			ggoOper.SetReqAttr	.txtInvNo, "Q"
			ggoOper.SetReqAttr	.txtIncotermsCd, "Q"
			ggoOper.SetReqAttr	.txtTransCd, "Q"
			ggoOper.SetReqAttr	.txtBankCd, "Q"
			ggoOper.SetReqAttr	.txtDvryPlce, "Q"
			ggoOper.SetReqAttr	.txtApplicantCd, "Q"
			ggoOper.SetReqAttr	.txtManuCd, "Q"
			ggoOper.SetReqAttr	.txtAgentCd, "Q"
			ggoOper.SetReqAttr	.txtOrigin, "Q"
			ggoOper.SetReqAttr	.txtPackingCd, "Q"
			ggoOper.SetReqAttr	.txtInspectCd, "Q"
			ggoOper.SetReqAttr	.txtDisCity, "Q"
			ggoOper.SetReqAttr	.txtDisPort, "Q"
			ggoOper.SetReqAttr	.txtLoadPort, "Q"
			ggoOper.SetReqAttr	.txtShipment, "Q"
		Else
			'첫번째 Tab
			ggoOper.SetReqAttr	.txtPoNo2, "D"
			ggoOper.SetReqAttr	.txtPoDt, "N"
			ggoOper.SetReqAttr	.txtGroupCd, "N"
			ggoOper.SetReqAttr	.txtSupplierCd, "N"
			ggoOper.SetReqAttr	.txtCurr, "N"
			ggoOper.SetReqAttr	.txtXch, "D"
			ggoOper.SetReqAttr	.txtVatType, "D"
			ggoOper.SetReqAttr	.txtPayTermCd, "N"
			ggoOper.SetReqAttr	.txtPayDur, "D"
			ggoOper.SetReqAttr	.txtPayTermstxt, "D"
			ggoOper.SetReqAttr	.txtPayTypeCd, "D"
			ggoOper.SetReqAttr	.txtSuppSalePrsn, "D"
			ggoOper.SetReqAttr	.txtTel, "D"
			ggoOper.SetReqAttr	.txtRemark, "D"
			ggoOper.SetReqAttr	.rdoMergPurFlg1, "D"
			ggoOper.SetReqAttr	.rdoMergPurFlg2, "D"

			if .hdnImportflg.value = "Y" then
			    ggoOper.SetReqAttr	.txtDvryDt, "N"
			else     
			    ggoOper.SetReqAttr	.txtDvryDt, "D"
			end if	
			
			'두번째 Tab
			ggoOper.SetReqAttr	.txtOffDt, "N"
			'ggoOper.SetReqAttr	.txtDvryDt, "N"
			ggoOper.SetReqAttr	.txtExpiryDt, "D"
			ggoOper.SetReqAttr	.txtInvNo, "D"
			ggoOper.SetReqAttr	.txtIncotermsCd, "N"
			ggoOper.SetReqAttr	.txtTransCd, "N"
			ggoOper.SetReqAttr	.txtBankCd, "D"
			ggoOper.SetReqAttr	.txtDvryPlce, "D"
			ggoOper.SetReqAttr	.txtApplicantCd, "N"
			ggoOper.SetReqAttr	.txtManuCd, "D"
			ggoOper.SetReqAttr	.txtAgentCd, "D"
			ggoOper.SetReqAttr	.txtOrigin, "D"
			ggoOper.SetReqAttr	.txtPackingCd, "D"
			ggoOper.SetReqAttr	.txtInspectCd, "D"
			ggoOper.SetReqAttr	.txtDisCity, "D"
			ggoOper.SetReqAttr	.txtDisPort, "D"
			ggoOper.SetReqAttr	.txtLoadPort, "D"
			ggoOper.SetReqAttr	.txtShipment, "D"
			
			If UCase(Trim(frm1.txtCurr.value)) = UCase(Parent.gCurrency) then
				Call ggoOper.SetReqAttr(frm1.txtXch,"Q")
			Else
				Call ggoOper.SetReqAttr(frm1.txtXch,"D")
			End If 
		End if 
	End With
End Function 

'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD

		
	If Kubun = 0 Then

		strTemp = ReadCookie("PoNo")
		
		If strTemp = "" then Exit Function
		
		frm1.txtPoNo.value = strTemp
		
		WriteCookie "PoNo" , ""
		
		Call MainQuery()
	
	elseIf Kubun = 1 Then
		
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If
		
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
		WriteCookie "PoNo" , frm1.txtPoNo.value
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)
	
	elseIf Kubun = 2 Then
	
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
	    WriteCookie "Process_Step" , "PO"
		WriteCookie "Po_No" , Trim(frm1.txtPoNo.value)
		WriteCookie "Pur_Grp", Trim(frm1.txtGroupCd.Value)
		WriteCookie "Po_Cur", Trim(frm1.txtCurr.Value)
		WriteCookie "Po_Xch", Trim(frm1.txtXch.Value)
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)
				
	End IF
End Function

'------------------------------------------------------------------------------------------
'Radio에서 Click을 할 경우 flag를 Setting
'------------------------------------------------------------------------------------------
Sub Setchangeflg()
	lgBlnFlgChgValue = True	
End Sub

'------------------------------------------------------------------------------------------
'사용자가 Radio Button을 Click할 때 마다 숨겨진 hdnRelease를 Setting
'------------------------------------------------------------------------------------------
Sub Changeflg()
	if frm1.rdoRelease(0).checked = true then
		frm1.hdnRelease.value= "N"
	else
		frm1.hdnRelease.value= "Y"
	end if 
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                        
    lgBlnFlgChgValue = False                                         
    IsOpenPop = False												
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	lgOpenFlag	= False
	lgTabClickFlag	= False
    Call SetToolbar("1110100000001111")
    frm1.rdoRelease(0).checked = true
    frm1.txtOffDt.text = EndDate
    frm1.txtPoDt.text = EndDate
    frm1.hdnCurr.value = Parent.gCurrency   
    frm1.btnCfm.disabled = true
    frm1.btnSend.disabled = true
    frm1.txtGroupCd.Value = Parent.gPurGrp
    frm1.txtXch.Text = ""
	frm1.txtApplicantCd.value = Parent.gCompany 
	frm1.txtApplicantNm.value = Parent.gCompanyNm
	frm1.btnCfm.value = "확정"
	frm1.txtPoNo.focus
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
  	frm1.txtGroupCd.value = lgPGCd
	End If
	Set gActiveElement = document.activeElement
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'*********************************************************************************************************
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	
	gSelframeFlg = TAB1
	
   	frm1.txtPoNo.focus
	Set gActiveElement = document.activeElement
End Function

Function ClickTab2()
	If lgOpenFlag = False Then 
		Call POTypeRef()		
		Exit Function
	End If

	If gSelframeFlg = TAB2 Then Exit Function

	IF frm1.hdnImportflg.value<>"Y" then
		Call DisplayMsgBox("17a007", "X", "X", "X")
		lgOpenFlag = False		
		Exit Function
	End if 

	Call changeTabs(TAB2)

	lgOpenFlag	= False
	lgTabClickFlag = False
	gSelframeFlg = TAB2
	frm1.txtOffDt.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetClickflag, ResetClickflag()  -----------------------------
'	Name : SetClickflag, ResetClickflag()
'	Description :  
'---------------------------------------------------------------------------------------------------------
Function SetClickflag()

	lgTabClickFlag = True	
	
End Function

Function ResetClickflag()

	lgTabClickFlag = False
	
End Function

'------------------------------------------  OpenMpOrderRef()  -------------------------------------------------
'	Name : OpenMpOrderRef()
'	Description : OpenMpOrderRef PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenMpOrderRef()
	Dim strRet
	Dim strParam
	
	if frm1.rdoRelease(1).checked = true then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	strParam = Parent.gColSep
	strParam = strParam & Trim(frm1.txtPoDt.Text)

	strRet = window.showModalDialog("m3011ra1.asp", strParam, _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtMaintNo.value = strRet(0)
		Call OnLineQuery()
		lgBlnFlgChgValue = true
	End If	
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
		Dim strRet
		Dim arrParam(2)
		Dim iCalledAspName
		Dim IntRetCD
		
		If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		lblnWinEvent = True
		
		arrParam(0) = "N"  'Return Flag
		arrParam(1) = "N"  'Release Flag
		arrParam(2) = ""  'STO Flag
		
		iCalledAspName = AskPRAspName("M3111PA1_KO441")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
			lblnWinEvent = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


		lblnWinEvent = False
	
		If strRet(0) = "" Then
			frm1.txtPoNo.focus
			Exit Function
		Else
			frm1.txtPoNo.value = strRet(0)
			frm1.txtPoNo.focus
		End If	
		Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPotypeCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주형태"					
	arrParam(1) = "M_CONFIG_PROCESS"			
	
	arrParam(2) = Trim(frm1.txtPotypeCd.Value)
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and Ret_FLG <>" & FilterVar("Y", "''", "S") & " "							
	'arrParam(4) = "USAGE_FLG='Y'"							
	arrParam(5) = "발주형태"					
	
    arrField(0) = "PO_TYPE_CD"					
    arrField(1) = "PO_TYPE_NM"					
    
    arrHeader(0) = "발주형태"				
    arrHeader(1) = "발주형태명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPotypeCd.focus
		Exit Function
	Else
		frm1.txtPoTypeCd.Value    = arrRet(0)		
		frm1.txtPoTypeNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True

		Call PotypeRef()
		frm1.txtPotypeCd.focus
	End If	
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  OpenCurr()  -------------------------------------------------
'	Name : OpenCurr()
'	Description : OpenCurr PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCurr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtCurr.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "화폐"						
	arrParam(1) = "B_Currency"					
	
	arrParam(2) = Trim(frm1.txtCurr.Value)
	'arrParam(3) = Trim(frm1.txtItemNm2.Value)	
		
	arrParam(4) = ""							
	arrParam(5) = "화폐"						
	
    arrField(0) = "Currency"					
    arrField(1) = "Currency_Desc"				
    
    arrHeader(0) = "화폐"					
    arrHeader(1) = "화폐명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCurr.focus
		Exit Function
	Else
		frm1.txtCurr.Value    = arrRet(0)		
		frm1.txtCurrNm.Value  = arrRet(1)		
		Call ChangeCurr()
		frm1.txtCurr.focus
		lgBlnFlgChgValue = True
	End If	
	Set gActiveElement = document.activeElement 
End Function


Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & "  AND IN_OUT_FLAG = " & FilterVar("O", "''", "S") & " "	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"	
    arrField(2) = "BP_RGST_NO"				
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"	
    arrHeader(2) = "사업자등록번호"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
	
		Call SpplRef()
		frm1.txtSupplierCd.focus
	End If	
	Set gActiveElement = document.activeElement 
End Function


Function OpenVat()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtVattype.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT형태"				
	arrParam(1) = "B_MINOR,b_configuration"	
	
	arrParam(2) = Trim(frm1.txtVattype.Value)
		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd "	
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT형태"					
	
    arrField(0) = "b_minor.MINOR_CD"			
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"	
    
    arrHeader(0) = "VAT형태"					
    arrHeader(1) = "VAT형태명"				
    arrHeader(2) = "VAT율"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtVattype.focus
		Exit Function
	Else
		frm1.txtVattype.Value    = arrRet(0)		
		frm1.txtVattypeNm.Value    = arrRet(1)		
		frm1.txtVatRt.Value = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		frm1.txtVattype.focus
		lgBlnFlgChgValue = True
	End If	
	Set gActiveElement = document.activeElement 
End Function


Function OpenBank()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBankCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "송금은행"	
	arrParam(1) = "B_Bank"				
	
	arrParam(2) = Trim(frm1.txtBankCd.Value)
	'arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "송금은행"			
	
    arrField(0) = "BANK_CD"	
    arrField(1) = "BANK_NM"	
    
    arrHeader(0) = "송금은행"		
    arrHeader(1) = "송금은행명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBankCd.focus
		Exit Function
	Else
		frm1.txtBankCd.Value= arrRet(0)		
		frm1.txtBankNm.Value= arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtBankCd.focus
	End If	
	Set gActiveElement = document.activeElement 
End Function 

Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)
		frm1.txtGroupNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtGroupCd.focus
	End If	
	Set gActiveElement = document.activeElement 
End Function 

Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	if Trim(frm1.txtPayTermCd.Value) = "" then
		Call DisplayMsgBox("17a002", "X","결제방법", "X")
		Exit Function
	End if

	If IsOpenPop = True Or UCase(frm1.txtPayTypeCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급유형"				
	arrParam(1) = "B_MINOR,B_CONFIGURATION," _
	& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""_
		& "And MINOR_CD= " & FilterVar(frm1.txtPayTermCd.value, "''", "S") & "  And SEQ_NO>=2)C"
	
	arrParam(2) = Trim(frm1.txtPayTypeCd.Value)
	
	arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
				& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("P", "''", "S") & " )"	
	arrParam(5) = "지급유형"					
	
	arrField(0) = "B_MINOR.MINOR_CD"						
	arrField(1) = "B_MINOR.MINOR_NM"				
    
    arrHeader(0) = "지급유형"				
    arrHeader(1) = "지급유형명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPayTypeCd.focus
		Exit Function
	Else
		frm1.txtPayTypeCd.Value = arrRet(0)
		frm1.txtPayTypeNm.Value = arrRet(1)
		lgBlnFlgChgValue 		= True
		frm1.txtPayTypeCd.focus
	End If	
	Set gActiveElement = document.activeElement 
End Function

Function OpenPaymeth()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPayTermCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "결제방법"					
	arrParam(1) = "B_Minor,b_configuration"		
	
	arrParam(2) = Trim(frm1.txtPayTermCd.Value)
	'arrParam(3) = Trim(frm1.txtPayNm.Value)	
	
	arrParam(4) = "b_minor.Major_Cd=" & FilterVar("B9004", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"						
	arrParam(5) = "결제방법"					
	
    arrField(0) = "b_minor.Minor_Cd"			
    arrField(1) = "b_minor.Minor_Nm"			
    arrField(2) = "b_configuration.REFERENCE"
    
    arrHeader(0) = "결제방법"				
    arrHeader(1) = "결제방법명"				
    arrHeader(2) = "Reference"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPayTermCd.focus
		Exit Function
	Else
		frm1.txtPayTermCd.focus
		frm1.txtPaytermCd.Value    = arrRet(0)		
		frm1.txtPaytermNm.Value    = arrRet(1)		
		frm1.txtReference.Value	   = arrRet(2)
		lgBlnFlgChgValue = True
		Call changePayterm()
	End If	
	Set gActiveElement = document.activeElement 
End Function


Function OpenMinorCode(MajorCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strtitle
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	'공통부분 
	arrParam(1) = "B_Minor"						
	
    arrField(0) = "Minor_Cd"					
    arrField(1) = "Minor_Nm"					
	arrParam(4) = "Major_Cd= " & FilterVar(MajorCode, "''", "S") & ""
    
    Select Case MajorCode
    Case "B9006"								
		if frm1.txtIncotermsCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtIncotermsCd.value)
		arrParam(0) = "가격조건"						
		arrParam(5) = "가격조건"						
		arrHeader(0) = "가격조건"					
		arrHeader(1) = "가격조건명"					
		
	Case "B9009"										
		if frm1.txtTransCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtTransCd.value)
		arrParam(0) = "운송방법"						
		arrParam(5) = "운송방법"						
		arrHeader(0) = "운송방법"					
		arrHeader(1) = "운송방법명"					
		
	Case "B9007"										
		if frm1.txtPackingCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtPackingCd.value)
		arrParam(0) = "포장조건"						
		arrParam(5) = "포장조건"						
		arrHeader(0) = "포장조건"					
		arrHeader(1) = "포장조건명"					
		
	Case "B9008"										
		if frm1.txtInspectCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtInspectCd.value)
		arrParam(0) = "검사방법"					
		arrParam(5) = "검사방법"					
		arrHeader(0) = "검사방법"				
		arrHeader(1) = "검사방법명"				
	
	Case "B9095"									
		if frm1.txtDvryPlce.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtDvryPlce.value)
		arrParam(0) = "인도장소"					
		arrParam(5) = "인도장소"					
		arrHeader(0) = "인도장소"						
		arrHeader(1) = "인도장소명"						
	
	Case "B9094"											
		if frm1.txtOrigin.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtOrigin.value)
		arrParam(0) = "원산지"							
		arrParam(5) = "원산지"							
		arrHeader(0) = "원산지"							
		arrHeader(1) = "원산지명"						
		
	Case "B9096"											
		if frm1.txtDisCity.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtDisCity.value)
		arrParam(0) = "도착도시"						
		arrParam(5) = "도착도시"						
		arrHeader(0) = "도착도시"					
		arrHeader(1) = "도착도시명"					
		
	Case "B9092"										
		if frm1.txtDisPort.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtDisPort.value)
		arrParam(0) = "도착항"						
		arrParam(5) = "도착항"						
		arrHeader(0) = "도착항"						
		arrHeader(1) = "도착항명"					
		
	Case "B9092-1"										
		if frm1.txtLoadPort.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtLoadPort.value)
		arrParam(0) = "선적항"						
		arrParam(5) = "선적항"						
		arrHeader(0) = "선적항"						
		arrHeader(1) = "선적항명"					
		arrParam(4) = "Major_Cd=" & FilterVar("B9092", "''", "S") & ""	
		
	Case "A1006"										
		if frm1.txtPaytypecd.ReadOnly = true then
			IsOpenPop = False
			Exit Function 
		End if
		arrParam(2) = Trim(frm1.txtPaytypeCd.value)
		arrParam(0) = "지급유형"						
		arrParam(5) = "지급유형"						
		arrHeader(0) = "지급유형"					
		arrHeader(1) = "지급유형명"					
		
    End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case MajorCode
			Case "B9006"								
				frm1.txtIncotermsCd.focus
			Case "B9009"								
				frm1.txtTransCd.focus
			Case "B9007"								
				frm1.txtPackingCd.focus
			Case "B9008"								
				frm1.txtInspectCd.focus
			Case "B9095"								
				frm1.txtDvryPlce.focus
			Case "B9094"								
				frm1.txtOrigin.focus
			Case "B9096"								
				frm1.txtDisCity.focus
			Case "B9092"								
				frm1.txtDisPort.focus
			Case "B9092-1"								
				frm1.txtLoadPort.focus
			Case "A1006"								
				frm1.txtPaytypeCd.focus
		End Select
		Exit Function
	Else
		Select Case MajorCode
			Case "B9006"								
				frm1.txtIncotermsCd.Value   = arrRet(0)		
				frm1.txtIncotermsNm.Value   = arrRet(1)		
				frm1.txtIncotermsCd.focus
				lgBlnFlgChgValue 			= True
			Case "B9009"								
				frm1.txtTransCd.Value    	= arrRet(0)		
				frm1.txtTransNm.Value    	= arrRet(1)		
				lgBlnFlgChgValue 			= True
				frm1.txtTransCd.focus
			Case "B9007"								
				frm1.txtPackingCd.Value    	= arrRet(0)		
				frm1.txtPackingNm.Value    	= arrRet(1)		
				frm1.txtPackingCd.focus
				lgBlnFlgChgValue 			= True
			Case "B9008"								
				frm1.txtInspectCd.Value    	= arrRet(0)		
				frm1.txtInspectNm.Value    	= arrRet(1)		
				frm1.txtInspectCd.focus
				lgBlnFlgChgValue 			= True
			Case "B9095"								
				frm1.txtDvryPlce.Value    	= arrRet(0)		
				frm1.txtDvryPlceNm.Value    = arrRet(1)		
				frm1.txtDvryPlce.focus
				lgBlnFlgChgValue 			= True
			Case "B9094"								
				frm1.txtOrigin.Value    	= arrRet(0)		
				frm1.txtOriginNm.Value    	= arrRet(1)		
				frm1.txtOrigin.focus
				lgBlnFlgChgValue 			= True
			Case "B9096"								
				frm1.txtDisCity.Value    	= arrRet(0)		
				frm1.txtDisCityNm.Value    	= arrRet(1)		
				frm1.txtDisCity.focus
				lgBlnFlgChgValue 			= True
			Case "B9092"								
				frm1.txtDisPort.Value    	= arrRet(0)		
				frm1.txtDisPortNm.Value    	= arrRet(1)		
				frm1.txtDisPort.focus
				lgBlnFlgChgValue 			= True
			Case "B9092-1"								
				frm1.txtLoadPort.Value	   	= arrRet(0)		
				frm1.txtLoadPortNm.Value   	= arrRet(1)		
				frm1.txtLoadPort.focus
				lgBlnFlgChgValue 			= True
			Case "A1006"								
				frm1.txtPaytypeCd.Value	   	= arrRet(0)		
				frm1.txtPaytypeNm.Value   	= arrRet(1)		
				frm1.txtPaytypeCd.focus
				lgBlnFlgChgValue 			= True
		End Select
	End If	
End Function

Function OpenBiz(strValue)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	'공통부분 
	arrParam(1) = "B_BIZ_PARTNER"						
	
    arrField(0) = "BP_Cd"				
    arrField(1) = "BP_Nm"				
    
    Select Case strValue
    Case "Appl"											
    	if UCase(frm1.txtApplicantCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		arrParam(2) = Trim(frm1.txtApplicantCd.value)
		arrParam(0) = "수입자"						
		arrParam(5) = "수입자"						
		arrHeader(0) = "수입자"						
		arrHeader(1) = "수입자명"					
		
	Case "Manu"											
		if UCase(frm1.txtManuCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		arrParam(2) = Trim(frm1.txtManuCd.value)
		arrParam(0) = "제조자"						
		arrParam(5) = "제조자"						
		arrHeader(0) = "제조자"						
		arrHeader(1) = "제조자명"					
		
	Case "Agent"										
		if UCase(frm1.txtAgentCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		arrParam(2) = Trim(frm1.txtAgentCd.value)
		arrParam(0) = "대행자"						
		arrParam(5) = "대행자"						
		arrHeader(0) = "대행자"						
		arrHeader(1) = "대행자명"					
		
    End Select
    
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case strValue
		Case "Appl"										
			frm1.txtApplicantCd.focus
		Case "Manu"										
			frm1.txtManuCd.focus
		Case "Agent"									
			frm1.txtAgentCd.focus
		End Select
		Exit Function
	Else
			
		Select Case strValue
		Case "Appl"										
			frm1.txtApplicantCd.Value    = arrRet(0)		
			frm1.txtApplicantNm.Value    = arrRet(1)		
			frm1.txtApplicantCd.focus
			lgBlnFlgChgValue = True
		Case "Manu"										
			frm1.txtManuCd.Value    = arrRet(0)		
			frm1.txtManuNm.Value    = arrRet(1)		
			frm1.txtManuCd.focus
			lgBlnFlgChgValue = True
		Case "Agent"									
			frm1.txtAgentCd.Value    = arrRet(0)		
			frm1.txtAgentNm.Value    = arrRet(1)		
			frm1.txtAgentCd.focus
			lgBlnFlgChgValue = True
		End Select
	End If	
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtPoAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtGrossPoAmt,.txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXch, .txtCurr.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec	
	End With
End Sub
'========================================================================================
' Function Name : ChangeCurr()
' Function Desc : 
'========================================================================================
Sub ChangeCurr()
	If UCase(Trim(frm1.txtCurr.value)) = UCase(Parent.gCurrency) Then
		frm1.txtXch.Text = 1
		Call ggoOper.SetReqAttr(frm1.txtXch,"Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtXch,"D")
		frm1.txtXch.Text = 0
	End If 
	Call CurFormatNumericOCX()
	
End Sub
'========================================================================================
' Function Name : changePayterm
' Function Desc : 
'========================================================================================
Sub changePayterm()
	frm1.txtPayTypeCd.Value = ""
	frm1.txtPayTypeNm.Value = ""
	frm1.txtPayDur.Text = 0	
End Sub

'========================================================================================
' Function Name : Sending
' Function Desc : 
'========================================================================================
Function Sending()
    Err.Clear                                           
    
    Sending = False                                     
    
	If LayerShowHide(1) = False Then Exit Function

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "SendingB2B"	
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)		
    
	Call RunMyBizASP(MyBizASP, strVal)								
	
    Sending = True                                                  
End Function

Function SendingOK()
	'msgbox "전송이 완료 되었습니다."	
End Function
'========================================================================================
' Function Name : OnLineQuery
' Function Desc : 주문서관리번호로 OnLine관련 조회 
'========================================================================================
Function OnLineQuery() 
    
    Err.Clear                                                       
    
    OnLineQuery = False                                             
    
	If LayerShowHide(1) = False Then Exit Function

    Dim strVal
    
    strVal = BIZ_OnLine_ID & "?txtMode=" & "OnLineLookUp"			
    strVal = strVal & "&txtMaintNo=" & Trim(frm1.txtMaintNo.value)
    
	Call RunMyBizASP(MyBizASP, strVal)								
	
    OnLineQuery = True                                              

End Function

'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================

Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation, parent.gLogoName
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
' Function Name : GetCollectTypeRef
' Function Desc : 
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
' Function Name : SetVatType
' Function Desc : 
'========================================================================================
Sub SetVatType()
	Dim VatType, VatTypeNm, VatRate

	VatType = frm1.txtVattype.value

	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

	frm1.txtVatTypeNm.value = VatTypeNm
	frm1.txtVatrt.text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
End Sub



'==========================================================================================
'   Event Name : OCX Event
'   Event Desc :
'==========================================================================================
Sub txtPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtPoDt.Focus
	End if
End Sub

Sub txtPoDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtOffDt_DblClick(Button)
	if Button = 1 then
		frm1.txtOffDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtOffDt.Focus
	End if
End Sub

Sub txtOffDt_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtXch_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtPayDur_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtDvryDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDvryDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtDvryDt.Focus
	End if
End Sub
Sub txtDvryDt_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub txtExpiryDt_DblClick(Button)
	if Button = 1 then
		frm1.txtExpiryDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtExpiryDt.Focus
	End if
End Sub
Sub txtExpiryDt_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub rdoVatFlg1_OnClick()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoVatFlg2_OnClick()
	lgBlnFlgChgValue = true	
End Sub
'==========================================================================================
'   Event Name : txtVat_Type_OnChange
'   Event Desc : 수주형태별로 무역정보 필수입력 처리 
'==========================================================================================
Sub txtVattype_OnChange()
	Call SetVatType()
End Sub

Sub setVatAmt()
	Dim sum
	  
	With frm1
	    sum = UNICDbl(.txtVatrt.text) * UNICDbl(.txtPoAmt.text)/100
	End With
End Sub 

Function ChkAuth()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i

  Err.Clear

	ChkAuth = False

	If CommonQueryRs("PO_NO", "m_pur_ord_hdr", " PO_NO=" & FilterVar(frm1.txtPoNo.value, "''", "S") & " And PUR_GRP=" & FilterVar(lgPGCd, "''", "S") , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)=False Then
		Exit Function
	End If
	ChkAuth = True
End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                
    Err.Clear                                                       
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    						
    Call InitVariables												
	
	If Not chkFieldByCell(frm1.txtPoNo, "A",1)	then
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
            
        Exit Function
    End If 

	'f lgPGCd <> "" then 
    '	If ChkAuth() = False Then
    '		Call DisplayMsgBox("210033", "X", "X", "X")
    '		Exit Function 
    '	End If
    'End If

    If DbQuery = False Then Exit Function
    Call Changeflg       
      
    FncQuery = True													
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                  
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ClickTab1()
    Call ggoOper.ClearField(Document, "1")                          
    Call ggoOper.ClearField(Document, "2")                          
    Call ggoOper.ClearField(Document, "3")                          
    Call ggoOper.LockField(Document, "N")                              
    Call ChangeTag(False)
    Call SetDefaultVal
    Call InitVariables													

    frm1.txtPoNo.focus
	Set gActiveElement = document.activeElement
    
    FncNew = True														
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD

    FncDelete = False
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    
    If IntRetCD = vbNo Then Exit Function
    														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                      
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True                                        
End Function

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                               
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
	IF frm1.hdnImportflg.value="Y" then
	
				
	    If Not chkEachFieldDomestic() Then                 
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        
	        Exit Function
	    End If

	    If Not chkEachFieldImport() Then                            
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        
	        Exit Function
	    End If

	else
	    If Not chkEachFieldDomestic() Then                            
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        
	        Exit Function
	    End If

	End if
		
	
	'# Tab 1
	
		
	If Not chkFieldLengthByCell (frm1.txtSuppSalePrsn,"A",1) then 
		Exit Function
	End If
	
	If Not chkFieldLengthByCell (frm1.txtTel,"A",1) then 
		Exit Function
	End If
	
	If Not chkFieldLengthByCell (frm1.txtPayTermstxt,"A",1) then 
		Exit Function
	End If

	If Not chkFieldLengthByCell (frm1.txtRemark,"A",1) then 
		Exit Function
	End If
	
	
	'# Tab 2
	
	If Not chkFieldLengthByCell (frm1.txtInvNo,"A",2) then 
		Exit Function
	End If	
	
	If Not chkFieldLengthByCell (frm1.txtShipment,"A",2) then 
		Exit Function
	End If
		
	if frm1.rdoVatFlg1.checked = true then
    	frm1.hdvatFlg.Value = "1"	'별도 
    else
    	frm1.hdvatFlg.Value = "2"	'포함 
    End if
	
    Call Changeflg
    
	If DbSave("ToolBar") = False Then Exit Function
    
    FncSave = True                                                     
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE											
    
    
    Call ggoOper.ClearField(Document, "1")                              
    Call ggoOper.LockField(Document, "N")								
    Call Changeflg
    Call ChangeTag(False)
    
    frm1.rdoRelease(0).checked = true
    Call SetToolbar("11101000000111")
	
	frm1.txtPoAmt.Text		= 0
	frm1.txtPoLocAmt.Text	= 0
	frm1.txtVatAmt.Text		= 0
	frm1.txtPoNo2.value = ""
	frm1.btnCfm.disabled = True
    frm1.btnSend.disabled = True
    lgBlnFlgChgValue = True
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
  	frm1.txtGroupCd.value = lgPGCd
	End If
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)										
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                               
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
 Function DbDelete() 
    Err.Clear                                                           
    
    DbDelete = False													
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003						
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
    
	If LayerShowHide(1) = False Then Exit Function    
	
	Call RunMyBizASP(MyBizASP, strVal)								
	
    DbDelete = True                                                 

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
 Function DbQuery() 
    
    Err.Clear                                                       
    
    DbQuery = False                                                 
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001					
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)	
    
    If LayerShowHide(1) = False Then Exit Function
    
	Call RunMyBizASP(MyBizASP, strVal)								
	
    DbQuery = True                                                  

End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

    Call setVatAmt   
    'Call ggoOper.LockField(Document, "Q")							
	Call LockHTMLField(frm1.txtPotypeCd , "P")
    frm1.btnCfm.disabled = False
    if frm1.rdoRelease(1).checked = true then
		Call ChangeTag(true)
		Call SetToolbar("11100000001111")
		if frm1.hdclsflg.value = "Y" then
    	    frm1.btnCfm.disabled = true
    	else
    	    frm1.btnCfm.disabled = False
    	end if
    	frm1.btnCfm.value = "확정취소"
    	frm1.btnSend.disabled = False
   	Else
		Call ChangeTag(False)
		frm1.txtPoNo.focus
		Call SetToolbar("11111000001111")
		frm1.btnCfm.value = "확정"
    	frm1.btnSend.disabled = True
	end if

	ggoOper.SetReqAttr	frm1.txtPoNo2, "Q"
	
	lgIntFlgMode = Parent.OPMD_UMODE										
    lgBlnFlgChgValue = False

	Call ClickTab1()

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
  	frm1.txtGroupCd.value = lgPGCd
	End If
	
	frm1.txtSupplierCd.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave(byval btnflg) 
    Err.Clear														

	DbSave = False													
    
    Dim strVal

	With frm1
	
		.txtMode.value = Parent.UID_M0002									
		.txtFlgMode.value = lgIntFlgMode
		
		if btnflg = "Cfm" then
			.txtMode.value = "Release" 				
		elseif btnflg = "UnCfm" then
			.txtMode.value = "UnRelease" 			
		end if
		
		if frm1.rdoMergPurFlg(0).Checked = True then
			frm1.hdnMergPurFlg.Value = "Y"
		else
			frm1.hdnMergPurFlg.Value = "N"
		end if
		
		If LayerShowHide(1) = False Then Exit Function
    		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                   
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()									
	lgBlnFlgChgValue = False
	Call MainQuery()	
End Function

'========================================================================================
' Function Name : chkEachFieldDomestic, chkEachFieldImport
' Function Desc : Manual check whether a value is entered at required field 
'========================================================================================
Function chkEachFieldDomestic()
	chkEachFieldDomestic = True
	
	If Not chkFieldByCell (frm1.txtPotypeCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtSupplierCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtPoDt, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtGroupCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtCurr, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtPayTermCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
End Function

Function chkEachFieldImport()
	chkEachFieldImport	= True
	
	If Not chkFieldByCell (frm1.txtDvryDt, "A",1) then 
		chkEachFieldImport = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtOffDt, "A",2) then 
		chkEachFieldImport = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtIncotermsCd, "A",2) then 
		chkEachFieldImport = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtTransCd, "A",2) then 
		chkEachFieldImport = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtApplicantCd, "A",2) then 
		chkEachFieldImport = False
		Exit Function
	End If
	
End Function


'========================================================================================
' Function Name : initFormatField()
' Function Desc : Manual Formatting fields as amount or date 
'========================================================================================
Function  initFormatField()
	
	call FormatDateField(frm1.txtPoDt)
	call FormatDateField(frm1.txtDvryDt)
	call FormatDateField(frm1.txtOffDt)
	call FormatDateField(frm1.txtExpiryDt)	
	
	call FormatDoubleSingleField(frm1.txtXch)
	call FormatDoubleSingleField(frm1.txtPoAmt)
	call FormatDoubleSingleField(frm1.txtPoLocAmt)
	call FormatDoubleSingleField(frm1.txtGrossPoAmt)
	call FormatDoubleSingleField(frm1.txtGrossPoLocAmt)
	call FormatDoubleSingleField(frm1.txtVatAmt)
	call FormatDoubleSingleField(frm1.txtVatrt)
	call FormatDoubleSingleField(frm1.txtPayDur)
	
	call LockobjectField(frm1.txtPoDt,"R")
	call LockobjectField(frm1.txtDvryDt,"O")
	call LockobjectField(frm1.txtOffDt,"R")
	call LockobjectField(frm1.txtExpiryDt,"O")
	
	call LockobjectField(frm1.txtXch,"O")
	call LockobjectField(frm1.txtPoAmt,"P")
	call LockobjectField(frm1.txtPoLocAmt,"P")
	call LockobjectField(frm1.txtGrossPoAmt,"P")
	call LockobjectField(frm1.txtGrossPoLocAmt,"P")
	call LockobjectField(frm1.txtVatAmt,"P")
	call LockobjectField(frm1.txtVatrt,"P")
	call LockobjectField(frm1.txtPayDur,"O")
		
End Function 
