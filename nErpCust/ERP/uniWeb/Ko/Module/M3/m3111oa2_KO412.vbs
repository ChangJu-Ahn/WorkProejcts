
Option Explicit		

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0        
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtSoNo1.className = "TD6" 	
	frm1.txtFrDt.Text	= StartDate
	frm1.txtToDt.Text	= EndDate
End Sub
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                       
    
    call FormatDateField(frm1.txtFrDt)	
    call FormatDateField(frm1.txtToDt)	
    call LockobjectField(frm1.txtFrDt,"O")
	call LockobjectField(frm1.txtToDt,"O")	
    
    Call InitVariables                        
    Call SetDefaultVal 
    Call CookiePage(0)   
    Call SetToolbar("1000000000001111")		
    frm1.txtBpCd.focus 
	Set gActiveElement = document.activeElement
End Sub

'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------

Function CookiePage(Byval Kubun)

	Dim strPo, arrVal, strPoDt
	Dim IntRetCD


	If Kubun = 0 Then

		strPo = ReadCookie("PoNo")
		strPoDt = ReadCookie("PoDt")
		
		If strPo = "" then Exit Function

		frm1.txtPoNo.value = strPo
		frm1.txtFrDt.text = strPoDt
		frm1.txtToDt.text = strPoDt
		
		WriteCookie "PoNo" , ""
		WriteCookie "PoDt" , ""
		
		Call btnPreview()
	End If
End Function	

'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : OpenBpCd PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공급처"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "공급처"						
    arrHeader(1) = "공급처명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If		
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim PoFlg
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "발주번호"						
	arrParam(1) = "M_PUR_ORD_HDR,B_Biz_Partner,B_PUR_GRP"					
	arrParam(2) = Trim(frm1.txtPoNo.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)	
	
	If 	frm1.rdoPoflg1.checked = true then 
		PoFlg	= "Y"
	Else
		PoFlg	= "N"	
	End if	
	
'	arrParam(4) = "M_PUR_ORD_HDR.IMPORT_FLG = '" & PoFlg & "' AND M_PUR_ORD_HDR.BP_CD = B_Biz_Partner.BP_CD AND M_PUR_ORD_HDR.PUR_GRP = B_PUR_GRP.PUR_GRP"
	arrParam(4) = "M_PUR_ORD_HDR.release_flg	= 'Y' AND M_PUR_ORD_HDR.BP_CD = B_Biz_Partner.BP_CD AND M_PUR_ORD_HDR.PUR_GRP = B_PUR_GRP.PUR_GRP"
	arrParam(5) = "발주번호"						
	
    arrField(0) = "ED15" & Parent.gColSep &"M_PUR_ORD_HDR.PO_NO"							
    arrField(1) = "ED7" & Parent.gColSep &"M_PUR_ORD_HDR.BP_CD"							
    arrField(2) = "ED15" & Parent.gColSep &"B_Biz_Partner.BP_NM"
    arrField(3) = "DD10" & Parent.gColSep & " M_PUR_ORD_HDR.PO_DT "
    arrField(4) = "F212" & Parent.gColSep & " M_PUR_ORD_HDR.TOT_PO_DOC_AMT "
    
    if Trim(frm1.txtBpCd.Value)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.BP_CD='" & Trim(frm1.txtBpCd.Value) & "'"    
	End if
	if Trim(frm1.txtPurGrpCd.Value)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PUR_GRP='" & Trim(frm1.txtPurGrpCd.Value) & "'"    
	End if
	
	if Trim(frm1.txtFrDt.Text)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT >= '" &UNIConvDate(Trim(frm1.txtFrDt.Text)) & "'"  
	Else
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT >='1900-01-01'"    
	End if
	
	if Trim(frm1.txtToDt.Text)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT <= '" & UNIConvDate(Trim(frm1.txtToDt.Text)) & "'" 
	Else
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT <='2999-12-31'" 
	End if
	'--arrParam(4) = arrParam(4) & " order by po_no " 
    arrField(5) = "ED6" & Parent.gColSep & "M_PUR_ORD_HDR.PO_CUR"    
    arrField(6) = "ED10" & Parent.gColSep & "B_PUR_GRP.PUR_GRP_NM"
    
    
    
    arrHeader(0) = "발주번호"						
    arrHeader(1) = "공급처"					
    arrHeader(2) = "공급처명"					
    arrHeader(3) = "발주일"					
    arrHeader(4) = "발주금액"					
    arrHeader(5) = "화폐"					
    arrHeader(6) = "구매그룹"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoNo.focus
		Exit Function
	Else
		frm1.txtPoNo.Value = arrRet(0)		
		frm1.txtPoNo.focus
	End If	
End Function
'------------------------------------------  OpenPoNo1()  -------------------------------------------------
'	Name : OpenPoNo1()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim PoFlg
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "발주번호"						
	arrParam(1) = "M_PUR_ORD_HDR,B_Biz_Partner,B_PUR_GRP"					
	arrParam(2) = Trim(frm1.txtPoNo1.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)	
	
	If 	frm1.rdoPoflg1.checked = true then 
		PoFlg	= "Y"
	Else
		PoFlg	= "N"	
	End if	
	
'	arrParam(4) = "M_PUR_ORD_HDR.IMPORT_FLG = '" & PoFlg & "' AND M_PUR_ORD_HDR.BP_CD = B_Biz_Partner.BP_CD AND M_PUR_ORD_HDR.PUR_GRP = B_PUR_GRP.PUR_GRP"
	arrParam(4) = "M_PUR_ORD_HDR.release_flg	= 'Y' AND M_PUR_ORD_HDR.BP_CD = B_Biz_Partner.BP_CD AND M_PUR_ORD_HDR.PUR_GRP = B_PUR_GRP.PUR_GRP"
	arrParam(5) = "발주번호"						
	
    arrField(0) = "ED15" & Parent.gColSep &"M_PUR_ORD_HDR.PO_NO"							
    arrField(1) = "ED7" & Parent.gColSep &"M_PUR_ORD_HDR.BP_CD"							
    arrField(2) = "ED15" & Parent.gColSep &"B_Biz_Partner.BP_NM"
    arrField(3) = "DD10" & Parent.gColSep & " M_PUR_ORD_HDR.PO_DT "
    arrField(4) = "F212" & Parent.gColSep & " M_PUR_ORD_HDR.TOT_PO_DOC_AMT "
    
    if Trim(frm1.txtBpCd.Value)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.BP_CD='" & Trim(frm1.txtBpCd.Value) & "'"    
	End if
	if Trim(frm1.txtPurGrpCd.Value)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PUR_GRP='" & Trim(frm1.txtPurGrpCd.Value) & "'"    
	End if
	
	if Trim(frm1.txtFrDt.Text)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT >= '" &UNIConvDate(Trim(frm1.txtFrDt.Text)) & "'"  
	Else
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT >='1900-01-01'"    
	End if
	
	if Trim(frm1.txtToDt.Text)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT <= '" & UNIConvDate(Trim(frm1.txtToDt.Text)) & "'" 
	Else
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT <='2999-12-31'" 
	End if
	'--arrParam(4) = arrParam(4) & " order by po_no " 
    arrField(5) = "ED6" & Parent.gColSep & "M_PUR_ORD_HDR.PO_CUR"    
    arrField(6) = "ED10" & Parent.gColSep & "B_PUR_GRP.PUR_GRP_NM"
    
    
    
    arrHeader(0) = "발주번호"						
    arrHeader(1) = "공급처"					
    arrHeader(2) = "공급처명"					
    arrHeader(3) = "발주일"					
    arrHeader(4) = "발주금액"					
    arrHeader(5) = "화폐"					
    arrHeader(6) = "구매그룹"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoNo1.focus
		Exit Function
	Else
		frm1.txtPoNo1.Value = arrRet(0)		
		frm1.txtPoNo1.focus
	End If	
End Function
'------------------------------------------  OpenPurGrpCd()  -------------------------------------------------
'	Name : OpenPurGrpCd()
'	Description : OpenPurGrpCd PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPurGrpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG='Y'"			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
	End If	
End Function 


'==========================================================================================
'   Event Name : txtFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End If
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
	dim var1,var2,var3,var4,var5, var6, var7
    Dim ObjName
    	
'    If Not chkField(Document, "1") Then									
'       Exit Function
'    End If

	   
    IF ChkKeyField() = False Then 
		Exit Function
    End if
    
    'IF chkLength() = False Then 
	'	Exit Function
	'End if

	with frm1
	    If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And trim(.txtFrDt.text) <> "" And trim(.txtToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","발주일", "X")			
			Exit Function
		End if   
	End with
	
	On Error Resume Next                    
	
	lngPos = 0
	
	If Ucase(frm1.txtBpCd.value) = "" Then
		var1 = "%"
	Else
		var1= Ucase(frm1.txtBpCd.value)
	End If
	
	If Ucase(frm1.txtPurGrpCd.value) = "" Then
		var2 = "%"
	Else
		var2 = Ucase(frm1.txtPurGrpCd.value)
	End If
	
	If Ucase(frm1.txtFrDt.text) = "" Then
		var3 = "1900-01-01"
	Else
		var3 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType)'uniCdate(frm1.txtFrDt.text)
	End If
	
	If Ucase(frm1.txtToDt.text) = "" Then
		var4 = "2999-12-31"
	Else
		var4 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
	End If
	
	var5 = Ucase(frm1.txtPoNo.value)
	
	If Ucase(frm1.txtPoNo1.value) = "" Then
		var6 = "ZZZZZ"
	Else
		var6 = Ucase(frm1.txtPoNo1.value)
	End If

	
	If frm1.rdoPoFlg1.checked = True then		
		var7 = "Y"
	else		
		var7 = "N"
	End if
	   		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
	strUrl = strUrl & "po_no|"			& var5
	strUrl = strUrl & "|po_no1|"		& var6
	strUrl = strUrl & "|bp_cd|"			& var1
	strUrl = strUrl & "|pur_grp|"		& var2
	strUrl = strUrl & "|fr_dt|"			& var3
	strUrl = strUrl & "|to_dt|"			& var4
	strUrl = strUrl & "|Gb_fg|"			& var7

    If lgBACd<>"" Then
        strUrl = strUrl & "|FR_BIZ_AREA|" & lgBACd 
        strUrl = strUrl & "|TO_BIZ_AREA|" & lgBACd 
    Else
        strUrl = strUrl & "|FR_BIZ_AREA|" & "" 
        strUrl = strUrl & "|TO_BIZ_AREA|" & "ZZZZZZZZZZ" 
    End If

    If lgPGCd<>"" Then
        strUrl = strUrl & "|FR_PUR_GRP|" & lgPGCd 
        strUrl = strUrl & "|TO_PUR_GRP|" & lgPGCd 
    Else
        strUrl = strUrl & "|FR_PUR_GRP|" & "" 
        strUrl = strUrl & "|TO_PUR_GRP|" & "ZZZZZZZZZZ" 
    End If

    If lgPOCd<>"" Then
        strUrl = strUrl & "|FR_PUR_ORG|" & lgPOCd 
        strUrl = strUrl & "|TO_PUR_ORG|" & lgPOCd 
    Else
        strUrl = strUrl & "|FR_PUR_ORG|" & "" 
        strUrl = strUrl & "|TO_PUR_ORG|" & "ZZZZZZZZZZ" 
    End If
	
    If lgPLCd<>"" Then
        strUrl = strUrl & "|FR_PLANT_CD|" & lgPLCd 
        strUrl = strUrl & "|TO_PLANT_CD|" & lgPLCd 
    Else
        strUrl = strUrl & "|FR_PLANT_CD|" & "" 
        strUrl = strUrl & "|TO_PLANT_CD|" & "ZZZZZZZZZZ" 
    End If
			
	ObjName = AskEBDocumentName("m3111oa2_KO412","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)		
			
	Call BtnDisabled(0)	
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : btnPreview
' Function Desc : 
'========================================================================================
Sub btnPreview() 
	Err.Clear                                                       
    
    Dim strVal
    dim var1,var2,var3,var4,var5, var6,var7
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim ObjName
    
    IF ChkKeyField() = False Then 
		Exit Sub
    End if
        
    'IF chkLength() = False Then 
	'	Exit Sub
	'End if
    
	with frm1
	     If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And trim(.txtFrDt.text) <> "" And trim(.txtToDt.text) <> "" then
			Call DisplayMsgBox("17a003", "X","발주일", "X")			
			Exit sub
		End if   
	End with
		
	If Ucase(frm1.txtBpCd.value) = "" Then
		var1 = "%"
	Else
		var1= Ucase(frm1.txtBpCd.value)
	End If
	
	If Ucase(frm1.txtPurGrpCd.value) = "" Then
		var2 = "%"
	Else
		var2 = Ucase(frm1.txtPurGrpCd.value)
	End If
	
	If Ucase(frm1.txtFrDt.text) = "" Then
		var3 = ("1900-01-01")
	Else
		var3 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	End If
	
	If Ucase(frm1.txtToDt.text) = "" Then
		var4 = ("2999-12-31")
	Else
		var4 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtToDt.text)
	End If
	
	var5 = Ucase(frm1.txtPoNo.value)
	
	If Ucase(frm1.txtPoNo1.value) = "" Then
		var6 = "ZZZZZ"
	Else
		var6 = Ucase(frm1.txtPoNo1.value)
	End If

	
	If frm1.rdoPoFlg1.checked = True then		
		var7 = "Y"
	else		
		var7 = "N"
	End if
	
	strUrl = strUrl & "po_no|"			& var5
	strUrl = strUrl & "|po_no1|"		& var6 	 	
	strUrl = strUrl & "|bp_cd|"			& var1
	strUrl = strUrl & "|pur_grp|"		& var2
	strUrl = strUrl & "|fr_dt|"			& var3
	strUrl = strUrl & "|to_dt|"			& var4
	strUrl = strUrl & "|Gb_fg|"			& var7

    If lgBACd<>"" Then
        strUrl = strUrl & "|FR_BIZ_AREA|" & lgBACd 
        strUrl = strUrl & "|TO_BIZ_AREA|" & lgBACd 
    Else
        strUrl = strUrl & "|FR_BIZ_AREA|" & "" 
        strUrl = strUrl & "|TO_BIZ_AREA|" & "ZZZZZZZZZZ" 
    End If

    If lgPGCd<>"" Then
        strUrl = strUrl & "|FR_PUR_GRP|" & lgPGCd 
        strUrl = strUrl & "|TO_PUR_GRP|" & lgPGCd 
    Else
        strUrl = strUrl & "|FR_PUR_GRP|" & "" 
        strUrl = strUrl & "|TO_PUR_GRP|" & "ZZZZZZZZZZ" 
    End If

    If lgPOCd<>"" Then
        strUrl = strUrl & "|FR_PUR_ORG|" & lgPOCd 
        strUrl = strUrl & "|TO_PUR_ORG|" & lgPOCd 
    Else
        strUrl = strUrl & "|FR_PUR_ORG|" & "" 
        strUrl = strUrl & "|TO_PUR_ORG|" & "ZZZZZZZZZZ" 
    End If
	
    If lgPLCd<>"" Then
        strUrl = strUrl & "|FR_PLANT_CD|" & lgPLCd 
        strUrl = strUrl & "|TO_PLANT_CD|" & lgPLCd 
    Else
        strUrl = strUrl & "|FR_PLANT_CD|" & "" 
        strUrl = strUrl & "|TO_PLANT_CD|" & "ZZZZZZZZZZ" 
    End If
	
	ObjName = AskEBDocumentName("m3111oa2_KO412","ebr")
	Call FncEBRPreview(ObjName, strUrl)
			
	Call BtnDisabled(0)	
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)   
	Set gActiveElement = document.activeElement
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
	
	If trim(frm1.txtBpCd.value) <> "" Then
		strWhere = " BP_CD = '" & FilterVar(frm1.txtBpCd.value, "","SNM") & "' "
	
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","공급처","X")
			frm1.txtBpCd.focus 
			frm1.txtBpNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtBpNm.value = strDataNm(0)
	End If
	
	If trim(frm1.txtPurGrpCd.value) <> "" Then
		strWhere = " PUR_GRP = '" & FilterVar(frm1.txtPurGrpCd.value, "","SNM") & "' "
	
		Call CommonQueryRs(" PUR_GRP_NM "," B_PUR_GRP ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","구매그룹","X")
			frm1.txtPurGrpCd.focus 
			frm1.txtPurGrpNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPurGrpNm.value = strDataNm(0)
	End If

	If trim(frm1.txtPoNo.value) <> "" Then
		strWhere = " PO_NO = '" & FilterVar(frm1.txtPoNo.value, "","SNM") & "' "
	
		Call CommonQueryRs(" PO_NO "," M_PUR_ORD_HDR ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","발주번호","X")
			frm1.txtPoNo.focus 
			ChkKeyField = False
			Exit Function
		End If	
	End If
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
'	Name : ChkKeyField()
'	Description : 
'=========================================================================================================
Function chkLength()
	chkLength = true
	If Not chkFieldLengthByCell(frm1.txtPoNo, "A",1) Then		
	   chkLength = false
       Exit Function
    End If
    
    If Not chkFieldLengthByCell(frm1.txtBpCd, "A",1) Then	
	   chkLength = false	
       Exit Function
    End If
    
    If Not chkFieldLengthByCell(frm1.txtPurGrpCd, "A",1) Then		
	   chkLength = false
       Exit Function
    End If
end function
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################

