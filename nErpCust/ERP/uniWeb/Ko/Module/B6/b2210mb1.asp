<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%

    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status

    DIM txtCo_Cd,lgOpModeCRUD
    
    Call HideStatusWnd                                                              '☜: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                          '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                          '☜: Read Operation Mode (CRUD)
	txtCo_Cd          = Request("txtCo_Cd")
    
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                       '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                 '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)               '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Const CompanyCoCd				= 0
	Const CompanyCoNm				= 1
	Const CompanyCoFullNm			= 2
    Const CompanyCoEngNm			= 3
	Const CompanyOwnRgstNo			= 4
	Const CompanyRepreNm			= 5
    Const CompanyRepreRgstNo		= 6
	Const CompanyFaxNo				= 7
	Const CompanyIndClass			= 8
	Const IndClassMinorMinorNm		= 9
	Const CompanyTelNo				= 10
	Const CompanyIndType			= 11
	Const IndTypeMinorMinorNm		= 12
	Const CompanyCountryCd			= 13
	Const CompanyFiscCnt			= 14
	Const CompanyLocCur				= 15
	Const CompanyFiscStartDt		= 16
	Const CompanyFoundationDt		= 17
	Const CompanyFiscEndDt			= 18
	Const FirstYearMonthString		= 19
	Const LastYearMonthString		= 20
	Const CompanyZipCode			= 21
	Const CompanyXchRateFg			= 22
	Const CompanyTaxCfg				= 23
	Const CompanyLocCurCfg			= 24
	Const CompanyAddr				= 25
	Const CompanyEngAddr			= 26
	Const CompanyTransStartDt		= 27
	Const CompanyCurOrgChangeId		= 28
	Const CompanyOpenAcctFg			= 29
	Const CompanycboQmdpalignopt	= 30
	Const CompanycboImdpalignopt	= 31
	Const CompanyXchErrorUseFg		= 32
	Const CompanyInvPostingFg		= 33
		
	'------ Developer Coding part (End   ) ------------------------------------------------------------------     
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                        '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                        '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                        '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Dim PB6S1000		
	Dim iarrData
	Dim indx
	Dim strYear, strMonth, strDay
	redim iarrData(CompanyXchErrorUseFg)

    On Error Resume Next
    Err.Clear

    Set PB6S1000 = server.CreateObject ("PB6SA10.cBLkUpCompanySvr")

	If CheckSYSTEMError(Err,True) = True Then
		Set PB6S1010 = nothing
		Exit Sub
    End If
    iarrData = PB6S1000.B_LOOK_UP_COMPANY_SVR(gStrGlobalCollection, txtCo_Cd)

    If CheckSYSTEMError(Err,True) = True Then
		Set PB6S1000 = nothing
		Exit Sub
    End If
    
    Set PB6S1000 = nothing 
    
	If isarray(iarrData) = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Exit Sub	
	End if

   	'iarrData(FirstYearMonthString) = UNIConvDateToYYYYMMDD(iarrData(FirstYearMonthString),gAPDateFormat,gServerDateType) 
   	'iarrData(LastYearMonthString) = UNIConvDateToYYYYMMDD(iarrData(LastYearMonthString),gAPDateFormat,gServerDateType)

	Call ExtractDateFrom(iarrData(FirstYearMonthString), "YYYYMM", "", strYear, strMonth, strDay)
	iarrData(FirstYearMonthString) = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
	'Call ExtractDateFrom(iarrData(LastYearMonthString), "YYYYMM", "", strYear, strMonth, strDay)
	'iarrData(LastYearMonthString) = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")

   	Response.Write "<Script Language=vbscript>  " & vbCr
   	Response.Write " with parent.frm1" & vbCr
   	Response.Write " .txtCO_CD_Body.value       = """ & UCase(Trim(ConvSPChars(iarrData(CompanyCoCd)))) & """						" & vbCr
   	Response.Write " .txtco_nm.value			= """ & Trim(ConvSPChars(iarrData(CompanyCoNm))) & """								" & vbCr											'☆: Company Name
   	Response.Write " .txtCo_FullNm.value		= """ & Trim(ConvSPChars(iarrData(CompanyCoFullNm))) & """							" & vbCr											'☆: Company Name
   	Response.Write " .txtCO_FULL_NM_Body.value	= """ & Trim(ConvSPChars(iarrData(CompanyCoFullNm))) & """							" & vbCr										'☆: Company FullName
   	Response.Write " .txteng_nm.value			= """ & Trim(ConvSPChars(iarrData(CompanyCoEngNm))) & """							" & vbCr										'☆: Plant Name
   	Response.Write " .txtrepre_rgst_no.value	= """ & Trim(ConvSPChars(iarrData(CompanyRepreRgstNo))) & """						" & vbCr									'☆: Currency Code
   	Response.Write " .txtrepre_nm.value			= """ & Trim(ConvSPChars(iarrData(CompanyRepreNm))) & """							" & vbCr										'☆: Currency Name
   	Response.Write " .txtown_rgst_no.value		= """ & Trim(ConvSPChars(iarrData(CompanyOwnRgstNo))) & """							" & vbCr
   	Response.Write " .txtfax_no.value			= """ & Trim(ConvSPChars(iarrData(CompanyFaxNo))) & """								" & vbCr
   	Response.Write " .txtInd_class.value		= """ & Trim(ConvSPChars(iarrData(CompanyIndClass))) & """							" & vbCr
   	Response.Write " .txtInd_class_nm.value		= """ & Trim(ConvSPChars(iarrData(IndClassMinorMinorNm))) & """						" & vbCr
   	Response.Write " .txttel_no.value			= """ & Trim(ConvSPChars(iarrData(CompanyTelNo))) & """								" & vbCr
   	Response.Write " .txtind_type.value			= """ & Trim(ConvSPChars(iarrData(CompanyIndType))) & """							" & vbCr
   	Response.Write " .txtind_type_nm.value		= """ & Trim(ConvSPChars(iarrData(IndTypeMinorMinorNm))) & """						" & vbCr
   	Response.Write " .txtCountryCd.value		= """ & Trim(ConvSPChars(iarrData(CompanyCountryCd))) & """							" & vbCr
   	Response.Write " .txtfisc_cnt.value			= """ & UNINumClientFormat(iarrData(CompanyFiscCnt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
   	Response.Write " .txtloc_cur.value			= """ & Trim(ConvSPChars(iarrData(CompanyLocCur))) & """							" & vbCr
   	Response.Write " .txtfisc_start_dt.text		= """ & UNIDateClientFormat(iarrData(CompanyFiscStartDt)) & """						" & vbCr
   	Response.Write " .txtfoundation_dt.text		= """ & UNIDateClientFormat(iarrData(CompanyFoundationDt)) & """					" & vbCr
   	Response.Write " .txtfisc_end_dt.text		= """ & UNIDateClientFormat(iarrData(CompanyFiscEndDt)) & """						" & vbCr
   	Response.Write " .txtFirstDeprYyyymm.text	= """ & UNIMonthClientFormat(iarrData(FirstYearMonthString)) & """					" & vbCr
   	'Response.Write " .txtLastDeprYyyymm.text	= """ & UNIMonthClientFormat(iarrData(LastYearMonthString)) & """	" & vbCr
   	Response.Write " .txtzip_code.value			= """ & Trim(ConvSPChars(iarrData(CompanyZipCode))) & """							" & vbCr
   	Response.Write " .cboxch_rate_fg.value		= """ & Trim(iarrData(CompanyXchRateFg)) & """										" & vbCr
   	Response.Write " .cboTaxPolicy.value  		= """ & Trim(iarrData(CompanyTaxCfg)) & """											" & vbCr
   	Response.Write " .cboCurPolicy.value     	= """ & Trim(iarrData(CompanyLocCurCfg)) & """										" & vbCr
   	Response.Write " .cboQmdpalignopt.value     = """ & Trim(iarrData(CompanycboQmdpalignopt)) & """								" & vbCr
   	Response.Write " .cboImdpalignopt.value     = """ & Trim(iarrData(CompanycboImdpalignopt)) & """								" & vbCr
   	Response.Write " .txtaddr.value				= """ & Trim(ConvSPChars(iarrData(CompanyAddr))) & """								" & vbCr
   	Response.Write " .txteng_addr.value			= """ & Trim(ConvSPChars(iarrData(CompanyEngAddr))) & """							" & vbCr
   	Response.Write " .txtTransStartDt.text		= """ & UNIDateClientFormat(iarrData(CompanyTransStartDt)) &					 """" & vbCr
   	Response.Write " .txtCurOrgChangeID.Value	= """ & Trim(ConvSPChars(iarrData(CompanyCurOrgChangeId))) & """					" & vbCr
   	Response.Write " .cboOpenAcctFg.Value		= """ & Trim(iarrData(CompanyOpenAcctFg)) & """										" & vbCr
   	Response.Write " .cboXchErrorUseFg.Value	= """ & Trim(iarrData(CompanyXchErrorUseFg)) & """									" & vbCr
   	Response.Write " .cboInvPostingFg.Value	    = """ & Trim(iarrData(CompanyInvPostingFg)) & """									" & vbCr   	

    Response.Write "End with					" & vbcr
    Response.Write "Parent.DbQueryOk			" & vbcr
    Response.Write "</Script>                   " & vbCr
    
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()

	Dim PB6S1000
	Dim iarrData
	Dim lgIntFlgMode
	
    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear        

	ReDim iarrData(CompanyInvPostingFg)
    
	iarrData(CompanyCoNm)				= Request("txtco_nm")						'법인약명 
	iarrData(CompanyCoFullNm)			= Request("txtCO_FULL_NM_Body")				'법인명 
	iarrData(CompanyCoEngNm)			= Request("txteng_nm")						'법인영문명   
	iarrData(CompanyOwnRgstNo)			= Request("txtown_rgst_no")
	iarrData(CompanyRepreNm)			= Request("txtrepre_nm")
	iarrData(CompanyRepreRgstNo)		= Request("txtrepre_rgst_no")				'사업자등록번호 
	iarrData(CompanyFaxNo)				= Request("txtfax_no")
	iarrData(CompanyIndClass)			= Request("txtind_class") 
	iarrData(CompanyTelNo)				= Request("txttel_no")
	iarrData(CompanyIndType)			= Request("txtind_type")  
	iarrData(CompanyCountryCd)			= Request("txtCountryCd")
	iarrData(CompanyFiscCnt)			= UNIConvNum(Request("txtfisc_cnt"),0)
	iarrData(CompanyLocCur)				= Request("txtloc_cur")
	iarrData(CompanyFiscStartDt)		= UNIConvDate(Request("txtfisc_start_dt"))
	iarrData(CompanyFoundationDt)		= UNIConvDate(Request("txtfoundation_dt"))
	iarrData(CompanyFiscEndDt)			= UNIConvDate(Request("txtfisc_end_dt"))
	iarrData(FirstYearMonthString)		= Request("hFirstDeprYyyymm")
	'iarrData(LastYearMonthString)		= Request("hLastDeprYyyymm")
	iarrData(CompanyZipCode)			= Request("txtzip_code")
	iarrData(CompanyXchRateFg)			= Request("cboxch_rate_fg")
	iarrData(CompanyTaxCfg)				= Request("cboTaxPolicy") 
	iarrData(CompanyLocCurCfg)			= Request("cboCurPolicy") 
	iarrData(CompanyAddr)				= Request("txtaddr")
	iarrData(CompanyEngAddr)			= Request("txteng_addr")	
	iarrData(CompanyTransStartDt)		= UNIConvDate(Request("txttrans_st_dt"))
	iarrData(CompanyCurOrgChangeId)		= Request("txtCurOrgChangeID")
	iarrData(CompanyCoCd)				= txtCo_Cd 									'법인코드 
	iarrData(CompanyOpenAcctFg)			= Request("cboOpenAcctFg") 					'미결관리구분 
	iarrData(CompanycboQmdpalignopt)	= Request("cboQmdpalignopt") 				'조회용소수점자리수 
	iarrData(CompanycboImdpalignopt)	= Request("cboImdpalignopt") 				'입력용소수점자리수 
	iarrData(CompanyXchErrorUseFg)		= Request("cboXchErrorUseFg") 				'환율재계산불가 
	iarrData(CompanyInvPostingFg)		= Request("cboInvPostingFg") 				'재고포스팅방법	

    Set PB6S1000 = server.CreateObject ("PB6SA10.cBMngCompanySvr")   

    If CheckSYSTEMError(Err,True) = True Then
		Set PB6S1010 = Nothing
		Exit Sub
    End If

    lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: Read Operayion Mode (CREATE, UPDATE)
     
    Select Case lgIntFlgMode
		Case  OPMD_CMODE                                                            '☜ : Create
			  Call PB6S1000.B_MANAGE_COMPANY_SVR(gStrGlobalCollection,"CREATE",iarrData)
        Case  OPMD_UMODE															'☜ : Update
			  Call PB6S1000.B_MANAGE_COMPANY_SVR(gStrGlobalCollection,"UPDATE",iarrData)
    End Select

    If CheckSYSTEMError(Err,True) = True Then
		Set PB6S1000 = Nothing
		Exit Sub	
    End If

    Set PB6S1000 = Nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    
End Sub   

'============================================================================================================
' Name : SubBizDelete
' Desc : DELETE Data 
'============================================================================================================
SUB SubBizDelete()
	Dim PB6S1010
	Dim iarrData

    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear        

   	Redim iarrData(CompanyInvPostingFg)

	iarrData(CompanyCoCd) = Request("txtCO_CD")	 									'법인코드 

    Set PB6S1010 = server.CreateObject ("PB6SA10.cBManageCompanySvr")    

    If CheckSYSTEMError(Err, True) = True Then					
		Set PB6S1010 = Nothing
		Exit Sub
    End If    

    Call PB6S1010.B_MANAGE_COMPANY_SVR(gStrGlobalCollection,"DELETE",iarrData)

    If CheckSYSTEMError(Err,True) = True Then
		Set PB6S1010 = Nothing
		Exit Sub
    End If

    Set PB6S1010 = Nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub
%>


