<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Call HideStatusWnd

    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim lgErrorStatus
    Dim lgErrorPos
    Dim lgOpModeCRUD
    Dim txtVatNo
    Dim lgCurrency
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Trim(Request("txtMode"))                                           '☜: Read Operation Mode (CRUD)
	txtVatNo          = Trim(Request("txtVatNo"))
  
    '///////////상수 선언..
	Const EA_vat_no             = 0
    Const EA_issued_dt          = 1
    Const EA_own_rgst_no        = 2
    Const EA_ref_no             = 3
    Const EA_io_fg              = 4
    Const EA_vat_type           = 5
    Const EA_vat_type_nm        = 6
    Const EA_vat_rate           = 7
    Const EA_doc_cur            = 8
    Const EA_xch_rate           = 9
    Const EA_net_amt            = 10
    Const EA_net_loc_amt        = 11
    Const EA_vat_amt            = 12
    Const EA_vat_loc_amt        = 13
    Const EA_made_vat_fg        = 14
    Const EA_conf_fg            = 15
    Const EA_gl_no              = 16
    Const EA_item_seq           = 17
    Const EA_tmp_gl_no          = 18
    Const EA_tmp_item_seq       = 19
    Const EA_vat_desc           = 20
    Const EA_insrt_user_id      = 21
    Const EA_insrt_dt           = 22
    Const EA_updt_user_id       = 23
    Const EA_updt_dt            = 24
    Const EA_report_biz_area_cd = 25
    Const EA_report_biz_area_nm = 26
    Const EA_bp_cd              = 27
    Const EA_acct_cd            = 28
    Const EA_ar_no              = 29
    Const EA_ap_no              = 30
    Const Ea_biz_area_cd        = 31
	Const Ea_biz_area_nm        = 32
	Const Ea_bp_nm              = 33
	Const Ea_acct_nm            = 34
 
	Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)       
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
   End Select

'============================================================================================================
' Name : SubBizQuerySingle
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next  
    Err.Clear                                                                        '☜: Clear Error status

	Dim iPAVG070											'조회/입력/수정용 ComProxy Dll사용변수 
	Dim iArrData

'	lgStrPrevKey = Request("lgStrPrevKey")
	
	Set iPAVG070 = Server.CreateObject("PAVG070.cALkUpVatSvr")
  
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
        
    iArrData= iPAVG070.A_LOOKUP_VAT_SVR(gStrGlobalCollection, txtVatNo)

    If CheckSYSTEMError(Err,True) = True Then
		Set iPAVG070	 = nothing		
		Exit Sub
    End If
    
    Set iPAVG070 = Nothing 
    
	If IsArray(iArrData) = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Exit Sub	
	End If
	
	lgCurrency = ConvSPChars(iArrData(EA_doc_cur))
	
	Response.Write "<Script Language=vbscript>																											 " & vbCr
   	Response.Write " with parent																														 " & vbCr
   	Response.Write ".frm1.txtVatNo.value		   = """ & ConvSPChars(iArrData(EA_vat_no))															& """" & vbCr
	Response.Write ".frm1.txtVatNo1.value	       = """ & ConvSPChars(iArrData(EA_vat_no))															& """" & vbCr
	Response.Write ".frm1.txtOwnRgst.value	       = """ & ConvSPChars(iArrData(EA_own_rgst_no))													& """" & vbCr
	Response.Write ".frm1.txtAcctCd.value	       = """ & ConvSPChars(iarrData(EA_acct_cd))														& """" & vbCr
	Response.Write ".frm1.txtAcctNm.value	       = """ & ConvSPChars(iarrData(EA_acct_nm))														& """" & vbCr
	Response.Write ".frm1.txtAPNo.value		       = """ & ConvSPChars(iarrData(EA_ap_no))															& """" & vbCr
	Response.Write ".frm1.txtARNo.value		       = """ & ConvSPChars(iarrData(EA_ar_no))															& """" & vbCr
	Response.Write ".frm1.txtBizArea.value	       = """ & ConvSPChars(iarrData(EA_biz_area_cd))													& """" & vbCr
	Response.Write ".frm1.txtBizAreaNm.value       = """ & ConvSPChars(iarrData(EA_biz_area_nm))													& """" & vbCr
	Response.Write ".frm1.txtBpCd.value		       = """ & ConvSPChars(iarrData(EA_bp_cd))															& """" & vbCr
	Response.Write ".frm1.txtBpNm.value		       = """ & ConvSPChars(iarrData(EA_bp_nm))															& """" & vbCr
	Response.Write ".frm1.txtDocCur.value	       = """ & ConvSPChars(iarrData(EA_doc_cur))														& """" & vbCr
	Response.Write ".frm1.txtGLNo.value		       = """ & ConvSPChars(iarrData(EA_gl_no))															& """" & vbCr
	Response.Write ".frm1.txtRefNo.value	       = """ & ConvSPChars(iarrData(EA_ref_no))															& """" & vbCr
	Response.Write ".frm1.txtReportBizArea.value   = """ & ConvSPChars(iarrData(EA_report_biz_area_cd))												& """" & vbCr
	Response.Write ".frm1.txtReportBizAreaNm.value = """ & ConvSPChars(iarrData(EA_report_biz_area_nm))												& """" & vbCr
	Response.Write ".frm1.txtTempGLNo.value	       = """ & ConvSPChars(iarrData(EA_tmp_gl_no))														& """" & vbCr
	Response.Write ".frm1.txtVatRate.Text	       = """ & UNIConvNumDBToCompanyByCurrency(iarrData(EA_vat_rate), gCurrency,ggExchRateNo, "X", "X") & """" & vbCr
	Response.Write ".frm1.txtVatType.value	       = """ & ConvSPChars(iarrData(EA_vat_type))														& """" & vbCr
	Response.Write ".frm1.txtVatTypeNm.value       = """ & ConvSPChars(iarrData(EA_vat_type_nm))													& """" & vbCr
	Response.Write ".frm1.txtIssuedDt.text	       = """ & UNIDateClientFormat(ConvSPChars(iArrData(EA_issued_dt)))									& """" & vbCr
	Response.Write ".frm1.cboMadeVatFg.value       = """ & ConvSPChars(iarrData(EA_made_vat_fg))													& """" & vbCr
	Response.Write ".frm1.cboIoFg.value		       = """ & ConvSPChars(iarrData(EA_io_fg))															& """" & vbCr
	Response.Write ".frm1.cboConfFg.value	       = """ & ConvSPChars(iarrData(EA_conf_fg))														& """" & vbCr

	Response.Write ".frm1.txtXchRate.text	       = """ & UNIConvNumDBToCompanyByCurrency(iarrData(EA_xch_rate), lgCurrency,ggExchRateNo, "X", "X")				 & """" & vbCr
	Response.Write ".frm1.txtNetAmt.text	       = """ & UNIConvNumDBToCompanyByCurrency(iarrData(EA_net_amt), lgCurrency,ggAmtOfMoneyNo, "X", "X")				 & """" & vbCr
	Response.Write ".frm1.txtNetLocAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(iarrData(EA_net_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".frm1.txtVatAmt.text	       = """ & UNIConvNumDBToCompanyByCurrency(iarrData(EA_vat_amt), lgCurrency,ggAmtOfMoneyNo, "X", "X")				 & """" & vbCr
	Response.Write ".frm1.txtVatLocAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(iarrData(EA_vat_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".frm1.htxtVatNo.value          = """ & ConvSPChars(iarrData(EA_vat_no))																			 & """" & vbCr
	Response.Write ".DbQueryOk																																			  " & vbCr
	Response.Write "End With																																			  " & vbCr
	Response.Write "</Script>																																			  " & vbCr
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Update
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next  
    Err.Clear	
    
    Dim iPAVG070											'조회/입력/수정용 ComPlus Dll사용변수 
	Dim iUpdData
	
	Const IA_acct_cd            = 0
	Const IA_doc_cur            = 1							
	Const IA_io_fg		        = 2
	Const IA_issued_dt	        = 3
	Const IA_made_vat_fg        = 4
	Const IA_net_amt	        = 5
	Const IA_net_loc_amt        = 6
	Const IA_own_rgst_no        = 7
	Const IA_ref_no		        = 8
	Const IA_vat_amt	        = 9
	Const IA_vat_loc_amt        = 10
	Const IA_vat_no		        = 11
	Const IA_vat_rate	        = 12
	Const IA_vat_type	        = 13
	Const IA_xch_rate	        = 14
	Const IA_bp_cd		        = 15
	Const IA_biz_area_cd	    = 16
	Const IA_report_biz_area_cd	= 17
	
	Set iPAVG070 = Server.CreateObject("PAVG070.cAUpdVatSvr")
  
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

    Redim iUpdData(IA_report_biz_area_cd)

	iUpdData(IA_acct_cd)            = UCase(Trim(Request("txtAcctCd")))
	iUpdData(IA_doc_cur)            = UCase(Trim(Request("txtDocCur")))
	iUpdData(IA_io_fg)              = Trim(Request("cboIoFg"))
	iUpdData(IA_issued_dt)          = UNIConvDate(Request("txtIssuedDt"))
	iUpdData(IA_made_vat_fg)        = UCase(Trim(Request("cboMadeVatFg")))
	iUpdData(IA_net_amt)            = UNIConvNum(Request("txtNetAmt"),0)
	iUpdData(IA_net_loc_amt)        = UNIConvNum(Request("txtNetLocAmt"),0)
	iUpdData(IA_own_rgst_no)        = UCase(Trim(Request("txtOwnRgst")))
	iUpdData(IA_ref_no)             = UCase(Trim(Request("txtRefNo")))
	iUpdData(IA_vat_amt)            = UNIConvNum(Request("txtVatAmt"),0)
	iUpdData(IA_vat_loc_amt)        = UNIConvNum(Request("txtVatLocAmt"),0)
	iUpdData(IA_vat_no)             = Trim(Request("txtVatNo"))
	iUpdData(IA_vat_rate)           = UNIConvNum(Request("txtVatRate"),0)
	iUpdData(IA_vat_type)           = Trim(Request("txtVatType"))
	iUpdData(IA_xch_rate)           = UNIConvNum(Request("txtXchRate"),0)
	iUpdData(IA_bp_cd)	            = UCase(Trim(Request("txtBpCd")))
	iUpdData(IA_biz_area_cd)        = UCase(Trim(Request("txtBizArea")))
	iUpdData(IA_report_biz_area_cd) = UCase(Trim(Request("txtReportBizArea")))

    Call iPAVG070.A_UPDATE_VAT_SVR(gStrGlobalCollection,iUpdData)										'☜: Protect system from crashing

    If CheckSYSTEMError(Err,True) = True Then
        Set iPAVG070 = nothing
        Exit Sub	
    End If

    Set iPAVG070 = Nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    
End Sub
%>

