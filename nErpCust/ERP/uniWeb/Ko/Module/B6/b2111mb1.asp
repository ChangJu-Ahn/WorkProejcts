<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                    '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()
    
    Dim txtBizAreaCd
    Dim lgarrData
	Dim txtPrevNextFlg
	Dim iStatusCodeOfPrevNext
		
    Const biz_area_cd = 0
    Const biz_area_nm = 1
    Const biz_area_full_nm = 2
    Const biz_area_eng_nm = 3
    Const own_rgst_no = 4
    Const repre_nm = 5
    Const tax_office_cd = 6
    Const tax_office_nm = 7
    Const ind_type = 8
    Const ind_type_nm = 9
    Const ind_class = 10
    Const ind_class_nm = 11
    Const fax_no = 12
    Const tel_no = 13
'    Const tax_flag = 14
    Const zip_code = 14
    Const report_biz_area_cd = 15
    Const report_biz_area_nm = 16
    Const addr1 = 17
    Const addr2 = 18
    Const eng1_addr = 19
    Const eng2_addr = 20
    Const eng3_addr = 21
    
    Dim lgOpModeCRUD
                             
	ReDim lgarrData(eng3_addr)
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    'lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	txtBizAreaCd      = Request("txtBizAreaCd")
	txtPrevNextFlg = Request("PrevNextFlg")


    'Single
'    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
'   Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
'    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
	Dim PB6SA05Data
	
'    Response.Write  "<<" & gStrGlobalCollection & ">>"
	Set PB6SA05Data = server.CreateObject ("PB6SA05.cBLkUpBizAreaSvr") 
	
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " with parent.frm1" & vbCr
	Response.Write " .txtBizAreaNm.value       = """ & ConvSPChars(lgarrData(biz_area_nm)) & """		" & vbCr
	Response.Write "End with					" & vbcr
	Response.Write "</Script>                   " & vbCr  
	lgarrData = PB6SA05Data.B_LOOKUP_BIZ_AREA_SVR(gStrGlobalCollection,txtBizAreaCd,txtPrevNextFlg,iStatusCodeOfPrevNext)
	
	If CheckSYSTEMError(Err,True) = True Then
		Set PB6SA05Data = nothing		
		Exit Sub
    End If

	Set PB6SA05Data = Nothing															'¢Ð: Unload Component

	If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
		Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
	End If

	Set PB6SA05Data = nothing    

'     strTaxFg=  lgarrData(tax_flag)		 			 				 			 		 						 				        		 				 		 
   
   	Response.Write "<Script Language=vbscript>  " & vbCr
   	Response.Write " with parent.frm1" & vbCr
	Response.Write " .txtBizAreaCd.value       = """ & ConvSPChars(lgarrData(biz_area_cd)) & """		" & vbCr
	Response.Write " .txtBizAreaNm.value       = """ & ConvSPChars(lgarrData(biz_area_nm)) & """		" & vbCr
	Response.Write " .txtBizAreaCd_Body.value  = """ & ConvSPChars(lgarrData(biz_area_cd)) & """		" & vbCr
    Response.Write " .txtBizAreaNm_Body.value  = """ & ConvSPChars(lgarrData(biz_area_nm)) & """		" & vbCr
	Response.Write " .txtBizAreaFullNm.value   = """ & ConvSPChars(lgarrData(biz_area_full_nm)) & """		" & vbCr
    Response.Write " .txtBizAreaEngNm.value    = """ & ConvSPChars(lgarrData(biz_area_eng_nm)) & """		" & vbCr
    Response.Write " .txtOwnRgstNo.value       = """ & ConvSPChars(lgarrData(own_rgst_no)) & """		" & vbCr
    Response.Write " .txtRepreNm.value         = """ & ConvSPChars(lgarrData(repre_nm)) & """		" & vbCr
    Response.Write " .txtTaxOfficeCd.value     = """ & ConvSPChars(lgarrData(tax_office_cd)) & """		" & vbCr
    Response.Write " .txtTaxOfficeNm.value     = """ & ConvSPChars(lgarrData(tax_office_nm)) & """		" & vbCr
    Response.Write " .txtInd_Type.value        = """ & ConvSPChars(lgarrData(ind_type)) & """		" & vbCr
    Response.Write " .txtInd_Type_Nm.value     = """ & ConvSPChars(lgarrData(ind_type_nm)) & """		" & vbCr
    Response.Write " .txtInd_Class.value       = """ & ConvSPChars(lgarrData(ind_class)) & """		" & vbCr
    Response.Write " .txtInd_Class_Nm.value    = """ & ConvSPChars(lgarrData(ind_class_nm)) & """		" & vbCr
    Response.Write " .txtFaxNo.value           = """ & ConvSPChars(lgarrData(fax_no)) & """		" & vbCr
    Response.Write " .txtTelNo.value           = """ & ConvSPChars(lgarrData(tel_no)) & """		" & vbCr
    
 '     if strTaxFg = "Y" then
 '   Response.Write " .Rb_TaxFgY.checked        =  true 		" & vbCr
 '     else
 '   Response.Write " .Rb_TaxFgN.checked        =  true  	" & vbCr
 '     end if
        
    Response.Write " .txtZipCode.value         = """ & ConvSPChars(lgarrData(zip_code)) & """		" & vbCr
    Response.Write " .txtReportBizArea.value   = """ & ConvSPChars(lgarrData(report_biz_area_cd)) & """		" & vbCr
    Response.Write " .txtReportBizAreaNm.value = """ & ConvSPChars(lgarrData(report_biz_area_nm)) & """		" & vbCr
    Response.Write " .txtAddr1.value            = """ & ConvSPChars(lgarrData(addr1)) & """		" & vbCr
    Response.Write " .txtAddr2.value            = """ & ConvSPChars(lgarrData(addr2)) & """		" & vbCr
    Response.Write " .txtEng1Addr.value         = """ & ConvSPChars(lgarrData(eng1_addr)) & """		" & vbCr
    Response.Write " .txtEng2Addr.value         = """ & ConvSPChars(lgarrData(eng2_addr)) & """		" & vbCr
    Response.Write " .txtEng3Addr.value         = """ & ConvSPChars(lgarrData(eng3_addr)) & """		" & vbCr
    Response.Write "End with					" & vbcr
    Response.Write "Parent.DbQueryOk			" & vbcr
    Response.Write "</Script>                   " & vbCr
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear  
    
    Dim lgIntFlgMode                                                                      '¢Ð: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '¢Ð : Update
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
   
    Dim PB6SA05Data
    Dim iCommandSent
    
    iCommandSent = "CREATE"

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgarrData(biz_area_cd)        =  Request("txtBizAreaCd_Body")
    lgarrData(biz_area_nm)        =  Request("txtBizAreaNm_Body")
    lgarrData(biz_area_full_nm)   =  Request("txtBizAreaFullNm")
    lgarrData(biz_area_eng_nm)    =  Request("txtBizAreaEngNm")
    lgarrData(own_rgst_no)        =  Request("txtOwnRgstNo")
    lgarrData(repre_nm)           =  Request("txtRepreNm")
    lgarrData(tax_office_cd)      =  Request("txtTaxOfficeCd")
    lgarrData(tax_office_nm)      =  Request("txtTaxOfficeNm")
    lgarrData(ind_type)           =  Request("txtInd_Type")
    lgarrData(ind_type_nm)        =  Request("txtInd_Type_Nm")
    lgarrData(ind_class)          =  Request("txtInd_class")
    lgarrData(ind_class_nm)       =  Request("txtInd_class_Nm")
    lgarrData(fax_no)             =  Request("txtFaxNo")
    lgarrData(tel_no)             =  Request("txtTelNo")
    'lgarrData(tax_flag)           =  Request("Radio1")
    lgarrData(zip_code)           =  Request("txtZipCode")
    lgarrData(report_biz_area_cd) =  Request("txtReportBizArea")
    lgarrData(report_biz_area_nm) =  Request("txtReportBizAreaNm")
    lgarrData(addr1)               =  Request("txtAddr1")
    lgarrData(addr2)               =  Request("txtAddr2")
    lgarrData(eng1_addr)           =  Request("txtEng1Addr")
    lgarrData(eng2_addr)           =  Request("txtEng2Addr")
    lgarrData(eng3_addr)           =  Request("txtEng3Addr")
	
    Set PB6SA05Data = server.CreateObject ("PB6SA05.cbBMngBizAreaSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
     
    Call PB6SA05Data.B_MANAGE_BIZ_AREA_SVR(gStrGlobalCollection, iCommandSent, lgarrData )
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PB6SA05Data = nothing
		Exit Sub	
    End If
	
    Set PB6SA05Data = nothing
    	
   
    Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    Dim PB6SA05Data 
    Dim iCommandSent
    
    iCommandSent = "UPDATE"
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
	lgarrData(biz_area_cd)        =  Request("txtBizAreaCd_Body")
    lgarrData(biz_area_nm)        =  Request("txtBizAreaNm_Body")
    lgarrData(biz_area_full_nm)   =  Request("txtBizAreaFullNm")
    lgarrData(biz_area_eng_nm)    =  Request("txtBizAreaEngNm")
    lgarrData(own_rgst_no)        =  Request("txtOwnRgstNo")
    lgarrData(repre_nm)           =  Request("txtRepreNm")
    lgarrData(tax_office_cd)      =  Request("txtTaxOfficeCd")
    lgarrData(tax_office_nm)      =  Request("txtTaxOfficeNm")
    lgarrData(ind_type)           =  Request("txtInd_Type")
    lgarrData(ind_type_nm)        =  Request("txtInd_Type_Nm")
    lgarrData(ind_class)          =  Request("txtInd_class")
    lgarrData(ind_class_nm)       =  Request("txtInd_class_Nm")
    lgarrData(fax_no)             =  Request("txtFaxNo")
    lgarrData(tel_no)             =  Request("txtTelNo")
    'lgarrData(tax_flag)           =  Request("Radio1")
    lgarrData(zip_code)           =  Request("txtZipCode")
    lgarrData(report_biz_area_cd) =  Request("txtReportBizArea")
    lgarrData(report_biz_area_nm) =  Request("txtReportBizAreaNm")
    lgarrData(addr1)               =  Request("txtAddr1")
    lgarrData(addr2)               =  Request("txtAddr2")
    lgarrData(eng1_addr)           =  Request("txtEng1Addr")
    lgarrData(eng2_addr)           =  Request("txtEng2Addr")
    lgarrData(eng3_addr)           =  Request("txtEng3Addr")
	
    Set PB6SA05Data = server.CreateObject ("PB6SA05.cbBMngBizAreaSvr") 
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PB6SA05Data = nothing
		Exit Sub
    End If
       
    Call PB6SA05Data.B_MANAGE_BIZ_AREA_SVR(gStrGlobalCollection, iCommandSent, lgarrData )
    
     If CheckSYSTEMError(Err,True) = True Then
		Set PB6SA05Data = nothing
		Exit Sub	
    End If
	
    Set PB6SA05Data = nothing
   
    Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    Dim PB6SA05Data
    Dim iCommandSent
    
    iCommandSent = "DELETE"

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    Set PB6SA05Data = server.CreateObject ("PB6SA05.cbBMngBizAreaSvr") 
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If
          
    Call PB6SA05Data.B_MANAGE_BIZ_AREA_SVR(gStrGlobalCollection, iCommandSent, ,txtBizAreaCd)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PB6SA05Data = nothing
		Exit Sub
    End If
	
    Set PB6SA05Data = nothing
    
    Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
    Response.Write "</Script>                   " & vbCr

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>