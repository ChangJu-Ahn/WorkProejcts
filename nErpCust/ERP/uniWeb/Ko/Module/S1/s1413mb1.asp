<%@ LANGUAGE=VBSCript%>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1413MB1
'*  4. Program Name         : 담보등록 
'*  5. Program Desc         : 담보등록 
'*  6. Comproxy List        : PS1G114.dll, PS1G115.dll, PS1G116.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/12/09 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    Dim lgOpModeCRUD
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
	
	'======================
	' for Call Com Agent  : Added by Oh, Sang Eun
	'======================
	Dim pvCB 
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "CHECK"                                                                 '☜: Check	
             Call SubBizCheck()    
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iS1G115
    Dim iWarrentNo
    
    Dim EG1_S_Collateral
    
    Const S046_bp_cd               = 0
    Const S046_bp_nm               = 1
    Const S046_del_type_nm_minor_nm      = 2
    Const S046_warnt_type_nm_minor_nm    = 3
    Const S046_sales_grp             = 4
    Const S046_sales_grp_nm          = 5
    Const S046_collateral_no        = 6
    Const S046_collateral_type      = 7
    Const S046_proffer_nm           = 8
    Const S046_proffer_rgst_no      = 9
    Const S046_relationship         = 10
    Const S046_estate_place         = 11
    Const S046_appraiser_nm         = 12
    Const S046_est_dt               = 13
    Const S046_asgn_seq             = 14
    Const S046_cur                  = 15
    Const S046_est_amt              = 16
    Const S046_able_amt             = 17
    Const S046_asgn_amt             = 18
    Const S046_asgn_dt              = 19
    Const S046_expiry_dt            = 20
    Const S046_warnt_org_nm         = 21
    Const S046_stock_no             = 22
    Const S046_org_tel_no           = 23
    Const S046_lender_nm            = 24
    Const S046_floor_space          = 25
    Const S046_ground_space         = 26
    Const S046_del_dt               = 27
    Const S046_del_type             = 28
    Const S046_credit_chk_day       = 29
    Const S046_remark               = 30
    Const S046_ext1_qty             = 31
    Const S046_ext2_qty             = 32
    Const S046_ext1_amt             = 33
    Const S046_ext2_amt             = 34
    Const S046_ext1_cd              = 35
    Const S046_ext2_cd              = 36
    
    On Error Resume Next

    Err.Clear 
    
    iWarrentNo = Trim(Request("txtWarrentNo"))
    
    If Request("txtWarrentNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
       Call ServerMesgBox("조회 조건값이 비어있습니다.", vbInformation, I_MKSCRIPT)              
       Exit Sub
	End If

    Set iS1G115 = Server.CreateObject("PS1G115.cLookupcollateralsvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
    
    EG1_S_Collateral = iS1G115.S_LOOKUP_COLLATERAL_SVR(gStrGlobalCollection,"QUERY",UCase(iWarrentNo))
      
	If CheckSYSTEMError(Err,True) = True Then
       Set iS1G115 = Nothing
       Exit Sub
    End If  
    
    Set iS1G115 = Nothing
	'-----------------------
	'Display result data
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	
	Response.Write ".txtCustomer.Value			= """ & ConvSPChars(EG1_S_Collateral(S046_bp_cd))            & """" & vbCr
	Response.Write ".txtCustomerNm.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_bp_nm))            & """" & vbCr
    Response.Write ".txtDelTypeNm.Value			= """ & ConvSPChars(EG1_S_Collateral(S046_del_type_nm_minor_nm))   & """" & vbCr
	Response.Write ".txtWarrentTypeNm.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_warnt_type_nm_minor_nm)) & """" & vbCr
	Response.Write ".txtSalesGrp.Value		    = """ & ConvSPChars(EG1_S_Collateral(S046_sales_grp))          & """" & vbCr
	Response.Write ".txtSalesGrpNm.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_sales_grp_nm))       & """" & vbCr 
    
	Response.Write ".txtWarrentNo.Value		    = """ & ConvSPChars(EG1_S_Collateral(S046_collateral_no))        & """" & vbCr
	Response.Write ".txtWarrentNo1.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_collateral_no))        & """" & vbCr	    
    Response.Write ".txtWarrentType.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_collateral_type))      & """" & vbCr
	Response.Write ".txtProffer.value           = """ & ConvSPChars(EG1_S_Collateral(S046_proffer_nm))           & """" & vbCr
	Response.Write ".txtProfferRgstNo.Value     = """ & ConvSPChars(EG1_S_Collateral(S046_proffer_rgst_no))      & """" & vbCr
	Response.Write ".txtRelationShip.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_relationship))         & """" & vbCr            
	Response.Write ".txtRocation.value	        = """ & ConvSPChars(EG1_S_Collateral(S046_estate_place))         & """" & vbCr            
	Response.Write ".txtEstimatePlace.Value		= """ & ConvSPChars(EG1_S_Collateral(S046_appraiser_nm))         & """" & vbCr                                                                          
	Response.Write ".txtEstimateDt.text 		= """ & UNIDateClientFormat(EG1_S_Collateral(S046_est_dt))       & """" & vbCr                               
	Response.Write ".txtAsgnSeq.text            = """ & ConvSPChars(EG1_S_Collateral(S046_asgn_seq))             & """" & vbCr   
	Response.Write ".txtCurrency.Value          = """ & ConvSPChars(EG1_S_Collateral(S046_cur))                  & """" & vbCr  
	Response.Write ".txtEstimateAmt.text      	= """ & UNINumClientFormat(EG1_S_Collateral(S046_est_amt), ggAmtOfMoney.DecPoint, 0)       & """" & vbCr                                                 
	Response.Write ".txtWarrentAbleAmt.text    	= """ & UNINumClientFormat(EG1_S_Collateral(S046_able_amt), ggAmtOfMoney.DecPoint, 0)      & """" & vbCr   
	Response.Write ".txtWarrentAsignAmt.text    = """ & UNINumClientFormat(EG1_S_Collateral(S046_asgn_amt), ggAmtOfMoney.DecPoint, 0)      & """" & vbCr                  
	Response.Write ".txtAsignDt.text 		    = """ & UNIDateClientFormat(EG1_S_Collateral(S046_asgn_dt))      & """" & vbCr                                                
	Response.Write ".txtExpiryDt.text 		    = """ & UNIDateClientFormat(EG1_S_Collateral(S046_expiry_dt))    & """" & vbCr                                             		                                                                                                                  
	Response.Write ".txtWarrentOrgNm.value		= """ & ConvSPChars(EG1_S_Collateral(S046_warnt_org_nm))         & """" & vbCr  
	Response.Write ".txtStockNo.Value		    = """ & ConvSPChars(EG1_S_Collateral(S046_stock_no))             & """" & vbCr                                            
	Response.Write ".txtOrgTelNo.Value		    = """ & ConvSPChars(EG1_S_Collateral(S046_org_tel_no))           & """" & vbCr  
	
	Response.Write ".txtLenderNm.Value		    = """ & ConvSPChars(EG1_S_Collateral(S046_lender_nm))            & """" & vbCr                                                                                                                                                                                             
	Response.Write ".txtFloorSpace.text			= """ & UNINumClientFormat(EG1_S_Collateral(S046_floor_space), ggQty.DecPoint, 0)          & """" & vbCr          
	Response.Write ".txtGroundSpace.text		= """ & UNINumClientFormat(EG1_S_Collateral(S046_ground_space), ggQty.DecPoint , 0)         & """" & vbCr   
	Response.Write ".txtDelDt.text 			    = """ & UNIDateClientFormat(EG1_S_Collateral(S046_del_dt))       & """" & vbCr       
	Response.Write ".txtDelType.value			= """ & ConvSPChars(EG1_S_Collateral(S046_del_type))             & """" & vbCr     
	
	If EG1_S_Collateral(S046_del_type)   <> "" Then                                                                       
	   Response.Write ".chkDeleteFlg.checked = True " & vbCr
	End If   
         
	Response.Write ".txtCreditCheckDt.text		= """ & UNINumClientFormat(EG1_S_Collateral(S046_credit_chk_day),ggQty.DecPoint, 0)      & """" & vbCr 
	Response.Write ".txtRemark.value			= """ & ConvSPChars(EG1_S_Collateral(S046_remark))                & """" & vbCr 	                                                         
	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
Response.Write ggAmtOfMoney.DecPoint	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    Dim pvCommand
    Dim itxtFlgMode
    Dim iS1G114
    Dim E1_S_Collateral
    
    Dim I1_S_Collateral
    
    Const C_I1_collateral_no   = 0
    Const C_I1_collateral_type = 1
    Const C_I1_proffer_nm      = 2
    Const C_I1_proffer_rgst_no = 3
    Const C_I1_relationship    = 4
    Const C_I1_estate_place    = 5
    Const C_I1_est_dt          = 6
    Const C_I1_asgn_seq        = 7
    Const C_I1_appraiser_nm    = 8
    Const C_I1_cur             = 9
    Const C_I1_est_amt         = 10
    Const C_I1_able_amt        = 11
    Const C_I1_asgn_amt        = 12
    Const C_I1_asgn_dt         = 13
    Const C_I1_expiry_dt       = 14
    Const C_I1_warnt_org_nm    = 15
    Const C_I1_stock_no        = 16
    Const C_I1_org_tel_no      = 17
    Const C_I1_lender_nm       = 18
    Const C_I1_floor_space     = 19
    Const C_I1_ground_space    = 20
    Const C_I1_del_dt          = 21
    Const C_I1_del_type        = 22
    Const C_I1_credit_chk_day  = 23
    Const C_I1_remark          = 24
    Const C_I1_ext1_qty        = 25
    Const C_I1_ext2_qty        = 26
    Const C_I1_ext1_amt        = 27
    Const C_I1_ext2_amt        = 28
    Const C_I1_ext1_cd         = 29
    Const C_I1_ext2_cd         = 30
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    ReDim I1_S_Collateral(C_I1_ext2_cd)

    I1_S_Collateral(C_I1_collateral_no)    = UCase(Trim(Request("txtWarrentNo1")))    
    I1_S_Collateral(C_I1_collateral_type)  = UCase(Trim(Request("txtWarrentType")))
    I1_S_Collateral(C_I1_proffer_nm)       = Trim(Request("txtProffer"))
    I1_S_Collateral(C_I1_proffer_rgst_no)  = UCase(Trim(Request("txtProfferRgstNo")))
    I1_S_Collateral(C_I1_relationship)     = Trim(Request("txtRelationship"))
    I1_S_Collateral(C_I1_estate_place)     = Trim(Request("txtRocation"))  
    I1_S_Collateral(C_I1_est_dt)           = UNIConvDate(Request("txtEstimateDt"))
	I1_S_Collateral(C_I1_asgn_seq)         = UNIConvNum(Trim(Request("txtAsgnSeq")),0)
    I1_S_Collateral(C_I1_appraiser_nm)     = Trim(Request("txtEstimatePlace"))
    I1_S_Collateral(C_I1_cur)              = UCase(Trim(Request("txtCurrency")))
    I1_S_Collateral(C_I1_est_amt)          = UNIConvNum(Trim(Request("txtEstimateAmt")),0)
    I1_S_Collateral(C_I1_able_amt)         = UNIConvNum(Trim(Request("txtWarrentAbleAmt")),0)
    I1_S_Collateral(C_I1_asgn_amt)         = UNIConvNum(Trim(Request("txtWarrentAsignAmt")),0)
    I1_S_Collateral(C_I1_asgn_dt)          = UNIConvDate(Trim(Request("txtAsignDt")))
    I1_S_Collateral(C_I1_expiry_dt)        = UNIConvDate(Trim(Request("txtExpiryDt")))
    I1_S_Collateral(C_I1_warnt_org_nm)     = Trim(Request("txtWarrentOrgNm"))
    I1_S_Collateral(C_I1_stock_no)         = Trim(Request("txtStockNo"))
    I1_S_Collateral(C_I1_org_tel_no)       = Trim(Request("txtOrgTelNo"))
    I1_S_Collateral(C_I1_lender_nm)        = Trim(Request("txtLenderNm"))
    
  	If Len(Request("txtFloorSpace")) Then
	    I1_S_Collateral(C_I1_floor_space)  = UNIConvNum(Request("txtFloorSpace"),0)
    End If
    
    If Len(Request("txtGroundSpace")) Then
        I1_S_Collateral(C_I1_ground_space) = UNIConvNum(Request("txtGroundSpace"),0)
	End If
	
	I1_S_Collateral(C_I1_del_dt)           = UNIConvDate(Request("txtDelDt"))
	I1_S_Collateral(C_I1_del_type)         = UCase(Trim(Request("txtDelType")))
	I1_S_Collateral(C_I1_credit_chk_day)   = UNIConvNum(Trim(Request("txtCreditCheckDt")), 0)
	I1_S_Collateral(C_I1_remark)           = Trim(Request("txtRemark"))
    
	itxtFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    If itxtFlgMode = OPMD_CMODE Then
		pvCommand = "CREATE"
    ElseIf itxtFlgMode = OPMD_UMODE Then'
		pvCommand = "UPDATE"
    End If

    Set iS1G114 = Server.CreateObject("PS1G114.CMaintcollateralsvr")

	If CheckSYSTEMError(Err,True) = True Then
       Set iS1G114 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    Dim reqtxtSalesGrp
    Dim reqtxtCustomer
    reqtxtSalesGrp = Request("txtSalesGrp")
    reqtxtCustomer = Request("txtCustomer")
    
    '===================
    ' For Call Com Agent
    '===================
    pvCB = "F"
    
    E1_S_Collateral = iS1G114.S_MAINT_COLLATERAL_SVR (pvCB, gStrGlobalCollection, pvCommand,I1_S_Collateral,Trim(reqtxtCustomer), Trim(reqtxtSalesGrp),"")
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS1G114 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS1G114 = Nothing

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"           & vbCr

	If E1_S_Collateral <>"" Then
	    Response.Write ".frm1.txtWarrentNo.value = """ & ConvSPChars(E1_S_Collateral)     & """" & vbCr
	else
	    Response.Write ".frm1.txtWarrentNo.value =  .frm1.txtWarrentNo1.value	    	     " & vbCr
	    
	end if	
    Response.Write ".DbSaveOk"                  & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    Dim iS1G114
    
    Dim I1_S_Collateral
        
    Const C_I1_collateral_no   = 0    
    
    On Error Resume Next
    Err.Clear
                                                                   '☜: Protect system from crashing
    ReDim I1_S_Collateral(C_I1_collateral_no)
    
    I1_S_Collateral(C_I1_collateral_no)    = FilterVar(UCase(Trim(Request("txtWarrentNo"))),"","SMM") 
    If I1_S_Collateral(C_I1_collateral_no) = "" Then										'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Exit Sub
	End If
   
    Set iS1G114 = Server.CreateObject("PS1G114.CMaintcollateralsvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
    Dim reqtxtInsrtUserId
    reqtxtInsrtUserId = Request("txtInsrtUserId")
    
    '===================
    ' For Call Com Agent
    '===================
    pvCB = "F"
    
    Call iS1G114.S_MAINT_COLLATERAL_SVR(pvCB, gStrGlobalCollection,"DELETE",I1_S_Collateral,"","",UCase(Trim(reqtxtInsrtUserId)))
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS1G114 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

	Set iS1G114 = Nothing                                                                    '☜: Unload Comproxy
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"                & vbCr
    Response.Write ".DbDeleteOk"                & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr
	 
End Sub


Sub SubBizCheck()

    Dim strChkMode
    Dim lgIntFlgMode
    Dim iS1G116

    Dim I2_ief_supplied

    Dim I1_S_Collateral
    Const S052_cur           = 0
    Const S052_asgn_amt      = 1
    Const S052_asgn_dt       = 2
    Const S052_collateral_no = 3
    Const S052_del_type      = 4

    
    On Error Resume Next
    Err.Clear 
    strChkMode = Request("txtCHKMode")
    If strChkMode = "SAVE" Then
        lgIntFlgMode = CInt(Request("txtFlgMode"))	
		If lgIntFlgMode = OPMD_CMODE Then
			I2_ief_supplied = "C"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			I2_ief_supplied = "U"
		End If
	Else
		I2_ief_supplied = "D"
    End If
    Redim I1_S_Collateral(S052_del_type)
    I1_S_Collateral(S052_cur)              = UCase(Trim(Request("txtCurrency")))
    I1_S_Collateral(S052_asgn_amt)         = UNIConvNum(Trim(Request("txtWarrentAsignAmt")),0) 
    I1_S_Collateral(S052_asgn_dt)          = UNIConvDate(Trim(Request("txtAsignDt")))
    I1_S_Collateral(S052_collateral_no)    = UCase(Trim(Request("txtWarrentNo1")))
	I1_S_Collateral(S052_del_type)         = UCase(Trim(Request("txtDelType")))

    Set iS1G116 = Server.CreateObject("PS1G116.cCHeckcreditlimitsvr") 
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If     
	Dim reqtxtCustomer
	reqtxtCustomer = Request("txtCustomer")
    Call iS1G116.S_CHECK_CREDIT_LIMIT_SVR(gStrGlobalCollection,I1_S_Collateral,I2_ief_supplied, UCase(Trim(reqtxtCustomer))) 

	If InStr(UCase(Err.Description), "B_MESSAGE201928") Then

		Response.Write "<Script Language=vbscript>"                       & vbCr
		Response.Write "	Dim msgCreditlimit"                           & vbCr
		Response.Write "	msgCreditlimit = parent.DisplayMsgBox(""201928"", parent.parent.VB_YES_NO, ""X"", ""X"")" & vbCr
		Response.Write "	If msgCreditlimit = vbYes Then "              & vbCr
		Response.Write "		If """ & strChkMode & """ = ""DELETE"" Then " & vbCr
		Response.Write "			parent.DbDelete"                      & vbCr
		Response.Write "		Else"                                     & vbCr
		Response.Write "			parent.DbSave"                        & vbCr
		Response.Write "		End If"                                   & vbCr
		Response.Write "	End If"                                       & vbCr
		Response.Write "</script>"                                        & vbCr
		
		Set iS1G116 = Nothing	
		Exit sub
	Else
		If CheckSYSTEMError(Err,True) = True Then
			Call SetErrorStatus                                                           '☆: Mark that error occurs
			Set iS1G116 = Nothing		                                                 '☜: Unload Comproxy DLL
			Exit Sub
		End If  
    End if
    Set iS1G116 = Nothing
    
    If  strChkMode  =  "SAVE"  Then 
	    Call SubBizSave    
    Else
		Call SubBizDelete
    End If
	 
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

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub

%>

