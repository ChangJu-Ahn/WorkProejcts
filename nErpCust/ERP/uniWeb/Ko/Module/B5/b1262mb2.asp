<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : B1262MB2
'*  4. Program Name         : 매출거래처형태등록 
'*  5. Program Desc         : 매출거래처형태등록 
'*  6. Comproxy List        : PB5GS42.dll, PB5GS43.dll     
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2000/03/27
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									
'*                            this mark(⊙) Means that "may  change"									
'*                            this mark(☆) Means that "must change"									
'* 13. History              : 
'**********************************************************************************************
%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->

<%
    Dim lgOpModeCRUD
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()

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
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim pB5GS43    
    
	Dim I1_child_b_biz_partner
	Dim I2_parent_b_biz_partner
	Dim I3_b_biz_partner_ftn

	Dim E1_child_b_biz_partner
	Dim E2_parent_b_biz_partner
	Dim E3_b_biz_partner_ftn
    
    Const EG1_E1_b_biz_partner_ftn_partner_ftn = 0
    Const EG1_E1_b_biz_partner_ftn_bp_prsn_nm = 1
    Const EG1_E1_b_biz_partner_ftn_bp_contact_pt = 2
    Const EG1_E1_b_biz_partner_ftn_default_flag = 3
    Const EG1_E1_b_biz_partner_ftn_usage_flag = 4
   
    Const EG1_E2_child_b_biz_partner_bp_cd = 0
    Const EG1_E2_child_b_biz_partner_bp_nm = 1

    Const EG1_E3_parent_b_biz_partner_bp_cd = 0
    Const EG1_E3_parent_b_biz_partner_bp_nm = 1

    On Error Resume Next

    Err.Clear 

    If Request("txtBp_cd1") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If

    If Request("txtPartner_cd1") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Exit Sub
	End If

    If Request("txtRadioType") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
       Exit Sub
	End If

	I2_parent_b_biz_partner  = Trim(Request("txtBp_cd1"))
	I1_child_b_biz_partner   = Trim(Request("txtPartner_cd1"))
	I3_b_biz_partner_ftn     = Request("txtRadioType")
	
    Set pB5GS43 = Server.CreateObject("PB5GS43.cBLkBizPartnerFtnSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  

    Call pB5GS43.B_LOOKUP_BIZ_PARTNR_FTN_SVR(gStrGlobalCollection, I1_child_b_biz_partner, _
											I2_parent_b_biz_partner, I3_b_biz_partner_ftn, _
											E1_child_b_biz_partner, E2_parent_b_biz_partner, E3_b_biz_partner_ftn)
      
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent.frm1"           & vbCr
		Response.Write ".txtBp_nm1.Value			= """ & ConvSPChars(E2_parent_b_biz_partner(EG1_E3_parent_b_biz_partner_bp_nm))       & """" & vbCr
		Response.Write ".txtPartner_nm1.Value		= """ & ConvSPChars(E1_child_b_biz_partner(EG1_E2_child_b_biz_partner_bp_nm))		& """" & vbCr
		Response.Write ".txtBp_cd1.focus " & vbCr
		Response.Write "End With"          & vbCr
		Response.Write "</Script>"         & vbCr

       Set pB5GS43 = Nothing
       Exit Sub
    End If  
    
    Set pB5GS43 = Nothing


	'-----------------------
	'Display result data
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	
	Response.Write ".txtBp_cd2.Value			= """ & ConvSPChars(E2_parent_b_biz_partner(EG1_E3_parent_b_biz_partner_bp_cd))       & """" & vbCr
	Response.Write ".txtBp_nm2.Value			= """ & ConvSPChars(E2_parent_b_biz_partner(EG1_E3_parent_b_biz_partner_bp_nm))       & """" & vbCr
    Response.Write ".txtPartner_cd2.Value		= """ & ConvSPChars(E1_child_b_biz_partner(EG1_E2_child_b_biz_partner_bp_cd))		& """" & vbCr
	Response.Write ".txtPartner_nm2.Value		= """ & ConvSPChars(E1_child_b_biz_partner(EG1_E2_child_b_biz_partner_bp_nm))		& """" & vbCr
    Response.Write ".txtBp_prsn_nm.Value		= """ & ConvSPChars(E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_bp_prsn_nm))		& """" & vbCr
	Response.Write ".txtBp_contact_pt.Value		= """ & ConvSPChars(E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_bp_contact_pt))		& """" & vbCr
    
	If E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_partner_ftn)   = "SSH" Then 
	   Response.Write ".rdoParttype21.checked = True " & vbCr
	elseif	E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_partner_ftn)   = "SBI" Then    
	   Response.Write ".rdoParttype22.checked = True " & vbCr
	elseif	E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_partner_ftn)   = "SPA" Then    	   
	   Response.Write ".rdoParttype23.checked = True " & vbCr	   	   
	End If   
         
	If E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_usage_flag)   = "Y" Then 
	   Response.Write ".rdoUsage_flag1.checked = True " & vbCr
	elseif	E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_usage_flag)   = "N" Then    
	   Response.Write ".rdoUsage_flag2.checked = True " & vbCr
	End If   

	If E3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_default_flag)   = "Y" Then 
	   Response.Write ".chkPartner.checked = True " & vbCr
	else
	   Response.Write ".chkPartner.checked = False " & vbCr
	End If   

	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write ".txtBp_cd1.focus " & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()

    Dim pB5GS42   
    Dim iCommandSent
    Dim itxtFlgMode
    
	Dim I1_child_b_biz_partner
	Dim I2_parent_b_biz_partner
	Dim I3_b_biz_partner_ftn

    Const EG1_E1_b_biz_partner_ftn_partner_ftn = 0
    Const EG1_E1_b_biz_partner_ftn_bp_prsn_nm = 1
    Const EG1_E1_b_biz_partner_ftn_bp_contact_pt = 2
    Const EG1_E1_b_biz_partner_ftn_default_flag = 3
    Const EG1_E1_b_biz_partner_ftn_usage_flag = 4

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status        
    
	ReDim I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_usage_flag)
	
	I2_parent_b_biz_partner = UCase(Trim(Request("txtBp_cd2")))
	I1_child_b_biz_partner  = UCase(Trim(Request("txtPartner_cd2")))
	I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_partner_ftn)   = Trim(Request("txtRadioType"))	
	I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_bp_prsn_nm)    = Trim(Request("txtBp_prsn_nm"))	
	I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_bp_contact_pt) = Trim(Request("txtBp_contact_pt"))
	I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_default_flag)  = Trim(Request("txtCheck"))
	I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_usage_flag)    = Request("txtRadioFlag")				
	
	itxtFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    If itxtFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf itxtFlgMode = OPMD_UMODE Then
    		iCommandSent = "UPDATE"
    End If

    Set pB5GS42 = Server.CreateObject("PB5GS42.cBHBizPartnerFtnSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Set pB5GS42 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    Call pB5GS42.B_MAINT_BIZ_PARTNER_FTN_SVR(gStrGlobalCollection, iCommandSent, I1_child_b_biz_partner, I2_parent_b_biz_partner, I3_b_biz_partner_ftn)
    
	If CheckSYSTEMError(Err,True) = True Then
       Set pB5GS42 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set pB5GS42 = Nothing	
   			
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr   
    Response.Write "</Script> "  
    
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    Dim pB5GS42   
    Dim iCommandSent

	Dim I1_child_b_biz_partner
	Dim I2_parent_b_biz_partner
	Dim I3_b_biz_partner_ftn

    Const EG1_E1_b_biz_partner_ftn_partner_ftn = 0
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Redim I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_partner_ftn)

    If Request("txtBp_cd2") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	    Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If

    If Request("txtPartner_cd2") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Exit Sub
	End If

    If Request("txtRadioType") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
       Exit Sub
	End If
	
	I2_parent_b_biz_partner = UCase(Trim(Request("txtBp_cd2")))
	I1_child_b_biz_partner  = UCase(Trim(Request("txtPartner_cd2")))
	I3_b_biz_partner_ftn(EG1_E1_b_biz_partner_ftn_partner_ftn) = Request("txtRadioType")    

    iCommandSent = "DELETE"
    
    Set pB5GS42 = Server.CreateObject("PB5GS42.cBHBizPartnerFtnSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
               
	call pB5GS42.B_MAINT_BIZ_PARTNER_FTN_SVR(gStrGlobalCollection, iCommandSent, I1_child_b_biz_partner, I2_parent_b_biz_partner, I3_b_biz_partner_ftn)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set pB5GS42 = Nothing
		Exit Sub
	End If     
    '-----------------------
	'Result data display area
	'----------------------- 
	Set pB5GS42 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbDeleteOk "    & vbCr   
    Response.Write "</Script> "  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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

%>

