<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->

<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s1912mb1
'*  4. Program Name         : ATP설정 
'*  5. Program Desc         : ATP설정 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/05/18
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Sonbumyeol
'* 10. Modifier (Last)      : Sonbumyeol
'* 11. Comment              :
'=======================================================================================================
	
	Dim lgOpModeCRUD
    
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB") 
	Call HideStatusWnd                                                                 '☜: Hide Processing message
    
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
Sub SubBizQuery()
	
	Dim iD2GS67
	Dim I1_b_plant
	Dim E1_b_plant
	Dim EG1_s_atp_config
	
	Const C_plant_cd = 0
	Const C_plant_nm = 1
    
    Const C_atp_chk_flag = 0
    Const C_atp_area_lvl = 1
    Const C_atp_area_lvl_nm = 2
    Const C_onhand_stk_lvl = 3
    Const C_onhand_stk_lvl_nm = 4
    Const C_planned_gi_lvl = 5
    Const C_planned_gi_lvl_nm = 6
    Const C_planned_gr_lvl = 7
    Const C_planned_gr_lvl_nm = 8
    Const C_req_qty_lvl = 9
	
	On Error Resume Next
    Err.Clear 
    
    I1_b_plant = Trim(Request("txtconPlant_cd"))
    
    Set iD2GS67 = Server.CreateObject("PD2GS67.cSLkAtpConfigSvr") 
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
    
    Call iD2GS67.S_LOOKUP_ATP_CONFIG_SVR(gStrGlobalCollection, I1_b_plant, E1_b_plant, EG1_s_atp_config) 
    
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "parent.frm1.txtconPlant_nm.value	= """ & ConvSPChars(E1_b_plant(C_plant_nm))  & """" & vbCr
	Response.Write "</Script>"         & vbCr
	
    If CheckSYSTEMError(Err,True) = True Then
       Set iD2GS67 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iD2GS67 = Nothing
	
	'-----------------------
	'Display result data
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	
	Response.Write ".txtconPlant_nm.value		= """ & ConvSPChars(E1_b_plant(C_plant_nm))			  & """" & vbCr
	Response.Write ".txtPlant_cd.value			= """ & ConvSPChars(E1_b_plant(C_plant_cd))			  & """" & vbCr			
	Response.Write ".txtPlant_nm.value			= """ & ConvSPChars(E1_b_plant(C_plant_nm))			  & """" & vbCr			
	
	If EG1_s_atp_config(C_atp_chk_flag)			= "Y" Then
		Response.Write ".rdoATP_flag1.checked	= True													  	   " & vbCr
		Response.Write ".txtRadioflag.value		= .rdoATP_flag1.value										   " & vbCr
	ElseIf EG1_s_atp_config(C_atp_chk_flag)		= "N" Then		
		Response.Write ".rdoATP_flag2.checked	= True														   " & vbCr
		Response.Write ".txtRadioflag.value		= .rdoATP_flag2.value										   " & vbCr
	End IF
	
	Response.Write ".txtAtp_area_lvl.value		= """ & ConvSPChars(EG1_s_atp_config(C_atp_area_lvl))	   & """" & vbCr
	Response.Write ".txtAtp_area_lvl_nm.value	= """ & ConvSPChars(EG1_s_atp_config(C_atp_area_lvl_nm))   & """" & vbCr		
	Response.Write ".txtOnhand_stk_lvl.value	= """ & ConvSPChars(EG1_s_atp_config(C_onhand_stk_lvl))    & """" & vbCr		
	Response.Write ".txtOnhand_stk_lvl_nm.value	= """ & ConvSPChars(EG1_s_atp_config(C_onhand_stk_lvl_nm)) & """" & vbCr		
	Response.Write ".txtPlaned_gi_lvl.value		= """ & ConvSPChars(EG1_s_atp_config(C_planned_gi_lvl))    & """" & vbCr		
	Response.Write ".txtPlaned_gi_lvl_nm.value	= """ & ConvSPChars(EG1_s_atp_config(C_planned_gi_lvl_nm)) & """" & vbCr		
	Response.Write ".txtPlaned_gr_lvl.value		= """ & ConvSPChars(EG1_s_atp_config(C_planned_gr_lvl))    & """" & vbCr
	Response.Write ".txtPlaned_gr_lvl_nm.value	= """ & ConvSPChars(EG1_s_atp_config(C_planned_gr_lvl_nm)) & """" & vbCr				
	
	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
    
End Sub	


'============================================================================================================
Sub SubBizSave()
	
	Dim iCommandSent
	Dim lgIntFlgMode
    Dim I1_b_plant
    Dim I2_s_atp_config
    Dim E1_s_atp_config
    Dim E2_b_plant
    Dim E3_atp_area_lvl_b_minor
    Dim E4_onhand_stk_b_minor
    Dim E5_planned_gi_b_minor
    Dim E6_planned_gr_b_minor
    
    Dim iD2GS66 
    
    Const atp_chk_flag = 0
    Const planned_gi_lvl = 1
    Const planned_gr_lvl = 2
    Const onhand_stk_lvl = 3
    Const req_qty_lvl = 4
    Const insrt_user_id = 5
    Const insrt_dt = 6
    Const updt_user_id = 7
    Const updt_dt = 8
    Const atp_area_lvl = 9
    Const ext1_qty = 10
    Const ext2_qty = 11
    Const ext3_qty = 12
    Const ext1_amt = 13
    Const ext2_amt = 14
    Const ext3_amt = 15
    Const ext1_cd = 16
    Const ext2_cd = 17
    Const ext3_cd = 18
    
    On Error Resume Next
    Err.Clear 
    
    Redim I2_s_atp_config(18)
	
    I1_b_plant = UCase(Trim(Request("txtPlant_cd")))
    
    I2_s_atp_config(planned_gi_lvl) = UCase(Trim(Request("txtPlaned_gi_lvl")))
    I2_s_atp_config(planned_gr_lvl) = UCase(Trim(Request("txtPlaned_gr_lvl")))
    I2_s_atp_config(onhand_stk_lvl) = UCase(Trim(Request("txtOnhand_stk_lvl")))
    I2_s_atp_config(atp_area_lvl) = UCase(Trim(Request("txtAtp_area_lvl")))
    I2_s_atp_config(atp_chk_flag) = UCase(Trim(Request("txtRadioFlag")))
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 
	
    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
        
    Set iD2GS66 = Server.CreateObject("PD2GS66.cSAtpConfigSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iD2GS66 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Call iD2GS66.S_MAINT_ATP_CONFIG_SVR(gStrGlobalCollection, iCommandSent, _
                                        I1_b_plant, I2_s_atp_config, _
                                        E1_s_atp_config, E2_b_plant, _
                                        E3_atp_area_lvl_b_minor, E4_onhand_stk_b_minor, _
                                        E5_planned_gi_b_minor, E6_planned_gr_b_minor)
    
    If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iD2GS66 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iD2GS66 = Nothing
    
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"           & vbCr
	Response.Write ".DbSaveOk"                  & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr
	Response.End																				'☜: Process End

End Sub


'============================================================================================================
Sub SubBizDelete()
	
	Dim lgIntFlgMode
	Dim iCommandSent
    Dim I1_b_plant
    
    Dim iD2GS66 
    
    On Error Resume Next
    Err.Clear 
    
    I1_b_plant = Trim(Request("txtconPlant_cd"))
        
    Set iD2GS66 = Server.CreateObject("PD2GS66.cSAtpConfigSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iD2GS66 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Call iD2GS66.S_MAINT_ATP_CONFIG_SVR(gStrGlobalCollection, "DELETE", I1_b_plant)
    
    If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iD2GS66 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iD2GS66 = Nothing
    
	Response.Write "<Script Language=vbscript>"	& vbCr
	Response.Write "With parent"				& vbCr
	Response.Write ".DbDeleteOk"                & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr
	Response.End		
	
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
    
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

End Sub

%>

