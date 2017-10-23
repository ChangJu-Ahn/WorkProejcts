<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1251mb1
'*  4. Program Name         : 구매그룹등록 
'*  5. Program Desc         : 구매그룹등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%

	Dim lgOpModeCRUD
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    
    Dim iB26019																	'☆ : 입력/수정용 ComProxy Dll 사용 변수															'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim strDt
	Dim strDefFrDate
	Dim strDefToDate

	Dim I1_b_pur_grp
	Dim E1_b_cost_center
	Dim E2_b_pur_org
	Dim E3_b_pur_grp

    Const M028_E1_cost_cd = 0
    Const M028_E1_cost_nm = 1
	Redim E1_b_cost_center(M028_E1_cost_nm)

    Const M028_E2_pur_org = 0
    Const M028_E2_pur_org_nm = 1
	Redim E2_b_pur_org(M028_E2_pur_org_nm)
	
    Const M028_E3_pur_grp = 0
    Const M028_E3_pur_grp_nm = 1
    Const M028_E3_usage_flg = 2
    Const M028_E3_valid_fr_dt = 3
    Const M028_E3_valid_to_dt = 4
    Const M028_E3_ext1_cd = 5
    Const M028_E3_ext2_cd = 6
    Const M028_E3_ext3_cd = 7
    Const M028_E3_ext4_cd = 8
    Redim E3_b_pur_grp(M028_E3_ext4_cd)
	
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	I1_b_pur_grp = UCase(Trim(Request("txtGroupCd1")))
	
    Set iB26019 = Server.CreateObject("PB2G619.cBLookupPurGrpS") 
      
    Call  iB26019.B_LOOKUP_PUR_GRP_SVR(gStrGlobalCollection, I1_b_pur_grp, E1_b_cost_center, E2_b_pur_org, E3_b_pur_grp)  
	
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iB26019 = Nothing												'☜: ComProxy Unload
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Parent.frm1.txtGroupNm1.value = """" " & vbCr
		Response.Write "</Script>" & vbCr
		
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if
    
	'-----------------------
	'Result data display area
	'----------------------- 
	strDefFrDate = "1900-01-01"
	strDefToDate = "2999-12-31"
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.frm1.txtGroupCd1.value = """ & ConvSPChars(E3_b_pur_grp(M028_E3_pur_grp))      & """" & vbCr
	Response.Write "Parent.frm1.txtGroupNm1.value = """ & ConvSPChars(E3_b_pur_grp(M028_E3_pur_grp_nm))   & """" & vbCr
	Response.Write "Parent.frm1.txtGroupCd2.value = """ & ConvSPChars(E3_b_pur_grp(M028_E3_pur_grp))      & """" & vbCr
	Response.Write "Parent.frm1.txtGroupNm2.value = """ & ConvSPChars(E3_b_pur_grp(M028_E3_pur_grp_nm))   & """" & vbCr
	Response.Write "Parent.frm1.txtOrgCd2.value   = """ & ConvSPChars(E2_b_pur_org(M028_E2_pur_org))      & """" & vbCr
	Response.Write "Parent.frm1.txtOrgNm2.value   = """ & ConvSPChars(E2_b_pur_org(M028_E2_pur_org_nm))   & """" & vbCr
	Response.Write "Parent.frm1.txtCostCd.value   = """ & ConvSPChars(E1_b_cost_center(M028_E1_cost_cd))  & """" & vbCr
	Response.Write "Parent.frm1.txtCostNm.value   = """ & ConvSPChars(E1_b_cost_center(M028_E1_cost_nm))  & """" & vbCr
	
	Response.Write "If """ & E3_b_pur_grp(M028_E3_usage_flg)  & """" & "=""Y""" & " Then " & vbCr
	Response.Write "	Parent.frm1.rdoUseflg(0).checked= true "      & vbCr
	Response.Write "Else Parent.frm1.rdoUseflg(1).checked= true" & vbCr
	Response.Write "End If "                                      & vbCr

	Response.Write "Parent.frm1.txtFrDt.Text  = """ & UNIDateClientFormat(E3_b_pur_grp(M028_E3_valid_fr_dt)) & """" & vbCr
	Response.Write "Parent.frm1.txtToDt.Text  = """ & UNIDateClientFormat(E3_b_pur_grp(M028_E3_valid_to_dt)) & """" & vbCr
	
	Response.Write "Parent.lgNextNo  = """" "                    & vbCr
	Response.Write "Parent.lgPrevNo  = """" "                    & vbCr
	Response.Write "Parent.DbQueryOk "           & vbCr
    Response.Write "</Script>" & vbCr
	 
    Set iB26019 = Nothing															'☜: Unload Comproxy

End Sub	


'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim ib26011
	Dim iCommandSent
	Dim lgIntFlgMode		
	
	Dim I1_b_cost_center
	Dim I3_b_pur_org
	Dim I2_b_pur_grp
	
	'import_b_pur_grp
	Const M006_I4_pur_grp = 0
    Const M006_I4_pur_grp_nm = 1
    Const M006_I4_usage_flg = 2
    Const M006_I4_valid_fr_dt = 3
    Const M006_I4_valid_to_dt = 4
    Const M006_I4_ext1_cd = 5
    Const M006_I4_ext2_cd = 6
    Const M006_I4_ext3_cd = 7
    Const M006_I4_ext4_cd = 8
    ReDim I2_b_pur_grp(M006_I4_ext4_cd)

	    
    lgIntFlgMode = CInt(Request("txtFlgMode")) 
    
    I2_b_pur_grp(M006_I4_pur_grp)	 = UCase(Trim(Request("txtGroupCd2")))
    I2_b_pur_grp(M006_I4_pur_grp_nm) = Trim(Request("txtGroupNm2"))
    I3_b_pur_org					 = Trim(UCase(Request("txtORGCd2")))
    I1_b_cost_center				 = Trim(Request("txtCostCd"))
    I2_b_pur_grp(M006_I4_usage_flg)	 = Request("txtUseflg") 
    I2_b_pur_grp(M006_I4_valid_fr_dt)	= UNIConvDate(Request("txtFrDt"))
    I2_b_pur_grp(M006_I4_valid_to_dt)	= UNIConvDate(Request("txtToDt"))
    
    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "Create"	
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "Update"
	End If
  
   	Set ib26011 = Server.CreateObject("PB2G611.cBMaintPurGrpS")      
  
   	Call ib26011.B_MAINT_PUR_GRP_SVR(gStrGlobalCollection, iCommandSent, I1_b_cost_center, I3_b_pur_org, I2_b_pur_grp)
        
	If CheckSYSTEMError(Err,True) = true then 		
		Set ib26011 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
 	End if
		
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
    Response.Write "</Script>"                  & vbCr 

    Set ib26011 = Nothing                                                   '☜: Unload Comproxy

End Sub	
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear   
     
    Dim ib26011	
    Dim iCommandSent
    
    Dim I2_b_pur_grp															
	
	'import_b_pur_grp
	Const M006_I4_pur_grp = 0
    ReDim I2_b_pur_grp(M006_I4_pur_grp)

    I2_b_pur_grp(M006_I4_pur_grp) = Trim(Request("txtGroupCd1"))                                                                    '☜: Clear Error status
	
	Set ib26011 = Server.CreateObject("PB2G611.cBMaintPurGrpS") 
	
	If CheckSYSTEMError(Err,True) = true then 		
		Set ib26011 = Nothing												'☜: ComProxy Unload
		Exit Sub													'☜: 비지니스 로직 처리를 종료함 
	End if
		
    iCommandSent = "Delete"
    
    call ib26011.B_MAINT_PUR_GRP_SVR(gStrGlobalCollection, iCommandSent, , , I2_b_pur_grp)
    
    
	If CheckSYSTEMError(Err,True) = true then 		
		Set ib26011 = Nothing												'☜: ComProxy Unload
		Exit Sub													'☜: 비지니스 로직 처리를 종료함 
	End if

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "parent.DbDeleteOk "           & vbCr
    Response.Write "</Script>"                  & vbCr
	        
    Set ib26011 = Nothing                                                   '☜: Unload Comproxy

End Sub


%>
