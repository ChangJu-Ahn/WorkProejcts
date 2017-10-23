<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
On Error Resume Next                                                             
Err.Clear                                                                        '☜: Clear Error status

Call LoadBasisGlobalInf()

Call HideStatusWnd                                                               '☜: Hide Processing message

 
lgOpModeCRUD  = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

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

On Error Resume Next
Err.Clear 

Dim PB6CS83
Dim imp_b_sales_org
Dim exp_b_sales_org
Dim exp_upper_b_sales_org

imp_b_sales_org = Trim(Request("txtSales_Org2"))

Const C_exp_b_sales_org_sales_org = 0
Const C_exp_b_sales_org_sales_org_nm = 1
Const C_exp_b_sales_org_sales_org_full_nm = 2
Const C_exp_b_sales_org_sales_org_eng_nm = 3
Const C_exp_b_sales_org_head_usr_nm = 4
Const C_exp_b_sales_org_lvl = 5
Const C_exp_b_sales_org_end_org_flag = 6
Const C_exp_b_sales_org_usage_flag = 7

Const C_exp_upper_b_sales_org_sales_org = 0
Const C_exp_upper_b_sales_org_sales_org_nm = 1

Set PB6CS83 = CreateObject("PB6CS83.cBLkSalesOrgSvr")
Dim reqtxtPrevNext
reqtxtPrevNext = Request("txtPrevNext")
Call PB6CS83.B_LOOKUP_SALES_ORG_SVR(gStrGlobalCollection,reqtxtPrevNext, _
imp_b_sales_org, exp_b_sales_org, exp_upper_b_sales_org)

'Response.Write "11111111"&Err.Description
    
If CheckSYSTEMError(Err,True) = True Then
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	Response.Write ".txtSales_Org_nm1.value  = """ & "" & """" & vbCr 
	Response.Write ".txtSales_Org1.focus"           & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
    Set PB6CS83 = Nothing
    Exit Sub
End If  

Set PB6CS83 = Nothing
'-----------------------
'Display result data
'----------------------- 
 Response.Write "<Script Language=vbscript>" & vbCr
 Response.Write "With parent.frm1"           & vbCr
 
 Response.Write ".txtSales_Org1.Value  = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org))            & """" & vbCr
 Response.Write ".txtSales_Org_nm1.Value  = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org_nm))            & """" & vbCr
    
 Response.Write ".txtSales_Org2.Value  = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org))   & """" & vbCr
 Response.Write ".txtlvl.Value    = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_lvl))      & """" & vbCr                  
 Response.Write ".txtSales_Org_nm2.Value  = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org_nm))          & """" & vbCr
 Response.Write ".txtSales_Org_Fullnm.Value = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org_full_nm))       & """" & vbCr 
    
 Response.Write ".txtSales_Org_Engnm.Value = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org_eng_nm))        & """" & vbCr
 Response.Write ".txtUpper_Sales_Org.Value = """ & ConvSPChars(exp_upper_b_sales_org(C_exp_upper_b_sales_org_sales_org))        & """" & vbCr     
 Response.Write ".txtUpper_Sales_OrgNm.Value = """ & ConvSPChars(exp_upper_b_sales_org(C_exp_upper_b_sales_org_sales_org_nm))      & """" & vbCr
 Response.Write ".txtHead_usr_nm.value       = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_head_usr_nm))           & """" & vbCr
   
 
 If exp_b_sales_org(C_exp_b_sales_org_end_org_flag) = "Y" Then                                                                       
  Response.Write ".rdoEndOrgFlagY.checked = True " & vbCr
  Response.Write ".txtEndOrgFlag.value = Y " & vbCr
 ElseIf exp_b_sales_org(C_exp_b_sales_org_end_org_flag) = "N" Then
  Response.Write ".rdoEndOrgFlagN.checked = True " & vbCr
  Response.Write ".txtEndOrgFlag.value = N " & vbCr
 End IF
    
    If exp_b_sales_org(C_exp_b_sales_org_usage_flag) = "Y" Then                                                                       
  Response.Write ".rdoUsage_flag1.checked = True " & vbCr
  Response.Write ".txtRadio.value = Y " & vbCr
 ElseIf exp_b_sales_org(C_exp_b_sales_org_usage_flag) = "N" Then
  Response.Write ".rdoUsage_flag2.checked = True " & vbCr
  Response.Write ".txtRadio.value = N " & vbCr
 End IF
          
 Response.Write ".txtSales_Org_nm1.value  = """ & ConvSPChars(exp_b_sales_org(C_exp_b_sales_org_sales_org_nm))      & """" & vbCr 
 Response.Write "parent.DbQueryOk" & vbCr
 Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
    
End Sub 
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

On Error Resume Next                                 
Err.Clear                                           
 
Dim pvCommand
Dim itxtFlgMode
Dim PB6CS82
Dim imp_b_sales_org
Dim imp_upper_b_sales_org
  
Const C_exp_b_sales_org_sales_org = 0
Const C_exp_b_sales_org_sales_org_nm = 1
Const C_exp_b_sales_org_sales_org_full_nm = 2
Const C_exp_b_sales_org_sales_org_eng_nm = 3
Const C_exp_b_sales_org_head_usr_nm = 4
Const C_exp_b_sales_org_lvl = 5
Const C_exp_b_sales_org_end_org_flag = 6
Const C_exp_b_sales_org_usage_flag = 7 

ReDim imp_b_sales_org(C_exp_b_sales_org_usage_flag)

imp_b_sales_org(C_exp_b_sales_org_sales_org)         = UCase(Trim(Request("txtSales_Org2")))
imp_b_sales_org(C_exp_b_sales_org_sales_org_nm)      = Trim(Request("txtSales_Org_nm2"))
imp_b_sales_org(C_exp_b_sales_org_sales_org_full_nm) = Trim(Request("txtSales_Org_Fullnm"))
imp_b_sales_org(C_exp_b_sales_org_sales_org_eng_nm)  = Trim(Request("txtSales_Org_Engnm"))
imp_b_sales_org(C_exp_b_sales_org_head_usr_nm)       = Trim(Request("txtHead_usr_nm"))
imp_b_sales_org(C_exp_b_sales_org_lvl)               = Trim(Request("txtlvl"))
imp_b_sales_org(C_exp_b_sales_org_end_org_flag)      = Trim(Request("txtEndOrgFlag"))
imp_b_sales_org(C_exp_b_sales_org_usage_flag)        = Trim(Request("txtRadio"))

imp_upper_b_sales_org = Trim(Request("txtUpper_Sales_Org"))
 
itxtFlgMode = CInt(Request("txtFlgMode"))          '☜: 저장시 Create/Update 판별 

If itxtFlgMode = OPMD_CMODE Then
 pvCommand = "CREATE"
ElseIf itxtFlgMode = OPMD_UMODE Then'
 pvCommand = "UPDATE"
End If
 
Set PB6CS82 = CreateObject("PB6CS82.cBSalesOrgSvr")

Call PB6CS82.B_MAINT_SALES_ORG_SVR(gStrGlobalCollection, pvCommand, _
imp_b_sales_org, imp_upper_b_sales_org)
   
If CheckSYSTEMError(Err,True) = True Then
    Set PB6CS82 = Nothing
    Exit Sub
End If  
    
Set PB6CS82 = Nothing

 Response.Write "<Script Language=vbscript>" & vbCr
 Response.Write "With parent"    & vbCr
    Response.Write ".DbSaveOk()"                & vbCr
 Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr

End Sub 

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

On Error Resume Next                                 
Err.Clear                                           
 
Dim PB6CS82 
Dim imp_b_sales_org
Dim imp_upper_b_sales_org

Const C_exp_b_sales_org_sales_org = 0

ReDim imp_b_sales_org(C_exp_b_sales_org_sales_org)

imp_b_sales_org(C_exp_b_sales_org_sales_org) = Trim(Request("txtSales_Org2"))
imp_upper_b_sales_org = Trim(Request("txtUpper_Sales_Org"))

Set PB6CS82 = CreateObject("PB6CS82.cBSalesOrgSvr")

Call PB6CS82.B_MAINT_SALES_ORG_SVR(gStrGlobalCollection, "DELETE", _
imp_b_sales_org, imp_upper_b_sales_org)

If CheckSYSTEMError(Err,True) = True Then
   Set PB6CS82 = Nothing                                                   '☜: Unload Comproxy DLL
   Exit Sub
End If  
    
Set PB6CS82 = Nothing                                                    '☜: Unload Comproxy

 Response.Write "<Script Language=vbscript>" & vbCr
 Response.Write "With parent"                & vbCr
    Response.Write ".DbDeleteOk()"              & vbCr
 Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr
  
End Sub

%>

