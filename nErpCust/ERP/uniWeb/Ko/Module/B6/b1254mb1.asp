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
	Call HideStatusWnd               
    '---------------------------------------Common-----------------------------------------------------------
 
    lgOpModeCRUD = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

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

    Dim pB27019
    Dim iSalesGrp

    Dim E1_b_biz_area 
    Dim E2_b_cost_center 
    Dim E3_b_sales_grp 
    Dim E4_b_sales_org    
    
    Const EG1_E1_b_sales_grp_sales_grp1 = 0
    Const EG1_E1_b_sales_grp_sales_grp_nm1 = 1
    Const EG1_E1_b_sales_grp_usage_flag1 = 2
    Const EG1_E1_b_sales_grp_sales_grp_full_nm1 = 3
    Const EG1_E1_b_sales_grp_sales_grp_eng_nm1 = 4
    
    Const EG1_E2_b_sales_org_sales_org1 = 0
    Const EG1_E2_b_sales_org_sales_org_nm1 = 1
    
    Const EG1_E3_b_cost_center_cost_cd1 = 0
    Const EG1_E3_b_cost_center_cost_nm1 = 1
    
    On Error Resume Next
    Err.Clear 
    
    iSalesGrp = Trim(Request("txtSales_Grp2"))
    
    Set pB27019 = CreateObject("PB6CS81.cBLookupSalesGrpSvr")   

   If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
    
    Call pB27019.B_LOOKUP_SALES_GRP_SVR ( gStrGlobalCollection, Trim("Query"), cstr(isalesGrp), _
                                               E1_b_biz_area, E2_b_cost_center, E3_b_sales_grp , E4_b_sales_org )
    
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "parent.frm1.txtSales_Grp_nm1.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp_nm1))  & """" & vbCr
    Response.Write "</Script>"         & vbCr
    
 If CheckSYSTEMError(Err,True) = True Then
       Set pB27019 = Nothing
       Exit Sub
    End If  
    
    Set pB27019 = Nothing
    
%>
<Script Language=vbscript>
 With parent.frm1

  If "<%=E3_b_sales_grp(EG1_E1_b_sales_grp_usage_flag1)%>" = .rdoUsage_flag1.value Then
   .rdoUsage_flag1.checked = True
   .txtRadio.value = .rdoUsage_flag1.value
  ElseIf "<%=E3_b_sales_grp(EG1_E1_b_sales_grp_usage_flag1)%>" = .rdoUsage_flag2.value Then
   .rdoUsage_flag2.checked = True
   .txtRadio.value = .rdoUsage_flag2.value   
  End IF
  
 End With
</Script>
<%    
    
 '-----------------------
 'Display result data
 '----------------------- 
 Response.Write "<Script Language=vbscript>" & vbCr
 Response.Write "With parent.frm1"           & vbCr

 Response.Write ".txtSales_Grp1.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp1))            & """" & vbCr
 Response.Write ".txtSales_Grp_nm1.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp_nm1))            & """" & vbCr

    Response.Write ".txtSales_Grp2.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp1))   & """" & vbCr
 Response.Write ".txtSales_Grp_nm2.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp_nm1)) & """" & vbCr

 Response.Write ".txtSales_Org_Fullnm.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp_full_nm1))          & """" & vbCr
 Response.Write ".txtSales_Org_Engnm.Value = """ & ConvSPChars(E3_b_sales_grp(EG1_E1_b_sales_grp_sales_grp_eng_nm1))       & """" & vbCr 

 Response.Write ".txtCost_center.Value = """ & ConvSPChars(E2_b_cost_center(EG1_E3_b_cost_center_cost_cd1))            & """" & vbCr
 Response.Write ".txtCost_center_nm.Value = """ & ConvSPChars(E2_b_cost_center(EG1_E3_b_cost_center_cost_nm1))            & """" & vbCr
    Response.Write ".txtSales_Org.Value = """ & ConvSPChars(E4_b_sales_org(EG1_E2_b_sales_org_sales_org1))   & """" & vbCr
    Response.Write ".txtSales_Org_Nm.Value = """ & ConvSPChars(E4_b_sales_org(EG1_E2_b_sales_org_sales_org_nm1))   & """" & vbCr
 
 Response.Write "parent.DbQueryOk" & vbCr
 Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
 
 'Response.Write " Test Ending "
 
End Sub 

'============================================================================================================
Sub SubBizSave()

    Dim pvCommand
    Dim itxtFlgMode
    Dim iS27011 
    
    Dim E1_S_sales_grp

    Dim I1_b_cost_center
    Dim I3_b_sales_grp
    Dim I4_b_sales_org
    
    Redim I1_b_cost_center(0)   
    ReDim I3_b_sales_grp(4)
    ReDim I4_b_sales_org(0)
    
    Dim I1_b_cost_center1
    Dim I3_b_sales_grp1
    Dim I4_b_sales_org1    
    
    Const C_I1_cost_cd = 0
    Const C_I3_b_sales_grp = 0
    Const C_I3_sales_grp_nm = 1
    Const C_I3_usage_flag = 2
    Const C_I3_sales_grp_full_nm = 3
    Const C_I3_sales_grp_eng_nm = 4
    Const C_I4_sales_org = 0

    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status
   
    I1_b_cost_center(C_I1_cost_cd) = UCase(Trim(Request("txtCost_center")))
    I3_b_sales_grp(C_I3_b_sales_grp) = UCase(Trim(Request("txtSales_Grp2")))
    I3_b_sales_grp(C_I3_sales_grp_nm) = Trim(Request("txtSales_Grp_nm2"))
    I3_b_sales_grp(C_I3_usage_flag) =  Trim(Request("txtRadio"))
    I3_b_sales_grp(C_I3_sales_grp_full_nm) = Trim(Request("txtSales_Org_Fullnm"))
    I3_b_sales_grp(C_I3_sales_grp_eng_nm) = Trim(Request("txtSales_Org_Engnm"))
    I4_b_sales_org(C_I4_sales_org) = UCase(Trim(Request("txtSales_Org")))

    I1_b_cost_center1 = I1_b_cost_center
    I3_b_sales_grp1 = I3_b_sales_grp
    I4_b_sales_org1 = I4_b_sales_org

    itxtFlgMode = CInt(Request("txtFlgMode"))

    If itxtFlgMode = OPMD_CMODE Then
  pvCommand = "CREATE"
    ElseIf itxtFlgMode = OPMD_UMODE Then
  pvCommand = "UPDATE"
    End If
    
    Set iS27011 = server.CreateObject("PB6CS80.cB27011MaintSalesGrpSvr")
           
 If CheckSYSTEMError(Err,True) = True Then
       Set iS27011 = Nothing                                                   '☜: Unload Comproxy DLL
       Exit Sub
    End If      
       
    E1_S_sales_grp = iS27011.B_MAINT_SALES_GRP_SVR (gStrGlobalCollection, pvCommand, _
                                        I1_b_cost_center1, I3_b_sales_grp1, I4_b_sales_org1)
 
    
 If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS27011 = Nothing                                                   '☜: Unload Comproxy DLL
       Exit Sub
    End If
    
    Set iS27011 = Nothing

 Response.Write "<Script Language=vbscript>" & vbCr
 Response.Write "With parent"           & vbCr
 'b1256ma1에서 같이쓰이므로 구분짓기위해 아래사항을 안하면 b1256ma1에서 에러가 남.
 '20021224 강준구 
 if Trim(Request("txtprogramId")) <> "b1256ma1" then
	If E1_S_sales_grp <>"" Then
	    Response.Write ".frm1.txtSales_Grp2.value = """ & ConvSPChars(E1_S_sales_grp)     & """" & vbCr
	else
	    Response.Write ".frm1.txtSales_Grp2.value =  .frm1.txtSales_Grp1.value           " & vbCr
	end if
end if	
    Response.Write ".DbSaveOk"                  & vbCr
    Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr

End Sub 
     

'============================================================================================================
Sub SubBizDelete()

    Dim iS27011 
    
    Dim E1_S_sales_grp    
    Dim L1_S_sales_grp
    Dim L1_S_sales_grp1
    
    Const C_I1_Sales_grp = 0 

    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    Redim L1_S_sales_grp(C_I1_Sales_grp) 
   
    L1_S_sales_grp(C_I1_Sales_grp) = Trim(Request("txtSales_Grp2"))
     
    If L1_S_sales_grp(C_I1_Sales_grp) = "" Then  
  Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
 End If
 
 L1_S_sales_grp1 = L1_S_sales_grp
 
    Set iS27011 = server.CreateObject("PB6CS80.cB27011MaintSalesGrpSvr")
           
 If CheckSYSTEMError(Err,True) = True Then
       Set iS27011 = Nothing                                                   '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    E1_S_sales_grp = iS27011.B_MAINT_SALES_GRP_SVR (gStrGlobalCollection, "DELETE", _
                                        "", L1_S_sales_grp1, "")
    
 If CheckSYSTEMError(Err,True) = True Then
       Set iS27011 = Nothing                                                   '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
 Set iS27011 = Nothing                                                                    '☜: Unload Comproxy
 
 Response.Write "<Script Language=vbscript>" & vbCr
 Response.Write "With parent"                & vbCr
 Response.Write ".DbDeleteOk"                & vbCr
 Response.Write "End With"                   & vbCr
 Response.Write "</Script>"                  & vbCr
      
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

