<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
 
 call LoadBasisGlobalInf()
 
 Dim lgOpModeCRUD
 
 On Error Resume Next															    '☜: Protect system from crashing
 Err.Clear																			'☜: Clear Error status

 Call HideStatusWnd

 lgOpModeCRUD = Request("txtMode")													 '☜: Read Operation Mode (CRUD)

 Select Case lgOpModeCRUD
         Case CStr(UID_M0003)                                                         '☜: Delete
              Call SubBizDelete()
          
 End Select

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()             '☜: 삭제 요청 
  Dim iPM5G111                ' 수출 B/L Header 삭제용 Object
  Dim I3_m_bl_hdr
  Dim str_txtBeneficiary
  Dim str_txtApplicant
  Dim str_txtPurGrp
  
  On Error Resume Next
  Err.Clear 

  Set iPM5G111 = Server.CreateObject("PM5G111.cMMaintImportBlHdrS")

  If CheckSYSTEMError(Err,True) = True Then
	Set iPM5G111 = Nothing
	Exit Sub
  End If
  
  Redim I3_m_bl_hdr(0)
  I3_m_bl_hdr(0)=UCase(Trim(Request("txtBLNo")))
  
 str_txtBeneficiary = UCase(Trim(Request("txtBeneficiary")))
 str_txtApplicant   = UCase(Trim(Request("txtApplicant")))
 str_txtPurGrp      = UCase(Trim(Request("txtPurGrp")))
 
 
  Call iPM5G111.M_MAINT_IMPORT_BL_HDR_SVR(gStrGlobalCollection, "DELETE", _
                                     str_txtBeneficiary, str_txtApplicant, _
                                     I3_m_bl_hdr, str_txtPurGrp)

  If CheckSYSTEMError(Err,True) = True Then
	Set iPM5G111 = Nothing
	Exit Sub
  End If

  Set iPM5G111 = Nothing              '☜: Unload Comproxy


  Response.Write "<Script Language=VBScript>" & vbCr
  Response.Write " With parent"    & vbCr
  Response.Write "  .DbDeleteOk"   & vbCr
  Response.Write " End With"      & vbCr
  Response.Write "</Script>"      & vbCr

  Set iPM5G111 = Nothing              '☜: Unload Comproxy

End Sub

%>
