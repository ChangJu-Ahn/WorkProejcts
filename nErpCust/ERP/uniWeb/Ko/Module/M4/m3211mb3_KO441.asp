<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%	
call LoadBasisGlobalInf()

	Dim lgOpModeCRUD
	On Error Resume Next
				'☜: Protect system from crashing
	Err.Clear 
				'☜: Clear Error status

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")
				'☜: Read Operation Mode (CRUD)


	'Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	'Dim M32119	
	'Dim M32111

	'strMode = Request("txtMode")													'☜ : 현재 상태를 받음 
	Select Case lgOpModeCRUD
	       ' Case CStr(UID_M0001)                                                         '☜: Query
	       '      Call SubBizQuery()
	       ' Case CStr(UID_M0002)
	        '     Call SubBizSave()
	        Case CStr(UID_M0003)                                                         '☜: Delete
	             Call SubBizDelete()
	End Select

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================

Sub SubBizDelete()
	Dim iPM4G111
	
	Dim I1_b_biz_partner
    Dim I2_b_biz_partner
    Dim I3_b_pur_grp
    Dim I4_m_lc_hdr
	Dim I5_b_bank
    Dim I6_b_bank
    Dim I7_b_bank
    Dim I8_b_bank
    Dim I9_b_bank
    Dim I10_s_wks_user
    
    Const M468_I12_lc_no = 0
    
    Redim I4_m_lc_hdr(91)
    
	On Error Resume Next
	Err.Clear 

	If Request("txtLCNo") = "" Then											'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		exit sub
	End If

    I4_m_lc_hdr(M468_I12_lc_no) = UCase(Trim(Request("txtLCNo")))

	Set iPM4G111 = Server.CreateObject("PM4G111.cMMaintLcHdrS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    Call iPM4G111.M_MAINT_LC_HDR_SVR(gStrGlobalCollection,"DELETE",I1_b_biz_partner,I2_b_biz_partner, _
            I3_b_pur_grp,I4_m_lc_hdr,I5_b_bank,I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_s_wks_user)       

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM4G111 = Nothing

	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.DbDeleteOk "      & vbCr   
    Response.Write "</Script> "            
End Sub	
%>
