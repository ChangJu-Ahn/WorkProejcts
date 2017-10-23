<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%

	Dim lgOpModeCRUD

	On Error Resume Next														'��: Protect system from crashing
	Err.Clear 																	'��: Clear Error status
	
	Call LoadBasisGlobalInf()
	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")										'��: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD
	     Case CStr(UID_M0003)                                                   '��: Delete
	          Call SubBizDelete()
	End Select

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================

	Sub SubBizDelete()

		Dim OBJ_PM42111																' Master L/C Header Save�� Object
		Dim expMCcHdrCcNo
		Dim iCommandSent
		
        Dim I1_b_pur_grp
        Dim I2_b_biz_partner
        Dim I3_b_biz_partner
        Dim I5_m_cc_hdr

		Const M410_I5_cc_no = 0														'import m_cc_hdr
		
		Redim I5_m_cc_hdr(M410_I5_cc_no)

		On Error Resume Next
		Err.Clear																'��: Protect system from crashing

		If Request("txtCCNo") = "" Then									
			Call DisplayMsgBox("700112", vbInformation,	"", "",	I_MKSCRIPT)
			Exit Sub 
		End If
		
		iCommandSent = "DELETE"
		
		I5_m_cc_hdr(M410_I5_cc_no) = UCase(Trim(Request("txtCCNo")))
		
		Set OBJ_PM42111 = Server.CreateObject("PM6G111.cMMaintImportCcHdrS")

		If CheckSYSTEMError(Err,True) = True Then
			Exit Sub
		End If

		expMCcHdrCcNo =  OBJ_PM42111.M_MAINT_IMPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, _
            I1_b_pur_grp, I2_b_biz_partner, I3_b_biz_partner, I5_m_cc_hdr)
                        
		If CheckSYSTEMError(Err,True) = True Then
			Set OBJ_PM42111 = Nothing
			Exit Sub
		End If

		Set OBJ_PM42111 = Nothing														'��: Unload Comproxy
		'-----------------------
		'Result data display area
		'-----------------------
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent"                & vbCr
		Response.Write "    .DbDeleteOk"            & vbCr
		Response.Write "End With"                   & vbCr
		Response.Write "</Script>"                  & vbCr
		
		
	End Sub																				'��: Process End
%>
