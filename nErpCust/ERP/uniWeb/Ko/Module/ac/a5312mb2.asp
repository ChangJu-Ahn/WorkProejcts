<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
	Dim strYYYYMM
	Dim strVerCd
	Dim iStrData

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear  
    
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")										'☜: Clear Error status
    Call HideStatusWnd 
																						'☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""																'☜: Set to space
    lgOpModeCRUD      = Request("txtMode")												'☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
													'☜: Batch
															'☜: Query
     Call SubBizBatch()

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
	Dim iPACG060																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim I1_exch_result_posting
	Dim iCommandSent

	Const A064_I1_exch_result_posting_yyyymm = 0
	Const A064_I1_exch_result_posting_module_cd = 1
	Const A064_I1_exch_result_posting_biz_area_cd = 2
	Const A064_I1_exch_result_posting_gl_dt = 3
	Const A064_I1_exch_result_posting_org_change_id = 4
	Const A064_I1_exch_result_posting_dept_cd = 5
	Const A064_I1_exch_result_posting_temp_gl_no = 6
	Const A064_I1_exch_result_posting_gl_no = 7
	Const A064_I1_exch_result_posting_reverse_gl_dt = 8
	Const A064_I1_exch_result_posting_rev_temp_gl_no = 9
	Const A064_I1_exch_result_posting_rev_gl_no = 10
	
	On Error Resume Next 

	Redim I1_exch_result_posting(A064_I1_exch_result_posting_rev_gl_no)

	I1_exch_result_posting(A064_I1_exch_result_posting_yyyymm)         = Trim(Request("txtYYYYMM"))
	I1_exch_result_posting(A064_I1_exch_result_posting_module_cd)      = Trim(Request("txtModuleCd"))
	I1_exch_result_posting(A064_I1_exch_result_posting_biz_area_cd)    = Trim(Request("txtBizAreaCd"))
	I1_exch_result_posting(A064_I1_exch_result_posting_gl_dt)          = Trim(Request("txtGLDt"))
	I1_exch_result_posting(A064_I1_exch_result_posting_org_change_id)  = Trim(Request("txtOrgChangeId"))
	I1_exch_result_posting(A064_I1_exch_result_posting_dept_cd)        = Trim(Request("txtDeptCd"))
	I1_exch_result_posting(A064_I1_exch_result_posting_temp_gl_no)     = Trim(Request("txtTempGLNo"))
    I1_exch_result_posting(A064_I1_exch_result_posting_gl_no)          = Trim(Request("txtGLNo")) 
	I1_exch_result_posting(A064_I1_exch_result_posting_reverse_gl_dt)  = Trim(Request("txtRevGLDt"))
	I1_exch_result_posting(A064_I1_exch_result_posting_rev_temp_gl_no) = Trim(Request("txtRevTempGLNo"))
	I1_exch_result_posting(A064_I1_exch_result_posting_rev_gl_no)      = Trim(Request("txtRevGLNo"))
	
	Select Case lgOpModeCRUD
        Case CStr(UID_M0002)															'☜: Query
			iCommandSent= "C"
        Case CStr(UID_M0003)															'☜: Query
            iCommandSent= "D"
    End Select															
 
	Set iPACG060 = Server.CreateObject("PACG060.cAExchangeJobToGLSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	
	Call iPACG060.A_Exchange_ToGLSvr (gStrGlobalCollection, iCommandSent, I1_exch_result_posting )
	
	If CheckSYSTEMError(Err,True) = True Then
		Set iPACG060 = Nothing		
		Response.End 
	End If
    
	Set iPACG060 = Nothing 
End Sub
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
		Case "<%=UID_M0002%>"                                                         '☜ : Query
			With Parent
				.ExeReflectOk
			End With
		Case "<%=UID_M0003%>"                                                         '☜ : Batch
			Parent.ExeCancleOk
    End Select
</Script>
