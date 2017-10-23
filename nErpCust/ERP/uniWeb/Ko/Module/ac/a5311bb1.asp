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

    'Multi SpreadSheet
    lgLngMaxRow       = Request("txtMaxRows")											'☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)					'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	
	Const C_SHEETMAXROWS_D  = 100        
	lgMaxCount = CInt(C_SHEETMAXROWS_D)													'☜: Max fetched data at a time

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)															'☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)															'☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)															'☜: Delete
             Call SubBizDelete()
        Case CStr(UID_M0006)															'☜: Batch
           Call SubBizBatch()
    End Select

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
	Dim iPACG060																	'☆ : 조회용 ComProxy Dll 사용 변수 

	Dim I1_yyyymm  
	Dim I2_ver_cd

	'#########################################################################################################
	'												2.2. 요청 변수 처리 
	'##########################################################################################################

	I1_yyyymm = Trim(Request("txthWork_dt"))
	I2_ver_cd = UCASE(Trim(Request("txtVerCd")))

	'#########################################################################################################
	'												2.3. 업무 처리 
	'##########################################################################################################

	Set iPACG060 = Server.CreateObject("PACG060.cAExchangeJobBatchWithSpSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

	Call iPACG060.A_EXCHANGE_JOB_BATCH_WITH_SP_SVR (gStrGlobalCollection, I1_yyyymm,I2_ver_cd,Trim(Request("txtSpread")))

	If CheckSYSTEMError(Err,True) = True Then
		Set iPACG060 = Nothing		
		Exit Sub
	End If
	    
	Set iPACG060 = Nothing 
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iPACG060																	'☆ : 조회용 ComProxy Dll 사용 변수 

	Dim I1_yyyymm  
	Dim I2_ver_cd
	Dim EG1_exchange_job_info
	
	Dim iLngRow

    On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear 
	
    Const A65_EG1_exchange_job_info_chk_fg = 0
    Const A65_EG1_exchange_job_info_progress_fg = 1
    Const A65_EG1_exchange_job_info_progress_nm = 2
	Const A65_EG1_exchange_job_info_module_cd = 3
	Const A65_EG1_exchange_job_info_module_nm = 4
	Const A65_EG1_exchange_job_info_error_cnt = 5

	'#########################################################################################################
	'												2.2. 요청 변수 처리 
	'##########################################################################################################

'	LngMaxRow = Cint(Request("txtMaxRows"))
	I1_yyyymm = Trim(Request("txtYYYYMM"))
	I2_ver_cd = Trim(Request("txtVerCd"))

	'#########################################################################################################
	'												2.3. 업무 처리 
	'##########################################################################################################

	Set iPACG060 = Server.CreateObject("PACG060.cALkupExchangeJobSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

	Call iPACG060.A_LIST_EXCHANGE_JOB_SVR (gStrGlobalCollection, I1_yyyymm,I2_ver_cd,EG1_exchange_job_info)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPACG060 = Nothing		
		Exit Sub
	End If
	    
	Set iPACG060 = Nothing 

	'#########################################################################################################
	'												2.4. HTML 결과 생성부 
	'##########################################################################################################
	Response.Write "<Script Language=vbscript>										" & vbcr
	Response.Write " With parent.frm1                                               " & vbcr 
'	Response.Write " .txtWork_dt.Text			= """ & ConvSPChars(I1_yyyymm) & """" & vbcr
'	Response.Write " .txtVerCd.Value			= """ & ConvSPChars(I2_ver_cd) & """" & vbcr
	Response.Write " .txthWork_dt.Value			= """ & ConvSPChars(I1_yyyymm) & """" & vbcr
	Response.Write " .txthVerCd.Value			= """ & ConvSPChars(I2_ver_cd) & """" & vbcr
	Response.Write " End With														" & vbcr		    
	Response.write "</Script>														" & vbcr  

	iStrData = ""

	For iLngRow = 0 To UBound(EG1_exchange_job_info,1)
		iStrData = iStrData & Chr(11) & EG1_exchange_job_info(iLngRow, A65_EG1_exchange_job_info_chk_fg)
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exchange_job_info(iLngRow, A65_EG1_exchange_job_info_progress_fg))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exchange_job_info(iLngRow, A65_EG1_exchange_job_info_progress_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exchange_job_info(iLngRow, A65_EG1_exchange_job_info_module_cd))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exchange_job_info(iLngRow, A65_EG1_exchange_job_info_module_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exchange_job_info(iLngRow, A65_EG1_exchange_job_info_error_cnt))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & iLngRow + 1 
		iStrData = iStrData & Chr(11) & Chr(12)
	Next                                                         '☜: Release RecordSSet
End Sub

%>

<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
		Case "<%=UID_M0001%>"                                                         '☜ : Query
			With Parent
				.ggoSpread.Source     = .frm1.vspdData
				.ggoSpread.SSShowData   "<%=iStrData%>"
				.DBQueryOk
			End With
		Case "<%=UID_M0006%>"                                                         '☜ : Batch
			Parent.ExeReflectOk
    End Select
</Script>
