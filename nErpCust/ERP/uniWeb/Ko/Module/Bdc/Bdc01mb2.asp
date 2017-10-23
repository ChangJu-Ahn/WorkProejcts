<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->

<%	
Dim importArray
Dim szSpData1
Dim szSpData2
Dim szSpData3
Dim szSpData4
Dim pBDC001																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim szRetValue

Const C_import_b_process_id  = 0
Const C_import_b_process_nm  = 1
Const C_import_b_use_flag    = 2
Const C_import_b_join_method = 3
Const C_import_b_tran_flag   = 4
Const C_import_b_start_row   = 5
Const C_import_b_run_time    = 6

ReDim importArray(C_import_b_run_time)										'⊙: Single 데이타 저장 

Call LoadBasisGlobalInf() 

Call HideStatusWnd

importArray(C_import_b_process_id)	= Trim(Ucase(Request("txtProcID")))
importArray(C_import_b_process_nm)	= Trim(Request("txtNProcNm"))
importArray(C_import_b_use_flag)	= Trim(Request("hUseFlag"))
importArray(C_import_b_join_method) = Trim(Request("hJoinMethod"))
importArray(C_import_b_tran_flag)	= Trim(Request("hTranFlag"))
importArray(C_import_b_start_row)	= Trim(Request("txtStartRow"))
importArray(C_import_b_run_time)	= Trim(Request("txtRunTime"))

szSpData1 = Request("txtSpread")
szSpData2 = Request("txtSpread1")
szSpData3 = Request("txtSpread2")
szSpData4 = Request("txtSpread3")

On Error Resume Next
Set pBDC001 = Server.CreateObject("BDC001.cBDCMaster")	    	    

If Err.Number <> 0 Then
	Set pBDC001 = Nothing
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
	Response.End															'☜: 비지니스 로직 처리를 종료함 
End If

szRetValue = pBDC001.CreateProcess( gStrGlobalCollection, _
									importArray, _
									CStr(szSpData1), _
									CStr(szSpData2), _
									CStr(szSpData3), _
									CStr(szSpData4))

If CheckSYSTEMError(Err,True) = True Then
	Set pBDC001 = Nothing													'☜: ComProxy Unload
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
	Response.End															'☜: 비지니스 로직 처리를 종료함 
End If

Set pBDC001 = Nothing														'☜: Unload Comproxy
%>
<SCRIPT LANGUAGE=vbscript>
	With parent
		.frm1.txtProcID.value = "<%=Trim(szRetValue)%>"
		.frm1.hProcID.value = "<%=Trim(szRetValue)%>"
		.DbSaveOk
	End With
</SCRIPT>
