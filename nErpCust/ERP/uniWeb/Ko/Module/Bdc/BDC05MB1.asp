<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
	Response.Expires = -1                               '¢Ð: will expire the response immediately
	Response.Buffer = True                              '¢Ð: The server does not send output to the client until all of the ASP 
														'    scripts on the current page have been processed
'	On Error Resume Next                                '¢Ð: Protect system from crashing
'	Err.Clear                                           '¢Ð: Clear Error status
	Dim lgErrorStatus, lgErrorPos, lgObjConn, lgObjRs
	Dim lgOpModeCRUD, iStrData
	
	Call LoadBasisGlobalInf()
	Call HideStatusWnd
	'---------------------------------------Common-----------------------------------------------------------
	lgErrorStatus   = "NO"
	lgErrorPos      = ""                                '¢Ð: Set to space
	lgOpModeCRUD    = Request("txtMode")                '¢Ð: Read Operation Mode (CRUD)

	Call SubOpenDB(lgObjConn)                           '¢Ð: Make a DB Connection
    Call SubBizQuery()
    Call SubCloseDB(lgObjConn)                          '¢Ð: Close DB Connection

Sub SubBizQuery()
	Dim iKey1
	Dim lgCARD_DD
	Dim lgStrSQL
	Dim TmpBuffer
	Dim strVal
	Dim LngRow
	
'	On Error Resume Next                                '¢Ð: Protect system from crashing
'	Err.Clear                                           '¢Ð: Clear Error status

	lgStrSQL = " SELECT job_id, job_title, job_state "& _
			   " FROM b_bdc_jobs " & _
			   " WHERE job_id = " & Filtervar(Request("txtJobId"), "''", "S")
			  
	'========================== Query Job ID ========================================
	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then           'If data not exists
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent.frm1 "                                   & vbCr
		Response.Write "	.txtJobNm.value = """"" & vbCr
		Response.Write "	.txtJobCd.focus " & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Response.End
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
	Else
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent.frm1 "                                   & vbCr
		Response.Write "	.txtJobNm.value = """ & lgObjRs("job_title") & """" & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		
		If Trim(lgObjRs("job_state")) <> "D" Then
			Response.Write "<Script Language=vbscript>"                     & vbCr
			Response.Write "With Parent.frm1 "                                   & vbCr
			Response.Write "	.txtJobNm.value = """"" & vbCr
			Response.Write "	.txtJobID.focus " & vbCr
			Response.Write "End With "                                      & vbCr
			Response.Write "</Script>"    
			'Response.End
			'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
		End If
	End If		   
	
	'========================== Query Job Result ========================================
	lgStrSQL = "SELECT seq_no, action_time, hresult, com_name, method_name " & _
			   "FROM   b_bdc_detail " & _
			   "WHERE  job_id = " & Filtervar(Request("txtJobId"), "''", "S") & _
			   " ORDER BY seq_no "

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then           'If data not exists
		Response.End
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
	Else
		
		ReDim TmpBuffer(0)		'lgObjRs.RecordCount
		
		LngRow = 0
		
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
            strVal =  Chr(11) & lgObjRs("seq_no")
            strVal = strVal & Chr(11) & lgObjRs("action_time")
            strVal = strVal & Chr(11) & lgObjRs("hresult")
            strVal = strVal & Chr(11) & lgObjRs("com_name")
            strVal = strVal & Chr(11) & Replace(lgObjRs("method_name"), Chr(10), "")
            strVal = strVal & Chr(11) & Chr(12)
            
            Redim Preserve TmpBuffer(LngRow)
            TmpBuffer(LngRow)  = strVal 
            LngRow = LngRow + 1
			lgObjRs.MoveNext
		Loop
		
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData "        & vbCr
		
		Response.Write "    .ggoSpread.SSShowData """ & replace(Join(TmpBuffer, ""),chr(13),"") & """" & vbCr
		Response.Write "    .frm1.vspdData.ReDraw = True "              & vbCr   
		Response.Write "    .DbQueryOk  "                               & vbCr    
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
	End If

	Call SubCloseRs(lgObjRs)                                        '¢Ð : Release RecordSSet
End Sub
%>
