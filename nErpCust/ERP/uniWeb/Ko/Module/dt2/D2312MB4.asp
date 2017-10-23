<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"		-->
<!-- #Include file="../../inc/adovbs.inc"				-->
<!-- #Include file="../../inc/lgSvrVariables.inc"	-->
<!-- #Include file="../../inc/incServeradodb.asp"	-->
<!-- #Include file="../../inc/incSvrDate.inc"		-->
<!-- #Include file="../../inc/incSvrNumber.inc"		-->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Call HideStatusWnd															'☜: Hide Processing message
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q","M", "NOCOOKIE", "MB")

	Dim gtaxBillNos
	Dim nMaxRows, i
	Dim aRowData

    Dim successCount
    Dim failCount

    successCount = 0
    failCount = 0

	On Error Resume Next															'☜: Protect system from crashing
	Err.Clear																		'☜: Clear Error status

	lgErrorStatus = "NO"
	lgErrorPos    = ""															'☜: Set to space

	Call SubOpenDB(lgObjConn)

	gtaxBillNos = Split(Request("txtSpread"), gRowSep)
	
	nMaxRows = UBound(gtaxBillNos) - 1

	For i = 0 To nMaxRows
		aRowData = Split(gtaxBillNos(i), gColSep)
		lgStrSQL = "INSERT dt_simple_issue_transmit_log" & vbCrLf & _
					  "		 (inv_type, inv_no, dt_inv_no, process_date, success_flag, error_desc, " & vbCrLf & _
					  "		  re_flag, change_reason, change_remark, change_remark2, change_remark3, " & vbCrLf & _
					  "		  insert_user_id, insert_date) " & vbCrLf & _
					  "VALUES ('MM', " & _
					  				Filtervar(aRowData(0),"''","S") & "," & _
					  			  	Filtervar(aRowData(1),"''","S") & ", GETDATE(), " & _
									Filtervar(aRowData(2),"''","S") & "," & _
									Filtervar(aRowData(3),"''","S") & ", 'Y'," & _
									Filtervar(aRowData(4),"''","S") & ", " & _
									Filtervar(aRowData(5),"''","S") & ", " & _
									Filtervar(aRowData(6),"''","S") & ", " & _
									Filtervar(aRowData(7),"''","S") & ", " & _
									Filtervar(trim(gUsrId),"''","S") & ", GETDATE())"

        If aRowData(2) = "Y" Then
            successCount = successCount + 1
        Else 
            failCount = failCount + 1
        End If

		lgObjConn.Execute lgStrSQL, , adCmdText + adExecuteNoRecords
	Next

	Call SubCloseDB(lgObjConn)
%>
<Script Language=vbscript>
    MsgBox "총 " & "<%= nMaxRows + 1 %>" &"건 중 성공 " & "<%= successCount %>" & " 건, 실패 " & "<%= failCount %>" & " 건이 전송되었습니다."

	parent.FncQuery()
</Script>