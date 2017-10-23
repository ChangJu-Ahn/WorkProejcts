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
	Call HideStatusWnd															'��: Hide Processing message
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

	On Error Resume Next															'��: Protect system from crashing
	Err.Clear																		'��: Clear Error status

	lgErrorStatus = "NO"
	lgErrorPos    = ""															'��: Set to space

	Call SubOpenDB(lgObjConn)

	gtaxBillNos = Split(Request("txtSpread"), gRowSep)
	
	nMaxRows = UBound(gtaxBillNos) - 1

	For i = 0 To nMaxRows
		aRowData = Split(gtaxBillNos(i), gColSep)
		lgStrSQL = "INSERT dt_simple_issue_transmit_log" & vbCrLf & _
					  "		 (inv_type, inv_no, dt_inv_no, process_date, success_flag, error_desc, " & vbCrLf & _
					  "		  re_flag, change_reason, change_remark, change_remark2, change_remark3, " & vbCrLf & _
					  "		  insert_user_id, insert_date) " 	  & vbCrLf & _
					  "VALUES (" & Filtervar(aRowData(0),"''","S") & "," & _
					  					Filtervar(aRowData(1),"''","S") & "," & _
										Filtervar(aRowData(2),"''","S") & ", GETDATE(), " & _
										Filtervar(aRowData(3),"''","S") & "," & _
										Filtervar(aRowData(4),"''","S") & ", 'N'," & _
										Filtervar(aRowData(5),"''","S") & "," & _
										Filtervar(aRowData(6),"''","S") & "," & _
										Filtervar(aRowData(7),"''","S") & "," & _
										Filtervar(aRowData(8),"''","S") & "," & _
										Filtervar(trim(gUsrId),"''","S") & ", GETDATE())" & vbCrLf & _
					  vbCrLf & _
           		  "UPDATE a_vat" & vbCrLf & _
 					  " 	SET issue_dt_kind = (SELECT minor_cd FROM b_configuration WHERE major_cd = 'DT004' AND reference = 'Y')," & vbCrLf & _
					  "     	 issue_dt_fg = 'Y'," & vbCrLf & _
                 "     	 updt_user_id = " & Filtervar(trim(gUsrId),"''","S") & "," & vbCrLf & _
                 "     	 updt_dt = GETDATE()" & vbCrLf & _
					  " WHERE ref_no = (SELECT tax_doc_no FROM s_tax_bill_hdr WHERE tax_bill_no = " & Filtervar(aRowData(1),"''","S") & ")"

        If aRowData(3) = "Y" Then
            successCount = successCount + 1
        Else 
            failCount = failCount + 1
        End If

		lgObjConn.Execute lgStrSQL, , adCmdText + adExecuteNoRecords
	Next

	Call SubCloseDB(lgObjConn)
%>
<Script Language=vbscript>
    MsgBox "�� " & "<%= nMaxRows + 1 %>" &"�� �� ���� " & "<%= successCount %>" & " ��, ���� " & "<%= failCount %>" & " ���� ���۵Ǿ����ϴ�."

	parent.FncQuery()
</Script>