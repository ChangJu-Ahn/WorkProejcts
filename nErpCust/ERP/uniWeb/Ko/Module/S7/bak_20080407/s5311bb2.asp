<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","BB")%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
Call HideStatusWnd                                                               '☜: Hide Processing message
Call SubCreateCommandObject(lgObjComm)
Call SubBizBatch()
Call SubCloseCommandObject(lgObjComm)
  
'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizBatch()

    Dim intRetCD

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IntRetCD = 0

    With lgObjComm
		.CommandTimeout = 0
		
		IF Request("txtWkFlag") = "E" Then
			.CommandText = "dbo.usp_s_CreateAVat_Batch"
		Else
			.CommandText = "dbo.usp_s_DeleteAVat_Batch"
		End If
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_issued_dt", adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtFromDt")))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_issued_dt", adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtToDt")))

	    ' 발행처 
	    IF Len(Trim(Request("txtBillToParty"))) Then
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bp_cd", adVarXChar,adParamInput,10,Replace(Request("txtBillToParty"), "'", "''"))
		Else
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bp_cd", adVarXChar,adParamInput,10)
	    End If
	    ' 세금신고사업장 
	    IF Len(Trim(Request("txtTaxbizArea"))) Then
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@report_biz_area", adVarXChar,adParamInput,10,Replace(Request("txtTaxbizArea"), "'", "''"))
		Else
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@report_biz_area", adVarXChar,adParamInput,10)
	    End If
	    ' 영업그룹 
	    IF Len(Trim(Request("txtSalesGrp"))) Then
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sales_grp", adVarXChar,adParamInput,4,Replace(Request("txtSalesGrp"), "'", "''"))
		Else
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sales_grp", adVarXChar,adParamInput,4)
	    End If
	    ' User ID
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",  adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))

		IF Request("txtWkFlag") = "E" Then
			' 작업일 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@insrt_dt", adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtWorkDt")))
			' 호출여부 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CalledFlag",  adXChar,adParamInput,13,"N")
		End If

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With
    
    If CheckSYSTEMError(Err,True) = True Then
       IntRetCD = -1
       ObjectContext.SetAbort
       Exit Sub
    End If
    
    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
    
    If CDbl(intRetCD) = 0 Then
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Parent.ExeReflectOk  " & vbCr
       Response.Write  " </Script>                  " & vbCr
    Else
       Call DisplayMsgBox(IntRetCd, vbInformation, "", "", I_MKSCRIPT)
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Call parent.SetFocusToDocument(""M"")  " & vbCr
       Response.Write  "       parent.frm1.txtFromDt.Focus  " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If

End Sub	

'============================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

