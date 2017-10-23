<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 공장별 일별 품목 판매계획확정 
'*  3. Program ID           : S2216BB1
'*  4. Program Name         : 
'*  5. Program Desc         : 판매계획관리 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/01/15
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
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
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim intRetCD

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IntRetCD = 0

    With lgObjComm
		.CommandTimeout = 0
		If Request("txtWorkType") = "Y" Then
			.CommandText = "dbo.usp_s_CfmSSpItemByPlantDaily"
        Else
			.CommandText = "dbo.usp_s_CancelCfmSSpItemByPlantDaily"
        End If
        .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
		.Parameters.Append .CreateParameter("@fr_sp_period", adVarXChar,adParamInput,8,Replace(Request("txtFrSpPeriod"), "'", "''"))
		If Request("txtWorkType") = "Y" Then
			.Parameters.Append .CreateParameter("@to_sp_period", adVarXChar,adParamInput,8,Replace(Request("txtToSpPeriod"), "'", "''"))
		    .Parameters.Append .CreateParameter("@fc_sp_period", adVarXChar,adParamInput,8,Replace(Request("txtFcSpPeriod"), "'", "''"))
		End If
	    .Parameters.Append .CreateParameter("@loc_exp_flag", adXChar,adParamInput,1,"1")
	    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,Replace(Request("txtSalesGrp"), "'", "''"))
	    .Parameters.Append .CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
	    .Parameters.Append .CreateParameter("@grp_flag", adXChar,adParamInput,1,Request("txtGrpFlag"))
		.Parameters.Append .CreateParameter("@called_flag", adXChar,adParamInput,1,"N")

        .Execute ,, adExecuteNoRecords
        
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
    End If

End Sub	

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

