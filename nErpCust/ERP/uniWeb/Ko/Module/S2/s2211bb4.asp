<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2211BA4
'*  4. Program Name			: 품목별공장배분비일괄생성 
'*  5. Program Desc         : 품목별공장배분비일괄생성 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/01/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB")%>


<%
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
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
		.CommandText = "dbo.usp_s_CreateBPlantRateByItem"
        .CommandType = adCmdStoredProc
        
	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
	        
	    
		If Len(Request("txtItemGroupCd")) Then
			.Parameters.Append .CreateParameter("@item_group_cd", adVarXChar,adParamInput,10,Replace(Request("txtItemGroupCd"), "'", "''"))
		Else
			.Parameters.Append .CreateParameter("@item_group_cd", adVarXChar,adParamInput,10)
		End If
		
		If Len(Request("txtItemAcct")) Then
						
			.Parameters.Append .CreateParameter("@item_acct", adXChar,adParamInput,2,Replace(Request("txtItemAcct"), "'", "''"))
		Else
			.Parameters.Append .CreateParameter("@item_acct", adXChar,adParamInput,2)
		End If

		If Len(Request("txtItemCd")) Then
			.Parameters.Append .CreateParameter("@item_cd", adVarXChar,adParamInput,18,Replace(Request("txtItemCd"), "'", "''"))
		Else
			.Parameters.Append .CreateParameter("@item_cd", adVarXChar,adParamInput,18)
		End If

	    .Parameters.Append .CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))

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

