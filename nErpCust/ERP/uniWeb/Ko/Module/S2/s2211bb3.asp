<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2211BB3
'*  4. Program Name         : 판매계획기간정보생성 
'*  5. Program Desc         : 판매계획기간정보생성 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/10
'*  8. Modified date(Last)  : 2003/01/10
'*  9. Modifier (First)     : Park yongsik
'* 10. Modifier (Last)      : Park yongsik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

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
        .CommandText = "dbo.usp_s_CreateSalesPlanPeriod"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_dt",  adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtFromDt")))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_dt",  adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtToDt")))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_type", adXChar,adParamInput,1,Request("txtSpType"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@method", adXChar,adParamInput,2,Request("txtMethod"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
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

