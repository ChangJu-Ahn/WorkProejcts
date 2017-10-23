<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 영업조직/그룹정보등록 
'*  3. Program ID           : b1256mb1
'*  4. Program Name         : 영업조직/그룹정보등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seong Bae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Call LoadBasisGlobalInf()
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
    Dim iStrUpperSalesOrg

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IntRetCD = 0

    With lgObjComm
		.CommandTimeout = 0
		
		Select Case Request("txtFlag")
		' 영업 그룹의 상위 조직이 Update되는 경우 
		Case "GRP"
			.CommandText = "UPDATE dbo.b_sales_grp " & _
						  "SET sales_org =  " & FilterVar(Request("txtSalesOrg"), "''", "S") & ", " & _
						  "	   updt_user_id =  " & FilterVar(Request("txtUserId"), "''", "S") & ", " & _
						  "    updt_dt = GETDATE() " & _
						  "WHERE sales_grp =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""
		' 영업 조직의 상위 조직만 Update되는 경우 
		Case "ORG1"
			.CommandText = "UPDATE dbo.b_sales_org " & _
						  "SET upper_sales_org =  " & FilterVar(Request("txtUpperSalesOrg"), "''", "S") & ", " & _
						  "	   updt_user_id =  " & FilterVar(Request("txtUserId"), "''", "S") & ", " & _
						  "    updt_dt = GETDATE() " & _
						  "WHERE sales_org =  " & FilterVar(Request("txtSalesOrg"), "''", "S") & ""
		' 영업 조직의 상위 조직과 Level이 변경되나 하위 조직이 존재하지 않는 경우 
		Case "ORG2"
			If Len(Trim(Request("txtUpperSalesOrg"))) = 0 Then
				iStrUpperSalesOrg = "NULL"
			Else
				iStrUpperSalesOrg = " " & FilterVar(Request("txtUpperSalesOrg"), "''", "S") & ""
			End If
			
			.CommandText = "UPDATE dbo.b_sales_org " & _
						  "SET upper_sales_org = " & iStrUpperSalesOrg & ", " & _
						  "	   lvl = " & Request("txtSalesOrgNewLvl") & ", " & _
						  "	   end_org_flag = CASE  " & FilterVar(Request("txtEndOrgFlag"), "''", "S") & " WHEN " & FilterVar("N", "''", "S") & "  THEN end_org_flag ELSE " & FilterVar("Y", "''", "S") & "  END, " & _
						  "	   updt_user_id =  " & FilterVar(Request("txtUserId"), "''", "S") & ", " & _
						  "    updt_dt = GETDATE() " & _
						  "WHERE sales_org =  " & FilterVar(Request("txtSalesOrg"), "''", "S") & ""
		' 하위 조직을 포함한 영업조직의 Level 변경 
		Case "ORG3"
			.CommandText = "dbo.usp_s_ChangeSalesOrgHierarchy"
	        .CommandType = adCmdStoredProc

			lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sales_org", adVarXChar,adParamInput,4,Replace(Request("txtSalesOrg"), "'", "''"))
		    
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_cur_lvl", adTinyInt,adParamInput,,Request("txtSalesOrgCurLvl"))
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_new_lvl", adTinyInt,adParamInput,,Request("txtSalesOrgNewLvl"))

			' 상위조직 
			IF Len(Trim(Request("txtUpperSalesOrg"))) = 0 Then
			    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@upper_sales_org", adVarXChar,adParamInput,4)
			Else
			    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@upper_sales_org", adVarXChar,adParamInput,4,Replace(Request("txtUpperSalesOrg"), "'", "''"))
			End If
			' User ID
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))

		' 영업조직 트리 삭제 
		Case "ORG4"
			.CommandText = "dbo.usp_s_DeleteSalesOrgHierarchy"
	        .CommandType = adCmdStoredProc
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sales_org", adVarXChar,adParamInput,4,Trim(Request("txtSalesOrg")))
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_lvl", adTinyInt,adParamInput,,Int(Request("txtSalesOrgCurLvl")))
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@end_org_flag", adXChar,adParamInput,1,Request("txtEndOrgFlag"))
		End Select
		
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
       Response.Write  "       Parent.DbSaveOk		" & vbCr
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

