<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(�μ�Ȯ��)
'*  3. Program ID           : B2403mb2.asp
'*  4. Program Name         : B2403mb2.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B24032ConfirmDept
'*  7. Modified date(First) : 2000/10/30
'*  8. Modified date(Last)  : 2002/11/30
'*  9. Modifier (First)     : Hwnag Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              : EAB�� ADO�� ������ 
'**********************************************************************************************

Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
    
Dim intRetCD
Dim OrgNm																	     '�����̸����� 

Call LoadBasisGlobalInf()

Call HideStatusWnd                                                               '��: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '��: Set to space

'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
		
 Call SubCreateCommandObject(lgObjComm)
 ''Call LookUpPupUpOrgNm()
 Call SubBizBatch()
 Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

   
    Dim strMsg_cd
    Dim OrgId
    Dim UsrId
    Dim iConfirm
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	OrgId = Request("txtOrgId")
	UsrId = Request("txtUsrId")
	iConfirm = Request("txtConfirm")
	 
	 
    With lgObjComm
        .CommandText = "usp_change_dept"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@orgid"     ,adVarXChar,adParamInput, 5, OrgId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@chg_fg"     ,adVarXChar,adParamInput, 1, iConfirm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@userid"     ,adVarXChar,adParamInput, 13, UsrId)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msgcd"        ,adVarXChar,adParamOutput, 6)
	    
	    lgObjComm.Execute ,, adExecuteNoRecords
        
	End With
	
    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        if  IntRetCD < 0 Then								'SP�� ���� ���������� Return ���� -1�� ��� 
            strMsg_cd = lgObjComm.Parameters("@msgcd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_cd, "", I_MKSCRIPT)
            
            IntRetCD = -1
            Exit Sub
        else
            IntRetCD = 1									'SP ���� ���� 
        end if
    Else													'SP ���� ���� ERROR �߻��� ó��         
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End If
    
End Sub

'============================================================================================================
' Name : LookUpPupUpOrgNm
' Desc : Find name of organization 
'============================================================================================================
Sub LookUpPupUpOrgNm()	

End Sub	


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       Call DisplayMsgBox(Err.number & " : " & Err.Description, vbInformation, "", "", I_MKSCRIPT)
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          Call DisplayMsgBox(Err.number & " : " & Err.Description, vbInformation, "", "", I_MKSCRIPT)             'Can not create(Demo code)
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If
End Sub
%>

<Script Language="VBScript">
    With Parent
        If Trim("<%=lgErrorStatus%>") = "NO" Then      
            IF  "<%=CInt(intRetCD)%>" >= 0 Then
               	''.frm1.txtOrgNm.value = "<%=ConvSPChars(OrgNm)%>"
				.Batch_OK
            Else      
                .LookUp_OK          
            End If      
        Else
            .LookUp_OK
        End If   
   End with
       
</Script>	

