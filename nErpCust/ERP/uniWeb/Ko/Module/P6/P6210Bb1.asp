<%@ LANGUAGE=VBSCript%>
<% Server.ScriptTimeout = 300 %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
    Dim intRetCD

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "BB")

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space

     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strCast_Cd
    Dim strReportDtFr, strReportDtTo
    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	
	strPlantCd = Request("txtPlantCd")
	strCast_Cd = Request("txthCast_Cd")
	If IsNull(strCast_Cd) Or Trim(strCast_Cd) = "" Then
		strCast_Cd = "%"
	Else
		
		StrCast_Cd = Request("txthCast_Cd")
	End If
	
    strReportDtFr   = Request("txtReportDtFr")
    strReportDtTo   = Request("txtReportDtTo")
      
    With lgObjComm
        .CommandText = "usp_y_fin_ad_accnt"

        .CommandType =  adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd"  ,adVarXChar,adParamInput  , 10 , strPlantCd)        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@report_dt_fr"  ,adVarXChar,adParamInput  , 10 , strReportDtFr)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@report_dt_to"  ,adVarXChar,adParamInput  , 10 , strReportDtTo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cast_cd"    ,adVarXChar,adParamInput  , 18, strCast_Cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarXChar,adParamInput  , 10, gUsrId)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput , 6 )
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput , 60)

        lgObjComm.Execute ,, adExecuteNoRecords
    End With
    
    
    
    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
'        if  IntRetCD < 0 then
'            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
'            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            
'            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
            
'            IntRetCD = -1
'            Exit Sub
'        else
'            IntRetCD = 1
'        end if
    Else           
        Call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
End Sub	

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language="VBScript">

   If Trim("<%=lgErrorStatus%>") = "NO" Then
      With Parent
           IF  "<%=CInt(intRetCD)%>" >= 0 Then
               .ExeReflectOk
           Else
               .ExeReflectNo
           End If
      End with
   End If
       
</Script>	
