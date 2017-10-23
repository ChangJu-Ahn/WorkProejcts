<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%

    Dim intRetCD

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "BB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    
	lgOpModeCRUD      = Trim(Request("txtMode"))
    lgKeyStream       = Split(Request("lgKeyStream"),gColSep)    
	
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strMsg_cd 

    Dim strApply_dt,strPost_apply_dt,strPay_grd,strPay_grd_nm,strBase_amt
    Dim strPay_grd1,strPay_grd2,strAllow_cd,strAllow_amt,strAllow_rate,strEnd_type
    Dim PBase_amt, PAllow_amt, PAllow_rate

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strApply_dt         = Trim(lgKeyStream(0))
    strPost_apply_dt    = Trim(lgKeyStream(1))

    strPay_grd          = Trim(lgKeyStream(2))
    If strPay_grd = "" Then
    	strPay_grd =  "%"
    End if
    		
    strPay_grd_nm       = Trim(lgKeyStream(3))
    If strPay_grd_nm = "" Then
   		strPay_grd_nm =  "%"
   	End if

    strPay_grd1         = Trim(lgKeyStream(4))
    If strPay_grd1 = "" Then
   		strPay_grd1 =  "0"
   	End if

    If Not IsNull(strPay_grd1) And Len(strPay_grd1) = 1 Then
        strPay_grd1 = "0" + strPay_grd1
    End If

    strPay_grd2         = Trim(lgKeyStream(5))
    If strPay_grd2 = "" Then
    	strPay_grd2 =  "zz"
    End if
    
    If Not IsNull(strPay_grd2) And Len(strPay_grd2) = 1 Then
        strPay_grd2 = "0" + strPay_grd2
    End If

    strAllow_cd = Trim(lgKeyStream(6))
    
    If Trim(lgKeyStream(7)) = "" Then        
        strAllow_amt        =  0
    Else
        strAllow_amt        = UNIConvNum(lgKeyStream(7), 0)
    End If
    
    If Trim(lgKeyStream(8)) = "" Then
        strAllow_rate       = 0
    Else
        strAllow_rate       = UNIConvNum(lgKeyStream(8), 0)
    End If
    
    If Trim(lgKeyStream(9)) = ""  Then
        strBase_amt         = 0
    Else
        strBase_amt         = UNIConvNum(lgKeyStream(9), 0)
    End If
        
    strEnd_type         = Trim(lgKeyStream(10))
    
    With lgObjComm
    
        Select Case Cdbl(lgOpModeCRUD)
            Case Cdbl(UID_M0002)                                    '==>실행 sp
                .CommandText = "usp_hdf011b1_1"
                .CommandType = adCmdStoredProc       

	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"           ,adVarXChar,adParamInput, 13, gUsrId)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@apply_dt_s"       ,adVarXChar,adParamInput, 8,  strApply_dt)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@post_apply_dt_s"  ,adVarXChar,adParamInput, 8,  strPost_apply_dt)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_grd1"    ,adVarXChar,adParamInput, 2,  strPay_grd)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_pay_grd2"      ,adVarXChar,adParamInput, 3,  strPay_grd1)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_pay_grd2"      ,adVarXChar,adParamInput, 3,  strPay_grd2)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_allow_cd"    ,adVarXChar,adParamInput, 3,  strAllow_cd)
	            
                Set PBase_amt = lgObjComm.CreateParameter("@base_amt",adNumeric,adParamInput)
                PBase_amt.Precision = 18
                PBase_amt.NumericScale = 4
                PBase_amt.Value = strBase_amt
                lgObjComm.Parameters.Append PBase_amt

	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@end_type"         ,adVarXChar,adParamInput, 1,  strEnd_type)

                Set PAllow_amt = lgObjComm.CreateParameter("@para_allow_amt",adNumeric,adParamInput)
                PAllow_amt.Precision = 18                           '==>numeric(18,4)를 전체자리 18과 소수부분4를 따로 설정해준다.
                PAllow_amt.NumericScale = 4
                PAllow_amt.Value = strAllow_amt
                lgObjComm.Parameters.Append PAllow_amt

                Set PAllow_rate = lgObjComm.CreateParameter("@para_allow_rate",adNumeric,adParamInput)
                PAllow_rate.Precision = 4
                PAllow_rate.NumericScale = 2
                PAllow_rate.Value = strAllow_rate
                lgObjComm.Parameters.Append PAllow_rate

                lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"           ,adVarXChar,adParamoutput, 6)
                lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"         ,adVarXChar,adParamOutput,60)
                
            Case Cdbl(UID_M0003)                                      '==>삭제 sp
                .CommandText = "usp_hdf011b1_2"
                .CommandType = adCmdStoredProc       
                
                lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"           ,adVarXChar,adParamInput, 13, gUsrId)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@apply_dt_s"       ,adVarXChar,adParamInput, 8,  strApply_dt)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@post_apply_dt_s"  ,adVarXChar,adParamInput, 8,  strPost_apply_dt)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_grd1"    ,adVarXChar,adParamInput, 8,  strPay_grd)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_pay_grd2"      ,adVarXChar,adParamInput, 3,  strPay_grd1)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_pay_grd2"      ,adVarXChar,adParamInput, 3,  strPay_grd2)
	            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"           ,adVarXChar,adParamoutput, 6)
                lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"         ,adVarXChar,adParamOutput,60)
                
        End Select
        
        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        If  Cdbl(IntRetCD) < 0 Then
			strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
            
            IntRetCD = -1
            Exit Sub
        Else
            IntRetCD = 1
        End if
    Else           
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
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
