<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
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

    lgOpModeCRUD      = Request("txtMode")
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space

	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         
             Call SubCreateCommandObject(lgObjComm)
			 Call SubBizBatch()
             Call SubCloseCommandObject(lgObjComm)
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strconBp_cd
    Dim strconItem_cd
    Dim strPlantCode
    Dim strFr_dt
    Dim strTo_dt

    Dim strVol_flag
    Dim strQty_flag
    Dim strRetire_flag
    Dim strSave_flag
    Dim strLoan_flag

    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strconBp_cd = Request("txtconBp_cd")
    strconItem_cd = Request("txtconItem_cd")
    strPlantCode = Request("txtPlantCode")
    if  strPlantCode = "" then
        strPlantCode = "%"
    end if
    if  strconItem_cd = "" then
        strconItem_cd = "%"
    end if

    strFr_dt = Request("txtFr_dt")
    strTo_dt = Request("txtTo_dt")
    strVol_flag = Request("txtVol_flag")
    strQty_flag = Request("txtQty_flag")

    With lgObjComm
        .CommandText = "usp_s_price_calc_ko441"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_fr_dt",adXChar,adParamInput,Len(strFr_dt), strFr_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_to_dt" ,adXChar,adParamInput,Len(strTo_dt), strTo_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_bp_cd"   ,adXChar,adParamInput,Len(strconBp_cd), strconBp_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_plant_cd",adXChar,adParamInput,Len(strPlantCode), strPlantCode)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_item_cd"  ,adXChar,adParamInput,Len(strconItem_cd), strconItem_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_vol_flag"   ,adXChar,adParamInput,Len(strVol_flag), strVol_flag)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@p_qty_flag" ,adXChar,adParamInput,Len(strQty_flag), strQty_flag)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id"     ,adXChar,adParamInput,Len(gUsrId), gUsrId)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)
        
        lgObjComm.Execute ,, adExecuteNoRecords

    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        if  IntRetCD < 0 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
            IntRetCD = -1
            Exit Sub
        else
            IntRetCD = 1
        end if
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub


%>

<Script Language="VBScript">

    With Parent
		Select Case "<%=lgOpModeCRUD %>"
		    Case "<%=UID_M0001%>"
			    IF Trim("<%=lgErrorStatus%>") = "NO" AND "<%=CInt(intRetCD)%>" >= 0 Then
			        .ExeReflectOk
			    Else
			        .ExeReflectNo
			    End If
		End Select
    End with	   
</Script>