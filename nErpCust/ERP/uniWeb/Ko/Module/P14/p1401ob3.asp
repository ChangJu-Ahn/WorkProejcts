<%@ LANGUAGE=VBSCript  TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->

<%
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "OB")

Dim strCode																	'�� : Lookup �� �ڵ� ���� ���� 
Dim CallType
Dim strPlantCd
Dim strItemCdFrom
Dim strItemCdTo
Dim strBomNo
Dim strExpFlg
Dim btnType
Dim strBaseDt
Dim strSpId

Call HideStatusWnd                                                               '��: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '��: Set to space

Call SubOpenDB(lgObjConn)       
Call SubCreateCommandObject(lgObjComm)
'Call SubDelete
Call SubMakeParameter()
Call SubBomExplode()
'Call SubBizSaveSingleUpdate()
		
Call SubCloseCommandObject(lgObjComm)
Call SubCloseDB(lgObjConn)      

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtSpId.value = """ & strSpId & """" & vbCrLf
	If btnType = 0 Then
		Response.Write "Call parent.PrevExecOk()" & vbCrLf
	ElseIf btnType = 1  Then
		Response.Write "Call parent.PrintExecOk()" & vbCrLf
	End If
Response.Write "</Script>" & vbCrLf
Response.End

'============================================================================================================
' Name : SubMakeParameter
' Desc : Make SP Parameter
'============================================================================================================
Sub SubMakeParameter()
	
	If CInt(Request("txtMode")) <> UID_M0001 Then
		Response.End 
	End If
	
	strPlantCd = "" & Request("txtPlantCd")									' ��ȸ�� Ű 
	strItemCdFrom = "" & Request("txtItemCdFrom")							' ��ȸ�� ����Ű 
	strItemCdTo = "" & Request("txtItemCdTo")								' ��ȸ�� ����Ű 
	
	IF Request("txtBomNo") = "" Then
		strBomNo = " "
	Else
		strBomNo = Request("txtBomNo")
	End If
	
	strExpFlg = Request("rdoPrintType")
	btnType = Request("BtnType")
	strBaseDt = Request("txtBaseDt")
	    
End Sub     
'============================================================================================================
' Name : SubBomExplode
' Desc : Query Data from Db
'============================================================================================================
Sub SubBomExplode()

    Dim strMsg_cd
    Dim strMsg_text
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    With lgObjComm
        .CommandText = "usp_multi_BOM_explode_main"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, strExpFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@root_item_cd_from",	advarXchar,adParamInput,18, strItemCdFrom)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@root_item_cd_to",	advarXchar,adParamInput,18, strItemCdTo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4,strBomNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	advarXchar,adParamInput,10,strBaseDt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_qty",	adInteger,adParamInput,2,1)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	advarXchar,adParamOutput,60)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	advarXchar,adParamOutput,13)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With
    
    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            strSpId = lgObjComm.Parameters("@user_id").Value
            
            If strMsg_cd <> MSG_OK_STR Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
				Response.End
			End If
            IntRetCD = -1
            Exit Sub
        Else
			IntRetCD = 1
        End if
    Else           
        Call SvrMsgBox(Err.Description, VBInformation, I_MKSCRIPT)
        Response.End
'        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
End Sub	

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	
	lgStrSQL = ""
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  P_MULTI_BOM_FOR_EXPLOSION"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " REMARK = " &  FilterVar(gUsrId ,"''","S")
    lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
    lgStrSQL = lgStrSQL & " AND USER_ID   = " &  FilterVar(strSpId ,"''", "S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubDelete
' Desc : Delete Temp Table Data from Db
'============================================================================================================
Sub SubDelete()

	lgStrSQL = ""
	'-------------------------
	' ������ temp table ���� 
	'-------------------------
	lgStrSQL = "DELETE FROM P_MULTI_BOM_FOR_EXPLOSION "
	lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
	lgStrSQL = lgStrSQL & " AND REMARK = " & FilterVar(gUsrId ,"''","S")
	
	'---------- Developer Coding part (End  ) ---------------------------------------------------------------
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
		
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
