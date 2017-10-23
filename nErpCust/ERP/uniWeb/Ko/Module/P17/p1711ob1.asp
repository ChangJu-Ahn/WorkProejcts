<%@ LANGUAGE=VBSCript  TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "OB")

Call HideStatusWnd

Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim CallType
Dim strPlantCd
Dim strItemCd
Dim strBomNo
Dim strExpFlg
Dim btnType
Dim strBaseDt
Dim strSpId

'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space

Call SubOpenDB(lgObjConn)       
Call SubCreateCommandObject(lgObjComm)
'	Call SubDelete
Call SubMakeParameter()
Call SubBomExplode()
'Call SubBizSaveSingleUpdate()
		
Call SubCloseCommandObject(lgObjComm)
Call SubCloseDB(lgObjConn)      

Response.Write "<Script Language = VBScript>" & vbcrLf
	Response.Write "parent.frm1.txtSpId.value = """ & strSpId & """" & vbCrLf
	If btnType = 0 Then
		Response.Write "Call parent.PrevExecOk()" & vbcrLf
	Else
		Response.Write "Call parent.PrintExecOk()" & vbcrLf
	End If
Response.Write "</Script>" & vbcrLf

Response.End

'============================================================================================================
' Name : SubMakeParameter
' Desc : Make SP Parameter
'============================================================================================================
Sub SubMakeParameter()
	
	If CInt(Request("txtMode")) <> UID_M0001 Then
		Response.End 
	End If
	
	strPlantCd = "" & Request("txtPlantCd")									' 조회할 키 
	strItemCd = "" & Request("txtItemCd")									' 조회할 상위키 
	
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
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, strExpFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, strItemCd)
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
			End If
            IntRetCD = -1
            Exit Sub
        Else
			IntRetCD = 1
        End if
    Else           
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
End Sub	

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	lgStrSQL = ""
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  P_BOM_FOR_EXPLOSION"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " REMARK = " &  FilterVar(gUsrId ,"''","S")
    lgStrSQL = lgStrSQL & " WHERE plant_cd = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
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
	' 생성된 temp table 삭제 
	'-------------------------
	lgStrSQL = "DELETE FROM p_bom_for_explosion "
	lgStrSQL = lgStrSQL & " WHERE plant_cd = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
	lgStrSQL = lgStrSQL & " AND remark = " & FilterVar(gUsrId ,"''","S")
	
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