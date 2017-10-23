<% Option Explicit %>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	
	Call LoadBasisGlobalInf()
	
'	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
'    Dim lgLngMaxRow
'    Dim lgObjConn
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
	
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizCopy()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iPC1G045Q
    Dim iStrData
 
    Dim exportData
    Dim iLngRow,iLngCol
        
    Const C_VerCd = 0
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear                                                                            '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------


	Set iPC1G045Q = Server.CreateObject("PC1G045.cCListDstRlByCcSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If    
	
	Call iPC1G045Q.C_LIST_DSTB_RULE_BY_CC_SVR(gStrGloBalCollection, Trim(Request("txtVerCd")), exportData)
    
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus  
       Set iPC1G045Q = Nothing
       Exit Sub
       
    End If    
    
    Set iPC1G045Q = Nothing
	
	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportData, 1) 		
		For iLngCol = 0 To UBound(exportData, 2)
		    If iLngCol = 0 Or iLngCol = 2 Then
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, iLngCol)))
				iStrData = iStrData & Chr(11) & ""
			Else
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, iLngCol)))
			End If	
		Next
		iStrData = iStrData & Chr(11) & iLngRow
		iStrData = iStrData & Chr(11) & Chr(12)
	Next
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.frm1.htxtVerCd.value = """ & Trim(Request("txtVerCd"))    & """" & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr 
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim iPC1G045S

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    Set iPC1G045S = Server.CreateObject("PC1G045.cCMngDstRlByCcSvr")
	
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    
		    
	Call iPC1G045S.C_MANAGE_DSTB_RULE_BY_CC_SVR(gStrGloBalCollection,Trim(Request("txtVerCd")),Request("txtSpread"),Request("txtSpread3"))		
			
	If CheckSYSTEMError(Err, True) = True Then					
	   Call SetErrorStatus
	   Set iPC1G045S = Nothing
	   Exit Sub
	End If
    
    Set iPC1G045S = Nothing
	
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizCopy
' Desc : 
'============================================================================================================
Sub SubBizCopy()
	Dim iOldVerCd
	Dim iNewVerCd
	Dim IntRetCD
	Dim strMsg_cd
	
	iOldVerCd = Trim(Request("txtVerCd"))
	iNewVerCd = Trim(Request("txtNewVerCd"))

	Call SubCreateCommandObject(lgObjComm)	
    
	With lgObjComm
		.CommandTimeOut = 0
	    .CommandText = "usp_c_copy_dstb_rule_by_cc"
	    .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@old_ver_cd"  ,adVarXChar,adParamInput,6,iOldVerCd)
		.Parameters.Append .CreateParameter("@new_ver_cd"  ,adVarXChar,adParamInput,6,iNewVerCd)
	    .Parameters.Append .CreateParameter("@usrid"       ,adVarXChar,adParamInput,13, gUsrID)
	    .Parameters.Append .CreateParameter("@msgcd"       ,adVarXChar,adParamOutput,6)

	    .Execute ,, adExecuteNoRecords
	End With

	If  Err.number = 0 Then
	    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

	    If  IntRetCD <> 1 Then
	        strMsg_cd = lgObjComm.Parameters("@msgcd").Value
	        Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '¢Ð: Protect system from crashing   
			Response.end
	    End If
	Else
	    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	    Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End If

	Call SubCloseCommandObject(lgObjComm)
	
	Response.Write " <Script Language=vbscript>	                    " & vbCr
	Response.Write " With parent                                    " & vbCr
    Response.Write "	.frm1.txtVerCd.value = """ & iNewVerCd & """" & vbCr
    Response.Write "	.frm1.txtNewVerCd.value = """ & """         " & vbCr
    Response.Write " End With										" & vbCr
    Response.Write " </Script>										" & vbCr 	
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
        Case "MD"
        Case "MR"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MU"
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk()
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk()
          End If   
       Case "<%=UID_M0003%>"                                                         '¢Ð : Copy
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If                
    End Select              
</Script>