<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->

<%
	Call HideStatusWnd_uniSIMS
                                                               '¢Ð: Hide Processing message
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                                 '¢Ð: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '¢Ð: Save,Update
              Call SubBizSave()
         Case "UID_M0003"
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status


    iKey1 = FilterVar(lgKeyStream(2), "''", "S")
    iKey1 = iKey1 & " AND lang_cd = " & FilterVar(lgKeyStream(3), "''", "S")
    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
 
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
    
  
%>
<Script Language=vbscript>
    dim obj
       With Parent	

            .frm1.txtmenu_name.value = "<%=ConvSPChars(lgObjRs("menu_name"))%>"
            .frm1.txthref.value = "<%=ConvSPChars(lgObjRs("href"))%>"
            .frm1.txtmenu_level.value = "<%=ConvSPChars(lgObjRs("menu_level"))%>"
            .frm1.txtpro_auth.value = "<%=ConvSPChars(lgObjRs("pro_auth"))%>"
            .frm1.txtpro_use_flag.value = "<%=ConvSPChars(lgObjRs("pro_use_flag"))%>"
            .frm1.txtref_menu_id.value = "<%=ConvSPChars(lgObjRs("ref_menu_id"))%>"
            .frm1.txtoriginal_ref_id.value = "<%=ConvSPChars(lgObjRs("ref_menu_id"))%>"
            .frm1.txtoriginal_order.value = "<%=ConvSPChars(lgObjRs("orders"))%>"
			 .frm1.txtmenu_order.options.length =0  
			Set obj = Document.CreateElement("OPTION")	
			obj.Text = "<%=ConvSPChars(lgObjRs("orders"))%>"
			obj.Value = "<%=ConvSPChars(lgObjRs("orders"))%>"
			.frm1.txtmenu_order.Add(obj)
			.frm1.txtmenu_order.selectedIndex = 1
			Set obj = Nothing

       End With   
</Script>       
<% 
    End If
    Call SubCloseRs(lgObjRs)
End Sub    


'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
   
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)
	
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
              Call SubBizSaveSingleCreate()  
             
        Case  OPMD_UMODE                                                             '¢Ð : Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  E11000T"
    lgStrSQL = lgStrSQL & " WHERE menu_id = " & FilterVar(lgKeyStream(2), "''", "S")                              ' »ç¹øchar(10)
    lgStrSQL = lgStrSQL & "   AND lang_cd = " & FilterVar(lgKeyStream(3), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
dim txtref_menu_id
dim txtmenu_order
dim iCodeArr
dim iNameArr
'    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO E11000T("
    lgStrSQL = lgStrSQL & " lang_cd, "
    lgStrSQL = lgStrSQL & " menu_id, "
    lgStrSQL = lgStrSQL & " menu_name, "
    lgStrSQL = lgStrSQL & " href, "
    lgStrSQL = lgStrSQL & " menu_level, "
    lgStrSQL = lgStrSQL & " pro_auth, "
    lgStrSQL = lgStrSQL & " pro_use_flag, "
    lgStrSQL = lgStrSQL & " pro_type, "
    lgStrSQL = lgStrSQL & " ref_menu_id,orders ) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtlang_cd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtmenu_id"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtmenu_name"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txthref"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtmenu_level"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtpro_auth"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtpro_use_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtpro_type"), "''", "S") & ","
    
	if  Request("txtref_menu_id") = "" then
	    lgStrSQL = lgStrSQL & "null,"
	else 
	    lgStrSQL = lgStrSQL & FilterVar(Request("txtref_menu_id"), "''", "S") & ","
	end if

	lgStrSQL = lgStrSQL & FilterVar(Request("txtmenu_order"), "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  E11000T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " menu_name = " & FilterVar(Request("txtmenu_name"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " href = " & FilterVar(Request("txthref"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " menu_level = " & FilterVar(Request("txtmenu_level"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " pro_auth = " & FilterVar(Request("txtpro_auth"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " pro_use_flag = " & FilterVar(Request("txtpro_use_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " pro_type = " & FilterVar(Request("txtpro_type"), "''", "S") & ","
    
    if  Request("txtref_menu_id") = "" then    
		lgStrSQL = lgStrSQL & " ref_menu_id=null,"
    else 
		lgStrSQL = lgStrSQL & " ref_menu_id = " & FilterVar(Request("txtref_menu_id"), "''", "S") & ","
	end if
    
	lgStrSQL = lgStrSQL & " orders = " & FilterVar(Request("txtmenu_order"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE menu_id =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
    lgStrSQL = lgStrSQL & "   AND lang_cd =  " & FilterVar(lgKeyStream(3), "''", "S") & ""

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                    lgStrSQL = "Select *" 
                    lgStrSQL = lgStrSQL & " From  E11000T "
                    lgStrSQL = lgStrSQL & " WHERE menu_id = " & pCode 
                Case "P"
                    lgStrSQL = "Select TOP 1 uid, emp_no, password, pro_auth, dept_auth, user_auth," 
                    lgStrSQL = lgStrSQL & " (Select name from HAA010T where HAA010T.emp_no = E11002T.emp_no) as name" 
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                    lgStrSQL = "Select TOP 1 uid, emp_no, password, pro_auth, dept_auth, user_auth," 
                    lgStrSQL = lgStrSQL & " (Select name from HAA010T where HAA010T.emp_no = E11002T.emp_no) as name" 
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
             End Select
    End Select
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
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,false) = True Then
                    Call DisplayMsgBox("800479", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                End If
       
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,false) = True Then
                    Call DisplayMsgBox("800480", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                End If
    End Select
End Sub

%>

<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "UID_M0001"                                                         '¢Ð : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "UID_M0003"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
