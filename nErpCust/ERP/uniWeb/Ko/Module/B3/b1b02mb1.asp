<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<%
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
	Call LoadBasisGlobalInf

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '��: "P"(Prev search) "N"(Next search)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim lgKeyItemVal

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")

    Call SubMakeSQLStatements("R",iKey1)                                       '�� : Make sql statements
  
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then         'If data not exists
		If lgPrevNext = "" Then
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
			Call SetErrorStatus()
			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "With Parent" & vbCrLf
			    Response.Write ".frm1.txtItemNm.Value	= """"" & vbCrLf
				Response.Write ".frm1.txtItemCd.Focus" & vbCrLf                   'Set condition area
			Response.Write "End With" & vbCrLf
			Response.Write "</Script>" & vbCrLf

		ElseIf lgPrevNext = "P" Then
			Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '�� : This is the starting data. 
			lgPrevNext = ""
			Call SubBizQuery()
		ElseIf lgPrevNext = "N" Then
			Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '�� : This is the ending data.
			lgPrevNext = ""
			Call SubBizQuery()
		End If
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "With Parent" & vbCrLf
			Response.Write ".Frm1.txtItemCd.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_CD"))) & """" & vbCrLf'Set condition area
			Response.Write ".Frm1.txtItemNm.Value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf

		lgKeyItemVal = Trim(lgObjRs("ITEM_CD"))   
    End If
    
    Call SubCloseRs(lgObjRs)                                                    '�� : Release RecordSSet
    	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '��: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '�� : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '�� : Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	Dim lgStrSQL
	
	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE B_ITEM_IMAGE " 
    lgStrSQL = lgStrSQL & " WHERE ITEM_CD =  " & FilterVar(lgKeyStream(0), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	Call SubHandleError("SD", lgObjConn, lgObjRs, Err)
	
	'--------------------------------------------------------------------------------	
		
	lgStrSQL = "UPDATE B_ITEM "
	lgStrSQL = lgStrSQL & " SET ITEM_IMAGE_FLG = " & FilterVar("N", "''", "S") & " ,"
	lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S") & "," 
	lgStrSQL = lgStrSQL & " UPDT_DT = GETDATE() "  
	lgStrSQL = lgStrSQL & " WHERE ITEM_CD = " & FilterVar(lgKeyStream(0), "''", "S")
    
	lgObjConn.Execute lgStrSQL,, adCmdText
	Call SubHandleError("MC", lgObjConn, lgObjRs, Err)
	
	'---------- Developer Coding part (End  ) ---------------------------------------------------------------

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate(lgObjRs)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate(lgObjRs)
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case pMode 
      Case "R"
         Select Case  lgPrevNext 
	        Case ""
                  lgStrSQL = "Select * " 
                  lgStrSQL = lgStrSQL & " From  B_ITEM "
                  lgStrSQL = lgStrSQL & " WHERE ITEM_CD = " & pCode
            Case "P"
				  lgStrSQL = "Select TOP 1 B_ITEM.ITEM_CD, B_ITEM.ITEM_NM " 
                  lgStrSQL = lgStrSQL & " From  B_ITEM_IMAGE, B_ITEM "
                  lgStrSQL = lgStrSQL & " WHERE B_ITEM_IMAGE.ITEM_CD < " & pCode
                  lgStrSQL = lgStrSQL & " AND B_ITEM_IMAGE.ITEM_CD = B_ITEM.ITEM_CD " 	
                  lgStrSQL = lgStrSQL & " ORDER BY B_ITEM_IMAGE.ITEM_CD DESC "
            Case "N"
                  lgStrSQL = "Select TOP 1 B_ITEM.ITEM_CD, B_ITEM.ITEM_NM " 
                  lgStrSQL = lgStrSQL & " From  B_ITEM_IMAGE, B_ITEM "
                  lgStrSQL = lgStrSQL & " WHERE B_ITEM_IMAGE.ITEM_CD > " & pCode
                  lgStrSQL = lgStrSQL & " AND B_ITEM_IMAGE.ITEM_CD = B_ITEM.ITEM_CD "
                  lgStrSQL = lgStrSQL & " ORDER BY B_ITEM_IMAGE.ITEM_CD ASC "
	    End Select
      Case "C"
      Case "U"
			lgStrSQL = "Select * " 
            lgStrSQL = lgStrSQL & " From   B_ITEM "
            lgStrSQL = lgStrSQL & " Where  ITEM_CD =  " & pCode 	
      Case "D"
    End Select
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode, pConn, pRs, pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">

	Select Case "<%=lgOpModeCRUD%>"
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQuery("<%=ConvSPChars(lgKeyItemVal)%>")               
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select      
</Script>	
