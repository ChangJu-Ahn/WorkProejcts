<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
	Call LoadBasisGlobalInf

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim lgKeyItemVal

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
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

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")

    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements
  
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then         'If data not exists
		If lgPrevNext = "" Then
			Call DisplayMsgBox("Y60040", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 

			Call SetErrorStatus()
			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "With Parent" & vbCrLf
			    Response.Write ".frm1.txtCast_Nm.Value	= """"" & vbCrLf
				Response.Write ".frm1.txtCast_Cd.Focus" & vbCrLf                   'Set condition area
			Response.Write "End With" & vbCrLf
			Response.Write "</Script>" & vbCrLf

		ElseIf lgPrevNext = "P" Then
			Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : This is the starting data. 
			lgPrevNext = ""
			Call SubBizQuery()
		ElseIf lgPrevNext = "N" Then
			Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : This is the ending data.
			lgPrevNext = ""
			Call SubBizQuery()
		End If
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "With Parent" & vbCrLf
			Response.Write ".Frm1.txtCast_Cd.Value = """ & ConvSPChars(Trim(lgObjRs("Cast_Cd"))) & """" & vbCrLf'Set condition area
			Response.Write ".Frm1.txtCast_Nm.Value = """ & ConvSPChars(lgObjRs("Cast_Nm")) & """" & vbCrLf
			Response.Write ".Frm1.txtMemo.Value 	   = """ & ConvSPChars(lgObjRs("IMG_TEXT")) & """" & vbCrLf
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf

		lgKeyItemVal = Trim(lgObjRs("Cast_Cd"))   
    End If
    
    Call SubCloseRs(lgObjRs)                                                    '¢Ð : Release RecordSSet
    	
End Sub	
'============================================================================================================
' Name : SubBizQuery
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
	Dim lgStrSQL
	
	On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE Y_FAC_CAST_IMAGE " 
    lgStrSQL = lgStrSQL & " WHERE FAC_CAST_CD =  " & FilterVar(lgKeyStream(0), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	Call SubHandleError("SD", lgObjConn, lgObjRs, Err)
	
	'--------------------------------------------------------------------------------	
		
	lgStrSQL = "UPDATE Y_CAST "
	lgStrSQL = lgStrSQL & " SET PIC_FLAG = " & FilterVar("N", "''", "S") & " ,"
	lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & "," 
	lgStrSQL = lgStrSQL & " UPDT_DT = GETDATE() "  
	lgStrSQL = lgStrSQL & " WHERE Cast_Cd = " & FilterVar(lgKeyStream(0), "''", "S")
    
	lgObjConn.Execute lgStrSQL,, adCmdText
	Call SubHandleError("MC", lgObjConn, lgObjRs, Err)
	
	'---------- Developer Coding part (End  ) ---------------------------------------------------------------

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate(lgObjRs)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate(lgObjRs)
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
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
'                   lgStrSQL = "Select * " 
'                   lgStrSQL = lgStrSQL & " From  Y_CAST "
'                   lgStrSQL = lgStrSQL & " WHERE Cast_Cd = " & pCode
				lgStrSQL = "Select Y_CAST.Cast_Cd, Y_CAST.Cast_Nm, Y_FAC_CAST_IMAGE.REMRK IMG_TEXT " 
	            lgStrSQL = lgStrSQL & " FROM  Y_CAST  "
	            lgStrSQL = lgStrSQL & " LEFT OUTER JOIN  Y_FAC_CAST_IMAGE ON Y_CAST.Cast_Cd = Y_FAC_CAST_IMAGE.FAC_CAST_CD"
	            lgStrSQL = lgStrSQL & " Where  Cast_Cd =  " & pCode 	

            Case "P"
				  lgStrSQL = "Select TOP 1 Y_CAST.Cast_Cd, Y_CAST.Cast_Nm " 
                  lgStrSQL = lgStrSQL & " From  Y_FAC_CAST_IMAGE, Y_CAST "
                  lgStrSQL = lgStrSQL & " WHERE Y_FAC_CAST_IMAGE.FAC_CAST_CD < " & pCode
                  lgStrSQL = lgStrSQL & " AND Y_FAC_CAST_IMAGE.FAC_CAST_CD = Y_CAST.Cast_Cd " 	
                  lgStrSQL = lgStrSQL & " ORDER BY Y_FAC_CAST_IMAGE.FAC_CAST_CD DESC "
            Case "N"
                  lgStrSQL = "Select TOP 1 Y_CAST.Cast_Cd, Y_CAST.Cast_Nm " 
                  lgStrSQL = lgStrSQL & " From  Y_FAC_CAST_IMAGE, Y_CAST "
                  lgStrSQL = lgStrSQL & " WHERE Y_FAC_CAST_IMAGE.FAC_CAST_CD > " & pCode
                  lgStrSQL = lgStrSQL & " AND Y_FAC_CAST_IMAGE.FAC_CAST_CD = Y_CAST.Cast_Cd "
                  lgStrSQL = lgStrSQL & " ORDER BY Y_FAC_CAST_IMAGE.FAC_CAST_CD ASC "
	    End Select
      Case "C"
      Case "U"
			lgStrSQL = "Select Y_CAST.Cast_Cd, Y_CAST.Cast_Nm, Y_FAC_CAST_IMAGE.REMRK IMG_TEXT " 
            lgStrSQL = lgStrSQL & " FROM  Y_CAST  "
            lgStrSQL = lgStrSQL & " LEFT OUTER JOIN  Y_FAC_CAST_IMAGE ON Y_CAST.Cast_Cd = Y_FAC_CAST_IMAGE.FAC_CAST_CD"
            lgStrSQL = lgStrSQL & " Where  Cast_Cd =  " & pCode 	
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode, pConn, pRs, pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
