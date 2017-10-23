<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Dim lgGetsvrDateTime
	Dim lgFlag
     
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB") 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    
	lgFlag            = Request("txtFlag")
    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
	lgGetsvrDateTime = GetSvrDateTime
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    
     On Error Resume Next    
    Err.Clear                                                               '¢Ð: Clear Error status

    Call SubMakeSQLStatements("MR", lgFlag, "X",C_LIKE)                              '¢Ð : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 

       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
			If lgFlag = "1" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD1"))
				lgstrData = lgstrData & Chr(11) & ""
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM") )                 
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ADD_RATE"))
            
				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)

    			lgObjRs.MoveNext
    		Else
    			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DUTY_MM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCUM_RATE")  )                
            
				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)

    			lgObjRs.MoveNext
    		End If
          
			iDx =  iDx + 1
			If iDx > C_SHEETMAXROWS_D Then
			lgStrPrevKey = lgStrPrevKey + 1
			Exit Do
			End If   
                       
        Loop 
    End If
      If iDx <= C_SHEETMAXROWS_D Then
         lgStrPrevKey = ""
      End If   

      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub      

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
     On Error Resume Next   
    Err.Clear                                                                        '¢Ð: Clear Error status
  
    If lgFlag = "1" Then    
		lgStrSQL = "INSERT INTO HGA010T("
		lgStrSQL = lgStrSQL & " pay_grd1, " 
		lgStrSQL = lgStrSQL & " add_rate, " 
		lgStrSQL = lgStrSQL & " ISRT_DT ," 
		lgStrSQL = lgStrSQL & " ISRT_EMP_NO ," 
		lgStrSQL = lgStrSQL & " UPDT_DT ," 
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO )" 
		lgStrSQL = lgStrSQL & " VALUES(" 
    
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(1)), "''", "S")     & ","     
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"0","D")      & ","    
		lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        
		lgStrSQL = lgStrSQL & ")"  
    Else
		lgStrSQL = "INSERT INTO HGA020T("
		lgStrSQL = lgStrSQL & " DUTY_MM, " 
		lgStrSQL = lgStrSQL & " ACCUM_RATE, " 
		lgStrSQL = lgStrSQL & " ISRT_DT ," 
		lgStrSQL = lgStrSQL & " ISRT_EMP_NO ," 
		lgStrSQL = lgStrSQL & " UPDT_DT ," 
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO )" 
		lgStrSQL = lgStrSQL & " VALUES(" 
    
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(1))),"0","D")      & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"0","D")      & ","    
		lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        
		lgStrSQL = lgStrSQL & ")"  
    End If 
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
	On Error Resume Next 
	Err.Clear                                                                        '¢Ð: Clear Error status

	If lgFlag = "1" Then
		lgStrSQL = "UPDATE  HGA010T"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " ADD_RATE =       " & FilterVar(Trim(UCase(arrColVal(2))),"0","D")  & ","
		lgStrSQL = lgStrSQL & " UPDT_DT  =       " & FilterVar(lgGetSvrDateTime,NULL,"S") & ","
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO =    " & FilterVar(gUsrId, "''", "S")  
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & " PAY_GRD1    =    " & FilterVar(UCase(arrColVal(1)), "''", "S")
	Else
		lgStrSQL = "UPDATE  HGA020T"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " ACCUM_RATE =     " & FilterVar(Trim(UCase(arrColVal(2))),"0","D")  & ","
		lgStrSQL = lgStrSQL & " UPDT_DT  =       " & FilterVar(lgGetSvrDateTime,NULL,"S") & ","
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO =    " & FilterVar(gUsrId, "''", "S")  
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & " DUTY_MM     =    " & FilterVar(Trim(UCase(arrColVal(1))),"0","D")
	End If

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

    If lgFlag = "1" Then
		lgStrSQL = "DELETE  HGA010T"
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & " PAY_GRD1    =    " & FilterVar(UCase(arrColVal(1)), "''", "S")
    Else
		lgStrSQL = "DELETE  HGA020T"
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & " DUTY_MM     =    " & FilterVar(Trim(UCase(arrColVal(1))),"0","D")
    End If
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
     Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
           
               Case "R"
					If pCode = "1" Then
                       lgStrSQL = "Select TOP " & iSelCount 
                       lgStrSQL = lgStrSQL & " a.pay_grd1, b.minor_nm, a.add_rate "
                       lgStrSQL = lgStrSQL & " From  HGA010T a, B_MINOR b "
                       lgStrSQL = lgStrSQL & " Where b.major_cd = " & FilterVar("H0001", "''", "S") & " and a.pay_grd1 = b.minor_cd "
				    Else
					   lgStrSQL = "Select TOP " & iSelCount  & " * "
                       lgStrSQL = lgStrSQL & " From  HGA020T "
				    End If
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
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
				If "<%=lgFlag%>" = "1" Then
					.ggoSpread.Source     = .frm1.vspdData
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
				Else
					.ggoSpread.Source     = .frm1.vspdData1
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
				End If
				.DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
