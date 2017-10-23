<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 1000
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")   
  
    Call HideStatusWnd
 
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
    End Select
    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere, strFg
  
     On Error Resume Next    
    Err.Clear                                                               '¢Ð: Clear Error status


    strWhere =  FilterVar(lgKeyStream(0), "''", "S")  & " AND " & FilterVar(lgKeyStream(1), "''", "S") 
 
    If lgKeyStream(2) <> "" then
       strWhere = strWhere & " AND HDF020T.emp_no = " & FilterVar(lgKeyStream(2), "''", "S")
    End if
    
    If lgKeyStream(3) <> "" then
       strWhere = strWhere & " AND HAA010T.PAY_GRD1 >= " & FilterVar(lgKeyStream(3), "''", "S")
    End if

    If lgKeyStream(4) <> "" then
       strWhere = strWhere & " AND HAA010T.PAY_GRD1 <= " & FilterVar(lgKeyStream(4), "''", "S")
    End if
    
    If lgKeyStream(5) <> "" then
       strWhere = strWhere & " AND HDF020T.PROV_TYPE = " & FilterVar(lgKeyStream(5), "''", "S")
    End if
    			
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '¢Ð : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
       Call SetErrorStatus()
    Else
'       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
       
       
            If ConvSPChars(lgObjRs("PROV_TYPE")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If
                   
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME") )                 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLL_PSTN"))             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD_NM")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD1")) 
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
 '           If iDx > C_SHEETMAXROWS_D Then
  '             lgStrPrevKey = lgStrPrevKey + 1
   '            Exit Do
    '        End If   
                       
        Loop 
    End If
            
'      If iDx <= C_SHEETMAXROWS_D Then
  '       lgStrPrevKey = ""
 '     End If   

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
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
    Next

End Sub      
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  HDF020T"
    lgStrSQL = lgStrSQL & " SET "
    
    If arrColVal(3) = "1" Then 
		lgStrSQL = lgStrSQL & " PROV_TYPE = " & FilterVar("Y", "''", "S") & ","
	Else
		lgStrSQL = lgStrSQL & " PROV_TYPE = " & FilterVar("N", "''", "S") & ","
	End If

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "   EMP_NO = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
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
                       lgStrSQL = "Select TOP " & iSelCount
                       lgStrSQL = lgStrSQL & " HAA010T.NAME         ,"   
                       lgStrSQL = lgStrSQL & " HDF020T.EMP_NO       ,"   
                       lgStrSQL = lgStrSQL & " HAA010T.DEPT_CD      ,"   
                       lgStrSQL = lgStrSQL & " HAA010T.DEPT_NM      ,"   
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", HDF020T.ROLL_PSTN) ROLL_PSTN,"                         
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ", HDF020T.PAY_CD) PAY_CD_NM,"   
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S")  & ", HAA010T.PAY_GRD1) PAY_GRD1,"
                       lgStrSQL = lgStrSQL & " HDF020T.PROV_TYPE   "                                                 
                       lgStrSQL = lgStrSQL & " From HAA010T join HDF020T on HAA010T.emp_no = hdf020t.emp_no"
                       lgStrSQL = lgStrSQL & " Where HAA010T.INTERNAL_CD BETWEEN " & pCode
                       lgStrSQL = lgStrSQL & " Order by HAA010T.internal_cd, HAA010T.PAY_GRD1, HDF020T.EMP_NO "
'Response.Write lgStrSQL
'Response.End
 
          End Select             
    End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
    Response.Write "<BR> Commit Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    Response.Write "<BR> Abort Event occur"
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
        Case "MS"
                 Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
                 ObjectContext.SetAbort
                 Call SetErrorStatus
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
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If    
  
    End Select       
       
</Script>	
