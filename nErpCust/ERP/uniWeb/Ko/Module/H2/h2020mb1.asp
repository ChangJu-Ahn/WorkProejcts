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
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim strWhere 
    Dim frRoll_pstn, toRoll_pstn
  
    On Error Resume Next    
    Err.Clear                                                               '☜: Clear Error status

	If trim(lgKeyStream(0)) <> "" Then
		strWhere = " And HAA080T.PASS_FLAG = " & FilterVar(lgKeyStream(0), "''", "S")
	End If

	If trim(lgKeyStream(1)) = "0" Then
		strWhere = strWhere & " And HAA010T.RETIRE_DT IS NULL "
	ElseIf trim(lgKeyStream(1)) = "1" Then
		strWhere = strWhere & " And HAA010T.RETIRE_DT IS NOT NULL "
	End If

	If trim(lgKeyStream(2)) <> "" Then
		strWhere = strWhere & " And HDF020T.PAY_CD = " & FilterVar(lgKeyStream(2),"'%'", "S")
	End If	
    
    strWhere = strWhere & " And (HAA010T.internal_cd >= " & FilterVar(lgKeyStream(3), "''", "S")
    strWhere = strWhere & " And  HAA010T.internal_cd <= " & FilterVar(lgKeyStream(4), "''", "S") & ")"
    
	frRoll_pstn = trim(lgKeyStream(5))
	toRoll_pstn = trim(lgKeyStream(6))

	If frRoll_pstn = "" Then
		frRoll_pstn = "0"
	End If

	If toRoll_pstn = "" Then
		toRoll_pstn = "ZZZZ"
	End If
			
    strWhere = strWhere & " And (HAA010T.Roll_pstn >= " & FilterVar(frRoll_pstn, "''", "S")
    strWhere = strWhere & " And  HAA010T.Roll_pstn <= " & FilterVar(toRoll_pstn, "''", "S") & ")"
    	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '☜ : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PASS_FLAG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PASS_FLAG_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO") )                 
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLL_PSTN_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ENG_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RES_NO"))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PASSPORT_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAT_CD") )                 
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAT_NM"))
            
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("ISSUE_DT"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("EXPIRE_DT"),"")            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext
          
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
    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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
    on error resume next
    Err.Clear  
    
	lgStrSQL = "INSERT INTO HAA080T"
	lgStrSQL = lgStrSQL & "(EMP_NO, PASS_FLAG, NAT_CD, ISSUE_DT, EXPIRE_DT, PASSPORT_NO, REMARK, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT) "
	lgStrSQL = lgStrSQL & " VALUES(" 
    
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
				
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(5)),NULL,"S")		& ","    
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(6)),NULL,"S")		& ","    	    
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S")     & ","
	
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")			& "," 
	lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")			& "," 	
	lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
	lgStrSQL = lgStrSQL & ")"  
 
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    on error resume next
     Err.Clear  
    
    lgStrSQL = "UPDATE  HAA080T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " ISSUE_DT	=       " & FilterVar(UNIConvDate(arrColVal(5)),NULL,"S")	& ","
    lgStrSQL = lgStrSQL & " EXPIRE_DT	=       " & FilterVar(UNIConvDate(arrColVal(6)),NULL,"S")	& ","
    lgStrSQL = lgStrSQL & " PASSPORT_NO	=       " & FilterVar(UCase(arrColVal(7)), "''", "S")		& ","
    lgStrSQL = lgStrSQL & " REMARK		=       " & FilterVar(arrColVal(8), "''", "S")		& ","
        
    lgStrSQL = lgStrSQL & " UPDT_DT		=       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO =       " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO		=       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " PASS_FLAG	=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " NAT_CD		=       " & FilterVar(UCase(arrColVal(4)), "''", "S")

'Response.Write lgStrSQL
'Response.End 
  
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HAA080T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO		=       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " PASS_FLAG	=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " NAT_CD		=       " & FilterVar(UCase(arrColVal(4)), "''", "S")
'Response.Write lgStrSQL
'Response.End 
    
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
                       lgStrSQL = "Select TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & " PASS_FLAG,case when PASS_FLAG ='0' then '여권' when PASS_FLAG ='1' then '비자' end PASS_FLAG_NM, "   
                       lgStrSQL = lgStrSQL & " HAA080T.EMP_NO,HAA010T.NAME,HAA010T.DEPT_NM,dbo.ufn_GetCodeName('H0002', HAA010T.ROLL_PSTN) ROLL_PSTN_NM, "  
                       lgStrSQL = lgStrSQL & " HAA010T.ENG_NAME, HAA010T.RES_NO,HAA080T.PASSPORT_NO, " 
                       lgStrSQL = lgStrSQL & " HAA080T.NAT_CD,dbo.ufn_H_GetCodeName('B_COUNTRY', HAA080T.NAT_CD, '') NAT_NM,ISSUE_DT,EXPIRE_DT,REMARK "
                       lgStrSQL = lgStrSQL & " FROM (HAA080T join HAA010T on HAA080T.EMP_NO = HAA010T.EMP_NO ) JOIN HDF020T ON HAA080T.EMP_NO = HDF020T.EMP_NO "
                       lgStrSQL = lgStrSQL & " WHERE 1=1 "  & pCode 
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
        Case "MV"
                 Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MX"
                 Call DisplayMsgBox("800350", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MY"
                 Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MZ"
                 Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
		  Else
                Parent.DBQueryFail  		  	         
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
