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
    
	Call loadInfTB19029B( "I", "*","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                    '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
   
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
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
    Dim strWhere 
	Dim strBp_cd, strITEM_CD
  
    On Error Resume Next    
    Err.Clear                                                               '¢Ð: Clear Error status

	strITEM_CD = Trim(Request("txtconITEM_CD"))
    strWhere = ""


	If strITEM_CD <> "" Then
		strWhere = strWhere & " And MES_ITEM_CD >= " & FilterVar(strITEM_CD, "''", "S")
	End If


    	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '¢Ð : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF


            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MES_ITEM_CD") )                 
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TESTER"))
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TESTER_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_METHOD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_METHOD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PURE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HOT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INDEX_T"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SHOT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_DIE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELEASE_DATE"))
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
      Else
         if lgStrPrevKey = 1 Then
%>
<Script Language=vbscript>
       With Parent	
		  .frm1.txtHconBp_cd.value = .frm1.txtconBp_cd.value
		  .frm1.txtHconMES_ITEM_CD.value = .frm1.txtconMES_ITEM_CD.value

       End With          
</Script>       
<%     
         
         End if
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
    on error resume next
    Err.Clear  
    
	lgStrSQL = "INSERT INTO S_ITEM_PGM_MASTER_KO441"
	lgStrSQL = lgStrSQL & "( MES_ITEM_CD, TESTER, TEST_METHOD, PURE, HOT,INDEX_T, SHOT, TEST_DIE, RELEASE_DATE, REMARK,  "
	lgStrSQL = lgStrSQL & " INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT ) "
	lgStrSQL = lgStrSQL & " VALUES(" 

	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","	
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(5),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(6),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(7),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(8),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(9),0)    & ", '1900-01-01',"
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(10), "''", "S")     & ","
'	lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(8)),NULL,"S")	& ","				
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")			& "," 
	lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")			& "," 	
	lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
	lgStrSQL = lgStrSQL & ")"  
 'Call svrmsgbox (lgstrsql, vbinformation, i_mkscript)
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
    
    lgStrSQL = "UPDATE  S_ITEM_PGM_MASTER_KO441 "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " PURE		    =       " & UNIConvNum(arrColVal(5),0)	& ","
    lgStrSQL = lgStrSQL & " HOT		        =       " & UNIConvNum(arrColVal(6),0)	& ","
    lgStrSQL = lgStrSQL & " INDEX_T		    =       " & UNIConvNum(arrColVal(7),0)	& ","
    lgStrSQL = lgStrSQL & " SHOT		    =       " & UNIConvNum(arrColVal(8),0)	& ","
    lgStrSQL = lgStrSQL & " TEST_DIE		=       " & UNIConvNum(arrColVal(9),0)	& ","
    lgStrSQL = lgStrSQL & " TEST_METHOD 	=       " & FilterVar(UCase(arrColVal(4)), "''", "S") & ", "
    lgStrSQL = lgStrSQL & " REMARK 	        =       " & FilterVar(arrColVal(10),"''", "S") & ", "
 '   lgStrSQL = lgStrSQL & " RELEASE_DATE	=       " & FilterVar(UNIConvDate(arrColVal(8)),NULL,"S") & ", "	
        
    lgStrSQL = lgStrSQL & " UPDT_DT			=       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID	=       " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE					"
    lgStrSQL = lgStrSQL & " MES_ITEM_CD		=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TESTER			=       " & FilterVar(UCase(arrColVal(3)), "''", "S") 

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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  S_ITEM_PGM_MASTER_KO441 WHERE "
    lgStrSQL = lgStrSQL & " MES_ITEM_CD		=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TESTER			=       " & FilterVar(UCase(arrColVal(3)), "''", "S") 

'Response.Write lgStrSQL
'Response.End 
'call svrmsgbox(lgStrSQL, vbinformation, i_mkscript)
	
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
                       lgStrSQL = lgStrSQL & " MES_ITEM_CD, TESTER, PURE,HOT,INDEX_T, SHOT, TEST_DIE, TEST_METHOD, RELEASE_DATE, REMARK, "   
                       lgStrSQL = lgStrSQL & "    ITEM_NM = dbo.ufn_x_getcodename('b_mes_item', MES_ITEM_CD,''), "  
                       lgStrSQL = lgStrSQL & "    TESTER_NM = dbo.ufn_x_getcodename('b_user_minor', tester,'zz504') "  
                       lgStrSQL = lgStrSQL & " FROM S_ITEM_PGM_MASTER_KO441 "
                       lgStrSQL = lgStrSQL & " WHERE 1=1 "  & pCode 
                       lgStrSQL = lgStrSQL & " ORDER BY MES_ITEM_CD, TESTER " 

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
        Case "MX"
                 Call DisplayMsgBox("800350", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MY"
                 Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
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
		  Else
                Parent.DBQueryFail  		  	         
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select       
       
</Script>	
