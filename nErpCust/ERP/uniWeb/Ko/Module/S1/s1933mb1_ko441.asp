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
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
   
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim strWhere 
	Dim strBp_cd
  
    On Error Resume Next    
    Err.Clear                                                               '��: Clear Error status

	strBp_cd = Trim(Request("txtconBp_cd"))

	If strBp_cd <> "" Then
		strWhere = " And BP_CD = " & FilterVar(strBp_cd, "''", "S")
	End If
    	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '�� : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF


            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD") )                 
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TESTER"))
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TESTER_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_METHOD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_METHOD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_PRICE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CURRENCY"))
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("VALID_FROM_DT"))
          '  lgstrData = lgstrData & Chr(11) & ConvSPChars(UNIConvDate(lgObjRs("VALID_FROM_DT"),NULL,"S"))

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
            .Frm1.txtHconBp_cd.Value  = .Frm1.txtconBp_cd.Value                  'Set condition area
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
    Err.Clear                                                                        '��: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
    
	lgStrSQL = "INSERT INTO S_Test_Basic_Price_Cust_KO441"
	lgStrSQL = lgStrSQL & "( BP_CD, TESTER, TEST_METHOD, TEST_TYPE, CURRENCY, VALID_FROM_DT, BASIC_PRICE,  "
	lgStrSQL = lgStrSQL & " INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT ) "
	lgStrSQL = lgStrSQL & " VALUES(" 

	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","	'C_VALID_FROM_DT
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(7)),NULL,"S")	& ","				
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(8),0)    & ","
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")			& "," 
	lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")			& "," 	
	lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
	lgStrSQL = lgStrSQL & ")"  
' Call svrmsgbox (lgstrsql, vbinformation, i_mkscript)
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
    
    lgStrSQL = "UPDATE  S_Test_Basic_Price_Cust_KO441 "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " BASIC_PRICE		=       " & UNIConvNum(arrColVal(8),0)	& ","
        
    lgStrSQL = lgStrSQL & " UPDT_DT			=       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID	=       " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE					"
    lgStrSQL = lgStrSQL & " BP_CD			=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TESTER			=       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TEST_METHOD 	=       " & FilterVar(UCase(arrColVal(4)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TEST_TYPE 		=       " & FilterVar(UCase(arrColVal(5)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " CURRENCY   		=       " & FilterVar(UCase(arrColVal(6)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " VALID_FROM_DT	=       " & FilterVar(UNIConvDate(arrColVal(7)),NULL,"S")	

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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  S_Test_Basic_Price_Cust_KO441 WHERE "
    lgStrSQL = lgStrSQL & " BP_CD			=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TESTER			=       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TEST_METHOD 	=       " & FilterVar(UCase(arrColVal(4)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TEST_TYPE 		=       " & FilterVar(UCase(arrColVal(5)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " CURRENCY   		=       " & FilterVar(UCase(arrColVal(6)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " VALID_FROM_DT	=       " & FilterVar(UNIConvDate(arrColVal(7)),NULL,"S")	
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
                       lgStrSQL = lgStrSQL & " BP_CD, TESTER, TEST_METHOD, TEST_TYPE, CURRENCY, VALID_FROM_DT, BASIC_PRICE, "   
                       lgStrSQL = lgStrSQL & "    BP_NM = dbo.ufn_x_getcodename('b_biz_partner', BP_CD,''), "  
                       lgStrSQL = lgStrSQL & "    TESTER_NM = dbo.ufn_x_getcodename('b_user_minor', tester,'zz504') "  
                       lgStrSQL = lgStrSQL & " FROM S_Test_Basic_Price_Cust_KO441 "
                       lgStrSQL = lgStrSQL & " WHERE 1=1 "  & pCode 
                       lgStrSQL = lgStrSQL & " ORDER BY BP_CD, TESTER, VALID_FROM_DT " 


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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '�� : Query
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
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select       
       
</Script>	
