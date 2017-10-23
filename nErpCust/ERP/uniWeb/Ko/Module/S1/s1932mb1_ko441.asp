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
	Dim strBp_cd, strItem_cd
  
    On Error Resume Next    
    Err.Clear                                                               '¢Ð: Clear Error status

	strBp_cd = Trim(Request("txtconBp_cd"))
	strItem_cd = Trim(Request("txtconItem_cd"))

	strWhere = " And BP_CD = " & FilterVar(strBp_cd, "''", "S")

	If strItem_cd <> "" Then
		strWhere = strWhere & " And ITEM_CD = " & FilterVar(strItem_cd, "''", "S")
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


            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APPLY_YYMM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD") )                 
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRICE_CUR"))
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_APPLY_OPT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEST_APPLY_OPT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RETEST_RATE"))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EFR_RATE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("QEI_RATE") )                 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ADVANTAGE_SEC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_PRICE_OPT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_PRICE_OPT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OPERATOR"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUTPUT_RATE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("XCHG_RATE_OPT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("XCHG_RATE_OPT"))

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
		  .frm1.txtHconItem_cd.value = .frm1.txtconItem_cd.value
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
    
	lgStrSQL = "INSERT INTO S_Pricing_Confuguration_ITEM_KO441"
	lgStrSQL = lgStrSQL & "(APPLY_YYMM, BP_CD, ITEM_CD, PRICE_CUR, TEST_APPLY_OPT, RETEST_RATE, EFR_RATE, QEI_RATE, ADVANTAGE_SEC,   "
	lgStrSQL = lgStrSQL & " BASIC_PRICE_OPT, OPERATOR, OUTPUT_RATE, XCHG_RATE_OPT,  "
	lgStrSQL = lgStrSQL & " INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT ) "
	lgStrSQL = lgStrSQL & " VALUES(" 

	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","	'C_Apply_yymm
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(7),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(8),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(9),0)    & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(10),0)    & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(12)), "''", "S")     & ","
	lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(13),0)    & ","					'C_Output_rate
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(14)), "''", "S")     & ","
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
    
    lgStrSQL = "UPDATE  S_Pricing_Confuguration_ITEM_KO441 "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " PRICE_CUR	    =       " & FilterVar(UCase(arrColVal(5)), "''", "S")		& ","
    lgStrSQL = lgStrSQL & " TEST_APPLY_OPT	=       " & FilterVar(UCase(arrColVal(6)), "''", "S")		& ","
    lgStrSQL = lgStrSQL & " RETEST_RATE		=       " & UNIConvNum(arrColVal(7),0)	& ","
    lgStrSQL = lgStrSQL & " EFR_RATE		=       " & UNIConvNum(arrColVal(8),0)	& ","
    lgStrSQL = lgStrSQL & " QEI_RATE		=       " & UNIConvNum(arrColVal(9),0)	& ","
    lgStrSQL = lgStrSQL & " ADVANTAGE_SEC	=       " & UNIConvNum(arrColVal(10),0)	& ","
    lgStrSQL = lgStrSQL & " BASIC_PRICE_OPT	=       " & FilterVar(UCase(arrColVal(11)), "''", "S")		& ","
    lgStrSQL = lgStrSQL & " OPERATOR		=       " & FilterVar(UCase(arrColVal(12)), "''", "S")		& ","
    lgStrSQL = lgStrSQL & " OUTPUT_RATE		=       " & UNIConvNum(arrColVal(13),0)	& ","
    lgStrSQL = lgStrSQL & " XCHG_RATE_OPT	=       " & FilterVar(UCase(arrColVal(14)), "''", "S")		& ","
        
    lgStrSQL = lgStrSQL & " UPDT_DT			=       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID	=       " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE					"
    lgStrSQL = lgStrSQL & " BP_CD			=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " ITEM_CD			=       " & FilterVar(UCase(arrColVal(4)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " APPLY_YYMM		=       " & FilterVar(UCase(arrColVal(3)), "''", "S")

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

    lgStrSQL = "DELETE  S_Pricing_Confuguration_ITEM_KO441 WHERE "
    lgStrSQL = lgStrSQL & " BP_CD		=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " ITEM_CD			=       " & FilterVar(UCase(arrColVal(4)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " APPLY_YYMM	=       " & FilterVar(UCase(arrColVal(3)), "''", "S")
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
                       lgStrSQL = lgStrSQL & " BP_CD, ITEM_CD, APPLY_YYMM = LEFT(APPLY_YYMM,4)+'-'+RIGHT(APPLY_YYMM,2) , PRICE_CUR, TEST_APPLY_OPT, RETEST_RATE, EFR_RATE, QEI_RATE, "   
                       lgStrSQL = lgStrSQL & "    ADVANTAGE_SEC, BASIC_PRICE_OPT, OPERATOR, OUTPUT_RATE, XCHG_RATE_OPT, "  
                       lgStrSQL = lgStrSQL & "    BP_NM = dbo.ufn_x_getcodename('b_biz_partner', BP_CD,''), "  
                       lgStrSQL = lgStrSQL & "    ITEM_NM = dbo.ufn_x_getcodename('b_item', ITEM_CD,'') "  
                       lgStrSQL = lgStrSQL & " FROM S_Pricing_Confuguration_ITEM_KO441 "
                       lgStrSQL = lgStrSQL & " WHERE 1=1 "  & pCode 
                       lgStrSQL = lgStrSQL & " ORDER BY APPLY_YYMM, ITEM_CD " 


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
