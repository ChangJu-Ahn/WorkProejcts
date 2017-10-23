<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q","A","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("Q", "A","NOCOOKIE","MB")
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
	Const C_SHEETMAXROWS_D  = 100
    '---------------------------------------Common-----------------------------------------------------------
	Dim	iLoopCount
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)
    
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)


	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Call SubBizQueryMulti()   

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
        Case  OPMD_CMODE                                                             '¢Ð: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '¢Ð: Update
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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  B_MAJOR"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
	Dim strWhere
    Dim txtSecurityCd, txtSecurity_TypeCd, txtMajorCd, rdoGiFlag, rdoYiFlag
    Dim SECURITY_CD, SECURITY_TYPE, BIZ_AREA_CD
	Dim lgCurrency
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	
	If Trim(lgKeyStream(4)) = "Y" Then     
	    strWhere =  " and A.CALCU_YN = " & FilterVar("Y", "''", "S") & "  "
	Elseif Trim(lgKeyStream(4)) = "N" Then
	    strWhere =  " and A.CALCU_YN = " & FilterVar("N", "''", "S") & "  "
	End if    
  
    If Trim(lgKeyStream(3)) = "Y" Then     
	    strWhere =   strWhere & " and ISNULL(a.GL_NO, '') <> ''"
	ELSEif Trim(lgKeyStream(3)) = "N" then
	    strWhere =   strWhere & " and ISNULL(a.GL_NO, '') = ''"
	END IF    
    
	If Trim(lgKeyStream(0)) <> "" Then
		SECURITY_CD = FilterVar(lgKeyStream(0) & "%", "''", "S")    
		strWhere = strWhere & " and a.SECURITY_CD LIKE " & SECURITY_CD
		Call CommonQueryRs(" SECURITY_NM "," A_SECURITY "," SECURITY_CD =  " & FilterVar(lgKeyStream(0), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtSecurityCd = ""
		else   
		  txtSecurityCd = Trim(Replace(lgF0,Chr(11),""))
		end if    	    
	else 
		txtSecurityCd = ""
	End If
	
	If Trim(lgKeyStream(1)) <> "" Then
		SECURITY_TYPE = FilterVar(lgKeyStream(1), "''", "S")    
		strWhere = strWhere & " and a.SECURITY_TYPE = " & SECURITY_TYPE
		Call CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1031", "''", "S") & "  AND MINOR_CD=  " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtSecurity_TypeCd = ""
		else   
		  txtSecurity_TypeCd = Trim(Replace(lgF0,Chr(11),""))
		end if    	    
	else 
		txtSecurity_TypeCd = ""
	End If
	
	If Trim(lgKeyStream(2)) <> "" Then
		BIZ_AREA_CD = FilterVar(lgKeyStream(2), "''", "S")    
		strWhere = strWhere & " and a.BIZ_AREA_CD = " & BIZ_AREA_CD
		Call CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA "," BIZ_AREA_CD = " & FilterVar(lgKeyStream(2), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtMajorCd = ""
		else   
		  txtMajorCd = Trim(Replace(lgF0,Chr(11),""))
		end if    	    
	else 
		txtMajorCd = ""
	End If

	Call SubMakeSQLStatements("MR",strWhere,"X",pCode)                                 '¡Ù: Make sql statements
 
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iLoopCount = 0
         
        Do While Not lgObjRs.EOF
            iLoopCount =  iLoopCount + 1
            If iLoopCount > C_SHEETMAXROWS_D  Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If 
          
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
			lgCurrency = Trim(ConvSPChars(lgObjRs(3)))
 	    	lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(4), lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
  	    	lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(5), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") 
 	    	lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(6), lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
  	    	lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(7), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))
  	    	lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(9), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") 
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))
            If lgObjRs(11) = "Y" Then
				lgstrData = lgstrData & Chr(11) & "°è»ê"
            Else 
				lgstrData = lgstrData & Chr(11) & "¹Ì°è»ê"
            End If
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))
       
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iLoopCount
            lgstrData = lgstrData & Chr(11) & Chr(12)


 		    lgObjRs.MoveNext

       Loop 
    End If
   
    If iLoopCount <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If   

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet 
	
	
%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
      With Parent.Frm1
             .txtSecurityNm.Value     = "<%=txtSecurityCd%>"            
             .txtSecurity_TypeNm.Value     = "<%=txtSecurity_TypeCd%>"            
             .txtMajorName.Value  = "<%=txtMajorCd%>"        
      End With          
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>       
<%     

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

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
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    txtGlNo = FilterVar(lgKeyStream(0), "''", "S")

    lgStrSQL = "INSERT INTO B_MAJOR("
    lgStrSQL = lgStrSQL & " MAJOR_CD     ," 
    lgStrSQL = lgStrSQL & " MAJOR_NM     ," 
    lgStrSQL = lgStrSQL & " MINOR_LEN    ," 
    lgStrSQL = lgStrSQL & " TYPE         ," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtBizArea")), "''", "S") 
    lgStrSQL = lgStrSQL & ")"
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  M_LC_HDR"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " BIZ_AREA      = " & FilterVar(Request("txtBizArea"), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " PARTIAL_SHIP  = " & FilterVar(Request("cboYesNo"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE           "
    lgStrSQL = lgStrSQL & " LC_NO         = " & FilterVar(lgKeyStream(0), "''", "S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO B_MAJOR("
    lgStrSQL = lgStrSQL & " MAJOR_CD     ," 
    lgStrSQL = lgStrSQL & " MAJOR_NM     ," 
    lgStrSQL = lgStrSQL & " MINOR_LEN    ," 
    lgStrSQL = lgStrSQL & " TYPE         ," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))), "", "D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
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
    lgStrSQL = "UPDATE  B_MAJOR"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MAJOR_NM   = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " MINOR_LEN  = " & FilterVar(Trim(UCase(arrColVal(4))), "", "D")   & ","
    lgStrSQL = lgStrSQL & " TYPE       = " & FilterVar(UCase(arrColVal(5)), "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    
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
    lgStrSQL = "DELETE  B_MAJOR"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount


    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL = "Select * " 
                                   lgStrSQL = lgStrSQL & " From  M_LC_HDR "
                                   lgStrSQL = lgStrSQL & " WHERE LC_NO " & pComp & pCode 	
                             Case "P"
                                   lgStrSQL = "Select TOP 1 * " 
                                   lgStrSQL = lgStrSQL & " From  M_LC_HDR "
                                   lgStrSQL = lgStrSQL & " WHERE LC_NO < " & pCode 	
                                   lgStrSQL = lgStrSQL & " ORDER BY LC_NO DESC "
                             Case "N"
                                   lgStrSQL = "Select TOP 1 * " 
                                   lgStrSQL = lgStrSQL & " From  M_LC_HDR "
                                   lgStrSQL = lgStrSQL & " WHERE LC_NO > " & pCode 	
                                   lgStrSQL = lgStrSQL & " ORDER BY LC_NO ASC "
                        End Select
               Case "D"
                        lgStrSQL = "Select * " 
                        lgStrSQL = lgStrSQL & " From  M_LC_HDR "
                        lgStrSQL = lgStrSQL & " WHERE LC_NO " & pComp & pCode 	
               Case "U"
                        lgStrSQL = "Select * " 
                        lgStrSQL = lgStrSQL & " From  M_LC_HDR "
                        lgStrSQL = lgStrSQL & " WHERE LC_NO " & pComp & pCode 	
               Case "C"
                        lgStrSQL = "Select * " 
                        lgStrSQL = lgStrSQL & " From  M_LC_HDR "
                        lgStrSQL = lgStrSQL & " WHERE LC_NO " & pComp & pCode 	
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  M_LC_DTL "
                       lgStrSQL = lgStrSQL & " WHERE LC_NO  " & pComp & pCode 	
                       lgStrSQL = lgStrSQL & " AND   LC_SEQ " & pComp & pCode1 	
               Case "D"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  M_LC_DTL "
                       lgStrSQL = lgStrSQL & " WHERE LC_NO  " & pComp & pCode 	
                       lgStrSQL = lgStrSQL & " AND   LC_SEQ " & pComp & pCode1 	
               Case "R"
                       lgStrSQL = "SELECT A.SECURITY_CD, A.SECURITY_NM, B.MINOR_NM, A.DOC_CUR, ISNULL(A.BUY_AMT,0), ISNULL(A.LOC_BUY_AMT,0), " 
                       lgStrSQL = lgStrSQL & " ISNULL(A.PRICE_AMT,0), ISNULL(A.LOC_PRICE_AMT,0), ISNULL(A.CNT,0),  "
                       lgStrSQL = lgStrSQL & " (ISNULL(A.PRICE_AMT,0) * ISNULL(A.CNT,0)) AS SUM_SECURITY, "
                       lgStrSQL = lgStrSQL & " (ISNULL(A.LOC_PRICE_AMT,0) * ISNULL(A.CNT,0)) AS SUM_LOC_SECURITY, "
                       lgStrSQL = lgStrSQL & " A.CALCU_YN, CASE WHEN ISNULL(A.GL_NO,'') ='' THEN " & FilterVar("¹Ì½ÂÀÎ", "''", "S") & "  ELSE " & FilterVar("½ÂÀÎ", "''", "S") & "  END "
                       lgStrSQL = lgStrSQL & " FROM A_SECURITY A, B_MINOR B "
                       lgStrSQL = lgStrSQL & " WHERE A.SECURITY_TYPE *= B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("A1031", "''", "S") & "  "	
                       lgStrSQL = lgStrSQL & pCode
                       lgStrSQL = lgStrSQL & " Order By A.SECURITY_CD asc "
               Case "U"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  M_LC_DTL "
                       lgStrSQL = lgStrSQL & " WHERE LC_NO  " & pComp & pCode 	
                       lgStrSQL = lgStrSQL & " AND   LC_SEQ " & pComp & pCode1 	
           End Select    
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>" , "F"                               '¢Ð : Display data
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"                          '¢Ð : Next next data tag 
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
				Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=lgLngMaxRow + 1%>" , "<%=lgLngMaxRow + iLoopCount %>" ,.C_DOC_CUR ,.C_BUY_AMT ,   "A" ,"Q","X","X")

				Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=lgLngMaxRow + 1%>" , "<%=lgLngMaxRow + iLoopCount %>" ,.C_DOC_CUR ,.C_PRICE_AMT ,   "A" ,"Q","X","X")
				Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=lgLngMaxRow + 1%>" , "<%=lgLngMaxRow + iLoopCount %>" ,.C_DOC_CUR ,.C_PRICE_SUM ,   "A" ,"Q","X","X")

                .DBQueryOk        
	         End with
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
