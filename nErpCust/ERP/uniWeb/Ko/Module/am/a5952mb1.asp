<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    On Error Resume Next
    Err.Clear
    Call LoadBasisGlobalInf() 
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	Dim txtRegnm
    'Multi SpreadSheet
	Const C_SHEETMAXROWS_D  = 100

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)


    'Multi Multi SpreadSheet
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '¢Ð: "M"(Spread #1) "S"(Spread #2)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             lgCurrentSpd = lgKeyStream(1)
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection


'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
    Call SubBizQueryMulti()
End Sub	

'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim Currency_code

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
 	If lgKeyStream(0) <> "" Then
	 strWhere = "and  a.reg_cd = " & FilterVar(lgKeyStream(0), "''", "S") 
	 Call CommonQueryRs("a.minor_nm","b_minor a, b_major b","a.major_cd = b.major_cd and a.minor_type = " & FilterVar("S", "''", "S") & "  and a.major_cd = " & FilterVar("a1029", "''", "S") & "  and a.minor_cd = " & FilterVar(lgKeyStream(0), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If Trim(Replace(lgF0,Chr(11),"")) = "X" then
	       txtRegnm= ""
	   Else
	       txtRegnm= Trim(Replace(lgF0,Chr(11),""))
	   End if
    Else
	     txtRegnm= ""
    End If
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                   '¡Ù: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        'lgStrPrevKeyIndex = ""
        'If lgCurrentSpd = "M" Then
        '   Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        'End If   
        'Call SetErrorStatus()
    Else
        lgstrData = ""
        
        iDx = 1
        Do While Not lgObjRs.EOF
        
            		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    '---HIDDEN
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))


                    
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))
                    '--HIDDEN
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            		lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            		lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
    End If

    Call SubMakeSQLStatements("MK",strWhere,"X",C_EQGT)                                 '¡Ù: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    Else
        lgstrData1 = ""
        
        iDx = 1
        Do While Not lgObjRs.EOF
		            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(0))
                    lgstrData1 = lgstrData1 & Chr(11) & ""
                    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(1))
                    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData1 = lgstrData1 & Chr(11) & ""
                    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(3))
                    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(4))
                    lgstrData1 = lgstrData1 & Chr(11) & ""
                    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs(5))

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
            lgstrData1 = lgstrData1 & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
    End If
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtRegnm.value = "<%=ConvSPChars(txtRegnm)%>"
	END With
</SCRIPT>
<%
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear
    
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
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    if lgCurrentSpd = "M" then
        lgStrSQL = "INSERT INTO A_MONTHly_ACCT("
    	lgStrSQL = lgStrSQL & " REG_CD     ,"
    	lgStrSQL = lgStrSQL & " DR_CR_FG     ,"
    	lgStrSQL = lgStrSQL & " REF1    ,"
    	lgStrSQL = lgStrSQL & " REF2    ,"
    	lgStrSQL = lgStrSQL & " REF3         ,"
    	lgStrSQL = lgStrSQL & " REF4,"
    	lgStrSQL = lgStrSQL & " REF5,"
    	lgStrSQL = lgStrSQL & " REF6,"
    	lgStrSQL = lgStrSQL & " ACCT_CD,"
    	lgStrSQL = lgStrSQL & " INSRT_USER_ID     ,"
    	lgStrSQL = lgStrSQL & " INSRT_DT ,"
    	lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,  "
    	lgStrSQL = lgStrSQL & " UPDT_DT    )"
    	lgStrSQL = lgStrSQL & " VALUES("
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                      & ","
    	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,Null,"S") & ","
    	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                      & ","
    	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,Null,"S")   
    	lgStrSQL = lgStrSQL & ")"
 	
    else
        lgStrSQL = "INSERT INTO A_MONTHly_ACCT("
    	lgStrSQL = lgStrSQL & " REG_CD     ,"
    	lgStrSQL = lgStrSQL & " DR_CR_FG     ,"
    	lgStrSQL = lgStrSQL & " REF1    ,"
    	lgStrSQL = lgStrSQL & " REF2    ,"
    	lgStrSQL = lgStrSQL & " REF3         ,"
    	lgStrSQL = lgStrSQL & " REF4,"
    	lgStrSQL = lgStrSQL & " REF5,"
    	lgStrSQL = lgStrSQL & " REF6,"
    	lgStrSQL = lgStrSQL & " ACCT_CD,"   '9
    	
    	lgStrSQL = lgStrSQL & " INSRT_USER_ID     ,"
    	lgStrSQL = lgStrSQL & " INSRT_DT ,"
    	lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,  "
    	lgStrSQL = lgStrSQL & " UPDT_DT		,  "    
    	lgStrSQL = lgStrSQL & " EVAL_METH)"
    	lgStrSQL = lgStrSQL & " VALUES("
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & "" & FilterVar("*", "''", "S") & " "     & ","
    	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
  	
    	
    	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                      & ","
    	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,Null,"S") & ","
    	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                      & ","
    	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,Null,"S")  & "," 
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")        '10
    	lgStrSQL = lgStrSQL & ")"
    end if

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	 lgObjConn.Execute lgStrSQL,,adCmdText
	 
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    if lgCurrentSpd = "M" then
        lgStrSQL = "UPDATE  A_MONTHly_aCCT"
    	lgStrSQL = lgStrSQL & " SET "
    	lgStrSQL = lgStrSQL & " REG_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " DR_CR_FG = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " REF1    = " & FilterVar(UCase(arrColVal(4)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " REF2    = " & FilterVar(UCase(arrColVal(5)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " REF3    = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " REF4    = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " REF5    = " & FilterVar(UCase(arrColVal(8)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " REF6    = " & FilterVar(UCase(arrColVal(9)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " ACCT_CD  = " & FilterVar(UCase(arrColVal(10)), "''", "S")  & ","
    	lgStrSQL = lgStrSQL & " UPDT_USER_ID  = " & FilterVar(gUsrId, "''", "S")                & ","
    	lgStrSQL = lgStrSQL & " UPDT_DT		 = " & FilterVar(GetSvrDateTime,Null,"S")  		
    	lgStrSQL = lgStrSQL & " WHERE  "
    	lgStrSQL = lgStrSQL & " REG_CD       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND DR_CR_FG = " & FilterVar(UCase(arrColVal(11)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF1   = " & FilterVar(UCase(arrColVal(12)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF2   = " & FilterVar(UCase(arrColVal(13)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF3   = " & FilterVar(UCase(arrColVal(14)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF4   = " & FilterVar(UCase(arrColVal(15)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF5   = " & FilterVar(UCase(arrColVal(16)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF6   = " & FilterVar(UCase(arrColVal(17)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND ACCT_CD  = " & FilterVar(UCase(arrColVal(18)), "''", "S")
    	

    else
        lgStrSQL = "UPDATE  A_MONTHly_ACCT "
    	lgStrSQL = lgStrSQL & " SET "
    	lgStrSQL = lgStrSQL & " REG_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")   & ","
    	lgStrSQL = lgStrSQL & " DR_CR_FG = " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " REF1    =  " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " REF2    = " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " REF3    = " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " REF4    = " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " REF5    = " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " REF6    = " & FilterVar("*", "''", "S") & " , " 
    	lgStrSQL = lgStrSQL & " ACCT_CD  = " & FilterVar(UCase(arrColVal(3)), "''", "S")  & ","
    	lgStrSQL = lgStrSQL & " EVAL_METH  = " & FilterVar(UCase(arrColVal(5)), "''", "S")  & ","
    	lgStrSQL = lgStrSQL & " UPDT_USER_ID  = " & FilterVar(gUsrId, "''", "S")                & ","
    	lgStrSQL = lgStrSQL & " UPDT_DT		 = " & FilterVar(GetSvrDateTime,Null,"S")  		
    	lgStrSQL = lgStrSQL & " WHERE  "
    	lgStrSQL = lgStrSQL & " REG_CD       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND DR_CR_FG = " & FilterVar("*", "''", "S") & "  "
    	lgStrSQL = lgStrSQL & " AND REF1   = " & FilterVar("*", "''", "S") & "  " 
    	lgStrSQL = lgStrSQL & " AND REF2   = " & FilterVar("*", "''", "S") & "  " 
    	lgStrSQL = lgStrSQL & " AND REF3   = " & FilterVar("*", "''", "S") & "  " 
    	lgStrSQL = lgStrSQL & " AND REF4   = " & FilterVar("*", "''", "S") & "  " 
    	lgStrSQL = lgStrSQL & " AND REF5   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND REF6   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND ACCT_CD  = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    end if    
'---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    if lgCurrentSpd = "M" then
    	lgStrSQL = "DELETE  A_MONTHly_ACCT"
    	lgStrSQL = lgStrSQL & " WHERE  "
    	lgStrSQL = lgStrSQL & " REG_CD       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND DR_CR_FG = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF1   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF2   = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF3   = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF4   = " & FilterVar(UCase(arrColVal(7)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF5   = " & FilterVar(UCase(arrColVal(8)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND REF6   = " & FilterVar(UCase(arrColVal(9)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND ACCT_CD  = " & FilterVar(UCase(arrColVal(10)), "''", "S")
    else
    	lgStrSQL = "DELETE  A_MONTHly_ACCT"
    	lgStrSQL = lgStrSQL & " WHERE  "
    	lgStrSQL = lgStrSQL & " REG_CD       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    	lgStrSQL = lgStrSQL & " AND DR_CR_FG = " & FilterVar("*", "''", "S") & "  "
    	lgStrSQL = lgStrSQL & " AND REF1   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND REF2   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND REF3   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND REF4   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND REF5   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND REF6   = " & FilterVar("*", "''", "S") & " " 
    	lgStrSQL = lgStrSQL & " AND ACCT_CD  = " & 	FilterVar(UCase(arrColVal(3)), "''", "S")
    end if
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next
    Err.Clear
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext 
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
           
           Select Case Mid(pDataType,2,1)
               Case "C"
                      
               Case "K"
                      lgStrSQL = " Select Top " &iSelCount& " a.reg_cd reg_cd ,b.minor_nm minor_nm "
						lgStrSQL = lgStrSQL & " ,a.acct_cd acct_cd ,d.acct_nm acct_nm,isnull(EVAL_METH,''),isnull(f.minor_nm,'') "
						lgStrSQL = lgStrSQL & " From  a_monthly_acct a,b_minor b , b_major c , a_acct d, a_monthly_base e, b_minor f "
						lgStrSQL = lgStrSQL & " WHERE a.reg_Cd = b.minor_cd "
 						lgStrSQL = lgStrSQL & " and  c.major_Cd  = b.major_Cd "
 						lgStrSQL = lgStrSQL & " and  d.acct_Cd = a.acct_cd "
 						lgStrSQL = lgStrSQL & " and  a.eval_meth  *= f.MINOR_CD  "
 						lgStrSQL = lgStrSQL & " and f.MAJOR_CD = " & FilterVar("A1045", "''", "S") & "  "
                        lgStrSQL = lgStrSQL & " and a.reg_cd = e.reg_cd "
                        lgStrSQL = lgStrSQL & " and e.use_yn = " & FilterVar("Y", "''", "S") & "  "
 						lgStrSQL = lgStrSQL & " and  a.dr_cr_Fg = " & FilterVar("*", "''", "S") & "  "
 						lgStrSQL = lgStrSQL & " and  a.ref1 = " & FilterVar("*", "''", "S") & "   "
 						lgStrSQL = lgStrSQL & " and  a.ref2 = " & FilterVar("*", "''", "S") & "   "
 						lgStrSQL = lgStrSQL & " and  a.ref3 = " & FilterVar("*", "''", "S") & "   "
 						lgStrSQL = lgStrSQL & " and  a.ref4 = " & FilterVar("*", "''", "S") & "   "
 						lgStrSQL = lgStrSQL & " and  a.ref5 = " & FilterVar("*", "''", "S") & "   "
 						lgStrSQL = lgStrSQL & " and  a.ref6 = " & FilterVar("*", "''", "S") & "   "
 						lgStrSQL = lgStrSQL & " and  c.major_cd = " & FilterVar("a1029", "''", "S") & "  "
 						lgStrSQL = lgStrSQL &pCode
                        lgStrSQL = lgStrSQL & "  order by a.reg_cd "
               Case "R"
                       lgStrSQL = "Select Top " &iSelCount& " a.reg_cd reg_cd ,b.minor_nm minor_nm"
                       lgStrSQL = lgStrSQL & ",a.acct_cd acct_cd ,d.acct_nm acct_nm" 
                       lgStrSQL = lgStrSQL & " ,a.dr_cr_fg " 
                       lgStrSQL = lgStrSQL & ",a.ref1 "
'					   lgStrSQL = lgStrSQL &  ",(select minor_nm from b_minor where major_cd ='h0071' and minor_cd =a.ref1) ref1_nm "
                       lgStrSQL = lgStrSQL & ",a.ref2 " 
                       lgStrSQL = lgStrSQL & ",a.ref3 " 
                       lgStrSQL = lgStrSQL & ",a.ref4 " 
                       lgStrSQL = lgStrSQL & ",a.ref5 " 
                       lgStrSQL = lgStrSQL & ",a.ref6 " 
'                       lgStrSQL = lgStrSQL &  ",(select minor_nm from b_minor where minor_cd = a.ref6 and major_cd='h0040')  song_name" 
                       lgStrSQL = lgStrSQL & " ,a.eval_meth, f.minor_nm " 

                       lgStrSQL = lgStrSQL & " From  a_monthly_acct a,b_minor b , b_major c , a_acct d, a_monthly_base e, b_minor f "
                       lgStrSQL = lgStrSQL & " WHERE a.reg_Cd = b.minor_cd"
                       lgStrSQL = lgStrSQL & " and a.reg_cd = e.reg_cd"
                       lgStrSQL = lgStrSQL & " and e.use_yn = " & FilterVar("Y", "''", "S") & " "
                       lgStrSQL = lgStrSQL & " and  a.dr_cr_Fg <> " & FilterVar("*", "''", "S") & " "
 					  ' lgStrSQL = lgStrSQL &  " and  a.ref1 <> '*'"
 					  ' lgStrSQL = lgStrSQL &  " and  a.ref2 <> '*'"
 					  ' lgStrSQL = lgStrSQL &  " and  a.ref3 <> '*'"
 					  ' lgStrSQL = lgStrSQL &  " and  a.ref4 <> '*'"
 					  ' lgStrSQL = lgStrSQL &  " and  a.ref5 <> '*'"
 					  ' lgStrSQL = lgStrSQL &  " and  a.ref6 <> '*'"
                       lgStrSQL = lgStrSQL & "  and  c.major_Cd  = b.major_Cd"
                       lgStrSQL = lgStrSQL & "  and  d.acct_Cd =* a.acct_cd"
                       lgStrSQL = lgStrSQL & " and  a.eval_meth  *= f.MINOR_CD  "
 					   lgStrSQL = lgStrSQL & " and  f.MAJOR_CD = " & FilterVar("A1045", "''", "S") & "  "
                       lgStrSQL = lgStrSQL & "  and  c.major_cd = " & FilterVar("a1029", "''", "S") & " "
                       lgStrSQL = lgStrSQL &pCode
                       lgStrSQL = lgStrSQL & "  order by a.reg_cd "
              End Select             
    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  'Can not create(Demo code)
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
                .frm1.vspdData.ReDraw = False
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"				
				.InitData()    
				.SetSpreadColor1
				.frm1.vspdData.ReDraw = True
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKeyIndex1    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData1%>"
				.SetSpreadColor2

                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
