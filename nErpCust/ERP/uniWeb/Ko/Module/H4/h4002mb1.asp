<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    call LoadBasisGlobalInf()
        
    lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    
	
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection    
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = FilterVar(lgKeyStream(0), "''", "S")                                                      ':wk_type //근무조 
    strWhere = strWhere & " AND HCA030T.DEPT_CD = " & FilterVar(lgKeyStream(1), "''", "S")              ':dept_cd //부서코드 
    strWhere = strWhere & " AND HCA030T.HOL_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")             ':hol_type //휴일구분 
    strWhere = strWhere & " AND B_ACCT_DEPT.ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT)"
    strWhere = strWhere & " FROM B_ACCT_DEPT"
    strWhere = strWhere & " WHERE ORG_CHANGE_DT <= " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(3),NULL),"NULL", "S") & ")"        ' :apply_strt_dt//적용시작일 
	strWhere = strWhere & " AND HCA030T.APPLY_STRT_DT <= " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(3),NULL),"NULL", "S")      ':apply_strt_dt    

    Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("APPLY_STRT_DT"),NULL)            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_STRT_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_STRT_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("WK_STRT_HHMM")) 'FormatDateTime(lgObjRs("WK_STRT_HHMM"),4)
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_END_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_END_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("WK_END_HHMM"))'FormatDateTime(lgObjRs("WK_END_HHMM"),4)
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("BAS_HHMM"))'FormatDateTime(lgObjRs("BAS_HHMM"),4)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA1_STRT_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA1_STRT_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("RELA1_STRT_HHMM"))'FormatDateTime(lgObjRs("RELA1_STRT_HHMM"),4)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA1_END_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA1_END_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("RELA1_END_HHMM"))'FormatDateTime(lgObjRs("RELA1_END_HHMM"),4)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA2_STRT_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA2_STRT_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("RELA2_STRT_HHMM"))'FormatDateTime(lgObjRs("RELA2_STRT_HHMM"),4)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA2_END_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RELA2_END_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(lgObjRs("RELA2_END_HHMM"))'FormatDateTime(lgObjRs("RELA2_END_HHMM"),4)

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

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)
End Sub
'============================================================================================================
' Name : ConvToSSSSetTime(iVal)
' Desc : 
'============================================================================================================
Function ConvToSSSSetTime(iVal)
				
	If Trim(IVal) = ":" Or Trim(IVal) = "00:" Or Trim(IVal) = ":00" Or Trim(IVal) = "00:00" Then
		ConvToSSSSetTime = "00:00:00"
	Else		
		ConvToSSSSetTime = Trim(IVal) & ":00"		 
	End If
End Function
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()
    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                '☜: Split Row    data	
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
' Name : InsertSplitHHMM
' Desc : Split Time And Minute by ":"
'============================================================================================================
Function InsertSplitHHMM (HHMM)
Dim HHMMArr
Dim strCon

    HHMMArr = Split(HHMM,":")
    If UBound(HHMMArr) > 0 Then
        If FilterVar(HHMMArr(0), "''", "S") = "''" Then
            strCon = strCon & "" & FilterVar("00", "''", "S") & ","
        Else
            strCon = strCon & FilterVar(HHMMArr(0), "''", "S") & ","
        End If
        If FilterVar(HHMMArr(1), "''", "S") = "''" Then
            strCon = strCon & "" & FilterVar("00", "''", "S") & ","
        Else
            strCon = strCon & FilterVar(HHMMArr(1), "''", "S") & ","
        End If
        InsertSplitHHMM = strCon
    Else
        InsertSplitHHMM = "" & FilterVar("00", "''", "S") & "," & FilterVar("00", "''", "S") & ","
    End If
End Function
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim strCon
	lgStrSQL    = ""
    lgStrSQL = "INSERT INTO HCA030T("
    lgStrSQL = lgStrSQL & " WK_TYPE,DEPT_CD,HOL_TYPE,APPLY_STRT_DT,WK_CD,WK_STRT_TYPE,WK_STRT_HH,WK_STRT_MM," 
    lgStrSQL = lgStrSQL & " WK_END_TYPE,WK_END_HH,WK_END_MM,BAS_HH,BAS_MM,RELA1_STRT_TYPE,RELA1_STRT_HH,RELA1_STRT_MM,"
    lgStrSQL = lgStrSQL & " RELA1_END_TYPE,RELA1_END_HH,RELA1_END_MM,RELA2_STRT_TYPE,RELA2_STRT_HH,RELA2_STRT_MM,"
    lgStrSQL = lgStrSQL & " RELA2_END_TYPE,RELA2_END_HH,RELA2_END_MM,RELA3_STRT_TYPE,RELA3_STRT_HH,RELA3_STRT_MM,"
    lgStrSQL = lgStrSQL & " RELA3_END_TYPE,RELA3_END_HH,RELA3_END_MM,RELA4_STRT_TYPE,RELA4_STRT_HH,RELA4_STRT_MM,"
    lgStrSQL = lgStrSQL & " RELA4_END_TYPE,RELA4_END_HH,RELA4_END_MM,ISRT_DT,ISRT_EMP_NO,UPDT_DT,UPDT_EMP_NO)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NUll),"NULL","S")       & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)                                            & ","
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(8))
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)                                            & ","
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(10))
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(11))
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0)                                           & ","
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(13))
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(14),0)                                           & ","
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(15))
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(16),0)                                           & ","
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(17))
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(18),0)                                           & ","
    lgStrSQL = lgStrSQL & InsertSplitHHMM(arrColVal(19))    
    lgStrSQL = lgStrSQL & "0,"              'Real3_strt_tyep,Real3_strt_hh,Real3_strt_mm,Real3_end_type,Real3_end_hh,Real3_end_mm
    lgStrSQL = lgStrSQL & "" & FilterVar("00", "''", "S") & "," & FilterVar("00", "''", "S") & ","
    lgStrSQL = lgStrSQL & "0,"
    lgStrSQL = lgStrSQL & "" & FilterVar("00", "''", "S") & "," & FilterVar("00", "''", "S") & ","
    lgStrSQL = lgStrSQL & "0,"          'Real4_strt_tyep,Real4_strt_hh,Real4_strt_mm,Real4_end_type,Real4_end_hh,Real4_end_mm
    lgStrSQL = lgStrSQL & "" & FilterVar("00", "''", "S") & "," & FilterVar("00", "''", "S") & ","
    lgStrSQL = lgStrSQL & "0,"
    lgStrSQL = lgStrSQL & "" & FilterVar("00", "''", "S") & "," & FilterVar("00", "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                                            & ","
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)    
End Sub
'============================================================================================================
' Name : UpdateSplitHHMM (HHMM,Orderidx)
' Desc : Split Time And Minute by ":"
'============================================================================================================
Function UpdateSplitHHMM (HHMM,Orderidx)
	Dim HHMMArr
	Dim strCon

    HHMMArr = Split(HHMM,":")
    If UBound(HHMMArr) > 0 Then
        If Trim(UCase(Orderidx)) = "HH" Then
            If FilterVar(HHMMArr(0), "''", "S") = "''" Then
                strCon = strCon & "" & FilterVar("00", "''", "S") & ","
            Else
                strCon = strCon & FilterVar(HHMMArr(0), "''", "S") & ","
            End If
        Else
            If FilterVar(HHMMArr(1), "''", "S") = "''" Then
                strCon = strCon & "" & FilterVar("00", "''", "S") & ","
            Else
                strCon = strCon & FilterVar(HHMMArr(1), "''", "S") & ","
            End If
        End If
        UpdateSplitHHMM = strCon
    Else
        UpdateSplitHHMM = "" & FilterVar("00", "''", "S") & ","
    End If
End Function
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE HCA030T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " WK_STRT_TYPE = " & UNIConvNum(arrColVal(7),0)                   & ","
    lgStrSQL = lgStrSQL & " WK_STRT_HH = " & UpdateSplitHHMM(arrColVal(8),"HH")
    lgStrSQL = lgStrSQL & " WK_STRT_MM = " & UpdateSplitHHMM(arrColVal(8),"MM")
    lgStrSQL = lgStrSQL & " WK_END_TYPE = " & UNIConvNum(arrColVal(9),0)                    & ","
    lgStrSQL = lgStrSQL & " WK_END_HH = " & UpdateSplitHHMM(arrColVal(10),"HH")
    lgStrSQL = lgStrSQL & " WK_END_MM = " & UpdateSplitHHMM(arrColVal(10),"MM")
    lgStrSQL = lgStrSQL & " BAS_HH = " & UpdateSplitHHMM(arrColVal(11),"HH")
    lgStrSQL = lgStrSQL & " BAS_MM = " & UpdateSplitHHMM(arrColVal(11),"MM")
    lgStrSQL = lgStrSQL & " RELA1_STRT_TYPE = " & UNIConvNum(arrColVal(12),0)               & ","
    lgStrSQL = lgStrSQL & " RELA1_STRT_HH = " & UpdateSplitHHMM(arrColVal(13),"HH")
    lgStrSQL = lgStrSQL & " RELA1_STRT_MM = " & UpdateSplitHHMM(arrColVal(13),"MM")
    lgStrSQL = lgStrSQL & " RELA1_END_TYPE = " & UNIConvNum(arrColVal(14),0)               & ","
    lgStrSQL = lgStrSQL & " RELA1_END_HH = " & UpdateSplitHHMM(arrColVal(15),"HH")
    lgStrSQL = lgStrSQL & " RELA1_END_MM = " & UpdateSplitHHMM(arrColVal(15),"MM")
    lgStrSQL = lgStrSQL & " RELA2_STRT_TYPE = " & UNIConvNum(arrColVal(16),0)               & ","
    lgStrSQL = lgStrSQL & " RELA2_STRT_HH = " & UpdateSplitHHMM(arrColVal(17),"HH")
    lgStrSQL = lgStrSQL & " RELA2_STRT_MM = " & UpdateSplitHHMM(arrColVal(17),"MM")
    lgStrSQL = lgStrSQL & " RELA2_END_TYPE = " & UNIConvNum(arrColVal(18),0)                & ","
    lgStrSQL = lgStrSQL & " RELA2_END_HH = " & UpdateSplitHHMM(arrColVal(19),"HH")
    lgStrSQL = lgStrSQL & " RELA2_END_MM = " & UpdateSplitHHMM(arrColVal(19),"MM")
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgSvrDateTime,NULL,"S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE"
    lgStrSQL = lgStrSQL & " WK_TYPE = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND WK_CD = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DEPT_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND HOL_TYPE = " & FilterVar(UCase(arrColVal(4)), "''", "S")
  
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
	lgStrSQL =	""
    lgStrSQL = "DELETE  HCA030T"
    lgStrSQL = lgStrSQL & " WHERE"
    lgStrSQL = lgStrSQL & " WK_TYPE   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND WK_CD   = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DEPT_CD   = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND HOL_TYPE   = " & FilterVar(UCase(arrColVal(4)), "''", "S")


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
                       lgStrSQL = "SELECT TOP " & iSelCount  & " HCA030T.WK_TYPE,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0048", "''", "S") & ",HCA030T.WK_CD) WK_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0111", "''", "S") & ",HCA030T.WK_STRT_TYPE) WK_STRT_TYPE_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0111", "''", "S") & ",HCA030T.WK_END_TYPE) WK_END_TYPE_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0111", "''", "S") & ",HCA030T.RELA1_STRT_TYPE) RELA1_STRT_TYPE_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0111", "''", "S") & ",HCA030T.RELA2_STRT_TYPE) RELA2_STRT_TYPE_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0111", "''", "S") & ",HCA030T.RELA1_END_TYPE) RELA1_END_TYPE_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0111", "''", "S") & ",HCA030T.RELA2_END_TYPE) RELA2_END_TYPE_NM,"
                       lgStrSQL = lgStrSQL & " HCA030T.DEPT_CD,HCA030T.HOL_TYPE,HCA030T.APPLY_STRT_DT,"
                       lgStrSQL = lgStrSQL & " HCA030T.WK_CD,  HCA030T.WK_STRT_TYPE,HCA030T.WK_STRT_HH,"
                       lgStrSQL = lgStrSQL & " HCA030T.WK_STRT_MM,HCA030T.WK_END_TYPE,HCA030T.WK_END_HH,"
                       lgStrSQL = lgStrSQL & " HCA030T.WK_END_MM,HCA030T.BAS_HH,HCA030T.BAS_MM,HCA030T.RELA1_STRT_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA1_STRT_HH,HCA030T.RELA1_STRT_MM,HCA030T.RELA1_END_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA1_END_HH,HCA030T.RELA1_END_MM,HCA030T.RELA2_STRT_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA2_STRT_HH,HCA030T.RELA2_STRT_MM,HCA030T.RELA2_END_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA2_END_HH,HCA030T.RELA2_END_MM,HCA030T.RELA3_STRT_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA3_STRT_HH,HCA030T.RELA3_STRT_MM,HCA030T.RELA3_END_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA3_END_HH,HCA030T.RELA3_END_MM,HCA030T.RELA4_STRT_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA4_STRT_HH,HCA030T.RELA4_STRT_MM,HCA030T.RELA4_END_TYPE,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA4_END_HH,HCA030T.RELA4_END_MM,HCA030T.ISRT_DT,"
                       lgStrSQL = lgStrSQL & " HCA030T.ISRT_EMP_NO,HCA030T.UPDT_DT,HCA030T.UPDT_EMP_NO,"
                       lgStrSQL = lgStrSQL & " (HCA030T.WK_STRT_HH + ':' + HCA030T.WK_STRT_MM) AS WK_STRT_HHMM,(HCA030T.WK_END_HH +':'+ HCA030T.WK_END_MM) AS WK_END_HHMM,"'
                       lgStrSQL = lgStrSQL & " (HCA030T.BAS_HH +':'+ HCA030T.BAS_MM) AS BAS_HHMM,(HCA030T.RELA1_STRT_HH +':'+ HCA030T.RELA1_STRT_MM) AS RELA1_STRT_HHMM,"
                       lgStrSQL = lgStrSQL & " (HCA030T.RELA1_END_HH +':'+ HCA030T.RELA1_END_MM) AS RELA1_END_HHMM,(HCA030T.RELA2_STRT_HH +':'+ HCA030T.RELA2_STRT_MM) AS RELA2_STRT_HHMM,"
                       lgStrSQL = lgStrSQL & " (HCA030T.RELA2_END_HH +':'+ HCA030T.RELA2_END_MM) AS RELA2_END_HHMM,HCA030T.RELA3_STRT_HH +':'+ HCA030T.RELA3_STRT_MM,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA3_END_HH + HCA030T.RELA3_END_MM,HCA030T.RELA4_STRT_HH + HCA030T.RELA4_STRT_MM,"
                       lgStrSQL = lgStrSQL & " HCA030T.RELA4_END_HH + HCA030T.RELA4_END_MM"
                       lgStrSQL = lgStrSQL & " FROM HCA030T,B_ACCT_DEPT"
                       lgStrSQL = lgStrSQL & " WHERE HCA030T.DEPT_CD = B_ACCT_DEPT.dept_cd"
                       lgStrSQL = lgStrSQL & " AND HCA030T.WK_TYPE " & pComp & pCode
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"                                                                '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0002%>"                                                        '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else
          End If
    End Select
</Script>
