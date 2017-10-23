<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

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
    Dim strWhere
    Dim strallow_nm
    Dim strPAY_CD  
    Dim Temp  
    'Call svrmsgbox(lgKeyStream(5),0,1)   
    strWhere = FilterVar(lgKeyStream(0), "''", "S")  '비교년월ZZZ
    strWhere = strWhere & " OR HDF040T.pay_yymm LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ")"    '기준년 
       
    strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(HDF040T.emp_no," & FilterVar(lgKeyStream(5),"'%'", "S") & " )  between   " & FilterVar(lgKeyStream(2), "''", "S") & ""       '  internal_cd = min
    strWhere = strWhere & " AND   " & FilterVar(lgKeyStream(3), "''", "S") & ""       '  internal_cd = max
    strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(HDF040T.emp_no," & FilterVar(lgKeyStream(5),"'%'", "S") & " ) LIKE " & FilterVar(lgKeyStream(4) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'

    strWhere = strWhere & " AND HDA010T.allow_cd = HDF040T.allow_cd "
    strWhere = strWhere & " AND HDA010T.code_type = " & FilterVar("1", "''", "S") & "  "
    strWhere = strWhere & " AND HDF040T.prov_type = " & FilterVar("1", "''", "S") & "  "
    strWhere = strWhere & " AND HDF040T.emp_no = HAA010T.emp_no "
    strWhere = strWhere & " group by HDF040T.allow_cd , HDA010T.allow_nm "
    strWhere = strWhere & " order by HDF040T.allow_cd "        
  
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

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD") )  '수당코드 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))   '수당코드명 
                                            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BAS_GG_AMT"), ggAmtOfMoney.DecPoint,0) 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BAS_GG_TOTAL_AMT"), ggAmtOfMoney.DecPoint,0) 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM01_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM02_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM03_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM04_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM05_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM06_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM07_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM08_AMT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM09_AMT"), ggAmtOfMoney.DecPoint,0)  
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM10_AMT"), ggAmtOfMoney.DecPoint,0)  
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM11_AMT"), ggAmtOfMoney.DecPoint,0)  
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MM12_AMT"), ggAmtOfMoney.DecPoint,0)  
          
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
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

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
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
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
                        lgStrSQL = "Select hdf040t.allow_cd , hda010t.allow_nm , "
                        lgStrSQL = lgStrSQL & " sum(case when hdf040t.pay_yymm =  " & FilterVar(lgKeyStream(0), "''", "S") & " then allow else 0 end)/1000 as BAS_GG_AMT, "
                        lgStrSQL = lgStrSQL & " sum(case when hdf040t.pay_yymm LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " then allow else 0 end)/1000 as BAS_GG_TOTAL_AMT," 
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("01", "''", "S") & " then allow else 0 end)/1000 as MM01_AMT," 
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""      
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("02", "''", "S") & " then allow else 0 end)/1000 as MM02_AMT," 
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("03", "''", "S") & " then allow else 0 end)/1000 as MM03_AMT," 
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("04", "''", "S") & " then allow else 0 end)/1000 as MM04_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("05", "''", "S") & " then allow else 0 end)/1000 as MM05_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("06", "''", "S") & " then allow else 0 end)/1000 as MM06_AMT, " 
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("07", "''", "S") & " then allow else 0 end)/1000 as MM07_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("08", "''", "S") & " then allow else 0 end)/1000 as MM08_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("09", "''", "S") & " then allow else 0 end)/1000 as MM09_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("10", "''", "S") & " then allow else 0 end)/1000 as MM10_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("11", "''", "S") & " then allow else 0 end)/1000 as MM11_AMT, "
        
                        lgStrSQL = lgStrSQL & " sum(case when substring(hdf040t.pay_yymm, 1, 4) =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
                        lgStrSQL = lgStrSQL & " and substring(hdf040t.pay_yymm, 5, 2) = " & FilterVar("12", "''", "S") & " then allow else 0 end)/1000 as MM12_AMT " 
        
                        lgStrSQL = lgStrSQL & " From HDA010T ,HDF040T , HAA010T "
                        lgStrSQL = lgStrSQL & " Where (HDF040T.PAY_YYMM " & pComp & pCode
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
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    

</Script>
