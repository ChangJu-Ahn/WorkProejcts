<% Option Explicit%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    
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
    Dim strWhereHead
    Dim strWhereTab1
    Dim strWhereTab2
    Dim strWherefooter
    Dim txtFrom_dt
    Dim txtTo_dt
    Dim txtFr_internal_cd
    Dim txtTo_internal_cd
    Dim txtPay_grd1
    Dim gSelframeFlg
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If lgKeyStream(0) = "" Then
       txtFrom_dt = UniConvYYYYMMDDToDate(gDateFormat,"1900","01","01")
       txtFrom_dt = FilterVar(UNIConvDateCompanyToDB(txtFrom_dt,NULL),"NULL","S")
    Else
       txtFrom_dt = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
    End If
    
    If lgKeyStream(1) = "" Then
       txtTo_dt = UniConvYYYYMMDDToDate(gDateFormat,"2500","12","31")
       txtTo_dt = FilterVar(UNIConvDateCompanyToDB(txtTo_dt,NULL),"NULL","S")
    Else
       txtTo_dt = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(1)),NULL),"NULL","S")
    End If

    txtFr_internal_cd   = FilterVar(lgKeyStream(2),"'%'", "S")
    txtTo_internal_cd   = FilterVar(lgKeyStream(3),"'%'", "S")
    txtPay_grd1         = FilterVar(lgKeyStream(4),"'%'", "S")
    gSelframeFlg        = FilterVar(lgKeyStream(5),"" & FilterVar("1", "''", "S") & " ", "S")

    strWhereHead =txtFr_internal_cd
    strWhereHead =strWhereHead & " And dbo.ufn_H_get_internal_cd(emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") <= " & txtTo_internal_cd
    strWhereHead =strWhereHead & " And pay_grd1 LIKE " & txtPay_grd1
   
    strWhereTab1 = " AND entr_dt BETWEEN " & txtFrom_dt & " AND " & txtTo_dt   
    strWhereTab2 = " AND retire_dt BETWEEN " & txtFrom_dt & " AND " & txtTo_dt
    
    If gSelframeFlg= FilterVar("1", "''", "S")  Then
      strWherefooter = " ORDER BY entr_dt ASC, emp_no ASC "
    Else
      strWherefooter = " ORDER BY retire_dt ASC, emp_no ASC "
    End If

    If gSelframeFlg= FilterVar("1", "''", "S")  Then                                               'Tab 위치에 따라 SQL문을 변환 
        strWhere = strWhereHead & strWhereTab1 & strWherefooter
    Else
        strWhere = strWhereHead & strWhereTab2 & strWherefooter
    End If

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                      '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("entr_dt"),Null)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("resent_promote_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ocpt_type_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sex_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("retire_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("retire_resn_nm"))
            
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
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext 
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "C"
               Case "D"
               Case "R"               
                       lgStrSQL = "            Select TOP " & iSelCount
                       lgStrSQL = lgStrSQL & "					emp_no, "
                       lgStrSQL = lgStrSQL & "					name, "                       
'                       lgStrSQL = lgStrSQL & "					dbo.ufn_H_GetCodeName('H_CURRENT_DEPT',dept_cd,'') dept_nm, "
                       lgStrSQL = lgStrSQL & "  dept_nm, "                                               
                       lgStrSQL = lgStrSQL & "					dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm, "
                       lgStrSQL = lgStrSQL & "					entr_dt, resent_promote_dt,"                       
                       lgStrSQL = lgStrSQL & "					dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",ocpt_type) ocpt_type_nm, "
                       lgStrSQL = lgStrSQL & "					pay_grd1, "
                       lgStrSQL = lgStrSQL & "					dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",pay_grd1) pay_grd1_nm, "
                       lgStrSQL = lgStrSQL & "				    pay_grd2, "
                       lgStrSQL = lgStrSQL & "					dbo.ufn_GetCodeName(" & FilterVar("H0115", "''", "S") & ",sex_cd) sex_nm, "
                       lgStrSQL = lgStrSQL & "					retire_dt, "                       
                       lgStrSQL = lgStrSQL & "					dbo.ufn_GetCodeName(" & FilterVar("H0025", "''", "S") & ",retire_resn) retire_resn_nm "
                       lgStrSQL = lgStrSQL & " From  haa010t "
                       lgStrSQL = lgStrSQL & " Where dbo.ufn_H_get_internal_cd(emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") " & pComp & pCode
               Case "U"
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
