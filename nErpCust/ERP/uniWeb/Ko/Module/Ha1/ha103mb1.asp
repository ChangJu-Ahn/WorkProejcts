<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%	
	Dim lgGetSvrDateTime
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    lgGetSvrDateTime = GetSvrDateTime
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
              Call SubBizSaveMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strinternal_cd
    Dim StrWhere
    Dim arrYMD,strYMD
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0),"'%'", "S")    '사번 
    strinternal_cd = "" & FilterVar("%", "''", "S") & ""                          '자료관한이 필요하면 차후 ida.data_auth_lvl로 대치한다.
    strWhere = iKey1
    strWhere = strWhere & " And haa010t.internal_cd Like " & strinternal_cd
    strWhere = strWhere & " And hga040t.emp_no = haa010t.emp_no "
    strWhere = strWhere & " And ( haa010t.internal_cd  LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " ) "

    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                            '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("entr_dt"),Null)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("retire_dt"),Null)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("retire_pay_bas_dt"),Null)  
            'ufn_H_GetLongYYMMDD 에 함수 return 값에 한글 들어간 것을(xx년 xx월 xx일) - > YYYY/MM/DD로 return값을 수정한 후 mb단에서 '년월일 바꿈 
            arrYMD = Split(lgObjRs("long_day"),"/")  
            strYMD = arrYMD(0) & "년 " & arrYMD(1) & "월 " & arrYMD(2) & "일"
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strYMD)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ADJUST_DAY"), 0, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HONOR_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ETC_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PROV_YY_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EXACT_YY_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RETIRE_ANU"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RETIRE_INSUR"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ETC_SUB1"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ETC_SUB2"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ETC_SUB3"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ETC_SUB4"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("remark"))
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HGA040T         ("
    lgStrSQL = lgStrSQL & " emp_no          ," 
    lgStrSQL = lgStrSQL & " entr_dt         ," 
    lgStrSQL = lgStrSQL & " retire_dt       ," 
    lgStrSQL = lgStrSQL & " retire_pay_bas_dt ,"
    
    lgStrSQL = lgStrSQL & " adjust_day      ,"
    lgStrSQL = lgStrSQL & " honor_amt       ," 
    lgStrSQL = lgStrSQL & " etc_amt         ," 
    lgStrSQL = lgStrSQL & " prov_yy_mm      ," 
    lgStrSQL = lgStrSQL & " exact_yy_mm     ," 
    lgStrSQL = lgStrSQL & " retire_anu      ," 
    lgStrSQL = lgStrSQL & " retire_insur      ,"     
    lgStrSQL = lgStrSQL & " etc_sub1        ," 
    lgStrSQL = lgStrSQL & " etc_sub2        ," 
    lgStrSQL = lgStrSQL & " etc_sub3        ," 
    lgStrSQL = lgStrSQL & " etc_sub4        ," 
    lgStrSQL = lgStrSQL & " remark			,"     
    lgStrSQL = lgStrSQL & " Isrt_emp_no     ," 
    lgStrSQL = lgStrSQL & " isrt_dt         ," 
    lgStrSQL = lgStrSQL & " updt_emp_no     ," 
    lgStrSQL = lgStrSQL & " updt_dt         )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(4),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")    & ","    
    
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(6))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(7))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(8))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(9))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(10))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(11))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(12))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(13))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(14))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(15))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvNum(Trim(UCase(arrColVal(16))),0),"0","D")     & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(17)), "''", "S")     & ","   
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

Response.Write lgStrSQL
'Response.End
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HGA040T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " entr_dt          = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & " retire_dt        = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(4),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & " retire_pay_bas_dt= " & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")    & ","
        
    lgStrSQL = lgStrSQL & " adjust_day       = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(6))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " honor_amt        = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(7))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " etc_amt          = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(8))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " prov_yy_mm       = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(9))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " exact_yy_mm      = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(10))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " retire_anu       = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(11))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " retire_insur       = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(12))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " etc_sub1         = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(13))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " etc_sub2         = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(14))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " etc_sub3         = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(15))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " etc_sub4         = " & FilterVar(UniConvNum(Trim(UCase(arrColVal(16))),0),"0","D")     & ","
    lgStrSQL = lgStrSQL & " remark			 = " & FilterVar(UCase(arrColVal(17)), "''", "S")						& ","

    lgStrSQL = lgStrSQL & " updt_emp_no      = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt          = " & FilterVar(lgGetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " emp_no           = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And entr_dt      = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " And retire_dt    = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(4),NULL),"NULL","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HGA040T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " emp_no           = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And entr_dt      = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " And retire_dt    = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(4),NULL),"NULL","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
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
        
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           
           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  B_MAJOR "
                       lgStrSQL = lgStrSQL & " Where MAJOR_CD " & pComp & pCode
               Case "D"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  B_MAJOR "
                       lgStrSQL = lgStrSQL & " WHERE MAJOR_CD " & pComp & pCode
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount
                       lgStrSQL = lgStrSQL & " hga040t.emp_no,haa010t.name,hga040t.entr_dt,hga040t.retire_dt,hga040t.retire_pay_bas_dt,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_H_GetLongYYMMDD(hga040t.entr_dt,hga040t.retire_dt) long_day,"
                       lgStrSQL = lgStrSQL & " hga040t.adjust_day,hga040t.honor_amt,hga040t.etc_amt,hga040t.prov_yy_mm,hga040t.exact_yy_mm,hga040t.retire_anu, hga040t.retire_insur,"
                       lgStrSQL = lgStrSQL & " hga040t.etc_sub1,hga040t.etc_sub2,hga040t.etc_sub3,hga040t.etc_sub4, hga040t.remark,"
                       lgStrSQL = lgStrSQL & " hga040t.isrt_emp_no,hga040t.isrt_dt,hga040t.updt_emp_no,hga040t.updt_dt "
                       lgStrSQL = lgStrSQL & " From  hga040t, haa010t "
                       lgStrSQL = lgStrSQL & " Where hga040t.emp_no " & pComp & " "  & pCode
               Case "U"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  B_MAJOR "
                       lgStrSQL = lgStrSQL & " Where MAJOR_CD " & pComp & pCode
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
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
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
