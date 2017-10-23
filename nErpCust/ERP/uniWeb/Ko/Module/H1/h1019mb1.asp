<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
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
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
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
	Dim iDx
    Dim txtKey
    Dim iLcNo
    Dim strEmp_no
    Dim strname

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("SR","",C_EQ)                                       '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'If data not exists
    Else 
        %>
          <Script Language=vbscript>
                With Parent.Frm1
                     .txtFlgMode1.Value = "I"
                     .txtFlgMode2.Value = "I"
                     .txtFlgMode3.Value = "I"
                     .txtFlgMode4.Value = "I"
                End With 
          </Script>       
        <%     

        Do While Not lgObjRs.EOF
           Select Case ConvSPChars(lgObjRs("pay_type"))
              Case  "!"                                          '   급여 마감             
                  %>
                    <Script Language=vbscript>
                         With Parent.Frm1  
                              .txtclose_type1.Value = "<%=ConvSPChars(lgObjRs("close_type"))%>"
                              .txtclose_dt1.Year  = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",1,4)
                              .txtclose_dt1.Month = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",6,2)
                              .txtFlgMode1.Value = "U"
                         End With          
                    </Script>       
                  <%     
              Case  "@"                                          '   연월차시스템 마감 
                  %>
                  <Script Language=vbscript>
                         With Parent.Frm1
                              .txtclose_type2.Value = "<%=ConvSPChars(lgObjRs("close_type"))%>"
                              .txtclose_dt2.Year  = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",1,4)
                              .txtclose_dt2.Month = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",6,2)
                              .txtFlgMode2.Value = "U"
                         End With
                  </Script>       
                  <%
              Case  "*"                                          '   연말정산 마감             
                  %>
                    <Script Language=vbscript>
                         With Parent.Frm1  
                              .txtclose_type3.Value = "<%=ConvSPChars(lgObjRs("close_type"))%>"
                              .txtclose_dt3.Year    = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",1,4)
                              .txtFlgMode3.Value    = "U"
                         End With          
                    </Script>       
                  <%     
              Case  "#"                                          '   근태시스템 마감 
                  %>
                  <Script Language=vbscript>
                         With Parent.Frm1
                              .txtclose_type4.Value = "<%=ConvSPChars(lgObjRs("close_type"))%>"
                              .txtclose_dt4.Year    = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",1,4)
                              .txtclose_dt4.Month   = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",6,2)
                              .txtclose_dt4.Day     = Mid("<%=ConvSPChars(lgObjRs("close_dt"))%>",9,2)        
                              .txtFlgMode4.Value    = "U"
                         End With 
                  </Script>       
                  <%
           End Select
           lgObjRs.MoveNext
        Loop 
    End If
    Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
    Call SubBizQueryMulti()

End Sub	
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave() 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If Request("txtFlgMode1") = "U" Then
        Call SubBizSaveSingleUpdate(Request("txtclose_type1"), Request("txtclose_dt1"), "!")
    Else
        Call SubBizSaveSingleCreate(Request("txtclose_type1"), Request("txtclose_dt1"), "!")
    End If

    If Request("txtFlgMode2") = "U" Then
        Call SubBizSaveSingleUpdate(Request("txtclose_type2"),Request("txtclose_dt2"), "@")
    Else
        Call SubBizSaveSingleCreate(Request("txtclose_type2"), Request("txtclose_dt2"), "@")
    End If

    If Request("txtFlgMode3") = "U" Then
        Call SubBizSaveSingleUpdate(Request("txtclose_type3"), Request("txtclose_dt3"), "*")
    Else
        Call SubBizSaveSingleCreate(Request("txtclose_type3"), Request("txtclose_dt3"), "*")
    End If

    If Request("txtFlgMode4") = "U" Then
        Call SubBizSaveSingleUpdate(Request("txtclose_type4"), Request("txtclose_dt4"), "#")
    Else
        Call SubBizSaveSingleCreate(Request("txtclose_type4"), Request("txtclose_dt4"), "#")
    End If

End Sub	

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR","X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
         Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)    
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_type_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("close_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("close_type_nm"))
            lgstrData = lgstrData & Chr(11) & UniMonthClientFormat(lgObjRs("close_dt"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
            
        Loop 
	Else 
       lgStrPrevKey = ""
    End If
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If     
    
    Call SubHandleError("MR",lgObjRs,Err)
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
Sub SubBizSaveSingleCreate(iclose_type, iclose_dt, ipay_type)
    Dim strYear,strMonth,strDay
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If ipay_type = "#" Then         '근태마감일이면 일까지 포함 
		iclose_dt = UNIConvDateCompanyToDB(iclose_dt,NULL)
    ElseIf ipay_type = "*" Then         '연말정산마감이면 년도만 포함 
        iclose_dt = UniConvYYYYMMDDToDate(gServerDateFormat,iclose_dt,"12","31")
	Else
        Call ExtractDateFrom(iclose_dt,gDateFormatYYYYMM,gComDateType,strYear,strMonth,strDay)
        iclose_dt = UniConvYYYYMMDDToDate(gServerDateFormat,strYear,strMonth,"01")
    End if

    lgStrSQL = "INSERT INTO HDA270T( org_cd, pay_gubun, pay_type," 
    lgStrSQL = lgStrSQL & " close_type, close_dt, emp_no, input_dt)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & "" & FilterVar("Z", "''", "S") & " ,"
    lgStrSQL = lgStrSQL & FilterVar(ipay_type, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(iclose_type, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(iclose_dt, "NULL", "S") & ","    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : 시스템마감입력 
'============================================================================================================
Sub SubBizSaveSingleUpdate(iclose_type, iclose_dt, ipay_type)
    
    Dim strYear,strMonth,strDay    
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If ipay_type = "#" Then         '근태마감일이면 일까지 포함 
		iclose_dt = UNIConvDateCompanyToDB(iclose_dt,NULL)
    
	
    ElseIf ipay_type = "*" Then         '연말정산마감이면 년도만 포함 
        iclose_dt = UniConvYYYYMMDDToDate(gServerDateFormat,iclose_dt,"12","31")
    
	Else
        Call ExtractDateFrom(iclose_dt,gDateFormatYYYYMM,gComDateType,strYear,strMonth,strDay)    
        iclose_dt = UniConvYYYYMMDDToDate(gServerDateFormat,strYear,strMonth,"01")
    
    End if
	
    lgStrSQL = "           UPDATE HDA270T"
    lgStrSQL = lgStrSQL & "   SET close_type = " & FilterVar(iclose_type, "''", "S") & ","
    lgStrSQL = lgStrSQL & "       close_dt = " & FilterVar(iclose_dt, "NULL", "S")
    lgStrSQL = lgStrSQL & " WHERE org_cd = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND pay_gubun = " & FilterVar("Z", "''", "S") & " "
    lgStrSQL = lgStrSQL & "   AND pay_type = " & FilterVar(ipay_type, "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    Dim iclose_dt
    Dim strYear,strMonth,strDay
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call ExtractDateFrom(arrColVal(5),gDateFormatYYYYMM,gComDateType,strYear,strMonth,strDay)
    iclose_dt = UniConvYYYYMMDDToDate(gServerDateFormat,strYear,strMonth,"01")

    lgStrSQL = "INSERT INTO HDA270T( org_cd, pay_gubun, pay_type," 
    lgStrSQL = lgStrSQL & " close_type, close_dt, emp_no, input_dt)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & "" & FilterVar("Z", "''", "S") & " ,"
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(iclose_dt, "NULL", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    Dim iclose_dt
    Dim strYear,strMonth,strDay
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call ExtractDateFrom(arrColVal(5),gDateFormatYYYYMM,gComDateType,strYear,strMonth,strDay)
    iclose_dt = UniConvYYYYMMDDToDate(gServerDateFormat,strYear,strMonth,"01")

    lgStrSQL = "UPDATE  HDA270T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & "      close_type = " & FilterVar(arrColVal(4), "''", "S") & ","
    lgStrSQL = lgStrSQL & "      close_dt = " & FilterVar(iclose_dt, "NULL", "S")
    lgStrSQL = lgStrSQL & " WHERE org_cd = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND pay_gubun = " & FilterVar("Z", "''", "S") & " "
    lgStrSQL = lgStrSQL & "   AND pay_type = " & FilterVar(arrColVal(3), "''", "S")

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

    lgStrSQL = "DELETE  HDA270T"
    lgStrSQL = lgStrSQL & " WHERE org_cd = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND pay_gubun = " & FilterVar("Z", "''", "S") & " "
    lgStrSQL = lgStrSQL & "   AND pay_type = " & FilterVar(arrColVal(3), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select pay_type as pay_type, close_type as close_type, CONVERT(CHAR(21),close_dt, 20) as close_dt" 
                       lgStrSQL = lgStrSQL & " FROM  hda270t "
                       lgStrSQL = lgStrSQL & " WHERE org_cd = " & FilterVar("1", "''", "S") & "  and pay_gubun = " & FilterVar("Z", "''", "S") & "  and pay_type in (" & FilterVar("!", "''", "S") & " ," & FilterVar("@", "''", "S") & " ," & FilterVar("*", "''", "S") & " ," & FilterVar("#", "''", "S") & " ) "
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "            Select "
                       lgStrSQL = lgStrSQL & "			pay_type, "
                       lgStrSQL = lgStrSQL & "          dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ",pay_type) pay_type_nm, "
                       lgStrSQL = lgStrSQL & "			close_type, "
                       lgStrSQL = lgStrSQL & "          dbo.ufn_GetCodeName(" & FilterVar("H0104", "''", "S") & ",close_type) close_type_nm, "
                       lgStrSQL = lgStrSQL & "			close_dt "                       
                       lgStrSQL = lgStrSQL & " FROM  HDA270T "
                       lgStrSQL = lgStrSQL & " WHERE org_cd = " & FilterVar("1", "''", "S") & "  and pay_gubun = " & FilterVar("Z", "''", "S") & "  and pay_type not in (" & FilterVar("!", "''", "S") & " ," & FilterVar("@", "''", "S") & " ," & FilterVar("*", "''", "S") & " ," & FilterVar("#", "''", "S") & " ) "
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then              
	         With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
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
