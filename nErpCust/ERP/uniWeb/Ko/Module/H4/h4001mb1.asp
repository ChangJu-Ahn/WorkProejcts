<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
                                                                    '☜: Clear Error status
    call LoadBasisGlobalInf()
    lgSvrDateTime = GetSvrDateTime
    
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim iKey2
    Dim iFromDate
    Dim strYear, strMonth, strDay
    Dim lastMonthLastDate, tempYear, tempMonth, tempDay 
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iFromDate = UniConvYYYYMMDDToDate(gDateFormat, lgKeyStream(1), lgKeyStream(2), "01") 
	iFromTo   = UNIGetLastDay(iFromDate,gDateFormat)
    Call ExtractDateFrom(iFromTo, gDateFormat, gComDateType, strYear, strMonth, strDay)
    '2004.10.11 : 지난달 마지막날 가져오는 로직 추가 
    lastMonthLastDate =UniConvYYYYMMDDToDate(gDateFormat, lgKeyStream(1), lgKeyStream(2)-1, "01")
	lastMonthLastDate   = UNIGetLastDay(lastMonthLastDate,gDateFormat)    
    Call ExtractDateFrom(lastMonthLastDate, gDateFormat, gComDateType, tempYear, tempMonth, tempDay)    
    LastOfMonth = tempDay

    iKey1     = FilterVar(lgKeyStream(0),"''", "S")
    iFromDate = FilterVar(UniConvDateCompanyToDB(iFromDate, gDateFormat)     ,"''", "S")
    iFromTo   = FilterVar(UniConvDateCompanyToDB(iFromTo, gDateFormat)       ,"''", "S")
    iKey2     = FilterVar(lgKeyStream(3),"''", "S")

    Call SubMakeSQLStatements("R",iKey1,iFromDate,iFromTo,iKey2)                                '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
       Dim iWeek 
       Dim iX
       Dim iOther
       Dim iY
       Dim iClassName
       Dim iDay
       Dim IDecrease
       
       iX = CInt(lgObjRs("WEEK_DAY")) - 1
       
%>
<Script Language=vbscript>

	For CalCol = 0 To <%=(iX-1)%>  
		Parent.frm1.txtDate(CalCol).value     = CStr(<%=(LastOfMonth - iX)%> + CalCol + 1)
		Parent.frm1.txtDate(CalCol).className = "DummyDay"
		Parent.frm1.txtDate(CalCol).disabled  = True

		Parent.frm1.txtHoli(CalCol).value     = ""
		Parent.frm1.txtHoli(CalCol).disabled  = True

		Parent.frm1.txtDesc(CalCol).value     = ""
		Parent.frm1.txtDesc(CalCol).disabled  = True
		Parent.frm1.txtDesc(CalCol).title     = ""
	Next
       
</Script>       
<% 
              
       Do While Not lgObjRs.EOF
       
       iDay = UniConvDateDBToCompany(lgObjRs("Date"),"")
       Call ExtractDateFrom(iDay, gDateFormat, gComDateType, strYear, strMonth, strDay)
       iDay = Right("0" & strDay, 2)       
	   
       If ConvSPChars(lgObjRs("HOLI_TYPE")) = "H" Then
          iClassName = "red"
       ElseIf (iX + 1) Mod 7   = 0 Then       
          iClassName = "blue"
       Else   
          iClassName = "black"
       End If
              
%>
<Script Language=vbscript>
       
	     Parent.frm1.txtDate(CInt(<%=iX%>)).value       = "<%=iDay%>"
     	 Parent.frm1.txtDesc(CInt(<%=iX%>)).alt         = "<%=iDay%>" & "일의 사유"
	     Parent.frm1.txtDate(CInt(<%=iX%>)).className   = "Day"
	     Parent.frm1.txtDate(CInt(<%=iX%>)).disabled    = False
	     Parent.frm1.txtDate(CInt(<%=iX%>)).style.color = "<%=iClassName%>"
	
	     Parent.frm1.txtHoli(CInt(<%=iX%>)).value       = "<%=ConvSPChars(lgObjRs("HOLI_TYPE"))%>"
	     Parent.frm1.txtHoli(CInt(<%=iX%>)).disabled    = False
	
	     Parent.frm1.txtDesc(CInt(<%=iX%>)).value       = "<%=ConvSPChars(lgObjRs("remark"))%>"
	     Parent.frm1.txtDesc(CInt(<%=iX%>)).disabled    = False
	     Parent.frm1.txtDesc(CInt(<%=iX%>)).title       = "<%=ConvSPChars(lgObjRs("remark"))%>"	

</Script>       
<% 
      iX = iX + 1
           lgObjRs.MoveNext 
       Loop 

       iOther = iX
%>
<Script Language=vbscript>
       
	For CalCol = <%=iX%> to 41
		Parent.frm1.txtDate(CalCol).value     = CStr(CalCol - <%=iX%> + 1)
		Parent.frm1.txtDate(CalCol).className = "DummyDay"
		Parent.frm1.txtDate(CalCol).disabled  = True

		Parent.frm1.txtHoli(CalCol).value     = ""
		Parent.frm1.txtHoli(CalCol).disabled  = True

		Parent.frm1.txtDesc(CalCol).value     = ""
		Parent.frm1.txtDesc(CalCol).disabled  = True
		Parent.frm1.txtDesc(CalCol).title     = ""
	Next
       
</Script>       
<% 
       
      
    End If
    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update              
              Call SubBizSaveSingleUpdate()
    End Select

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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate(lgObjRs)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
	Dim iBaseYYYYMM
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    iBaseYYYYMM = lgKeyStream(1) & gServerDateType & lgKeyStream(2) & gServerDateType
    
    For i = 1 To Request("txtHoli").count

        lgStrSQL = "UPDATE  HCA020T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " HOLI_TYPE =   " & FilterVar(Request("txtHoli")(i)      , "''", "S") & ","
        lgStrSQL = lgStrSQL & " REMARK =      " & FilterVar(Request("txtDesc")(i)      , "''", "S") & ","
        lgStrSQL = lgStrSQL & " UPDT_DT  =    " & FilterVar(lgSvrDateTime,NULL,"S") & ","
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId                     , "''", "S")  
        lgStrSQL = lgStrSQL & " WHERE         " 
        lgStrSQL = lgStrSQL & " ORG_CD   =    " & FilterVar(lgKeyStream(0), "''", "S") & " AND "
        lgStrSQL = lgStrSQL & " WK_TYPE =     " & FilterVar(lgKeyStream(3), "''", "S") & " AND "
        lgStrSQL = lgStrSQL & " DATE =        " & FilterVar(iBaseYYYYMM & Request("txtDate")(i),NULL,"S")
  
        lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    Next
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements("R",iKey1,iFromDate,iFromTo,iKey2)
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode,pCode1,pCode2,pCode3)

    Select Case pMode 
      Case "R"
             lgStrSQL = "Select Date,WEEK_DAY,HOLI_TYPE,remark" 
             lgStrSQL = lgStrSQL & " From  HCA020T "
             lgStrSQL = lgStrSQL & " WHERE ORG_CD = " & pCode 
             lgStrSQL = lgStrSQL & "   AND WK_TYPE = " & pCode3
             lgStrSQL = lgStrSQL & "   AND DATE Between " & pCode1 & " AND " & pCode2 
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
             Parent.DBQueryOk        
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
