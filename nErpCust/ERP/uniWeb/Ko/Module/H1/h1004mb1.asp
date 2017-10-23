<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             lgCurrentSpd = lgKeyStream(2)
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQGT)                                 '☆: Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
        
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""

        iDx = 1
        Do While Not lgObjRs.EOF
			If lgCurrentSpd = "M" Then    
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD_NM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_KIND"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_KIND_nm"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_TYPE"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_TYPE_nm"))
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("LIMIT_AMT"), ggAmtOfMoney.DecPoint, 0)

				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALCU_TYPE"))

				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_MM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_MM_nm"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_DD"))
				lgstrData = lgstrData & Chr(11) & "~"
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_MM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_MM_nm"))
            
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_DD"))

				lgstrData = lgstrData & Chr(11) & "수당/"

				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_CALCU"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_CALCU_nm"))
				lgstrData = lgstrData & Chr(11) & "*"
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALCU_BAS_DD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALCU_BAS_DD_nm"))
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ALLOW_SEQ"), 0, 0)

				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
			else
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_TYPE"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_TYPE_nm"))

				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_MM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_MM_nm"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_DD"))
				lgstrData = lgstrData & Chr(11) & "~"
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_MM"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_MM_nm"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_DD"))

				lgstrData = lgstrData & Chr(11) & "수당/"

				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_CALCU"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_CALCU_nm"))
				lgstrData = lgstrData & Chr(11) & "*"
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALCU_BAS_DD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALCU_BAS_DD_nm"))
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ALLOW_SEQ"), 2, 0)

				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
			
			end if
		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
        Loop 
    else
            lgStrPrevKey = ""
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
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgCurrentSpd = Trim(UCase(arrColVal(2)))
    if lgCurrentSpd = "M" then
        lgStrSQL = "INSERT INTO HDA010T( CODE_TYPE, PAY_CD," 
        lgStrSQL = lgStrSQL & " ALLOW_CD, ALLOW_NM, ALLOW_KIND, TAX_TYPE," 
        lgStrSQL = lgStrSQL & " LIMIT_AMT, CALCU_TYPE, CRT_STRT_MM, CRT_STRT_DD," 
		lgStrSQL = lgStrSQL & " CRT_END_MM, CRT_END_DD, DAY_CALCU, CALCU_BAS_DD,"  
		lgStrSQL = lgStrSQL & " ALLOW_SEQ, ISRT_EMP_NO , ISRT_DT , UPDT_EMP_NO , UPDT_DT )" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        
		lgStrSQL = lgStrSQL & " " & FilterVar("1", "''", "S") & " ,"
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
        
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8), 0)  & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S") & ","
        
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0) & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S") & ","
        
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0) & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(13)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(14)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(15)), "''", "S") & ","

        lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(16), 0)					 & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                       & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")			 & "," 
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                       & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
        lgStrSQL = lgStrSQL & " )"
    end if

    lgObjConn.Execute lgStrSQL,,adCmdText    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgCurrentSpd = Trim(UCase(arrColVal(2)))

    if lgCurrentSpd = "M" then
        lgStrSQL = "UPDATE  HDA010T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " ALLOW_NM = " & FilterVar(arrColVal(5), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " ALLOW_KIND = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","

        lgStrSQL = lgStrSQL & " TAX_TYPE = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " LIMIT_AMT = " & UNIConvNum(arrColVal(8), 0)  & ","
        lgStrSQL = lgStrSQL & " CALCU_TYPE = " & FilterVar(UCase(arrColVal(9)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " CRT_STRT_MM = " & UNIConvNum(arrColVal(10),0)   & ","
        lgStrSQL = lgStrSQL & " CRT_STRT_DD = " & FilterVar(UCase(arrColVal(11)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " CRT_END_MM = " & UNIConvNum(arrColVal(12),0)   & ","
        lgStrSQL = lgStrSQL & " CRT_END_DD = " & FilterVar(UCase(arrColVal(13)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " DAY_CALCU = " & FilterVar(UCase(arrColVal(14)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " CALCU_BAS_DD = " & FilterVar(UCase(arrColVal(15)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " ALLOW_SEQ = " & UNIConvNum(arrColVal(16), 0) & ","
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
        lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime, "''", "S")
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & "     CODE_TYPE = " & FilterVar("1", "''", "S") & " "
        lgStrSQL = lgStrSQL & " AND PAY_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgStrSQL = lgStrSQL & " AND ALLOW_CD = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    else    
        lgStrSQL = "UPDATE  HDA010T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " ALLOW_NM = " & FilterVar(arrColVal(4), "''", "S")   & ","

        lgStrSQL = lgStrSQL & " TAX_TYPE = " & FilterVar(UCase(arrColVal(5)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " CRT_STRT_MM = " & UNIConvNum(arrColVal(6),0)   & ","
        lgStrSQL = lgStrSQL & " CRT_STRT_DD = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " CRT_END_MM = " & UNIConvNum(arrColVal(8),0)   & ","
        lgStrSQL = lgStrSQL & " CRT_END_DD = " & FilterVar(UCase(arrColVal(9)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " DAY_CALCU = " & FilterVar(UCase(arrColVal(10)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " CALCU_BAS_DD = " & FilterVar(UCase(arrColVal(11)), "''", "S")   & ","
        lgStrSQL = lgStrSQL & " ALLOW_SEQ = " & UNIConvNum(arrColVal(12), 0) & ","
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
        lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime, "''", "S")
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & "     CODE_TYPE = " & FilterVar("0", "''", "S") & " "
        lgStrSQL = lgStrSQL & " AND ALLOW_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    end if    
	
    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgCurrentSpd = Trim(UCase(arrColVal(2)))
    lgStrSQL = "DELETE  HDA010T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     CODE_TYPE = " & FilterVar("1", "''", "S") & " "
    lgStrSQL = lgStrSQL & " AND PAY_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Mid(pDataType,1,1)
        Case "M"
			If lgCurrentSpd = "M" Then
				pCode = " CODE_TYPE = " & FilterVar("1", "''", "S") & "  AND ALLOW_CD >= " & FilterVar(lgKeyStream(0), "''", "S") & " AND PAY_CD =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
			else
				pCode = " CODE_TYPE = " & FilterVar("0", "''", "S") & "  AND ALLOW_CD >= " & FilterVar(lgKeyStream(0), "''", "S") 
			end if
			
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"            
						lgStrSQL = "SELECT top " & iSelCount & " PAY_CD,  H_Pay_cd.Minor_nm Pay_cd_NM, "
						lgStrSQL = lgStrSQL & " ALLOW_CD,  ALLOW_NM,  ALLOW_KIND, "
						lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0087", "''", "S") & ",ALLOW_KIND) ALLOW_KIND_nm, "
						lgStrSQL = lgStrSQL & " TAX_TYPE, "
						lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0039", "''", "S") & ",TAX_TYPE) TAX_TYPE_nm, "
						lgStrSQL = lgStrSQL & " LIMIT_AMT,  CALCU_TYPE,  CRT_STRT_MM, "
						lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0088", "''", "S") & ",CRT_STRT_MM) CRT_STRT_MM_nm, "
						lgStrSQL = lgStrSQL & " CRT_STRT_DD,  CRT_END_MM, "
						lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0088", "''", "S") & ",CRT_END_MM) CRT_END_MM_nm, "
						lgStrSQL = lgStrSQL & " CRT_END_DD,  DAY_CALCU, "
						lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0089", "''", "S") & ",DAY_CALCU) DAY_CALCU_nm, "
						lgStrSQL = lgStrSQL & " CALCU_BAS_DD, "
						lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0090", "''", "S") & ",CALCU_BAS_DD) CALCU_BAS_DD_nm, "
						lgStrSQL = lgStrSQL & " ALLOW_SEQ  "
						lgStrSQL = lgStrSQL & " FROM  HDA010T, H_pay_cd "
						lgStrSQL = lgStrSQL & " WHERE " & pCode
						lgStrSQL = lgStrSQL & " And H_Pay_cd.Minor_cd = " & FilterVar("*", "''", "S") & "  "
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
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
				if "<%=lgCurrentSpd%>" = "M" then
					.ggoSpread.Source     = .frm1.vspdData
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
				else
					.ggoSpread.Source     = .frm1.vspdData1
					.lgStrPrevKey1    = "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
				end if
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
