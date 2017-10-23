<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
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
    Dim strRoll_pstn
    Dim strPay_grd1
    dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	dim strNat_cd
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey1 = iKey1 & " and EMP_NO = " & FilterVar(lgKeyStream(1), "''", "S")

    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
  
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1

        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL"))
            lgstrData = lgstrData & Chr(11) & ""
         
			call CommonQueryRs(" nat_cd "," HAA010T "," EMP_NO = " &  FilterVar(lgKeyStream(0), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
			strNat_cd = Replace(lgF0, Chr(11), "")  ' 주민번호 check를 위해서 
			if strNat_cd="KR" then            
				lgstrData = lgstrData & Chr(11) & ConvSPChars(Mid(lgObjRs("res_no"), 1, 6) & "-" & Mid(lgObjRs("res_no"), 7, 7))
			else 
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_RES_NO"))
			end if
			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAT_FLAG"))
			lgstrData = lgstrData & Chr(11) & ""
			
            If ConvSPChars(lgObjRs("BASE_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If

            If ConvSPChars(lgObjRs("PARIA_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If

            If ConvSPChars(lgObjRs("CHILD_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If

            If ConvSPChars(lgObjRs("INSUR_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If
            
            If ConvSPChars(lgObjRs("MEDI_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If

            If ConvSPChars(lgObjRs("EDU_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If
            
            If ConvSPChars(lgObjRs("CARD_YN")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If
                                                        
            lgstrData = lgstrData & Chr(11) & C_SHEETMAXROWS_D + iDx
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
	
    For iDx = 1 To C_SHEETMAXROWS_D
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

    lgStrSQL = "INSERT INTO HFA150T(YEAR_YY, EMP_NO, FAMILY_NAME, FAMILY_REL, FAMILY_RES_NO,NAT_FLAG, "
    lgStrSQL = lgStrSQL & " BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT) "

    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")			& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")		& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(12)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(13)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(14)), "''", "S")     & ","
    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgStrSQL = "UPDATE  HFA150T"
    lgStrSQL = lgStrSQL & " SET " 
'    lgStrSQL = lgStrSQL & " FAMILY_NAME		= " & FilterVar(arrColVal(4), "''", "S") & ","
 '   lgStrSQL = lgStrSQL & " FAMILY_REL		= " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " BASE_YN			= " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " PARIA_YN		= " & FilterVar(UCase(arrColVal(8)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " CHILD_YN		= " & FilterVar(UCase(arrColVal(9)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " INSUR_YN		= " & FilterVar(UCase(arrColVal(10)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " MEDI_YN			= " & FilterVar(UCase(arrColVal(11)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " EDU_YN			= " & FilterVar(UCase(arrColVal(12)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " CARD_YN			= " & FilterVar(UCase(arrColVal(13)), "''", "S") & ","

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO		= " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " UPDT_DT			= " & FilterVar(lgSvrDateTime, "''", "S")  

    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         EMP_NO		= " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   YEAR_YY		= " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   FAMILY_RES_NO = " & FilterVar(UCase(arrColVal(6)), "''", "S")
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HFA150T"
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         EMP_NO		= " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   YEAR_YY		= " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   FAMILY_RES_NO = " & FilterVar(UCase(arrColVal(4)), "''", "S")

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
                       lgStrSQL = "SELECT TOP " & iSelCount
                       lgStrSQL = lgStrSQL & "  FAMILY_NAME, FAMILY_REL, FAMILY_RES_NO, BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN, NAT_FLAG"
                       lgStrSQL = lgStrSQL & " FROM  HFA150T "
                       lgStrSQL = lgStrSQL & " WHERE YEAR_YY =" & pCode 
                       lgStrSQL = lgStrSQL & " ORDER BY FAMILY_REL, FAMILY_RES_NO ASC"
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
