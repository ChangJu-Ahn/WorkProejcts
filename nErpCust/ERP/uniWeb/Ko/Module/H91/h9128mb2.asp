<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    call LoadBasisGlobalInf()
    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection
 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim lgStrSQL
    Dim strYear

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	call CommonQueryRs(" isnull(max(YEAR_YY),0) "," HFA150T "," YEAR_YY <" &  FilterVar(lgKeyStream(1), "''", "S") & " and EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strYear = Replace(lgF0, Chr(11), "") 
	
	'If strYear > 0 Then ¸·À½..
	IF 1=2 then 
 
		lgStrSQL = "SELECT		FAMILY_NAME, FAMILY_REL, FAMILY_RES_NO, BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN,NAT_FLAG"
		lgStrSQL = lgStrSQL & " FROM  HFA150T "
		lgStrSQL = lgStrSQL & " WHERE YEAR_YY =" &  FilterVar(strYear, "''", "S") & " and EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
		lgStrSQL = lgStrSQL & " ORDER BY FAMILY_REL, FAMILY_RES_NO ASC"
    		
	Else
	
		lgStrSQL = "SELECT FAMILY_NAME,FAMILY_RES_NO,FAMILY_REL, BASE_YN, PARIA_YN,CHILD_YN, INSUR_YN,MEDI_YN, EDU_YN, CARD_YN, NAT_FLAG "
		lgStrSQL = lgStrSQL & " FROM (		"
		lgStrSQL = lgStrSQL & "			SELECT	family_nm FAMILY_NAME,  res_no FAMILY_RES_NO,  AA.REFERENCE FAMILY_REL "
		lgStrSQL = lgStrSQL & "				,'N' BASE_YN,CASE WHEN SUPP_CD= '3' THEN 'Y' ELSE 'N' END PARIA_YN,'N' CHILD_YN,'N' INSUR_YN,'N' MEDI_YN,'N' EDU_YN,'N' CARD_YN,'1' NAT_FLAG "
		lgStrSQL = lgStrSQL & "			FROM  HAA020T  H LEFT JOIN "
		lgStrSQL = lgStrSQL & " 			( SELECT A.MINOR_CD ,A.MINOR_NM,B.REFERENCE "
		lgStrSQL = lgStrSQL & " 			 FROM  B_MINOR A  JOIN B_CONFIGURATION B  ON  A.MAJOR_CD=B.MAJOR_CD AND A.MINOR_CD=B.MINOR_CD "
		lgStrSQL = lgStrSQL & " 			 WHERE A.MAJOR_CD='H0023' ) AA 	ON  H.rel_cd = AA.MINOR_CD  "
		lgStrSQL = lgStrSQL & "			WHERE emp_no =" & FilterVar(lgKeyStream(0), "''", "S")  
		lgStrSQL = lgStrSQL & "		union all "
		lgStrSQL = lgStrSQL & "			SELECT name FAMILY_NAME, res_no FAMILY_RES_NO, '0' FAMILY_REL,'N' BASE_YN, 'N' PARIA_YN,'N' CHILD_YN,'N' INSUR_YN,'N' MEDI_YN,'N' EDU_YN,'N' CARD_YN,'1' NAT_FLAG "
		lgStrSQL = lgStrSQL & "			FROM haa010t "
		lgStrSQL = lgStrSQL & "			WHERE emp_no =" & FilterVar(lgKeyStream(0), "''", "S")  & " and emp_no  not in ( select emp_no from HAA020T where rel_cd='00' and emp_no =" & FilterVar(lgKeyStream(0), "''", "S") & ")"
		lgStrSQL = lgStrSQL & "		) T "
		lgStrSQL = lgStrSQL & " ORDER BY FAMILY_REL, FAMILY_RES_NO ASC"
	End If
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL"))
            lgstrData = lgstrData & Chr(11) & ""        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_RES_NO"))

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
                        
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
             
        Loop 
        
		lgStrSQL = " DELETE HFA150T "
		lgStrSQL = lgStrSQL & " WHERE emp_no =" & FilterVar(lgKeyStream(0), "''", "S") & " and year_yy = " &  FilterVar(lgKeyStream(1), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
        
    End If
    
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBAutoQueryOk
	         End with
          End If   
    End Select    
    
       
</Script>	
