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
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    'On Error Resume Next
    Err.Clear

	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE,sBS_PL_FG
	Dim lgStrPrevKey

	' -- 그리드 컬럼 정의 
	Dim C_ACCT_CD
	Dim C_ACCT_BT
	Dim C_ACCT_NM
	Dim C_BS_PL_FG
	Dim C_BS_PL_NM
	Dim C_GP_CD
	Dim C_GP_BT
	Dim C_GP_NM

	lgErrorStatus   = "NO"
    lgErrorPos      = ""                                                           '☜: Set to space
    lgOpModeCRUD    = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")
    sBS_PL_FG  =   Request("cboBS_PL_FG") 
    

    lgLngMaxRow     = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 버전관리기능 추가 
	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
    
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_ACCT_CD	= 1
	C_ACCT_BT	= 2
	C_ACCT_NM	= 3
	C_BS_PL_FG	= 4
	C_BS_PL_NM	= 5
	C_GP_CD		= 6
	C_GP_BT		= 7
	C_GP_NM		= 8
End Sub


'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear


    lgStrSQL = "DELETE  TB_ACCT_MAPPING WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Trim(UCase(wgCO_CD)),"''","S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR"),"''","S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE"),"''","S") & vbCrLf
    
	PrintLog "SubBizDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Dim iCnt
    Dim iKey1, iKey2, iKey3, iIntMaxRows, iLngRow
    Dim iLoopMax
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        iIntMaxRows = C_SHEETMAXROWS_D * lgStrPrevKey

        lgstrData = ""
        iCnt = 1

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BS_PL_FG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BS_PL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GP_CD"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GP_NM"))
			lgstrData = lgstrData & Chr(11) & iIntMaxRows + iCnt

            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
			iCnt = iCnt + 1
            'If iCnt > C_SHEETMAXROWS_D Then
            '   lgStrPrevKey = lgStrPrevKey + 1
            '   Exit Do
            'End If   
        Loop 
    End If
    
    'If iCnt <= C_SHEETMAXROWS_D Then
    '   lgStrPrevKey = ""
    'End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1,pCode2,pCode3)

    Select Case pMode 
      Case "R"

			lgStrSQL =			  " SELECT A.ACCT_CD, B.ACCT_NM, A.BS_PL_FG, dbo.ufn_GetCodeName('W1081', A.BS_PL_FG) AS BS_PL_NM, " & vbCrLf
			lgStrSQL = lgStrSQL & "  	 A.GP_CD, C.GP_NM " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_ACCT_MAPPING A (NOLOCK)  " & vbCrLf
            lgStrSQL = lgStrSQL & " 	LEFT OUTER JOIN TB_WORK_6 B (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.ACCT_CD=B.ACCT_CD  " & vbCrLf
            lgStrSQL = lgStrSQL & " 	INNER JOIN dbo.ufn_TB_ACCT_GP('200703')  C  ON A.BS_PL_FG=C.BS_PL_FG AND A.GP_CD=C.GP_CD  " & vbCrLf
     
            lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf
            lgStrSQL = lgStrSQL & "  AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
            lgStrSQL = lgStrSQL & "  AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            IF sBS_PL_FG <> "" Then
				lgStrSQL = lgStrSQL & "   AND A.BS_PL_FG = " & sBS_PL_FG 	 & vbCrLf
			End If	
            

            lgStrSQL = lgStrSQL & "  ORDER BY A.ACCT_CD, A.GP_CD " & vbCrLf

    End Select
	'PrintLog "SubMakeSQLStatements: " & lgStrSQL
	

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

	PrintLog "SubBizSaveMulti.."
	
    'On Error Resume Next
    Err.Clear 
    
	PrintLog "1번째 그리드. .: " & Request("txtSpread") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
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
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_ACCT_MAPPING WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, ACCT_CD, BS_PL_FG,GP_CD "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_ACCT_CD))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_BS_PL_FG))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_GP_CD))),"''","S")	& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_ACCT_MAPPING WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " SET " 
    
	lgStrSQL = lgStrSQL & " BS_PL_FG	= " &  FilterVar(Trim(UCase(arrColVal(C_BS_PL_FG))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " GP_CD		= " &  FilterVar(Trim(UCase(arrColVal(C_GP_CD))),"''","S")	& ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND ACCT_CD = " & FilterVar(Trim(UCase(arrColVal(C_ACCT_CD))),"''","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_ACCT_MAPPING WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND ACCT_CD = " & FilterVar(Trim(UCase(arrColVal(C_ACCT_CD))),"''","S")
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
    lgErrorStatus     = "YES"
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
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "SD"
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .DBQueryOk
	         End with
	      Else
	      	Parent.FncNew
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