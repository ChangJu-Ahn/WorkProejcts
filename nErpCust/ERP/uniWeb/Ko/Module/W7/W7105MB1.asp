<%@ Transaction=required  CODEPAGE=949 Language=VBScript%>
<%Option Explicit%> 
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
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	' -- 그리드 컬럼 정의 
	Dim C_H1
	Dim C_H2
	Dim C_W_CD
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	DIm C_W_DESC

	lgErrorStatus   = "NO"
    lgOpModeCRUD    = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    lgPrevNext      = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow     = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    
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
Sub InitSpreadPosVariables()

	C_H1		= 1
	C_H2		= 2
	C_W1		= 3
	C_W_CD		= 4
	C_W2		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	C_W_DESC	= 9
	
End Sub

'========================================================================================
Sub SubBizQuery()
    'On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
    'On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_50A WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        iStrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
        
			iStrData = iStrData & Chr(11)			
			iStrData = iStrData & Chr(11)			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_CD"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W4"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_DESC"))		
			iStrData = iStrData & Chr(11) & iDx
			iStrData = iStrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
		    
		    iDx = iDx + 1

        Loop 
        
        lgObjRs.Close
        Set lgObjRs = Nothing
		       
    End If
    
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_CD, A.W1, A.W2, A.W3, A.W4, A.W5, A.W_DESC "
            lgStrSQL = lgStrSQL & " FROM TB_50A A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.W_CD ASC" & vbcrlf

                        
    End Select
	PrintLog "SubMakeSQLStatements = " & lgStrSQL
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
	

	call  SubBizSaveMultiDeleteA(arrColVal)
	
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
lgIntFlgMode = CInt(Request("txtFlgMode"))

      'Select Case lgIntFlgMode
	   '     Case  OPMD_CMODE                                                             '☜ : Create
         '           Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
	    '    Case  OPMD_UMODE   
          '          Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
        'End Select
        
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
	
	lgStrSQL = "INSERT INTO TB_50A WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W_CD, W1, W2, W3, W4, W5, W_DESC "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_CD))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_DESC))),"''","S")	& ","

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

    lgStrSQL = "UPDATE  TB_50A WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W_DESC	= " &  FilterVar(Trim(UCase(arrColVal(C_W_DESC ))),"''","S") & ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND W_CD = " & FilterVar(Trim(UCase(arrColVal(C_W_CD) )),"''","S")

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

    lgStrSQL = "DELETE  TB_50A WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_CD = " & FilterVar(Trim(UCase(arrColVal(C_W_CD) )),"''","S")
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

Sub SubBizSaveMultiDeleteA(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_50A WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	'lgStrSQL = lgStrSQL & "		AND W_CD = " & FilterVar(Trim(UCase(arrColVal(C_W_CD) )),"''","S")
	
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

'========================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    'On Error Resume Next
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
