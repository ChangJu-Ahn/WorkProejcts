<%@ Transaction=required Language=VBScript%>
<%Option Explicit%> 
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

	Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 

	Dim C_SEQ_NO
	Dim C_W1_CD
	Dim C_W1
	Dim C_W2
	Dim C_W2_CD
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_SEQ_NO	= 1
	C_W1_CD		= 2	' -- 1번 그리드 
	C_W1		= 3	' 구분 
	C_W2		= 4	' 근거법조항 
	C_W2_CD		= 5 ' 코드 
	C_W3		= 6	' 회사계상액 
	C_W4		= 7 ' 한도초과액 
	C_W5		= 8	' 차감액 
	C_W6		= 9	' 최저한세적용손금부인액 
	C_W7		= 10	' 손금불산입액계 
	C_W8		= 11	' 손금산입계 
End Sub

'========================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_5 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else

		    
		arrRs(iRow) = ""
				
		Do While Not lgObjRs.EOF
			arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W1_CD"))
			arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W1"))
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W2")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W2_CD")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W3")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W4")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W5")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W6")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W7")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W8")
			arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
			arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

			iLngRow = iLngRow + 1
			lgObjRs.MoveNext
		Loop 


		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_1 & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_1)       & """" & vbCr

		Response.Write "	.DbQueryOk                                      " & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
    
    End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1_CD, A.W1 , A.W2, A.W2_CD, A.W3, A.W4, A.W5, A.W6, A.W7, A.W8 "
            lgStrSQL = lgStrSQL & " FROM TB_5 A WITH (NOLOCK) "
            'lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN B_MINOR B WITH (NOLOCK) ON B.MAJOR_CD='W1054' AND A.W1_CD=B.MINOR_CD "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
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

    'On Error Resume Next
    Err.Clear 
	
	' 1번 그리드 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_1))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_1) ), gRowSep)                                 '☜: Split Row    data
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
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_5 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W1_CD, W1, W2, W2_CD, W3, W4, W5, W6, W7, W8 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1_CD))),"''","S")     & "," & vbCrLf
	'If Trim(UCase(arrColVal(C_W1_CD))) <> "" Then
	'	lgStrSQL = lgStrSQL & "''," & vbCrLf
	'Else
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," & vbCrLf
	'End If
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2_CD))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")     & "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_5 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1     = " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2     = " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2_CD  = " &  FilterVar(Trim(UCase(arrColVal(C_W2_CD))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3     = " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5     = " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W8     = " &  FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

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
		
	   Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") <> "NO" Then
             Parent.DBQueryFalse
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
<%
'   **************************************************************
'	1.4 Transaction 처러 이벤트 
'   **************************************************************

Sub	onTransactionCommit()
	' 트랜잭션 완료후 이벤트 처리 
End Sub

Sub onTransactionAbort()
	' 트랜잭선 실패(에러)후 이벤트 처리 
'PrintForm
'	' 에러 출력 
	Call SaveErrorLog(Err)	' 에러로그를 남긴 
	
End Sub
%>
