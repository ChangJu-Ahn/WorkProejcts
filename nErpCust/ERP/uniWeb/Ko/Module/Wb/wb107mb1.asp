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

	' -- 그리드 컬럼 정의 
	Dim C_W01	
	Dim C_W04	
	Dim C_W07	
	Dim C_W02	
	Dim C_W05	
	Dim C_W08
	Dim C_W06	
	Dim C_W09
	Dim C_W_SUM	
	Dim C_W19	
	Dim C_W10	
	Dim C_W11	
	Dim C_W12
	Dim C_W13	
	Dim C_W14	
	Dim C_W15	
	Dim C_W16	
	Dim C_W17	
	Dim C_W20	
	Dim C_W21
	Dim C_W18
	Dim C_W22
	Dim C_W23

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")
	lgIntFlgMode	= Request("txtHeadMode")
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
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_W01		= 0	' HTML상의 순서 
	C_W04		= 1
	C_W07		= 2	
	C_W02		= 3
	C_W05		= 4
	C_W08		= 5
	C_W06		= 6
	C_W09		= 7
	C_W_SUM		= 8
	C_W19		= 9
	C_W10		= 11
	C_W11		= 12
	C_W12		= 13
	C_W13		= 14
	C_W14		= 15
	C_W15		= 16
	C_W16		= 17
	C_W17		= 18
	C_W20		= 19
	C_W21		= 20
	C_W18		= 21 
	C_W22		= 22
	C_W23		= 10
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
    lgStrSQL =            "DELETE TB_51 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngCol
    Dim iRow, iKey1, iKey2, iKey3
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.IsRunEvents = True " & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.

		For iDx = C_W01 To C_W22 -1	
			lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").value = """ & lgObjRs(iDx) & """" & vbCrLf
			If Err.number <> 0 Then
				PrintLog "iDx=" & iDx
				Exit Sub
			End If
		Next 
		
		Response.Write lgstrData  &  vbCrLf
	
		Response.Write "	Call parent.QueryRadio()" & vbCr
		Response.Write "	.IsRunEvents = False " & vbCrLf	' 이벤트가 발생하게 한다.
		
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
            lgStrSQL = lgStrSQL & "   W01, W04, W07, W02, W05, W08, W06, W09, W_SUM, W19, W23, W10, W11, W12, W13, W14, W15, W16 "
            lgStrSQL = lgStrSQL & " , W17, W20, W21, W18, W22 "
            lgStrSQL = lgStrSQL & " FROM TB_51 WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
	
    End Select

	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

    On Error Resume Next
    Err.Clear 
    
	
	' 신규입력 

	PrintLog "txtSpread = " & Request("txtSpread")
			
	arrColVal = Split(Request("txtSpread"), gColSep)                                 '☜: Split Col   data
	lgLngMaxRow = UBound(arrColVal)
	
	PrintLog "SubBizSave = " & lgIntFlgMode & ";" & OPMD_CMODE
	
	If CDbl(lgIntFlgMode) =  OPMD_CMODE Then
		PrintLog Err.Description 
		Call SubBizSaveCreate(arrColVal)                            '☜: Create
	Else
		Call SubBizSaveUpdate(arrColVal)                            '☜: Update
	End If

End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_51 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W01, W02, W04, W05, W06, W07, W08, W09, W_SUM, W10, W11, W12, W13, W14, W15, W16 "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W17, W18, W19, W20, W21, W22, W23 "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W01))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W02))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W04))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W05))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W06))),"''","S")     & "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W07), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W08), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W09), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W_SUM), "0"),"0","D")     & "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S")     & "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W19))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W20))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W21))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W22))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W23))),"''","S")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreateH = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_51 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W01		= " &  FilterVar(Trim(UCase(arrColVal(C_W01))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W04		= " &  FilterVar(Trim(UCase(arrColVal(C_W04))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W07     = " &  FilterVar(UNICDbl(arrColVal(C_W07), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W02		= " &  FilterVar(Trim(UCase(arrColVal(C_W02))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W05		= " &  FilterVar(Trim(UCase(arrColVal(C_W05))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W08     = " &  FilterVar(UNICDbl(arrColVal(C_W08), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W06		= " &  FilterVar(Trim(UCase(arrColVal(C_W06))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W09     = " &  FilterVar(UNICDbl(arrColVal(C_W09), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W_SUM   = " &  FilterVar(UNICDbl(arrColVal(C_W_SUM), "0"),"0","D") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " W10     = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11     = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W15     = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W16     = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W17     = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W18     = " &  FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " W19		= " &  FilterVar(Trim(UCase(arrColVal(C_W19))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W20		= " &  FilterVar(Trim(UCase(arrColVal(C_W20))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W21		= " &  FilterVar(Trim(UCase(arrColVal(C_W21))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W22		= " &  FilterVar(Trim(UCase(arrColVal(C_W22))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W23		= " &  FilterVar(Trim(UCase(arrColVal(C_W23))),"''","S") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 


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
	'Call SaveErrorLog(Err)	' 에러로그를 남긴 
	
End Sub
%>
