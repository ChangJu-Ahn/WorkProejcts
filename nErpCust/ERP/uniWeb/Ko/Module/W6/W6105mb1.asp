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
	Dim	C_SEQ_NO
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W4_VAL
	Dim C_W5
	Dim C_W6

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
             Call SubBizSave()
          
             
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_SEQ_NO		= 1
	C_W1		= 2
	C_W2		= 3
	C_W3		= 4
	C_W4		= 5
	C_W4_VAL		= 6
	C_W5		= 7
	C_W6		= 8

End Sub


'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_JT2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 


	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngCol
    Dim iRow, iKey1, iKey2, iKey3, sW_TYPE
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")		' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
    

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
		
		iDx = 1

				lgstrData = lgstrData & " With parent.frm1.vspdData " & vbCr
				lgstrData = lgstrData & "	.Redraw = false " & vbCr
				Do While Not lgObjRs.EOF
					lgstrData = lgstrData & "	.Row = " &iDx & "" & vbCrLf
					lgstrData = lgstrData & "	.Col = 0 : .text = """" " & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_SEQ_NO & " : .text = """ & lgObjRs("SEQ_NO") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W1 & " : .text = """ & lgObjRs("W1") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W2 & " : .text = """ & lgObjRs("W2") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W3 & " : .text = """ & lgObjRs("W3") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W4 & " : .text = """ & lgObjRs("W4") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W4_VAL & " : .text = """ & lgObjRs("W4_VAL") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W5 & " : .text = """ & lgObjRs("W5") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W6 & " : .text = """ & lgObjRs("W6") & """" & vbCrLf
                    lgstrData = lgstrData & "	.Col = " & C_W6 + 1 & " : .text = """  &  iDx  & """" & vbCrLf
    
				If Err.number <> 0 Then
					PrintLog "iDx=" & iDx
					Exit Sub
				End If
		        iDx = iDx +1    
				lgObjRs.MoveNext
		
			Loop


		
		lgObjRs.Close
		Set lgObjRs = Nothing
			
		lgstrData = lgstrData & "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE" & vbCrLf
    End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write lgstrData  &  vbCrLf
		
	If lgstrData <> "" Then	
		Response.Write "	.Redraw = True " & vbCr
		Response.Write " End With " & vbCrLf	' With 문 종료 
	End If
	
	Response.Write " Call parent.DbQueryOk                                      " & vbCr
	Response.Write " </Script>                                          " & vbCr
	    
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
            lgStrSQL = lgStrSQL & "   SEQ_NO, W1, W2, W3, W4, W4_Val ,W5,W6  "
            lgStrSQL = lgStrSQL & " FROM TB_JT2 WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf

	 Case "Q"
			
	
    End Select

End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

    On Error Resume Next
    Err.Clear 
     lgIntFlgMode = CInt(Request("txtFlgMode")) 
	
	' 신규입력 

	PrintLog "txtSpread = " & Request("txtSpread")
			
	 arrRowVal = Split(Request("txtSpread"), gRowSep)                          '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
 
	 For iDx = 1 To lgLngMaxRow 
                arrColVal = Split(arrRowVal(iDx-1), gColSep)   
                
         
                                              '☜: Split Column data   
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

	Err.Clear
	dim iType


	lgStrSQL = "INSERT INTO TB_JT2 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , SEQ_NO, W1, W2, W3, W4, W4_VAL, W5,  w6 "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim((arrColVal(C_W1))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim((arrColVal(C_W2))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim((arrColVal(C_W3))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim((arrColVal(C_W4))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4_VAL), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")     & "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
  


End Sub




'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
	On Error Resume Next   
	Err.Clear
	dim iType


	
	

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

 
End Sub



'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_JT2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1        = " &  FilterVar(Trim((arrColVal(C_W1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2        = " &  FilterVar(Trim((arrColVal(C_W2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3        = " &  FilterVar(Trim((arrColVal(C_W3))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4        = " &  FilterVar(Trim((arrColVal(C_W4))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4_VAL    = " &  FilterVar(UNICDbl(arrColVal(C_W4_VAL), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5		  = " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6		  = " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & "," & vbCrLf  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function
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
