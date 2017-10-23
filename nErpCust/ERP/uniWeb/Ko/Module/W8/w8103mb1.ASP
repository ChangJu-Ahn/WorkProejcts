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

	Dim C_W1	' 그리드 행(디비의 열)
	Dim C_W2
	Dim C_W3
	Dim C_W4_1
	Dim C_W4_2
	Dim C_W5_1
	Dim C_W5_2
	Dim C_W6_1
	Dim C_W6_2
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10_1

	Dim C_W01	
	Dim C_W02	
	Dim C_W03	
	Dim C_W04	
	Dim C_W05	
	Dim C_W06	
	Dim C_W07	
	Dim C_W09	
	Dim C_W10	
	Dim C_W11	
	Dim C_W12	
	Dim C_W13	
	Dim C_W14	
	Dim C_W15	

	Dim C_W31	
	Dim C_W32	
	Dim C_W33	
	Dim C_W34	
	Dim C_W35	
	Dim C_W36	
	Dim C_W37	
	Dim C_W38	
	Dim C_W39	
	Dim C_W40	
	Dim C_W41	

	Dim C_W51	

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
	C_W1	= 0	
	C_W2	= 1
	C_W3	= 2
	C_W4_1	= 3
	C_W4_2	= 4
	C_W5_1	= 5
	C_W5_2	= 6
	C_W6_1	= 7
	C_W6_2	= 8
	C_W7	= 9
	C_W8	= 10
	C_W9	= 11
	C_W10_1	= 12
	
	C_W01	= 13
	C_W02	= 14
	C_W03	= 15
	C_W04	= 16
	C_W05	= 17
	C_W06	= 18
	C_W07	= 19
	C_W09	= 20
	C_W10	= 21
	C_W11	= 22
	C_W12	= 23
	C_W13	= 24
	C_W14	= 25
	C_W15	= 26

	C_W31	= 27
	C_W32	= 28
	C_W33	= 29
	C_W34	= 30
	C_W35	= 31
	C_W36	= 32
	C_W37	= 33
	C_W38	= 34
	C_W39	= 35
	C_W40	= 36
	C_W41	= 37

	C_W51	= 38
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
    lgStrSQL =            "DELETE TB_58 WITH (ROWLOCK) " & vbCrLf
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

		For iDx = C_W1 To C_W51 -1	
			If iDx = C_W32 Then
				lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").value = """ & lgObjRs(iDx) & "%"" '--" & lgObjRs.Fields(iDx).Name & "" & vbCrLf
			ElseIf  iDx = C_W5_1 OR iDx = C_W5_2 or iDx = C_W6_1 OR iDx = C_W6_2 or iDx=C_W8  Then
				lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").text = """ & UNIDateClientFormat(lgObjRs(iDx)) & """ '--" & lgObjRs.Fields(iDx).Name & "" & vbCrLf
			
			Else
				lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").value = """ & lgObjRs(iDx) & """ '--" & lgObjRs.Fields(iDx).Name & "" & vbCrLf
			End If
			
			If Err.number <> 0 Then
				PrintLog "iDx=" & iDx
				Exit Sub
			End If
		Next 
		
		Response.Write lgstrData  &  vbCrLf
		
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
			lgStrSQL = lgStrSQL & "   W1, W2, W3, W4_1, W4_2, W5_1, W5_2, W6_1, W6_2, W7, W8, W9, W10_1 "  & vbCrLf 
            lgStrSQL = lgStrSQL & " , W01, W02, W03, W04, W05, W06, W07, W09, W10, W11, W12, W13, W14, W15  "  & vbCrLf 
            lgStrSQL = lgStrSQL & " , W31, W32, W33, W34, W35, W36, W37, W38, W39, W40, W41, W51"  & vbCrLf 
            lgStrSQL = lgStrSQL & " FROM TB_58 WITH (NOLOCK) "  & vbCrLf 
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
	
	lgStrSQL = "INSERT INTO TB_58 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
	lgStrSQL = lgStrSQL & "  , W1, W2, W3, W4_1, W4_2, W5_1, W5_2, W6_1, W6_2, W7, W8, W9, W10_1 "
    lgStrSQL = lgStrSQL & " , W01, W02, W03, W04, W05, W06, W07, W09, W10, W11, W12, W13, W14, W15  "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W31, W32, W33, W34, W35, W36, W37, W38, W39, W40, W41, W51"  & vbCrLf 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W4_1))),"'1'","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W4_2))),"'3'","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W5_1))),"NULL","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W5_2))),"NULL","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6_1))),"","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6_2))),"","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")    & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8))),"NULL","S")    & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W10_1))),"''","S")    & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W01), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W02), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W03), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W04), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W05), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W06), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W07), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W09), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W32)), "0"),"0","D")    & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W34), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W35), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W36), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W37), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W38), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W39), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W40), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W41), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W51), "0"),"0","D")     & "," & vbCrLf

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

    lgStrSQL = "UPDATE  TB_58 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4_1    = " &  FilterVar(Trim(UCase(arrColVal(C_W4_1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4_2		= " &  FilterVar(Trim(UCase(arrColVal(C_W4_2))),"''","S") & "," & vbCrLf   
    lgStrSQL = lgStrSQL & " W5_1    = " &  FilterVar(Trim(UCase(arrColVal(C_W5_1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W5_2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6_1    = " &  FilterVar(Trim(UCase(arrColVal(C_W6_1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W6_2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W10_1	= " &  FilterVar(Trim(UCase(arrColVal(C_W10_1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W01     = " &  FilterVar(UNICDbl(arrColVal(C_W01), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W02     = " &  FilterVar(UNICDbl(arrColVal(C_W02), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W03     = " &  FilterVar(UNICDbl(arrColVal(C_W03), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W04     = " &  FilterVar(UNICDbl(arrColVal(C_W04), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W05     = " &  FilterVar(UNICDbl(arrColVal(C_W05), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W06     = " &  FilterVar(UNICDbl(arrColVal(C_W06), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W07     = " &  FilterVar(UNICDbl(arrColVal(C_W07), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W09     = " &  FilterVar(UNICDbl(arrColVal(C_W09), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W10     = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11     = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(RemovePercent(Request("txtW14")), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W15     = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & "," & vbCrLf

    lgStrSQL = lgStrSQL & " W31     = " &  FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W32     = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W32)), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W33     = " &  FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W34     = " &  FilterVar(UNICDbl(arrColVal(C_W34), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W35     = " &  FilterVar(UNICDbl(arrColVal(C_W35), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W36     = " &  FilterVar(UNICDbl(arrColVal(C_W36), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W37     = " &  FilterVar(UNICDbl(arrColVal(C_W37), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W38     = " &  FilterVar(UNICDbl(arrColVal(C_W38), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W39     = " &  FilterVar(UNICDbl(arrColVal(C_W39), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W40     = " &  FilterVar(UNICDbl(arrColVal(C_W40), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W41     = " &  FilterVar(UNICDbl(arrColVal(C_W41), "0"),"0","D") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " W51     = " &  FilterVar(UNICDbl(arrColVal(C_W51), "0"),"0","D") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 


	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
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
             'Parent.DBSaveOk
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
