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

	Dim C_W01	' 그리드 행(디비의 열)
	Dim C_W02	
	Dim C_W03	
	Dim C_W04	
	Dim C_W05	
	Dim C_W54	
	Dim C_W06	
	Dim C_W06_1
	Dim C_W07	
	Dim C_W08	
	Dim C_W09	
	Dim C_W10	
	Dim C_W10_1
	Dim C_W11	
	Dim C_W12	
	Dim C_W13	
	Dim C_W14	
	Dim C_W15	
	Dim C_W16	
	Dim C_W16_1
	Dim C_W17	
	Dim C_W18	
	Dim C_W19	
	Dim C_W20	
	Dim C_W21	
	Dim C_W22	
	Dim C_W23	
	Dim C_W24	
	Dim C_W25	
	Dim C_W26	
	Dim C_W27	
	Dim C_W28	
	Dim C_W29	
	Dim C_W30	
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
	Dim C_W42	
	Dim C_W43	
	Dim C_W44	
	Dim C_W45	
	Dim C_W46	
	Dim C_W55
	Dim C_W47	
	Dim C_W48	
	Dim C_W49	
	Dim C_W50	
	Dim C_W51	
	Dim C_W52	
	Dim C_W53

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
	
	' -- 그리드 1
	C_W01	= 0
	C_W02	= 1
	C_W03	= 2
	C_W04	= 3
	C_W05	= 4
	C_W54	= 5
	C_W06	= 6
	
	C_W06_1 = 7
	C_W07	= 8
	C_W08	= 9
	C_W09	= 10
	C_W10	= 11
	
	C_W10_1	= 12
	C_W11	= 13
	C_W12	= 14
	C_W13	= 15
	C_W14	= 16
	C_W15	= 17
	C_W16	= 18
	
	C_W16_1	= 19
	C_W17	= 20
	C_W18	= 21
	C_W19	= 22
	C_W20	= 23
	C_W21	= 24
	C_W22	= 25
	C_W23	= 26
	C_W24	= 27
	C_W25	= 28
	C_W26	= 29
	C_W27	= 30
	C_W28	= 31
	
	' -- 그리드2
	C_W29	= 32
	C_W30	= 33
	C_W31	= 34
	C_W32	= 35
	C_W33	= 36
	C_W34	= 37
	C_W35	= 38
	C_W36	= 39
	C_W37	= 40
	C_W38	= 41
	C_W39	= 42
	C_W40	= 43
	C_W41	= 44
	C_W42	= 45
	C_W43	= 46
	C_W44	= 47
	C_W45	= 48
	C_W46	= 49
	C_W55	= 50	' <-- 2003.03.07 개정추가 
	C_W47	= 51
	C_W48	= 52
	C_W49	= 53
	C_W50	= 54
	C_W51	= 55
	C_W52	= 56
	C_W53	= 57
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
    lgStrSQL =            "DELETE TB_3 WITH (ROWLOCK) " & vbCrLf
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

		For iDx = C_W01 To C_W53 	
		  ' if C_W48 > iDx Then
			 lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").value = """ & lgObjRs(iDx) & """" & vbCrLf
		  ' Else
		   '   lgstrData = lgstrData & "	.frm1.txtData52.value = """ & lgObjRs(iDx) & """" & vbCrLf
		   'End If	 
			 
			If Err.number <> 0 Then
				PrintLog "iDx=" & iDx
				Exit Sub
			End If
		Next 
		
		Response.Write lgstrData  &  vbCrLf
	
		Response.Write "	.frm1.txtW14.value = """ & lgObjRs(iLngCol-5) & """" & vbCr
		Response.Write "	.frm1.txtW35.value = """ & lgObjRs(iLngCol-4) & """" & vbCr
		Response.Write "	.frm1.txtW25_NM.value = """ & lgObjRs(iLngCol-3) & """" & vbCr
		Response.Write "	.frm1.txtW43_NM.value = """ & lgObjRs(iLngCol-2) & """" & vbCr
		
		' -- W55_1: 선박표준이익 추가(2005-12-29)
		Response.Write "	.frm1.txtW55_1.value = """ & lgObjRs(iLngCol-1) & """" & vbCr
		
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
            lgStrSQL = lgStrSQL & "   W01, W02, W03, W04, W05, W54, W06, W06 W06_1, W07, W08, W09, W10, W56 W10_1, W11, W12, W13, W14_CD, W15, W16, W16 W16_1  "	' W10 W10_1 을 W56 W10_1 로 변경
            lgStrSQL = lgStrSQL & " , W17, W18, W19, W20, W21, W22, W23, W24, W25, W26, W27, W28, W29, W30 "
            lgStrSQL = lgStrSQL & " , W31, W32, W33, W34, W35_CD  , W36, W37, W38, W39, W40, W41, W42, W43, W44, W45"
            lgStrSQL = lgStrSQL & " , W46, W55, W47, W48, W49, W50, W51, W52, W53, W14, W35, W25_NM, W43_NM, W55_1 "	' -- W55_1: 선박표준이익 추가(2005-12-29)
            lgStrSQL = lgStrSQL & " FROM TB_3 WITH (NOLOCK) "
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
	
	lgStrSQL = "INSERT INTO TB_3 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W01, W02, W03, W04, W05, W06, W07, W08, W09, W10, W11, W12, W13, W14_CD  , W14, W15, W16 "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W17, W18, W19, W20, W21, W22, W23, W24, W25_NM,   W25, W26, W27, W28, W29, W30 "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W31, W32, W33, W34, W35_CD,   W35, W36, W37, W38, W39, W40, W41, W42, W43_NM,   W43, W44, W45"  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W46, W55, W47, W48, W49, W50, W51, W52, W53, W55_1, W56 "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W01), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W02), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W03), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W04), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W05), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W06), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W07), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W08), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W09), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W14))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(Request("txtW14")), "0"),"0","D")    & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtW25_NM"))),"''","S") & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W34), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W35), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(Request("txtW35")), "0"),"0","D")    & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W36), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W37), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W38), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W39), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W40), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W41), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W42), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtW43_NM"))),"''","S") & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W43), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W44), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W45), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W46), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W55), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W47), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W48), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W49), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W50), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W51), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W52), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W53), "0"),"0","D")     & "," & vbCrLf
	
	' -- W55_1: 선박표준이익 추가(2005-12-29)
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW55_1"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10_1), "0"),"0","D")     & "," & vbCrLf

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

    lgStrSQL = "UPDATE  TB_3 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W01     = " &  FilterVar(UNICDbl(arrColVal(C_W01), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W02     = " &  FilterVar(UNICDbl(arrColVal(C_W02), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W03     = " &  FilterVar(UNICDbl(arrColVal(C_W03), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W04     = " &  FilterVar(UNICDbl(arrColVal(C_W04), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W05     = " &  FilterVar(UNICDbl(arrColVal(C_W05), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W54     = " &  FilterVar(UNICDbl(arrColVal(C_W54), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W06     = " &  FilterVar(UNICDbl(arrColVal(C_W06), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W07     = " &  FilterVar(UNICDbl(arrColVal(C_W07), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W08     = " &  FilterVar(UNICDbl(arrColVal(C_W08), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W09     = " &  FilterVar(UNICDbl(arrColVal(C_W09), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W10     = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11     = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14_CD  = " &  FilterVar(Trim(UCase(arrColVal(C_W14))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(RemovePercent(Request("txtW14")), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W15     = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W16     = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W17     = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W18     = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W19     = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W20     = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W21     = " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W22     = " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W23     = " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W24     = " &  FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W25_NM  = " &  FilterVar(Trim(UCase(Request("txtW25_NM"))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W25     = " &  FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W26     = " &  FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W27     = " &  FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W28     = " &  FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W29     = " &  FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W30     = " &  FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W31     = " &  FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W32     = " &  FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W33     = " &  FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W34     = " &  FilterVar(UNICDbl(arrColVal(C_W34), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W35_CD  = " &  FilterVar(Trim(UCase(arrColVal(C_W35))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W35     = " &  FilterVar(UNICDbl(RemovePercent(Request("txtW35")), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W36     = " &  FilterVar(UNICDbl(arrColVal(C_W36), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W37     = " &  FilterVar(UNICDbl(arrColVal(C_W37), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W38     = " &  FilterVar(UNICDbl(arrColVal(C_W38), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W39     = " &  FilterVar(UNICDbl(arrColVal(C_W39), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W40     = " &  FilterVar(UNICDbl(arrColVal(C_W40), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W41     = " &  FilterVar(UNICDbl(arrColVal(C_W41), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W42     = " &  FilterVar(UNICDbl(arrColVal(C_W42), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W43_NM  = " &  FilterVar(Trim(UCase(Request("txtW43_NM"))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W43     = " &  FilterVar(UNICDbl(arrColVal(C_W43), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W44     = " &  FilterVar(UNICDbl(arrColVal(C_W44), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W45     = " &  FilterVar(UNICDbl(arrColVal(C_W45), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W46     = " &  FilterVar(UNICDbl(arrColVal(C_W46), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W55     = " &  FilterVar(UNICDbl(arrColVal(C_W55), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W47     = " &  FilterVar(UNICDbl(arrColVal(C_W47), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W48     = " &  FilterVar(UNICDbl(arrColVal(C_W48), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W49     = " &  FilterVar(UNICDbl(arrColVal(C_W49), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W50     = " &  FilterVar(UNICDbl(arrColVal(C_W50), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W51     = " &  FilterVar(UNICDbl(arrColVal(C_W51), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W52     = " &  FilterVar(UNICDbl(arrColVal(C_W52), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W53     = " &  FilterVar(UNICDbl(arrColVal(C_W53), "0"),"0","D") & "," & vbCrLf
    
    ' -- W55_1: 선박표준이익 추가(2005-12-29)
    lgStrSQL = lgStrSQL & " W55_1   = " &  FilterVar(UNICDbl(Request("txtW55_1"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W56		= " &  FilterVar(UNICDbl(arrColVal(C_W10_1), "0"),"0","D") & "," & vbCrLf

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
