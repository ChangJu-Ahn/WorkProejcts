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

	Const TAB1 = 1																	'☜: Tab의 위치 
	Const TAB2 = 2
	Const TAB3 = 3
	Const TAB4 = 4

	Const TYPE_1	= 0		' 그리드를 구분짓기 위한 상수 
	Const TYPE_2_1	= 1		
	Const TYPE_2_2	= 2		 
	Const TYPE_3	= 3		
	Const TYPE_4	= 4		

	' -- 그리드 컬럼 정의 
	Dim C_1_W1
	Dim C_1_W2
	Dim C_1_W1_CD
	Dim C_1_W3
	Dim C_1_W4_1
	Dim C_1_W4_2
	Dim C_1_W4_3
	Dim C_1_W4_4
	Dim C_1_W4_5
	Dim C_1_W4_6
	Dim C_1_W5

	Dim C_2_SEQ_NO
	Dim C_2_W1
	Dim C_2_W1_BT
	Dim C_2_W1_NM
	Dim C_2_W2
	Dim C_2_W3_CD
	Dim C_2_W3
	Dim C_2_W4

	Dim C_3_W_TYPE
	Dim C_3_W1
	Dim C_3_W2
	Dim C_3_W3_1
	Dim C_3_W4_1
	Dim C_3_W5_1
	Dim C_3_W6_1
	Dim C_3_W7_1
	Dim C_3_W8_1
	Dim C_3_W3_2
	Dim C_3_W4_2
	Dim C_3_W5_2
	Dim C_3_W6_2
	Dim C_3_W7_2
	Dim C_3_W8_2
	Dim C_3_W9
	Dim C_3_W10
	Dim C_3_W11
	
	Dim C_4_W1
	Dim C_4_W2
	Dim C_4_W1_CD
	Dim C_4_W3
	Dim C_4_W4_1
	Dim C_4_W5_1
	Dim C_4_W4_2
	Dim C_4_W5_2
	Dim C_4_W4_3
	Dim C_4_W5_3
	Dim C_4_W4_4
	Dim C_4_W5_4
	Dim C_4_W4_5
	Dim C_4_W5_5
	Dim C_4_W4_6
	Dim C_4_W5_6
	Dim C_4_W6
	Dim C_4_W7
	Dim C_4_DESC1

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
	C_1_W1		= 1
	C_1_W2		= 2
	C_1_W1_CD	= 3
	C_1_W3		= 4
	C_1_W4_1	= 5
	C_1_W4_2	= 6
	C_1_W4_3	= 7
	C_1_W4_4	= 8
	C_1_W4_5	= 9
	C_1_W4_6	= 10
	C_1_W5		= 11

	C_2_SEQ_NO	= 1
	C_2_W1		= 2
	C_2_W1_BT	= 3
	C_2_W1_NM	= 4
	C_2_W2		= 5
	C_2_W3_CD	= 6
	C_2_W3		= 7
	C_2_W4		= 8

	C_3_W_TYPE	= 1
	C_3_W1		= 2
	C_3_W2		= 3
	C_3_W3_1	= 4
	C_3_W4_1	= 5
	C_3_W5_1	= 6
	C_3_W6_1	= 7
	C_3_W7_1	= 8
	C_3_W8_1	= 9
	C_3_W3_2	= 10
	C_3_W4_2	= 11
	C_3_W5_2	= 12
	C_3_W6_2	= 13
	C_3_W7_2	= 14
	C_3_W8_2	= 15
	C_3_W9		= 16
	C_3_W10		= 17
	C_3_W11		= 18

	C_4_W1		= 1
	C_4_W2		= 2
	C_4_W1_CD	= 3
	C_4_W3		= 4
	C_4_W4_1	= 5
	C_4_W5_1	= 6
	C_4_W4_2	= 7
	C_4_W5_2	= 8
	C_4_W4_3	= 9
	C_4_W5_3	= 10
	C_4_W4_4	= 11
	C_4_W5_4	= 12
	C_4_W4_5	= 13
	C_4_W5_5	= 14
	C_4_W4_6	= 15
	C_4_W5_6	= 16
	C_4_W6		= 17
	C_4_W7		= 18
	C_4_DESC1	= 19
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
    lgStrSQL =            "DELETE TB_48H2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_48D2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
    lgStrSQL = lgStrSQL & "DELETE TB_48D1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    'lgStrSQL = lgStrSQL & "DELETE TB_48H1 WITH (ROWLOCK) " & vbCrLf
	'lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	'lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	'lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
    lgStrSQL = lgStrSQL & "DELETE TB_48H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	'PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
	Dim arrRow(2), iType
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("RH",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "	Call Parent.FncNew()       " & vbCr
		Response.Write " </Script>	                        " & vbCr

        Call SetErrorStatus()
        
    Else
		lgstrData = "" : iLngRow = 1
        
		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W3 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W3") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W4_1 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_1") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W4_2 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_2") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W4_3 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_3") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W4_4 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_4") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W4_5 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_5") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W4_6 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_6") & """)" & vbCrLf
			lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_1 & ", " & C_1_W5 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5") & """)" & vbCrLf & vbCrLf

			lgObjRs.MoveNext
		Loop 

		lgObjRs.Close
		Set lgObjRs = Nothing
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	Call .MakeDefaultGrid(""Q"")       " & vbCr
		Response.Write lgstrData
		Response.Write " End With                                  " & vbCr
		Response.Write " </Script>	                        " & vbCr

		' 1번 감면분 콤보값 
	    Call SubMakeSQLStatements("RH1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do While Not lgObjRs.EOF
				iType = CDbl(lgObjRs("W_TYPE"))
				'Response.Write "	arrRet(" & CStr(iType-1) & ") = """ & ConvSPChars(lgObjRs("W_NM")) & ";"
				lgstrData = lgstrData & "	arrRet(" & CStr(iType-1) & ") = """ & ConvSPChars(lgObjRs("W_NM")) & """" & vbCrLf
				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent										" & vbCr
			Response.Write "	Dim arrRet(6)                                 " & vbCr
			Response.Write lgstrData & vbCr
			Response.Write "	Call .SetColHead(arrRet) " & vbCr
			Response.Write " End With                                  " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If
				
		' 2번 그리드 
	    Call SubMakeSQLStatements("RD1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do While Not lgObjRs.EOF
				iType = CDbl(lgObjRs("W_TYPE"))
				
				arrRs(iType) = arrRs(iType) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iType) = arrRs(iType) & Chr(11) & ConvSPChars(lgObjRs("W1"))
				arrRs(iType) = arrRs(iType) & Chr(11) & ""
				arrRs(iType) = arrRs(iType) & Chr(11) & ConvSPChars(lgObjRs("W1_NM"))
				arrRs(iType) = arrRs(iType) & Chr(11) & lgObjRs("W2")
				arrRs(iType) = arrRs(iType) & Chr(11) & ConvSPChars(lgObjRs("W3"))
				arrRs(iType) = arrRs(iType) & Chr(11) & ConvSPChars(lgObjRs("W3_NM"))
				arrRs(iType) = arrRs(iType) & Chr(11) & ConvSPChars(lgObjRs("W4"))
				arrRs(iType) = arrRs(iType) & Chr(11) & iLngRow
				arrRs(iType) = arrRs(iType) & Chr(11) & Chr(12)
				
				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 


			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent										" & vbCr
			Response.Write "	Call .InitSpreadComboBox2                                 " & vbCr
			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2_1 & ")" & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_2_1)       & """" & vbCr
			
			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2_2 & ")" & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_2_2)       & """" & vbCr
			
			Response.Write "	Call .SetSpreadTotalLine" & vbCr
			Response.Write " End With                                  " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If
		
		' 3번 그리드 
	    Call SubMakeSQLStatements("RD2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do While Not lgObjRs.EOF
			
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W3_1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W3_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W4_1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W4_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W5_1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W5_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W6_1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W6_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W7_1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W7_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W8_1 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W8_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W3_2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W3_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W4_2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W4_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W5_2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W5_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W6_2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W6_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W7_2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W7_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W8_2 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W8_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W9 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W9") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W10 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W10") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_3 & ", " & C_3_W11 & ", " & CDbl(lgObjRs("W_TYPE")) & ", """ & lgObjRs("W11") & """)" & vbCrLf
		
				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent										" & vbCr
			
			Response.Write lgstrData & vbCr

			Response.Write "	.frm1.cboW11.value = .GetGrid(" & TYPE_3 & ", " & C_3_W11 & ", 1) " & vbCr
			
			Response.Write " End With                                  " & vbCr
			Response.Write " </Script>	                        " & vbCr
			
		End If
			
		' 4번 그리드 
	    Call SubMakeSQLStatements("RH2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do While Not lgObjRs.EOF
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W3 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W3") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W4_1 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W5_1 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5_1") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W4_2 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W5_2 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5_2") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W4_3 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_3") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W5_3 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5_3") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W4_4 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_4") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W5_4 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5_4") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W4_5 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_5") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W5_5 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5_5") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W4_6 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W4_6") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W5_6 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W5_6") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W6 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W6") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_W7 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("W7") & """)" & vbCrLf
				lgstrData = lgstrData & "	Call .PutGrid(" & TYPE_4 & ", " & C_4_DESC1 & ", " & CDbl(lgObjRs("W1_CD")) & ", """ & lgObjRs("DESC1") & """)" & vbCrLf & vbCrLf

				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
		
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                   " & vbCr
			Response.Write lgstrData
			Response.Write " End With                                  " & vbCr
			Response.Write " </Script>	                        " & vbCr
		End If
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
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
      Case "RH"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W1_CD, A.W3, A.W4_1, A.W4_2, A.W4_3, A.W4_4, A.W4_5, A.W4_6, A.W5 "
            lgStrSQL = lgStrSQL & " FROM TB_48H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.SEQ_NO, A.W1, B.ITEM_NM W1_NM, A.W2, A.W3, dbo.ufn_GetCodeName('W1063', A.W3) W3_NM "
            lgStrSQL = lgStrSQL & " , A.W4, dbo.ufn_TB_48_GetCodeName(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ", A.W4) W4_NM "
            lgStrSQL = lgStrSQL & " FROM TB_48D1 A WITH (NOLOCK) "
            lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN TB_ADJUST_ITEM B WITH (NOLOCK) ON A.W1 = B.ITEM_CD "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

      Case "RD2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.W1, A.W2, A.W3_1, A.W4_1, A.W5_1, A.W6_1, A.W7_1, A.W8_1 "
            lgStrSQL = lgStrSQL & "						, A.W3_2, A.W4_2, A.W5_2, A.W6_2, A.W7_2, A.W8_2 "
            lgStrSQL = lgStrSQL & "	, A.W9, A.W10, A.W11 "
            lgStrSQL = lgStrSQL & " FROM TB_48D2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RH2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W1_CD, A.W3	, A.W4_1, A.W5_1, A.W4_2, A.W5_2, A.W4_3, A.W5_3 "
            lgStrSQL = lgStrSQL & "					, A.W4_4, A.W5_4, A.W4_5, A.W5_5, A.W4_6, A.W5_6 "
            lgStrSQL = lgStrSQL & " , A.W6, A.W7, A.DESC1 "
            lgStrSQL = lgStrSQL & " FROM TB_48H2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

      Case "RH1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.W_NM "
            lgStrSQL = lgStrSQL & " FROM TB_48H1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

    End Select

	'PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , iType

    'On Error Resume Next
    Err.Clear 
    
    For iType = TYPE_1 To TYPE_4
    
		' 그리드 
		PrintLog "txtSpread" & iType & " = " & Request("txtSpread" & CStr(iType))
			
		arrRowVal = Split(Request("txtSpread" & CStr(iType) ), gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow

		    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
			    
		    Select Case arrColVal(0)
		        Case "C"
		                Call SubBizSaveMultiCreate(iType, arrColVal)                            '☜: Create
		        Case "U"
		                Call SubBizSaveMultiUpdate(iType, arrColVal)                            '☜: Update
		        Case "D"
		                Call SubBizSaveMultiDelete(iType, arrColVal)                            '☜: Update
		    End Select
			    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit For
		    End If
			    
		Next
	Next
	
End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(Byval pType, Byref arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	Select Case pType
		Case TYPE_1
	
			lgStrSQL = "INSERT INTO TB_48H WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W1_CD, W3, W4_1, W4_2, W4_3, W4_4, W4_5, W4_6, W5 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_1_W1_CD))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W3), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W4_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W4_2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W4_3), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W4_4), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W4_5), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W4_6), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_1_W5), "0"),"0","D")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"

		Case TYPE_2_1

			lgStrSQL = "INSERT INTO TB_48D1 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W_TYPE , SEQ_NO, W1, W2, W3, W4 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & "'1'," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_2_SEQ_NO), "0"),"0","D")      & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_2_W1))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_2_W2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_2_W3_CD))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_2_W4))),"''","S")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"
			
		Case TYPE_2_2

			lgStrSQL = "INSERT INTO TB_48D1 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W_TYPE , SEQ_NO, W1, W2, W3, W4 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & "'2'," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_2_SEQ_NO), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_2_W1))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_2_W2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_2_W3_CD))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_2_W4))),"''","S")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"

		Case TYPE_3
		
			lgStrSQL = "INSERT INTO TB_48D2 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W_TYPE, W1, W2, W3_1, W4_1, W5_1, W6_1, W7_1, W8_1 " & vbCrLf
			lgStrSQL = lgStrSQL & "					, W3_2, W4_2, W5_2, W6_2, W7_2, W8_2 " & vbCrLf
			lgStrSQL = lgStrSQL & " , W9 , W10, W11 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			
			If arrColVal(C_3_W_TYPE) = "매출액 비율" Then
				lgStrSQL = lgStrSQL & "'1'," & vbCrLf
			Else
				lgStrSQL = lgStrSQL & "'2'," & vbCrLf
			End if
			
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W2)), "0"),"0","D")    & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W3_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W4_1)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W5_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W6_1)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W7_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W8_1)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W3_2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W4_2)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W5_2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W6_2)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W7_2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W8_2)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_3_W9), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W10)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_3_W11))),"''","S")     & "," & vbCrLf
			'Call svrmsgbox(lgStrSQL,0,1)
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"

		Case TYPE_4

			lgStrSQL = "INSERT INTO TB_48H2 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W1_CD, W3	, W4_1, W5_1, W4_2, W5_2, W4_3, W5_3 " & vbCrLf
			lgStrSQL = lgStrSQL & "				, W4_4, W5_4, W4_5, W5_5, W4_6, W5_6 " & vbCrLf
			lgStrSQL = lgStrSQL & " , W6, W7, DESC1 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_4_W1_CD))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W3), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W4_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_1)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W4_2), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_2)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W4_3), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_3)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W4_4), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_4)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W4_5), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_5)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W4_6), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_6)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_4_W6), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W7)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_4_DESC1))),"''","S")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"
					
	End Select
	PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(Byval pType, Byref arrColVal)
	dim i
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	Select Case pType
		Case TYPE_1
		
			lgStrSQL = "UPDATE  TB_48H WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_1_W3), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_1    = " &  FilterVar(UNICDbl(arrColVal(C_1_W4_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_2    = " &  FilterVar(UNICDbl(arrColVal(C_1_W4_2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_3    = " &  FilterVar(UNICDbl(arrColVal(C_1_W4_3), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_4    = " &  FilterVar(UNICDbl(arrColVal(C_1_W4_4), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_5    = " &  FilterVar(UNICDbl(arrColVal(C_1_W4_5), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_6	= " &  FilterVar(UNICDbl(arrColVal(C_1_W4_6), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(arrColVal(C_1_W5), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W1_CD = " & FilterVar(Trim(UCase(arrColVal(C_1_W1_CD))),"''","S") 	 & vbCrLf 
		
		Case TYPE_2_1
		
			lgStrSQL = "UPDATE  TB_48D1 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_2_W1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(UNICDbl(arrColVal(C_2_W2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_2_W3_CD))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(Trim(UCase(arrColVal(C_2_W4))),"''","S")& "," & vbCrLf
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W_TYPE = '1' " 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_2_SEQ_NO))),"0","N") 	 & vbCrLf 
									
		Case TYPE_2_2
		
			lgStrSQL = "UPDATE  TB_48D1 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_2_W1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(UNICDbl(arrColVal(C_2_W2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_2_W3))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(Trim(UCase(arrColVal(C_2_W4))),"''","S")& "," & vbCrLf
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W_TYPE = '2' " 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_2_SEQ_NO))),"0","N") 	 & vbCrLf 
					
		Case TYPE_3
		
			lgStrSQL = "UPDATE  TB_48D2 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(UNICDbl(arrColVal(C_3_W1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(UNICDbl(arrColVal(C_3_W2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3_1	= " &  FilterVar(UNICDbl(arrColVal(C_3_W3_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_1	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W4_1)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_1	= " &  FilterVar(UNICDbl(arrColVal(C_3_W5_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6_1	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W6_1)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W7_1	= " &  FilterVar(UNICDbl(arrColVal(C_3_W7_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W8_1	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W8_1)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3_2	= " &  FilterVar(UNICDbl(arrColVal(C_3_W3_2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_2	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W4_2)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_2	= " &  FilterVar(UNICDbl(arrColVal(C_3_W5_2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6_2	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W6_2)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W7_2	= " &  FilterVar(UNICDbl(arrColVal(C_3_W7_2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W8_2	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W8_2)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(UNICDbl(arrColVal(C_3_W9), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_3_W10)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W11		= " &  FilterVar(Trim(UCase(arrColVal(C_3_W11))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			
			If arrColVal(C_3_W_TYPE) = "매출액비율" Then
				lgStrSQL = lgStrSQL & "		AND W_TYPE = '1'," & vbCrLf
			Else
				lgStrSQL = lgStrSQL & "		AND W_TYPE = '2'," & vbCrLf
			End if
					
		Case TYPE_4	

			lgStrSQL = "UPDATE  TB_48H2 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_4_W3), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_1    = " &  FilterVar(UNICDbl(arrColVal(C_4_W4_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_1    = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_1)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_2    = " &  FilterVar(UNICDbl(arrColVal(C_4_W4_2), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_2    = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_2)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_3    = " &  FilterVar(UNICDbl(arrColVal(C_4_W4_3), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_3    = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_3)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_4    = " &  FilterVar(UNICDbl(arrColVal(C_4_W4_4), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_4    = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_4)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_5    = " &  FilterVar(UNICDbl(arrColVal(C_4_W4_5), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_5    = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_5)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4_6	= " &  FilterVar(UNICDbl(arrColVal(C_4_W4_6), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5_6	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W5_6)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(UNICDbl(arrColVal(C_4_W6), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_4_W7)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " DESC1	= " &  FilterVar(Trim(UCase(arrColVal(C_4_DESC1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W1_CD = " & FilterVar(Trim(UCase(arrColVal(C_4_W1_CD))),"''","S") 	 & vbCrLf 
			
	End Select
	
	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(Byval pType, Byref arrColVal)
	dim i
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	Select Case pType
		Case TYPE_2_1
		
			lgStrSQL = "DELETE  TB_48D1 "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W_TYPE = '1' " 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_2_SEQ_NO))),"0","N") 	 & vbCrLf 
		
		Case TYPE_2_2
			lgStrSQL = "DELETE  TB_48D1 "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W_TYPE = '2' " 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_2_SEQ_NO))),"0","N") 	 & vbCrLf 

	End Select
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
	
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
