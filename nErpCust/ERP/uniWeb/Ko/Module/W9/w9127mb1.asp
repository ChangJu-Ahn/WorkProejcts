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
	Dim lgStrPrevKey, lgCurrGrid


	' -- 서식명 
	Dim C_C_NM
	Dim C_C_CD

	Dim C_C01
	Dim C_C02
	Dim C_C03
	Dim C_C04
	Dim C_C05
	Dim C_C06
	Dim C_C07
	Dim C_C08
	Dim C_C09
	Dim C_C10
	Dim C_C11
	Dim C_C12

	' -- 행정보(서식)
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16
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
	Dim C_W47
	Dim C_W48
	Dim C_W49
	Dim C_W50
	Dim C_W51
	Dim C_W52
	Dim C_W53

	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	
    lgPrevNext			= Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow			= Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey		= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
    
    
Function  MinorQueryRs(byval strMajorcd, byval strMinorcd)
   Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
   Call CommonQueryRs("MINOR_NM"," B_MINOR "," MAJOR_CD = '"& strMajorcd &"' and minor_cd = '"& Trim(strMinorcd) &"' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   MinorQueryRs = replace(lgF0, chr(11),"")
   

End Function    

Function ChangevbCrLf(byval pData)
   ChangevbCrLf = Replace(Trim(pData), vbCrLf, """ & vbCrLf & """)

End Function

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 

	' -- 서식명 
	C_C_NM			= 1
	C_C_CD			= 2

	' -- 컬럼 정보 
	C_C01			= 3
	C_C02			= 4
	C_C03			= 5
	C_C04			= 6
	C_C05			= 7
	C_C06			= 8
	C_C07			= 9
	C_C08			= 10
	C_C09			= 11
	C_C10			= 12
	C_C11			= 13
	C_C12			= 14
	
	' -- 행정보(서식)
	C_W1			= 1
	C_W2			= 2
	C_W3			= 3
	C_W4			= 4
	C_W5			= 5
	C_W6			= 6
	C_W7			= 7
	C_W8			= 8
	C_W9			= 9
	C_W10			= 10
	C_W11			= 11
	C_W12			= 12
	C_W13			= 13
	C_W14			= 14
	C_W15			= 15
	C_W16			= 16
	C_W17			= 17
	C_W18			= 18
	C_W19			= 19
	C_W20			= 20
	C_W21			= 21
	C_W22			= 22
	C_W23			= 23
	C_W24			= 24
	C_W25			= 25
	C_W26			= 26
	C_W27			= 27
	
	C_W28			= 1
	C_W29			= 2 
	C_W30			= 3 
	C_W31			= 4 
	C_W32			= 5 
	C_W33			= 6 
	C_W34			= 7 
	C_W35			= 8 
	C_W36			= 9 
	C_W37			= 10
	C_W38			= 11
	C_W39			= 12
	C_W40			= 13
	C_W41			= 14
	C_W42			= 15
	C_W43			= 16
	C_W44			= 17
	C_W45			= 18
	C_W46			= 19
	C_W47			= 20
	C_W48			= 21
	C_W49			= 22 
	C_W50			= 23 
	C_W51			= 24 
	C_W52			= 25 
	C_W53			= 26 
	
End Sub


'========================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
   ' On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_A126 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, sData
    Dim iDx
    Dim iLoopMax,iLngCol,sW_TYPE,strMajor, arrRs
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

	Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	     lgStrPrevKey = ""
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		    
	Else
	    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
	    lgstrData = ""

		iLngCol = lgObjRs.Fields.Count
		
		sW_TYPE = "" : lgstrData = ""
		iDx = 1

			lgstrData = lgstrData & " With parent " & vbCr
			'lgstrData = lgstrData & "	.InitData " &  vbCrLf
			lgstrData = lgstrData & "	.frm1.vspdData.Redraw = false " & vbCr
			
			Do While Not lgObjRs.EOF
			
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W1 & ", """ & lgObjRs("W01") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W2 & ", """ & lgObjRs("W02") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W3 & ", """ & lgObjRs("W03") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W4 & ", """ & lgObjRs("W04") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W5 & ", """ & lgObjRs("W05") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W6 & ", """ & lgObjRs("W06") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W7 & ", """ & lgObjRs("W07") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W8 & ", """ & lgObjRs("W08") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W9 & ", """ & lgObjRs("W09") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W10 & ", """ & lgObjRs("W10") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W11 & ", """ & lgObjRs("W11") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W12 & ", """ & lgObjRs("W12") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W13 & ", """ & lgObjRs("W13") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W14 & ", """ & lgObjRs("W14") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W15 & ", """ & lgObjRs("W15") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W16 & ", """ & lgObjRs("W16") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W17 & ", """ & lgObjRs("W17") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W18 & ", """ & lgObjRs("W18") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W19 & ", """ & lgObjRs("W19") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W20 & ", """ & lgObjRs("W20") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W21 & ", """ & lgObjRs("W21") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W22 & ", """ & lgObjRs("W22") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W23 & ", """ & lgObjRs("W23") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W24 & ", """ & lgObjRs("W24") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W25 & ", """ & lgObjRs("W25") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W26 & ", """ & lgObjRs("W26") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W27 & ", """ & lgObjRs("W27") & """" & vbCrLf
				
				
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W27 & ", """ & lgObjRs("W27") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W28 & ", """ & lgObjRs("W28") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W29 & ", """ & lgObjRs("W29") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W30 & ", """ & lgObjRs("W30") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W31 & ", """ & lgObjRs("W31") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W32 & ", """ & lgObjRs("W32") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W33 & ", """ & lgObjRs("W33") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W34 & ", """ & lgObjRs("W34") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W35 & ", """ & lgObjRs("W35") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W36 & ", """ & lgObjRs("W36") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W37 & ", """ & lgObjRs("W37") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W38 & ", """ & lgObjRs("W38") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W39 & ", """ & lgObjRs("W39") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W40 & ", """ & lgObjRs("W40") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W41 & ", """ & lgObjRs("W41") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W42 & ", """ & lgObjRs("W42") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W43 & ", """ & lgObjRs("W43") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W44 & ", """ & lgObjRs("W44") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W45 & ", """ & lgObjRs("W45") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W46 & ", """ & lgObjRs("W46") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W47 & ", """ & lgObjRs("W47") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W48 & ", """ & lgObjRs("W48") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W49 & ", """ & lgObjRs("W49") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W50 & ", """ & lgObjRs("W50") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W51 & ", """ & lgObjRs("W51") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W52 & ", """ & lgObjRs("W52") & """" & vbCrLf
				lgstrData = lgstrData & "	.SetValue4Grid2 .C_C" & Right("0"&iDx,2) & "," & C_W53 & ", """ & lgObjRs("W53") & """" & vbCrLf
				
																


					
				iDx = iDx + 1
				lgObjRs.MoveNext
			Loop


		
		lgObjRs.Close
		Set lgObjRs = Nothing
			
		'lgstrData = lgstrData & "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE" & vbCrLf
    End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write lgstrData  &  vbCrLf
		
	If lgstrData <> "" Then	
		Response.Write "	.frm1.vspdData.Redraw = True " & vbCr
		Response.Write " End With " & vbCrLf	' With 문 종료 
		
		
	End If

'	Response.Write " Call parent.DbQueryOk                                      " & vbCr
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
            lgStrSQL = lgStrSQL & "  *"
            lgStrSQL = lgStrSQL & " FROM TB_A126 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

    End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i, sData

    'On Error Resume Next
    Err.Clear 

	sData = Request("txtSpread")
	PrintLog "1번째 그리드.. : " & sData
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow
		
			Response.Write "arrColVal=" & arrRowVal(iDx-1) & "<br>" & vbCrLf
			
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
		       Exit Sub
		    End If
		    
		Next
	End If

End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	'On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_A126 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO" 
	lgStrSQL = lgStrSQL & " , W01, W02, W03, W04, W05, W06 , W07 ,  W08 , W09 , W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22, W23, W24, W25, W26 "
	lgStrSQL = lgStrSQL & " , W27, W28, W29, W30, W31, W32, W33, W34, W35, W36, W37, W38, W39, W40, W41, W42, W43, W44, W45, W46, W47, W48, W49, W50,W51,W52,W53 "
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & " VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","

	For i = 1 To 54
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(i))),"0","N")     & ","
	Next


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
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_A126 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 

	For i = 1 To 53
		lgStrSQL = lgStrSQL & " W" & Right("0" & i,2) & "     = " &  FilterVar(Trim(UCase(arrColVal(i+1))),"0","N")     & ","
	Next
    
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(1))),"1","N") 	 & vbCrLf  

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_A125 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(1))),"''","S")  


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
    On Error Resume Next
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

          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk        
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