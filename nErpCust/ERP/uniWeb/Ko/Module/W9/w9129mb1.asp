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

	Dim C_C01
	Dim C_C01_POP
	Dim C_C02
	Dim C_C02_POP
	Dim C_C03
	Dim C_C03_POP
	Dim C_C04
	Dim C_C04_POP
	Dim C_C05
	Dim C_C05_POP
	Dim C_C06
	Dim C_C06_POP
	Dim C_C07
	Dim C_C07_POP
	Dim C_C08
	Dim C_C08_POP
	Dim C_C09
	Dim C_C09_POP
	Dim C_C10
	Dim C_C10_POP
	Dim C_C11
	Dim C_C11_POP
	Dim C_C12
	Dim C_C12_POP		

	' -- 행정보(서식)
	Dim C_W6
	Dim C_W6_1
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W12_1
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
	' -- 컬럼 정보 
	C_C01			= 3
	C_C01_POP		= 4
	C_C02			= 5
	C_C02_POP		= 6
	C_C03			= 7
	C_C03_POP		= 8
	C_C04			= 9
	C_C04_POP		= 10
	C_C05			= 11
	C_C05_POP		= 12
	C_C06			= 13
	C_C06_POP		= 14
	C_C07			= 15
	C_C07_POP		= 16
	C_C08			= 17
	C_C08_POP		= 18
	C_C09			= 19
	C_C09_POP		= 20
	C_C10			= 21
	C_C10_POP		= 22
	C_C11			= 23
	C_C11_POP		= 24
	C_C12			= 25
	C_C12_POP		= 26
	
	C_W6			= 1
	C_W6_1			= 2
	C_W7			= 3
	C_W8			= 4
	C_W9			= 5
	C_W10			= 6
	C_W11			= 7
	C_W12			= 8
	C_W12_1			= 9
	C_W13			= 10
	C_W14			= 11
	C_W15			= 12
	C_W16			= 13 
	C_W17			= 14
	C_W18			= 15
	C_W19			= 16
	C_W20			= 17
	C_W21			= 18
	C_W22			= 19
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
    On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_A128 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	'PrintLog "SubBizDelete = " & lgStrSQL 
	
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
			'lgstrData = lgstrData & "	.frm1.txtW4.value = """ & ConvSPChars(lgObjRs("W4")) & """" & vbCrLf
			'lgstrData = lgstrData & "	.frm1.txtW5.value = """ & ConvSPChars(lgObjRs("W5")) & """" & vbCrLf
			
			Do While Not lgObjRs.EOF
				If lgObjRs("W6") <> "" Then

					lgstrData = lgstrData & "	.ShowColumn .C_C" & Right("0"&iDx,2) & vbCrLf

					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W6 & ", """ & ConvSPChars(lgObjRs("W6")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W6_1 & ", """ & ConvSPChars(lgObjRs("W6_1")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W7 & ", """ & ChangevbCrLf(ConvSPChars(lgObjRs("W7"))) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W8 & ", """ & ConvSPChars(lgObjRs("W8")) & """" & vbCrLf

					If lgObjRs("W9") = "1" Then
						lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W9 & ", ""지점""" & vbCrLf
					Else
						lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W9 & ", ""사무소""" & vbCrLf
					End If

					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W10 & ", """ & UNIDateClientFormat(lgObjRs("W10")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W11 & ", """ & ChangevbCrLf(ConvSPChars(lgObjRs("W11"))) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W12 & ", """ & ConvSPChars(lgObjRs("W12")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W12_1 & ", """ & ConvSPChars(lgObjRs("W12_1")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W13 & ", """ & ConvSPChars(lgObjRs("W13")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W14 & ", """ & ConvSPChars(lgObjRs("W14")) & """" & vbCrLf
					'lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W15 & ", """ & ConvSPChars(lgObjRs("W15")) & """" & vbCrLf
					'lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W16 & ", """ & ConvSPChars(lgObjRs("W16")) & """" & vbCrLf
					'lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W17 & ", """ & ConvSPChars(lgObjRs("W17")) & """" & vbCrLf
					'lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W18 & ", """ & ConvSPChars(lgObjRs("W18")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W19 & ", """ & ConvSPChars(lgObjRs("W19")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W20 & ", """ & ConvSPChars(lgObjRs("W20")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetText4Grid .C_C" & Right("0"&iDx,2) & "," & C_W21 & ", """ & UNIDateClientFormat(lgObjRs("W21")) & """" & vbCrLf
					lgstrData = lgstrData & "	.SetValue4Grid .C_C" & Right("0"&iDx,2) & "," & C_W22 & ", """ & ConvSPChars(lgObjRs("W22")) & """" & vbCrLf
				
				End If
					
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
			lgStrSQL =			  " SELECT   CO_CD, FISC_YEAR,REP_TYPE, SEQ_NO, W4,W5,W6,COUNTRY_NM AS W6_1,	    " & vbcrlf
			lgStrSQL = lgStrSQL & " W7,W8,W9, W10,W11,W12,FULL_DETAIL_NM AS W12_1,W13,	     W14,	    " & vbcrlf 
			lgStrSQL = lgStrSQL & " CASE WHEN W15=0 THEN NULL ELSE W15 END W15," & vbcrlf		
			lgStrSQL = lgStrSQL & " CASE WHEN W16=0 THEN NULL ELSE W16 END W16,	" & vbcrlf    
			lgStrSQL = lgStrSQL & " CASE WHEN W17=0 THEN NULL ELSE W17 END W17,	" & vbcrlf    
			lgStrSQL = lgStrSQL & " CASE WHEN W18=0 THEN NULL ELSE W18 END W18,	      W19,	 " & vbcrlf    
			lgStrSQL = lgStrSQL & " CASE WHEN W20=0 THEN NULL ELSE W20 END W20,	     W21 ,W22 " & vbcrlf
            
            lgStrSQL = lgStrSQL & " FROM TB_A128 A WITH (NOLOCK)  " & vbcrlf
            lgStrSQL = lgStrSQL & "		LEFT JOIN UFN_TB_COUNTRY('200603') B ON A.W6=COUNTRY_CD  " & vbcrlf
			lgStrSQL = lgStrSQL & "		LEFT JOIN TB_STD_INCOME_RATE C (NOLOCK) ON A.W12=c.STD_INCM_RT_CD " 

			If wgFISC_YEAR >= "2006" Then
				lgStrSQL = lgStrSQL & "  AND c.ATTRIBUTE_YEAR='2005' "& vbCrLf
			Else
				lgStrSQL = lgStrSQL & "  AND c.ATTRIBUTE_YEAR='2003' "& vbCrLf
			End If
		
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

    End Select
	'PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i, sData

    On Error Resume Next
    Err.Clear 

	sData = Request("txtSpread")
	'PrintLog "1번째 그리드.. : " & sData
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
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
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_A128 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO" 
	lgStrSQL = lgStrSQL & " , W4, W5, W6 , W7 ,  W8 , W9 , W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22 "
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & " VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(1))),"1","N")     & ","
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtW4"))),"''","S")     & ","
	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")     & "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(6))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(7))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(8))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(9))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(10))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(11))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(12))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(13))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(14))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(15))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(16))),"0","N")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(17))),"NULL","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(18))),"0","N")     & ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"
 
	'PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
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

    lgStrSQL = "UPDATE  TB_A128 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 

	lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(Trim(UCase(Request("txtW4"))),"''","S")     & ","	& vbCrLf
	lgStrSQL = lgStrSQL & " W5     = " &   FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")     & "," & vbCrLf

    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W8     = " &  FilterVar(Trim(UCase(arrColVal(4))),"''","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W9     = " &  FilterVar(Trim(UCase(arrColVal(5))),"''","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W10    = " &  FilterVar(Trim(UCase(arrColVal(6))),"NULL","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11	   = " &  FilterVar(Trim(UCase(arrColVal(7))),"''","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12    = " &  FilterVar(Trim(UCase(arrColVal(8))),"''","S")     & "," & vbCrLf

    lgStrSQL = lgStrSQL & " W13    = " &  FilterVar(Trim(UCase(arrColVal(9))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14    = " &  FilterVar(Trim(UCase(arrColVal(10))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W15    = " &  FilterVar(Trim(UCase(arrColVal(11))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W16    = " &  FilterVar(Trim(UCase(arrColVal(12))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W17    = " &  FilterVar(Trim(UCase(arrColVal(13))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W18    = " &  FilterVar(Trim(UCase(arrColVal(14))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W19    = " &  FilterVar(Trim(UCase(arrColVal(15))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W20    = " &  FilterVar(Trim(UCase(arrColVal(16))),"0","N")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W21    = " &  FilterVar(Trim(UCase(arrColVal(17))),"NULL","S")     & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W22    = " &  FilterVar(Trim(UCase(arrColVal(18))),"0","N")     & "," & vbCrLf

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(1))),"1","N") 	 & vbCrLf  

	'PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
 
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

    lgStrSQL = "DELETE  TB_A128 WITH (ROWLOCK) "	 & vbCrLf
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