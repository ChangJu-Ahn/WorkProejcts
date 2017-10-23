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
	Const TYPE_1	= 1		' 그리드를 구분짓기 위한 상수 
	Const TYPE_2	= 2		
	Const TYPE_3	= 3		 
	Const TYPE_4	= 4		 


	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	' -- 그리드 컬럼 정의 
	Dim C_W1
	Dim C_W1_NM1
	Dim C_W1_NM2
	Dim C_W2
	Dim C_W2_NM
	Dim C_W3
	Dim C_W3_NM
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	
	Dim C_W8
	Dim C_W8_NM
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
	Dim C_W20_NM
	Dim C_W21
	Dim C_W22
	Dim C_W23
	Dim C_W24
	
	Dim C_101
	Dim C_102
	Dim C_103
	Dim C_104
	Dim C_105
	Dim C_106
	Dim C_107
	
	Dim C_108
	Dim C_109
	Dim C_110
	
	Dim C_111
	Dim C_112
	Dim C_113

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

	C_W1		= 1
	C_W1_NM1	= 2
	C_W1_NM2	= 3
	C_W2		= 4
	C_W2_NM		= 5
	C_W3		= 6
	C_W3_NM		= 7
	C_W4		= 8
	C_W5		= 9
	C_W6		= 10
	C_W7		= 11

	C_W8		= 1
	C_W8_NM		= 2
	C_W9		= 3
	C_W10		= 4
	C_W11		= 5
	C_W12		= 6
	C_W13		= 7

	C_W14		= 1
	C_W15		= 2
	C_W16		= 3
	C_W17		= 4
	C_W18		= 5
	C_W19		= 6

	C_W20		= 1
	C_W20_NM	= 2
	C_W21		= 3
	C_W22		= 4
	C_W23		= 5
	C_W24		= 6
	
	' 열에 대한 번호 지정 
	C_101		= 1
	C_102		= 2
	C_103		= 3
	C_104		= 4
	C_105		= 5
	C_106		= 6
	C_107		= 7

	C_108		= 1
	C_109		= 2
	C_110		= 3

	C_111		= 1
	C_112		= 2
	C_113		= 3
	
End Sub

'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_47B4 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_47B3 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_47B2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords


    lgStrSQL =            "DELETE TB_47B1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iStrData3, iStrData4, iIntMaxRows, iLngRow
    Dim iDx, blnArrQry(3)
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	For iDx = 0 To 3
		blnArrQry(iDx) = False
	Next
	
    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        iStrData = "" : blnArrQry(0) = True
        
        Do While Not lgObjRs.EOF

			iStrData = iStrData & "	.Row = " & RIGHT(lgObjRs("W1"), 1) & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W2		& "	: .Text = """ & lgObjRs("W2") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W2_NM	& "	: .Text = """ & lgObjRs("W2_NM") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W3		& "	: .Text = """ & lgObjRs("W3") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W3_NM	& "	: .Text = """ & lgObjRs("W3_NM") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W4		& "	: .Value = """ & lgObjRs("W4") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W5		& "	: .Value = """ & lgObjRs("W5") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W6		& "	: .Value = """ & lgObjRs("W6") & """" & vbCrLf
			iStrData = iStrData & "	.Col = " & C_W7		& "	: .Value = """ & lgObjRs("W7") & """" & vbCrLf

		    lgObjRs.MoveNext

        Loop 
        
        lgObjRs.Close
        Set lgObjRs = Nothing
	End If
	
	' 두번째 쿼리 
	Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	iStrData2 = ""
	
	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	  
	     lgStrPrevKey = ""
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
	        
	Else
	    blnArrQry(1) = True
	        
	    Do While Not lgObjRs.EOF
	        
			iStrData2 = iStrData2 & "	.Row = " & CInt(RIGHT(lgObjRs("W8"), 2)) - 7 & vbCrLf
			iStrData2 = iStrData2 & "	.Col = " & C_W8		& "	: .Text = """ & lgObjRs("W8") & """" & vbCrLf
			iStrData2 = iStrData2 & "	.Col = " & C_W9		& "	: .Value = """ & lgObjRs("W9") & """" & vbCrLf
			iStrData2 = iStrData2 & "	.Col = " & C_W10	& "	: .Value = """ & lgObjRs("W10") & """" & vbCrLf
			iStrData2 = iStrData2 & "	.Col = " & C_W11	& "	: .Value = """ & lgObjRs("W11") & """" & vbCrLf
			iStrData2 = iStrData2 & "	.Col = " & C_W12	& "	: .Value = """ & lgObjRs("W12") & """" & vbCrLf
			iStrData2 = iStrData2 & "	.Col = " & C_W13	& "	: .Value = """ & lgObjRs("W13") & """" & vbCrLf

		    lgObjRs.MoveNext
	
	    Loop 
	        
	    lgObjRs.Close
	    Set lgObjRs = Nothing
	End If
	
	' 세번째 쿼리 
	Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	iStrData3 = ""
		
	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		  
	     lgStrPrevKey = ""
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		        
	Else
	    blnArrQry(2) = True
	    Do While Not lgObjRs.EOF
		        
			iStrData3 = iStrData3 & "	.Row = 1 " & vbCrLf
			iStrData3 = iStrData3 & "	.Col = " & C_W14	& "	: .Value = """ & lgObjRs("W14") & """" & vbCrLf
			iStrData3 = iStrData3 & "	.Col = " & C_W15	& "	: .Value = """ & lgObjRs("W15") & """" & vbCrLf
			iStrData3 = iStrData3 & "	.Col = " & C_W16	& "	: .Value = """ & lgObjRs("W16") & """" & vbCrLf
			iStrData3 = iStrData3 & "	.Col = " & C_W17	& "	: .Value = """ & lgObjRs("W17") & """" & vbCrLf
			iStrData3 = iStrData3 & "	.Col = " & C_W18	& "	: .Value = """ & lgObjRs("W18") & """" & vbCrLf
			iStrData3 = iStrData3 & "	.Col = " & C_W19	& "	: .Value = """ & lgObjRs("W19") & """" & vbCrLf

		    lgObjRs.MoveNext
		
	    Loop 
		        
	    lgObjRs.Close
	    Set lgObjRs = Nothing
	End If

	' 네번째 쿼리 
	Call SubMakeSQLStatements("R4",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	iStrData4 = ""

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

	     lgStrPrevKey = ""
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()

	Else
		blnArrQry(1) = True
	    Do While Not lgObjRs.EOF

			iStrData4 = iStrData4 & "	.Row = " & RIGHT(lgObjRs("W20"), 1) & vbCrLf
			iStrData4 = iStrData4 & "	.Col = " & C_W20	& "	: .Text = """ & lgObjRs("W20") & """" & vbCrLf
			iStrData4 = iStrData4 & "	.Col = " & C_W21	& "	: .Value = """ & lgObjRs("W21") & """" & vbCrLf
			iStrData4 = iStrData4 & "	.Col = " & C_W22	& "	: .Value = """ & lgObjRs("W22") & """" & vbCrLf
			iStrData4 = iStrData4 & "	.Col = " & C_W23	& "	: .Value = """ & lgObjRs("W23") & """" & vbCrLf
			iStrData4 = iStrData4 & "	.Col = " & C_W24	& "	: .Value = """ & lgObjRs("W24") & """" & vbCrLf

		    lgObjRs.MoveNext

	    Loop 

	    lgObjRs.Close
	    Set lgObjRs = Nothing
	End If

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
     

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " 		parent.InitSpreadRow  " & vbCr
	Response.Write " With parent.frm1.vspdData1  " & vbCr
	Response.Write 			iStrData & vbCrLf
	Response.Write " End With                                           " & vbCr
	Response.Write " With parent.frm1.vspdData2  " & vbCr
	Response.Write 			iStrData2 & vbCrLf
	Response.Write " End With                                           " & vbCr
	Response.Write " With parent.frm1.vspdData3  " & vbCr
	Response.Write 			iStrData3 & vbCrLf
	Response.Write " End With                                           " & vbCr
	Response.Write " With parent.frm1.vspdData4  " & vbCr
	Response.Write 			iStrData4 & vbCrLf
	Response.Write " End With                                           " & vbCr
	Response.Write " With parent                                        " & vbCr
    
    If blnArrQry(0) = True Or blnArrQry(1) = True Or blnArrQry(2) = True Or blnArrQry(3) = True Then
		Response.Write "	.DbQueryOk                                      " & vbCr
	End If
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
            lgStrSQL = lgStrSQL & " A.W1, A.W2 "
            lgStrSQL = lgStrSQL & " , CASE WHEN A.W1 <= '104' THEN dbo.ufn_GetCodeName('W1057', A.W2) "
            lgStrSQL = lgStrSQL & "			WHEN A.W1 IN ('105', '106') THEN dbo.ufn_GetCodeName('W1058', A.W2) ELSE '' END AS W2_NM "
            lgStrSQL = lgStrSQL & " , A.W3, CASE WHEN A.W1 <= '104' THEN dbo.ufn_GetCodeName('W1057', A.W3) "
            lgStrSQL = lgStrSQL & "			WHEN A.W1 IN ('105', '106') THEN dbo.ufn_GetCodeName('W1058', A.W3) ELSE '' END AS W3_NM "
            lgStrSQL = lgStrSQL & " , A.W4, A.W5, A.W6, A.W7 "
            lgStrSQL = lgStrSQL & " FROM TB_47B1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.W1 ASC" & vbcrlf

      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W8, A.W9, A.W10, A.W11, A.W12, A.W13"
            lgStrSQL = lgStrSQL & " FROM TB_47B2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.W8 ASC" & vbcrlf

      Case "R3"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W14, A.W15, A.W16, A.W17, A.W18, A.W19 "
            lgStrSQL = lgStrSQL & " FROM TB_47B3 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

      Case "R4"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W20, A.W21, A.W22, A.W23, A.W24 "
            lgStrSQL = lgStrSQL & " FROM TB_47B4 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.W20 ASC" & vbcrlf

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
    
	PrintLog "1번째 그리드. .: " & Request("txtSpread1") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread1"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
		lgIntFlgMode = CInt(Request("txtFlgMode"))

        Select Case lgIntFlgMode
	        Case  OPMD_CMODE                                                             '☜ : Create
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
	        Case  OPMD_UMODE   
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

	PrintLog "2번째 그리드. .: " & Request("txtSpread2") 
	' --- 2번째 그리드 
	arrRowVal = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
		lgIntFlgMode = CInt(Request("txtFlgMode"))

        Select Case lgIntFlgMode
	        Case  OPMD_CMODE                                                             '☜ : Create
                    Call SubBizSaveMultiCreate2(arrColVal)                            '☜: Create
	        Case  OPMD_UMODE   
                    Call SubBizSaveMultiUpdate2(arrColVal)                            '☜: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
    
	PrintLog "3번째 그리드. .: " & Request("txtSpread3") 
	' --- 3번째 그리드 
	arrRowVal = Split(Request("txtSpread3"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
		lgIntFlgMode = CInt(Request("txtFlgMode"))

        Select Case lgIntFlgMode
	        Case  OPMD_CMODE                                                             '☜ : Create
                    Call SubBizSaveMultiCreate3(arrColVal)                            '☜: Create
	        Case  OPMD_UMODE   
                    Call SubBizSaveMultiUpdate3(arrColVal)                            '☜: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
    
	PrintLog "4번째 그리드. .: " & Request("txtSpread4") 
	' --- 4번째 그리드 
	arrRowVal = Split(Request("txtSpread4"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
		lgIntFlgMode = CInt(Request("txtFlgMode"))

        Select Case lgIntFlgMode
	        Case  OPMD_CMODE                                                             '☜ : Create
                    Call SubBizSaveMultiCreate4(arrColVal)                            '☜: Create
	        Case  OPMD_UMODE   
                    Call SubBizSaveMultiUpdate4(arrColVal)                            '☜: Update
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
	
	lgStrSQL = "INSERT INTO TB_47B1 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, W5, W6, W7 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D")		& ","

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
' Name : SubBizSaveCreate2
' Desc : 2번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_47B2 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W8, W9, W10, W11, W12, W13 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveCreate3
' Desc : 3번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate3(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_47B3 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W14, W15, W16, W17, W18, W19 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate3 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveCreate4
' Desc : 4번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate4(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_47B4 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W20, W21, W22, W23, W24 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W20))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate4 = " & lgStrSQL

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

    lgStrSQL = "UPDATE  TB_47B1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_W3 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D") & ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND W1 = " & FilterVar(Trim(UCase(arrColVal(C_W1) )),"''","S")

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate2
' Desc : 2번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_47B2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W11		= " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W12		= " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W13		= " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND W8 = " & FilterVar(Trim(UCase(arrColVal(C_W8) )),"''","S")

	PrintLog "SubBizSaveMultiUpdate2 = " & lgStrSQL

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate3
' Desc : 3번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate3(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_47B3 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W14		= " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W15		= " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W16		= " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W17		= " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W18		= " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W19		= " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf

	PrintLog "SubBizSaveMultiUpdate3 = " & lgStrSQL

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate4
' Desc : 4번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate4(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_47B4 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W21		= " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W22		= " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W23		= " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W24		= " &  FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D") & ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND W20 = " & FilterVar(Trim(UCase(arrColVal(C_W20) )),"''","S")

	PrintLog "SubBizSaveMultiUpdate4 = " & lgStrSQL

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
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SU"
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
