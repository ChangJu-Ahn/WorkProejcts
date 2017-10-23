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
   
	Const BIZ_MNU_ID = "W3105MA1"
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Dim C_SEQ_NO1
	Dim C_W16
	Dim C_W16_NM
	Dim C_W17
	Dim C_W18
	Dim C_W19
	Dim C_W20
	Dim C_W21
	Dim C_DESC1

	Dim C_SEQ_NO2
	Dim C_W22
	Dim C_W23
	Dim C_W23_NM
	Dim C_W24
	Dim C_W25
	Dim C_W26
	Dim C_W27
	Dim C_W28
	Dim C_W29
	Dim C_W30
	Dim C_W31
	Dim C_W32
	Dim C_DESC2

	lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR = Request("txtFISC_YEAR")
    sREP_TYPE = Request("cboREP_TYPE")

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
	C_SEQ_NO1	= 1	' -- 1번 그리드 
    C_W16		= 2
    C_W16_NM	= 3
    C_W17		= 4
    C_W18		= 5
    C_W19		= 6
    C_W20		= 7
    C_W21		= 8	
    C_DESC1		= 9
 
 	C_SEQ_NO2	= 1  ' -- 2번 그리드 
    C_W22		= 2 
    C_W23		= 3
    C_W23_NM	= 4
    C_W24		= 5
    C_W25		= 6
    C_W26		= 7
    C_W27		= 8
    C_W28		= 9
    C_W29		= 10
    C_W30		= 11
    C_W31		= 12
    C_W32		= 13
    C_DESC2		= 14
End Sub

'========================================================================================
'Sub SubBizQuery()
    'On Error Resume Next
 '   Err.Clear
'End Sub
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
    lgStrSQL =            "DELETE TB_34D1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_34D2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

    lgStrSQL = lgStrSQL & "DELETE TB_34H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 


 
	
PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	Call TB_15_DeleData("1", 0)
	Call TB_15_DeleData("2", 1)
	Call TB_15_DeleData("2", -1)
	Call TB_15_DeleData("3", 0)
 	
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
        Exit Sub
        
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        
        iDx = 1
        
        Do While Not lgObjRs.EOF

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write " 	.IsRunEvents = True                            " & vbCr
    Response.Write "	.frm1.txtW1.value = """ & ConvSPChars(lgObjRs("W1"))          & """" & vbCr
    Response.Write "	.frm1.txtW1_BEFORE.value = """ & ConvSPChars(lgObjRs("W1_BEFORE"))          & """" & vbCr
    Response.Write "	.frm1.txtW2_1_1.value = """ & ConvSPChars(lgObjRs("W2_1_1"))          & """" & vbCr
    Response.Write "	.frm1.txtW2_1_2.value = """ & ConvSPChars(lgObjRs("W2_1_2"))          & """" & vbCr
    Response.Write "	.frm1.txtW2_2.value = """ & ConvSPChars(lgObjRs("W2_2"))          & """" & vbCr
    Response.Write "	.frm1.txtW2_3.value = """ & ConvSPChars(lgObjRs("W2_3"))          & """" & vbCr
    Response.Write "	.frm1.txtW3.value = """ & ConvSPChars(lgObjRs("W3"))          & """" & vbCr
    Response.Write "	.frm1.txtW4.value = """ & ConvSPChars(lgObjRs("W4"))          & """" & vbCr
    Response.Write "	.frm1.txtW4_BEFORE.value = """ & ConvSPChars(lgObjRs("W4_BEFORE"))          & """" & vbCr
    Response.Write "	.frm1.txtW5.value = """ & ConvSPChars(lgObjRs("W5"))          & """" & vbCr
    Response.Write "	.frm1.txtW6.value = """ & ConvSPChars(lgObjRs("W6"))          & """" & vbCr
    Response.Write "	.frm1.txtW7.value = """ & ConvSPChars(lgObjRs("W7"))          & """" & vbCr
    Response.Write "	.frm1.txtW8.value = """ & ConvSPChars(lgObjRs("W8"))          & """" & vbCr
    Response.Write "	.frm1.txtW9.value = """ & ConvSPChars(lgObjRs("W9"))          & """" & vbCr
    Response.Write "	.frm1.txtW10.value = """ & ConvSPChars(lgObjRs("W10"))          & """" & vbCr
    Response.Write "	.frm1.txtW10_BEFORE.value = """ & ConvSPChars(lgObjRs("W10_BEFORE"))          & """" & vbCr
    Response.Write "	.frm1.txtW11.value = """ & ConvSPChars(lgObjRs("W11"))          & """" & vbCr
    Response.Write "	.frm1.txtW12.value = """ & ConvSPChars(lgObjRs("W12"))          & """" & vbCr
    Response.Write "	.frm1.txtW13.value = """ & ConvSPChars(lgObjRs("W13"))          & """" & vbCr
    Response.Write "	.frm1.txtW14.value = """ & ConvSPChars(lgObjRs("W14"))          & """" & vbCr
    Response.Write "	.frm1.txtW15.value = """ & ConvSPChars(lgObjRs("W15"))          & """" & vbCr
	Response.Write " 	.IsRunEvents = False                            " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr
   
		    lgObjRs.MoveNext

        Loop 
        
        lgObjRs.Close
        Set lgObjRs = Nothing

 
         ' 1번째 그리드 
        Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
'		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
'		    Call SetErrorStatus()
		    
		Else
		    Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W16"))			
				iStrData = iStrData & Chr(11) 	' W16_BT
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W16_NM"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W17"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W18"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W19"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W20"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W21"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("DESC1"))			 
				iStrData = iStrData & Chr(11) & iDx
				iStrData = iStrData & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
		        If iDx > C_SHEETMAXROWS_D Then
		           lgStrPrevKey = lgStrPrevKey + 1
		           Exit Do
		        End If               
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
			Response.Write " End With                                           " & vbCr
			Response.Write " </Script>                                          " & vbCr
		    
		End If  
		       
        ' 2번째 그리드 
        Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
'		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
'		    Call SetErrorStatus()
		    
		Else
		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData2 = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W22"))			
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W23"))	
				iStrData2 = iStrData2 & Chr(11) 	' W23_BT
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W23_NM"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W24"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W25"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W26"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W27"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W28"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W29"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W30"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W31"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W32"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("DESC2"))
				iStrData2 = iStrData2 & Chr(11) & iDx
				iStrData2 = iStrData2 & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
		        If iDx > C_SHEETMAXROWS_D Then
		           lgStrPrevKey = lgStrPrevKey + 1
		           Exit Do
		        End If               
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & iStrData2       & """" & vbCr	
			Response.Write " End With                                           " & vbCr
			Response.Write " </Script>                                          " & vbCr
		    
		End If        
    End If
    
     Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
     
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT TOP 1 "
            lgStrSQL = lgStrSQL & " A.W1, A.W1_BEFORE, A.W2_1_1, A.W2_1_2, W2_2, A.W2_3, A.W3, A.W4, A.W4_BEFORE "
            lgStrSQL = lgStrSQL & " , A.W5, A.W6, A.W7, A.W8, A.W9, A.W10, A.W10_BEFORE, A.W11, A.W12, A.W13, A.W14, A.W15 "

            lgStrSQL = lgStrSQL & " FROM TB_34H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W16, A.W16_NM, A.W17, W18, A.W19, A.W20, A.W21, A.DESC1 "

            lgStrSQL = lgStrSQL & " FROM TB_34D1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO" & vbcrlf
            
      Case "R3"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W22, A.W23, A.W23_NM, A.W24, A.W25, A.W26, A.W27 "
            lgStrSQL = lgStrSQL & " , A.W28, A.W29, A.W30, A.W31, A.W32, A.DESC2 "

            lgStrSQL = lgStrSQL & " FROM TB_34D2 A WITH (NOLOCK) "
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

	PrintLog "SubBizSaveMulti.."
	
    'On Error Resume Next
    Err.Clear 
    
    ' 헤더 저장 
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
	PrintLog "lgIntFlgMode = " & lgIntFlgMode
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select

	PrintLog "1번째 그리드. .: " & Request("txtSpread") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
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
	PrintLog "2번째 그리드.. : " & Request("txtSpread2")
	
	' --- 2번째 그리드 
	arrRowVal = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate2(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate2(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete2(arrColVal)                            '☜: Delete
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
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_34H WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, W1, W1_BEFORE "
    lgStrSQL = lgStrSQL & "  , W2_1_1, W2_1_2, W2_2, W2_3, W3, W4, W4_BEFORE "
    lgStrSQL = lgStrSQL & "  , W5, W6, W7, W8, W9, W10, W10_BEFORE "
    lgStrSQL = lgStrSQL & "  , W11, W12, W13, W14, W15 "
    lgStrSQL = lgStrSQL & "  , INSRT_USER_ID, UPDT_USER_ID ) " 
    lgStrSQL = lgStrSQL & " VALUES ( " 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","    
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1_BEFORE"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2_1_1"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW2_1_2"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2_2"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2_3"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4_BEFORE"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW10_BEFORE"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""
       
    lgStrSQL = lgStrSQL & "   ) " 

	PrintLog "SubBizSaveSingleCreate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	'	⑦ 한도초과액				Max ( ( ⑥ 계- ③ 한도액금액) , 0 ) 을 계산하여 입력함.								
	'			동 금액이  "0" 이 아닌경우 15-1호 서식에 (1)과목명 "대손충당금" (2) 금액은 동 금액								
	'			(3)소득처분에는 "유보(증가)"를 입력하고,								
	'			조정내용은 "대손충당금 한도초과액을 손금불산입하고 유보처분함."을 입력하고 경고하여 줌.	
	If UNICDbl(Request("txtW7"), "0") <> 0 Then
		Call TB_15_PushData("1", UNICDbl(Request("txtW7"), "0") )
	End If

	'	⑮과소환입·과다환입은 (13)에서 (14)을 차감하여 입력함.											
	'	  이 경우 (-)음수가 나오면 15-2호에 (1)과목명 "대손충당금" (2)금액에 동 금액 (3) 소득처분에 "유보(감소)"를 입력하고,											
	'	  조정내용은 "대손충당금 전기부인액을 손금산입하고 유보처분함"을 입력하고 경고하여줌.											
												
	'	(+)양수가 나오면 15-1호에 (1)과목명 "대손충당금" (2)금액에 동 금액의 절대값 (3) 소득처분에 "유보(증가)"를 입력함.											
	'	  조정내용은 "대손충당금 과다환입액을 익금불산입 하고 유보처분함"을 입력하고 경고하여줌.											
	If UNICDbl(Request("txtW15"), "0") <> 0 Then
	Call TB_15_PushData("2", UNICDbl(Request("txtW15"), "0") )
	End If

End Sub   

'============================================================================================================
' Name : TB_15_PushData
' Desc : 	'(1)	⑦ 한도초과액				Max ( ( ⑥ 계- ③ 한도액금액) , 0 ) 을 계산하여 입력함.								
	'			동 금액이  "0" 이 아닌경우 15-1호 서식에 (1)과목명 "대손충당금" (2) 금액은 동 금액								
	'			(3)소득처분에는 "유보(증가)"를 입력하고,								
	'			조정내용은 "대손충당금 한도초과액을 손금불산입하고 유보처분함."을 입력하고 경고하여 줌.	
	'(2)	⑮과소환입·과다환입은 (13)에서 (14)을 차감하여 입력함.											
	'	  이 경우 (-)음수가 나오면 15-2호에 (1)과목명 "대손충당금" (2)금액에 동 금액 (3) 소득처분에 "유보(감소)"를 입력하고,											
	'	  조정내용은 "대손충당금 전기부인액을 손금산입하고 유보처분함"을 입력하고 경고하여줌.											
	'	(+)양수가 나오면 15-1호에 (1)과목명 "대손충당금" (2)금액에 동 금액의 절대값 (3) 소득처분에 "유보(증가)"를 입력함.											
	'	  조정내용은 "대손충당금 과다환입액을 익금불산입 하고 유보처분함"을 입력하고 경고하여줌.											

	'(3)	저장시에 (29)부인액과 (32)부인액의 합을 15-1호에 (1)과목명에 대손금 (3)소득처분 유보(증가)로 하여 각각 입력하고, 
	'	조정내용은 "당기에 대손처리한 채권중 대손요건이 미비된 채권을 손금불산입하고 유보처분 함"을입력함.
'============================================================================================================
Sub TB_15_PushData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	
	Select Case pSeqNo
		Case "1"
			lgStrSQL = "EXEC usp_TB_15_PushData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("3401")),"''","S") & ", "		' 과목 코드 
			lgStrSQL = lgStrSQL & FilterVar(pAmt,"0","D")  & ", "			' 금액 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("400")),"''","S") & ", "			' 처분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("대손충당금 한도초과액을 손금불산입하고 유보처분함.")),"''","S") & ", "			' 조정내용 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
			
	

		Case "2"
			lgStrSQL = "EXEC usp_TB_15_PushData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			If pAmt > 0 Then
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("3401")),"''","S") & ", "		' 과목 코드 
				lgStrSQL = lgStrSQL & FilterVar(pAmt,"0","D")  & ", "			' 금액 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("400")),"''","S") & ", "			' 처분 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("대손충당금 과다환입액을 익금불산입 하고 유보처분함.")),"''","S") & ", "			' 조정내용 
			ElseIf pAmt < 0 Then
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2")),"''","S") & ", "			' 2호 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("3401")),"''","S") & ", "		' 과목 코드 
				lgStrSQL = lgStrSQL & FilterVar(pAmt * -1,"0","D")  & ", "			' 금액 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("100")),"''","S") & ", "			' 처분 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("대손충당금 전기부인액을 손금산입하고 유보처분함.")),"''","S") & ", "			' 조정내용 
			End If
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
			If pAmt = 0 Then lgStrSQL = ""
		Case "3"
			lgStrSQL = "EXEC usp_TB_15_PushData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("3402")),"''","S") & ", "		' 과목 코드 
			lgStrSQL = lgStrSQL & FilterVar(pAmt,"0","D")  & ", "			' 금액 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("400")),"''","S") & ", "			' 처분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("당기에 대손처리한 채권중 대손요건이 미비된 채권을 손금불산입하고 유보처분 함.")),"''","S") & ", "			' 조정내용 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	End Select

    	PrintLog "TB_15_PushData = " & lgStrSQL

	If lgStrSQL <> "" Then
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If	
End Sub


'============================================================================================================
' Name : TB_15_DeleData
' Desc : 	'(1)	⑦ 한도초과액				Max ( ( ⑥ 계- ③ 한도액금액) , 0 ) 을 계산하여 입력함.								
	'			동 금액이  "0" 이 아닌경우 15-1호 서식에 (1)과목명 "대손충당금" (2) 금액은 동 금액								
	'			(3)소득처분에는 "유보(증가)"를 입력하고,								
	'			조정내용은 "대손충당금 한도초과액을 손금불산입하고 유보처분함."을 입력하고 경고하여 줌.	
	'(2)	⑮과소환입·과다환입은 (13)에서 (14)을 차감하여 입력함.											
	'	  이 경우 (-)음수가 나오면 15-2호에 (1)과목명 "대손충당금" (2)금액에 동 금액 (3) 소득처분에 "유보(감소)"를 입력하고,											
	'	  조정내용은 "대손충당금 전기부인액을 손금산입하고 유보처분함"을 입력하고 경고하여줌.											
	'	(+)양수가 나오면 15-1호에 (1)과목명 "대손충당금" (2)금액에 동 금액의 절대값 (3) 소득처분에 "유보(증가)"를 입력함.											
	'	  조정내용은 "대손충당금 과다환입액을 익금불산입 하고 유보처분함"을 입력하고 경고하여줌.											

	'(3)	저장시에 (29)부인액과 (32)부인액의 합을 15-1호에 (1)과목명에 대손금 (3)소득처분 유보(증가)로 하여 각각 입력하고, 
	'	조정내용은 "당기에 대손처리한 채권중 대손요건이 미비된 채권을 손금불산입하고 유보처분 함"을입력함.
'============================================================================================================
Sub TB_15_DeleData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	
	Select Case pSeqNo
		Case "1"
			lgStrSQL = "EXEC usp_TB_15_DeleData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

		Case "2"
			lgStrSQL = "EXEC usp_TB_15_DeleData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			If pAmt > 0 Then
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2")),"''","S") & ", "			' 1호 
			ElseIf pAmt < 0 Then
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 2호 
			End If
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
			If pAmt = 0 Then lgStrSQL = ""
		Case "3"
			lgStrSQL = "EXEC usp_TB_15_DeleData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	End Select

	PrintLog "TB_15_DeleData = " & lgStrSQL
    If lgStrSQL <> "" Then
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If	
		
End Sub


'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_34H WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2_1_1 = " & FilterVar(UNICDbl(Request("txtW2_1_1"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2_1_2 = " & FilterVar(Request("txtW2_1_2"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2_2 = " & FilterVar(UNICDbl(Request("txtW2_2"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2_3 = " & FilterVar(UNICDbl(Request("txtW2_3"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4_BEFORE = " & FilterVar(UNICDbl(Request("txtW4_BEFORE"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W7 = " & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W8 = " & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W9 = " & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W10 = " & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W10_BEFORE = " & FilterVar(UNICDbl(Request("txtW10_BEFORE"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W11 = " & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W12 = " & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W13 = " & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W14 = " & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W15 = " & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "		  UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
 
	PrintLog "SubBizSaveSingleUpdate = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	'	⑦ 한도초과액				Max ( ( ⑥ 계- ③ 한도액금액) , 0 ) 을 계산하여 입력함.								
	'			동 금액이  "0" 이 아닌경우 15-1호 서식에 (1)과목명 "대손충당금" (2) 금액은 동 금액								
	'			(3)소득처분에는 "유보(증가)"를 입력하고,								
	'			조정내용은 "대손충당금 한도초과액을 손금불산입하고 유보처분함."을 입력하고 경고하여 줌.	
	If UNICDbl(Request("txtW7"), "0") <> 0 Then
		Call TB_15_PushData("1", UNICDbl(Request("txtW7"), "0") )
	ElseIf UNICDbl(Request("txtW7"), "0") <> 0 Then
		Call TB_15_DeleData("1", 0 )
	End If

	'	⑮과소환입·과다환입은 (13)에서 (14)을 차감하여 입력함.											
	'	  이 경우 (-)음수가 나오면 15-2호에 (1)과목명 "대손충당금" (2)금액에 동 금액 (3) 소득처분에 "유보(감소)"를 입력하고,											
	'	  조정내용은 "대손충당금 전기부인액을 손금산입하고 유보처분함"을 입력하고 경고하여줌.											
												
	'	(+)양수가 나오면 15-1호에 (1)과목명 "대손충당금" (2)금액에 동 금액의 절대값 (3) 소득처분에 "유보(증가)"를 입력함.											
	'	  조정내용은 "대손충당금 과다환입액을 익금불산입 하고 유보처분함"을 입력하고 경고하여줌.											

	Call TB_15_DeleData("2", UNICDbl(Request("txtW15"), "0") )
	If UNICDbl(Request("txtW15"), "0") <> 0 Then
	Call TB_15_PushData("2", UNICDbl(Request("txtW15"), "0") )
	End If
End Sub    

     
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_34D1 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W16, W16_NM, W17, W18, W19, W20, W21, DESC1 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"1","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W16))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W16_NM))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_DESC1))),"''","S")		& ","

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
' Name : SubBizSaveCreate
' Desc : 2번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_34D2 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W22, W23, W23_NM, W24, W25, W26, W27, W28 "  
	lgStrSQL = lgStrSQL & " , W29, W30, W31, W32, DESC2 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W22), ""),"NULL","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W23))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W23_NM))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W24))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W25))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_DESC2))),"''","S")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	'	저장시에 (29)부인액과 (32)부인액의 합을 15-1호에 (1)과목명에 대손금 (3)소득처분 유보(증가)로 하여 각각 입력하고, 
	'	조정내용은 "당기에 대손처리한 채권중 대손요건이 미비된 채권을 손금불산입하고 유보처분 함"을입력함.
	If UNICDbl(arrColVal(C_SEQ_NO2), "0") = 999999 Then
		Call TB_15_PushData("3", UNICDbl(arrColVal(C_W29), "0") + UNICDbl(arrColVal(C_W32), "0") )
	End If

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_34D1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W16	= " &  FilterVar(Trim(UCase(arrColVal(C_W16))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W16_NM	= " &  FilterVar(Trim(UCase(arrColVal(C_W16_NM))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W17 = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W18 = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W19 = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W20 = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W21 = " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " DESC1 = " &  FilterVar(Trim(UCase(arrColVal(C_DESC1))),"''","S") & ","
                   
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"0","D")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 2번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_34D2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W22	= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W22), ""),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W23	= " &  FilterVar(Trim(UCase(arrColVal(C_W23))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W23_NM	= " &  FilterVar(Trim(UCase(arrColVal(C_W23_NM))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W24 = " &  FilterVar(Trim(UCase(arrColVal(C_W24))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W25 = " &  FilterVar(Trim(UCase(arrColVal(C_W25))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W26 = " &  FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W27 = " &  FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W28 = " &  FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W29 = " &  FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W30 = " &  FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W31 = " &  FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W32 = " &  FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " DESC2 = " &  FilterVar(Trim(UCase(arrColVal(C_DESC2))),"''","S") & ","
                       
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"0","D")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf

	'	저장시에 (29)부인액과 (32)부인액의 합을 15-1호에 (1)과목명에 대손금 (3)소득처분 유보(증가)로 하여 각각 입력하고, 
	'	조정내용은 "당기에 대손처리한 채권중 대손요건이 미비된 채권을 손금불산입하고 유보처분 함"을입력함.
	If UNICDbl(arrColVal(C_SEQ_NO2), "0") = 999999 Then
		Call TB_15_PushData("3", UNICDbl(arrColVal(C_W29), "0") + UNICDbl(arrColVal(C_W32), "0") )
	End If

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_34D1 WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"0","D")  
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)


End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_34D2 WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"0","D")  
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	If UNICDbl(arrColVal(C_SEQ_NO2), "0") = 999999 Then
		Call TB_15_DeleData("3", 0 )
	End If

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
 