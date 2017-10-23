<%@ Transaction=required CODEPAGE=949 Language=VBScript%>
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
   
	Const BIZ_MNU_ID = "W4101MA1"
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
	Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.
	Const TYPE_3	= 2		
		
	Dim C_SEQ_NO	
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
	
	C_SEQ_NO	= 1	' -- 1번 그리드 
    C_W10		= 2	' 계정과목 
    C_W11		= 3 ' 금액 
    C_W12		= 4	' 감가상각누계액 
    C_W13		= 5	' 상각부인누계액 
    C_W14		= 6	' 가감계 
    C_W15		= 7	' 운휴설비가액 
    C_W16		= 8	' 가동설비가액 

	' C_SEQ_NO 포함 
	C_W17		= 2	' 설절액 
	C_W18		= 3	' 설정액 
	C_W19		= 4	' 장부상준비금 
	C_W20		= 5 ' 기중준비금 
	C_W21		= 6	' 준비금 
	C_W22		= 7	' 개체소요자금상당액 
	C_W23		= 8	' 미사용분 
	C_W24		= 9	' 개체소요자금상당액 
	C_W25		= 10 ' 기타 
	C_W26		= 11 ' 계 
	
	' C_SEQ_NO, C_W17 포함 
	C_W27		= 3	' 1차연도 
	C_W28		= 4	' 2차연도 
	C_W29		= 5	' 3차년도 
	C_W30		= 6 ' 계 
	C_W31		= 7	' 환입할금액합계 
	C_W32		= 8	' 회사환입액 
	C_W33		= 9	' 과소환입 
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
    lgStrSQL =            "DELETE TB_31_1D3 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_31_1D2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
    lgStrSQL = lgStrSQL & "DELETE TB_31_1D1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgStrSQL = lgStrSQL & "DELETE TB_31_1H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
 	' -- 15에 삭제 
 	Call TB_15_DeleData("", -1)
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

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1
        
		arrRs(iRow) = ""

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent.frm1                                   " & vbCr
		Response.Write "	parent.IsRunEvents = True " & vbCr
		Response.Write "	.txtW1.value = """ & lgObjRs("W1") & """" & vbCrLf
		Response.Write "	.txtW2.value = """ & lgObjRs("W2") & """"  & vbCrLf
		Response.Write "	.txtW2_VAL.value = """ & lgObjRs("W2_VAL") & """"  & vbCrLf
		Response.Write "	.txtW3.value = """ & lgObjRs("W3") & """"  &  vbCrLf
		Response.Write "	.txtW3_VAL.value = """ & lgObjRs("W3_VAL") & """"  &  vbCrLf
		Response.Write "	.txtW4.value = """ & lgObjRs("W4") & """"  &  vbCrLf
		Response.Write "	.txtW5.value = """ & lgObjRs("W5") & """"  &  vbCrLf
		Response.Write "	.txtW6.value = """ & lgObjRs("W6") & """"  &  vbCrLf
		Response.Write "	.txtW7.value = """ & lgObjRs("W7") & """"  &  vbCrLf
		Response.Write "	.txtW8.value = """ & lgObjRs("W8") & """"  &  vbCrLf
		Response.Write "	.txtW9.value = """ & lgObjRs("W9") & """"  &  vbCrLf
		Response.Write "	.txtDESC1.value = """ & lgObjRs("DESC1") & """"  &  vbCrLf
		Response.Write "	parent.IsRunEvents = False " & vbCr
		Response.Write " End With                                  " & vbCr
		Response.Write " </Script>	                        " & vbCr
		
		' 1번 그리드 
	    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then
  
		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1 : iRow = TYPE_1
		    
			arrRs(iRow) = ""
				
			Do While Not lgObjRs.EOF
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W10"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W11")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W12")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W13")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W14")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W15")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W16")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
				arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

		End If
		
		' 2번 그리드 
	    Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then
  
		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1 : iRow = TYPE_2
		    
			arrRs(iRow) = ""
				
			Do While Not lgObjRs.EOF
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W17"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W18")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W19")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W20")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W21")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W22")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W23")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W24")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W25")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W26")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
				arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

		End If
			
		' 3번 그리드 
	    Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1 : iRow = TYPE_3
		    
			arrRs(iRow) = ""
				
			Do While Not lgObjRs.EOF
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W17"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W27")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W28")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W29")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W30")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W31")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W32")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W33")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
				arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

		End If


		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_1 & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_1)       & """" & vbCr

		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2 & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_2)       & """" & vbCr

		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_3 & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_3)       & """" & vbCr				

	
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
			lgStrSQL =			  " SELECT  TOP 1 "
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W2_VAL, A.W3, A.W3_VAL, A.W4, A.W5, A.W6 "
            lgStrSQL = lgStrSQL & " , A.W7, A.W8, A.W9, A.DESC1 "
            lgStrSQL = lgStrSQL & " FROM TB_31_1H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W10, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16 "
            lgStrSQL = lgStrSQL & " FROM TB_31_1D1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO" & vbcrlf
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W17, A.W18, A.W19, A.W20, A.W21, A.W22, A.W23, A.W24, A.W25, A.W26 "
            lgStrSQL = lgStrSQL & " FROM TB_31_1D2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf
      Case "R3"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W17, A.W27, A.W28, A.W29, A.W30, A.W31, A.W32, A.W33 "
            lgStrSQL = lgStrSQL & " FROM TB_31_1D3 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf
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
    
    ' 헤더 저장 
    If CDbl(Request("txtHeadMode")) = OPMD_CMODE Then
		Call SubBizSaveCreate
	Else
		Call SubBizSaveUpdate
	End If
	
	' 1번 그리드 
	PrintLog "txtSpread1 = " & Request("txtSpread" & CStr(TYPE_1))
		
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

	' 2번 그리드 
	PrintLog "txtSpread2 = " & Request("txtSpread" & CStr(TYPE_2))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_2) ), gRowSep)                                 '☜: Split Row    data
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

	' 3번 그리드 
	PrintLog "txtSpread3 = " & Request("txtSpread" & CStr(TYPE_3))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_3) ), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
	For iDx = 1 To lgLngMaxRow

	    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
	    Select Case arrColVal(0)
	        Case "C"
	                Call SubBizSaveMultiCreate3(arrColVal)                            '☜: Create
	        Case "U"
	                Call SubBizSaveMultiUpdate3(arrColVal)                            '☜: Update
	        Case "D"
	                Call SubBizSaveMultiDelete3(arrColVal)                            '☜: Delete
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
Sub SubBizSaveCreate()
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_31_1H WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , W1, W2, W2_VAL, W3, W3_VAL, W4, W5, W6, W7, W8, W9, DESC1 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtW2"))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2_VAL"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtW3"))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW3_VAL"), "0"),"0","D")     & "," & vbCrLf

	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtDESC1"))),"''","S")     & "," & vbCrLf	
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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_31_1D1 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W10, W11, W12, W13, W14, W15, W16 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")     & "," & vbCrLf

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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_31_1D2 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W17, W18, W19, W20, W21, W22, W23, W24, W25, W26" & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate3(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	PrintLog "C_W27=" & arrColVal(C_W27)
	lgStrSQL = "INSERT INTO TB_31_1D3 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W17, W27, W28, W29, W30, W31, W32, W33 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate3 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	' -- 15호 푸쉬 
'	양수인 경우 15-1호에 (1)과목명 "중소기업투자준비금" (2)금액에는 같은 금액을 					
'	(3)소득처분에는 "유보"를 입력하고 조정내용은 " 중소기업투자준비금 과소환입액을 익금산입하고 					
'	유보처분함."을 입력하고 경고하여줌.					
'						
'	음수인 경우 15-2호에 (1)과목명 "중소기업투자준비금" (2)금액에는 계산된 금액의 절대값을 					
'	(3)소득처분에는 "유보"를 입력하고 조정내용은 " 중소기업투자준비금 과다환입액을 익금불산입하고 					
'	유보처분함."을 입력하고 경고하여줌.	

	If UNICDbl(arrColVal(C_SEQ_NO), 0) = 999999 Then
		Call TB_15_DeleData("", -1)
		If UNICDbl(arrColVal(C_W33), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W33), 0), 999999, "3101", "400", "중소기업투자준비금 과소환입액을 익금산입하고 유보처분함")
		ElseIf UNICDbl(arrColVal(C_W33), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W33), 0)), 999999, "3101", "100", "중소기업투자준비금 과다환입액을 익금불산입하고 유보처분함")
		End If
	End If
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveUpdate()
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_31_1H WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    'lgStrSQL = lgStrSQL & " W1     = " &  FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W1     = " &  FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2     = " &  FilterVar(Trim(UCase(Request("txtW2"))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2_VAL = " &  FilterVar(UNICDbl(Request("txtW2_VAL"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3     = " &  FilterVar(Trim(UCase(Request("txtW3"))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3_VAL = " &  FilterVar(UNICDbl(Request("txtW3_VAL"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5     = " &  FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W8     = " &  FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W9     = " &  FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " DESC1  = " &  FilterVar(Trim(UCase(Request("txtDESC1"))),"''","S")     & "," & vbCrLf	
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	'lgStrSQL = lgStrSQL & "		AND W17 = " & FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_31_1D1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W10     = " &  FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11     = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W15     = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W16     = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & "," & vbCrLf
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
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_31_1D2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W17     = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W18     = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W19     = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W20     = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W21     = " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W22     = " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W23     = " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W24     = " &  FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W25     = " &  FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W26     = " &  FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate2 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate3(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_31_1D3 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W17     = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W27     = " &  FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W28     = " &  FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W29     = " &  FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W30     = " &  FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W31     = " &  FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W32     = " &  FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W33     = " &  FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate3 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	' -- 15호 푸쉬 
'	양수인 경우 15-1호에 (1)과목명 "중소기업투자준비금" (2)금액에는 같은 금액을 					
'	(3)소득처분에는 "유보"를 입력하고 조정내용은 " 중소기업투자준비금 과소환입액을 익금산입하고 					
'	유보처분함."을 입력하고 경고하여줌.					
'						
'	음수인 경우 15-2호에 (1)과목명 "중소기업투자준비금" (2)금액에는 계산된 금액의 절대값을 					
'	(3)소득처분에는 "유보"를 입력하고 조정내용은 " 중소기업투자준비금 과다환입액을 익금불산입하고 					
'	유보처분함."을 입력하고 경고하여줌.	
	If UNICDbl(arrColVal(C_SEQ_NO), 0) = 999999 Then
		Call TB_15_DeleData("", -1)
		If UNICDbl(arrColVal(C_W33), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W33), 0), 999999, "3101", "400", "중소기업투자준비금 과소환입액을 익금산입하고 유보처분함")
		ElseIf UNICDbl(arrColVal(C_W33), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W33), 0)), 999999, "3101", "100", "중소기업투자준비금 과다환입액을 익금불산입하고 유보처분함")
		End If
	End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_31_1D1 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_31_1D2 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete3(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_31_1D3 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	' -- 15에 삭제 
	If UNICDbl(arrColVal(C_SEQ_NO), 0) = 999999 Then
 		Call TB_15_DeleData("", -1)
 	End If
End Sub

'============================================================================================================
' Name : 15호서식에 푸쉬 
' Desc :  
'============================================================================================================
Sub TB_15_PushData(Byval pType, Byval pAmt, Byval pSeqNo, Byval pAcctCd, Byval pCode, Byval pDesc)
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "				' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pType)),"''","S") & ", "			' 차/대 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pAcctCd)),"''","S") & ", "		' 과목 코드 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pAmt, "0"),"0","D")  & ", "			' 금액 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pCode)),"''","S") & ", "			' 처분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pDesc)),"''","S") & ", "			' 조정내용 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	PrintLog "TB_15_PushData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

Sub TB_15_DeleData(Byval pType, Byval pSeqNo)
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_DeleData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pType)),"''","S") & ", "			' 1호/2호 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "


	PrintLog "TB_15_DeleData = " & lgStrSQL
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
