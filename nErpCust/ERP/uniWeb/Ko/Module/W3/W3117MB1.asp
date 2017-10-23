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

	' -- 1번 합계 그리드 
	Dim C_SEQ_NO
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W7_NM
	Dim C_W8
	Dim C_W8_NM
	
	' -- 2,3번 가지급금/가수금 그리드 
	Dim C_CHILD_SEQ_NO
	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16

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

	'--1번그리드 
	C_SEQ_NO = 1
	C_W1 = 2
	C_W2 = 3
	C_W3 = 4
	C_W4 = 5
	C_W5 = 6
	C_W6 = 7
	C_W7 = 8
	C_W7_NM = 9
	C_W8 = 10
	C_W8_NM = 11

	'--2번 3번그리드 
	C_CHILD_SEQ_NO	= 2
	C_W9		= 3
	C_W10		= 4
	C_W11		= 5
	C_W12		= 6
	C_W13		= 7
	C_W14		= 8
	C_W15		= 9
	C_W16		= 10

	
End Sub

'========================================================================================
Sub SubBizQuery()
    'On Error Resume Next
    Err.Clear
End Sub
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
    lgStrSQL =            "DELETE TB_19BD1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_19BD2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

    lgStrSQL = lgStrSQL & "DELETE TB_19BH WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

'PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iStrData3, iIntMaxRows, iLngRow
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
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        iStrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
        
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W4"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W6"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7_NM"))	 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8_NM"))
			iStrData = iStrData & Chr(11) & iDx
			iStrData = iStrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
		    iDx = iDx + 1
              
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
		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    iStrData2 = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
		    
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("CHILD_SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W9"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W10"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W11"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W12"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W13"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W14"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W15"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W16"))			 
				iStrData2 = iStrData2 & Chr(11) & iDx
				iStrData2 = iStrData2 & Chr(11) & Chr(12)
			    lgObjRs.MoveNext
          		iDx = iDx + 1
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If  
		       
        ' 2번째 그리드 
        Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
'		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
'		    Call SetErrorStatus()
		    
		Else
'		    Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    iStrData3 = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
		    
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("CHILD_SEQ_NO"))
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W9"))	
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W10"))	
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W11"))		
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W12"))		
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W13"))		
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W14"))	
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W15"))	
				iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W16"))			 
				iStrData3 = iStrData3 & Chr(11) & iDx
				iStrData3 = iStrData3 & Chr(11) & Chr(12)
			    lgObjRs.MoveNext
             	iDx = iDx + 1
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If        
    End If
    
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    'Response.Write "	.lgPageNo = """ & iIntQueryCount           & """" & vbCr
    'Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)          & """" & vbCr
    'Response.Write "	.frm1.hCtrlCd.value =	""" & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_cd))          & """" & vbCr
    'Response.Write "	.frm1.txtCtrlNM.value = """ & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_nm))          & """" & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData2       & """" & vbCr	
    Response.Write "	.ggoSpread.Source = .frm1.vspdData3              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData3       & """" & vbCr	    
    Response.Write "	.DbQueryOk                                      " & vbCr
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
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W2, A.W3, A.W4, A.W5, A.W6, A.W7, dbo.ufn_GetCodeName('W1060', A.W7) W7_NM, A.W8, dbo.ufn_GetCodeName('W1059', A.W8) W8_NM "
            lgStrSQL = lgStrSQL & " FROM TB_19BH A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            'lgStrSQL = lgStrSQL & " ORDER BY  A.W2 ASC" & vbcrlf

      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.CHILD_SEQ_NO , A.W9, A.W10 as W10, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16 "
            lgStrSQL = lgStrSQL & " FROM TB_19BD1 A WITH (NOLOCK) "
            lgStrSQL = lgStrSQL & "		INNER JOIN TB_19BH B WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.REP_TYPE=B.REP_TYPE AND A.FISC_YEAR=B.FISC_YEAR AND A.SEQ_NO=B.SEQ_NO "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC, A.CHILD_SEQ_NO ASC" & vbcrlf	' 법인명 순서 
            
      Case "R3"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.CHILD_SEQ_NO , A.W9, A.W10, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16 "
            lgStrSQL = lgStrSQL & " FROM TB_19BD2 A WITH (NOLOCK) "
            lgStrSQL = lgStrSQL & "		INNER JOIN TB_19BH B WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.REP_TYPE=B.REP_TYPE AND A.FISC_YEAR=B.FISC_YEAR AND A.SEQ_NO=B.SEQ_NO "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC, A.CHILD_SEQ_NO ASC" & vbcrlf	' 법인명 순서 
            
    End Select
'	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

	'PrintLog "SubBizSaveMulti.."
	
    'On Error Resume Next
    Err.Clear 
    
	'PrintLog "1번째 그리드. .: " & Request("txtSpread") 
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
    
	'PrintLog "2번째 그리드.. : " & Request("txtSpread2")
	
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
 
 	'PrintLog "3번째 그리드.. : " & Request("txtSpread2")
	
	' --- 3번째 그리드 
	arrRowVal = Split(Request("txtSpread3"), gRowSep)                                 '☜: Split Row    data
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
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	'On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_19BH WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")	& ","	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S")	& ","	

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

'	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL

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
	
	lgStrSQL = "INSERT INTO TB_19BD1 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, CHILD_SEQ_NO, W9, W10, W11, W12, W13, W14, W15, W16 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_CHILD_SEQ_NO), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W10), ""),"NULL","S")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& "," 
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

'	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 3번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate3(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_19BD2 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, CHILD_SEQ_NO, W9, W10, W11, W12, W13, W14, W15, W16 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_CHILD_SEQ_NO), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W10), ""),"NULL","S")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& "," 

	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

'	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL

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

    lgStrSQL = "UPDATE  TB_19BH WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(Trim(UCase(arrColVal(C_W7 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(Trim(UCase(arrColVal(C_W8 ))),"''","S") & ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  

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
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_19BD1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
	lgStrSQL = lgStrSQL & " W9	= " &  FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& ","  
	lgStrSQL = lgStrSQL & " W10	= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W10), ""),"NULL","S")		& ","  
	lgStrSQL = lgStrSQL & " W11	= " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& "," 
    lgStrSQL = lgStrSQL & " W12 = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W13 = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W14 = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W15 = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W16 = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & ","  
                         
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND CHILD_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_CHILD_SEQ_NO))),"0","D") 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
'	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 2번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate3(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_19BD2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
	lgStrSQL = lgStrSQL & " W9	= " &  FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& ","  
	lgStrSQL = lgStrSQL & " W10	= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W10), ""),"NULL","S")		& ","  
	lgStrSQL = lgStrSQL & " W11	= " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& "," 
    lgStrSQL = lgStrSQL & " W12 = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W13 = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W14 = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W15 = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W16 = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & ","  
                         
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND CHILD_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_CHILD_SEQ_NO))),"0","D") 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_19BH WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	
'	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_19BD1 WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND CHILD_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_CHILD_SEQ_NO))),"0","D") 
   
'	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete3(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_19BD2 WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND CHILD_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_CHILD_SEQ_NO))),"0","D") 
   
'	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
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