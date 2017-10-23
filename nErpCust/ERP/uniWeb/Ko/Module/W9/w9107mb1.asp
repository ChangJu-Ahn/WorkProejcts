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

	Const TYPE_1	= 0		' 그리드를 구분짓기 위한 상수 
	Const TYPE_2	= 1		

	' -- 그리드 컬럼 정의 
	Dim	C_SEQ_NO
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
	Dim C_W29_1
	Dim C_W29_2
	Dim C_W30_1
	Dim C_W30_2
	'Dim C_W31_NM
	Dim C_W31

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
	C_SEQ_NO	= 1
	C_W1		= 2
	C_W2		= 3
	C_W3		= 4
	C_W4		= 5
	C_W5		= 6
	C_W6		= 7
	C_W7		= 8
	C_W8		= 9
	C_W9		= 10
	C_W10		= 11
	C_W11		= 12
	C_W12		= 13
	C_W13		= 14
	C_W14		= 15
	C_W15		= 16
	C_W16		= 17

	C_W17		= 2
	C_W18		= 3
	C_W19		= 4
	C_W20		= 5
	C_W21		= 6
	C_W22		= 7
	C_W23		= 8
	C_W24		= 9
	C_W25		= 10
	C_W26		= 11
	C_W27		= 12
	C_W28		= 13
	C_W29_1		= 14
	C_W29_2		= 15
	C_W30_1		= 16
	C_W30_2		= 17
	C_W31		= 18
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
    lgStrSQL =            "DELETE TB_52H2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_52H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
	Dim arrRow(2), iType, iStrData
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("RH",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1
        
		Do While Not lgObjRs.EOF
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W4"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W6"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W9"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W10"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W11"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W12"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W13"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W14"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W15"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W16"))			 
			iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext
		Loop 

		lgObjRs.Close
		Set lgObjRs = Nothing
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData0        " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
		Response.Write " End With                                  " & vbCr
		Response.Write " </Script>	                        " & vbCr

	End If
	
	iStrData = ""
	'TYPE_2 조회 
	Call SubMakeSQLStatements("RD",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

		lgstrData = ""
				
		Do While Not lgObjRs.EOF
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W17"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W18"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W19"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W20"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W21"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W22"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W23"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W24"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W25"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W26"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W27"))	 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W28"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W29_1"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W29_2"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W30_1"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W30_2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W31"))			 
			iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext
		Loop 

		lgObjRs.Close
		Set lgObjRs = Nothing
			
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent										" & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData1        " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
		Response.Write " End With                                  " & vbCr
		Response.Write " </Script>	                        " & vbCr
	Else
		Call SetErrorStatus()
	End If
				
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write "	Call parent.DbQueryOk        " & vbCr
	Response.Write " </Script>	                        " & vbCr

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
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W2, A.W3, A.W4, A.W5, A.W6, A.W7, A.W8, A.W9 "
            lgStrSQL = lgStrSQL & " , A.W10, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16"
            lgStrSQL = lgStrSQL & " FROM TB_52H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W17, A.W18, (CASE A.W19 WHEN '1' THEN '1' ELSE '0' END) W19, (CASE A.W19 WHEN '2' THEN '1' ELSE '0' END) W20 "
            lgStrSQL = lgStrSQL & " , A.W21, A.W22, A.W23, A.W24, A.W25"
            lgStrSQL = lgStrSQL & " , (CASE A.W26 WHEN '1' THEN '1' ELSE '0' END) W26, (CASE A.W26 WHEN '2' THEN '1' ELSE '0' END) W27 "
            lgStrSQL = lgStrSQL & " , A.W28, A.W29_1, A.W29_2, A.W30_1, A.W30_2, A.W31"
            lgStrSQL = lgStrSQL & " FROM TB_52H2 A WITH (NOLOCK) "
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
    Dim iDx , iType

    'On Error Resume Next
    Err.Clear 
    
    For iType = TYPE_1 To TYPE_2
    
		' 그리드 
		PrintLog "txtSpread = " & Request("txtSpread" & CStr(iType))
			
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
	
			lgStrSQL = "INSERT INTO TB_52H WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
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

		Case TYPE_2

			lgStrSQL = "INSERT INTO TB_52H2 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , SEQ_NO, W17, W18, W19, W21, W22, W23, W24, W25, W26, W28, W29_1, W29_2, W30_1, W30_2, W31 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S")     & "," & vbCrLf
			
			If Trim(UCase(arrColVal(C_W19))) = "1" Then
				lgStrSQL = lgStrSQL & "'1'," & vbCrLf
			ElseIf Trim(UCase(arrColVal(C_W20))) = "1" Then
				lgStrSQL = lgStrSQL & "'2'," & vbCrLf
			Else
				lgStrSQL = lgStrSQL & "'0'," & vbCrLf
			End If
	
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W21))),"NULL","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W23)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W25)), "0"),"0","D")     & "," & vbCrLf
			
			If Trim(UCase(arrColVal(C_W26))) = "1" Then
				lgStrSQL = lgStrSQL & "'1'," & vbCrLf
			ElseIf Trim(UCase(arrColVal(C_W27))) = "1" Then
				lgStrSQL = lgStrSQL & "'2'," & vbCrLf
			Else
				lgStrSQL = lgStrSQL & "'0'," & vbCrLf
			End If

			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W28))),"NULL","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W29_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W29_2)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W30_1), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W30_2)), "0"),"0","D")     & "," & vbCrLf

			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W31)), "0"),"0","D")     & "," & vbCrLf

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
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	Select Case pType
		Case TYPE_1
		
			lgStrSQL = "UPDATE  TB_52H WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W11		= " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W12		= " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W13		= " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W14		= " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W15		= " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W16		= " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 
		
		Case TYPE_2
		
			lgStrSQL = "UPDATE  TB_52H2 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W17		= " &  FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W18		= " &  FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S") & "," & vbCrLf

			If Trim(UCase(arrColVal(C_W19))) = "1" Then
				lgStrSQL = lgStrSQL & " W19		= '1'," & vbCrLf
			ElseIf Trim(UCase(arrColVal(C_W20))) = "1" Then
				lgStrSQL = lgStrSQL & " W19		= '2'," & vbCrLf
			Else
				lgStrSQL = lgStrSQL & " W19		= '0'," & vbCrLf
			End If
			
			lgStrSQL = lgStrSQL & " W21		= " &  FilterVar(Trim(UCase(arrColVal(C_W21))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W22		= " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W23		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W23)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W24		= " &  FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W25		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W25)), "0"),"0","D") & "," & vbCrLf
			
			If Trim(UCase(arrColVal(C_W26))) = "1" Then
				lgStrSQL = lgStrSQL & " W26		= '1'," & vbCrLf
			ElseIf Trim(UCase(arrColVal(C_W27))) = "1" Then
				lgStrSQL = lgStrSQL & " W26		= '2'," & vbCrLf
			Else
				lgStrSQL = lgStrSQL & " W26		= '0'," & vbCrLf
			End If			
			
			lgStrSQL = lgStrSQL & " W28		= " &  FilterVar(Trim(UCase(arrColVal(C_W28))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W29_1	= " &  FilterVar(UNICDbl(arrColVal(C_W29_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W29_2	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W29_2)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W30_1	= " &  FilterVar(UNICDbl(arrColVal(C_W30_1), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W30_2	= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W30_2)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W31		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W31)), "0"),"0","D") & "," & vbCrLf

			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 
									
	End Select
	
	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function

'========================================================================================
Sub SubBizSaveMultiDelete(Byval pType, Byref arrColVal)
    'On Error Resume Next
    Err.Clear


	Select Case pType
		Case TYPE_1

			lgStrSQL =            "DELETE TB_52H WITH (ROWLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 

		Case TYPE_2

			lgStrSQL =            "DELETE TB_52H2 WITH (ROWLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 

	End Select
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
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
