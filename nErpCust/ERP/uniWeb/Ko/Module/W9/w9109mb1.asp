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

	Const TYPE_1	= 0
	Const TYPE_2_1	= 1
	Const TYPE_2_2	= 2
	Const TYPE_3	= 3


	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6_1
	Dim C_W6_2
	Dim C_W89

	Dim C_SEQ_NO
	Dim C_W7
	Dim C_W8
	Dim C_W8_P
	Dim C_W8_NM
	Dim C_W9
	Dim C_W9_P
	Dim C_W9_NM
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13

	Dim C_W16
	Dim C_W17_1
	Dim C_W17
	Dim C_W17_P
	Dim C_W17_NM
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
	Dim C_W36_P
	Dim C_W36_NM

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
	' 그리드1
	
	' 그리드1
	C_W1		= 0
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W5		= 4
	C_W6_1		= 5
	C_W6_2		= 6
	C_W89		= 7
	
	C_SEQ_NO	= 1
	C_W7		= 2
	C_W8		= 3
	C_W8_P		= 4
	C_W8_NM		= 5
	C_W9		= 6
	C_W9_P		= 7
	C_W9_NM		= 8
	C_W10		= 9
	C_W11		= 10
	C_W12		= 11
	C_W13		= 12
	
	C_W16		= 1
	C_W17_1		= 2
	C_W17		= 3
	C_W17_P		= 4
	C_W17_NM	= 5
	C_W18		= 6
	C_W19		= 7
	C_W20		= 8
	C_W21		= 9
	C_W22		= 10
	C_W23		= 11
	C_W24		= 12
	C_W25		= 13
	C_W26		= 14
	C_W27		= 15
	C_W28		= 16
	C_W29		= 17
	C_W30		= 18
	C_W31		= 19
	C_W32		= 20
	C_W33		= 21
	C_W34		= 22
	C_W35		= 23
	C_W36		= 24
	C_W36_P		= 25
	C_W36_NM	= 26
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
    lgStrSQL =            "DELETE TB_54D1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	lgStrSQL = lgStrSQL & "DELETE TB_54D2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_54H WITH (ROWLOCK) " & vbCrLf
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
	Dim arrRow(2), iType, iStrData, iLngCol
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

	' TYPE_1
    Call SubMakeSQLStatements("RH",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        Response.Write " <Script Language=vbscript>	                        " & vbCr
        Response.Write "	Call parent.InitData()" & vbCrLf
        Response.Write " </Script>	                        " & vbCr
    Else
		iLngCol = lgObjRs.Fields.Count
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.IsRunEvents = True " & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.

		For iDx = C_W1 To C_W89
			Select Case iDx
				Case C_W4, C_W5,C_W6_1,C_W6_2
					lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").text = """ & lgObjRs(iDx) & """" & vbCrLf
				Case Else
					lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").value = """ & lgObjRs(iDx) & """" & vbCrLf
			End Select
			
			If Err.number <> 0 Then
				PrintLog "iDx=" & iDx
				Exit Sub
			End If
		Next 
		
		Response.Write lgstrData  &  vbCrLf
	
		Response.Write "	.IsRunEvents = False " & vbCrLf	' 이벤트가 발생하게 한다.
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr

		lgObjRs.Close
		Set lgObjRs = Nothing
			
		iStrData = ""
		'TYPE_2_1, TYPE_2_ 조회 
	    Call SubMakeSQLStatements("RD1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do Until lgObjRs.EOF
				Select Case CDbl(lgObjRs("SEQ_NO"))
					Case 1, 2, 3, 4, 5, 6, 7
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W7"))			
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W8"))	
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ""
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W8_NM"))
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W9"))	
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ""
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W9_NM"))
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W10"))	
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W11"))		
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W12"))		
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & ConvSPChars(lgObjRs("W13"))		
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & iIntMaxRows + iLngRow + 1
						arrRow(TYPE_2_1) = arrRow(TYPE_2_1) & Chr(11) & Chr(12)
					Case 8, 9, 10, 11, 12, 13, 14
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W7"))			
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W8"))	
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ""
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W8_NM"))
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W9"))	
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ""
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W9_NM"))
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W10"))	
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W11"))		
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W12"))		
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & ConvSPChars(lgObjRs("W13"))		
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & iIntMaxRows + iLngRow + 1
						arrRow(TYPE_2_2) = arrRow(TYPE_2_2) & Chr(11) & Chr(12)
				End Select
				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent										" & vbCr
			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2_1 & ")        " & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & arrRow(TYPE_2_1)       & """" & vbCr
			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2_2 & ")        " & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & arrRow(TYPE_2_2)       & """" & vbCr
			Response.Write " Call .ReTypeGrid(" & TYPE_2_1 & ", 1)                                  " & vbCr
			Response.Write " Call .ReTypeGrid(" & TYPE_2_2 & ", 6)                                  " & vbCr
			Response.Write " End With	                        " & vbCr
			Response.Write " </Script>	                        " & vbCr
		Else
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write "	Call parent.InitData()" & vbCrLf
			Response.Write " </Script>	                        " & vbCr
		End If

		iStrData = ""
		'TYPE_3 조회 
	    Call SubMakeSQLStatements("RD2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do Until lgObjRs.EOF
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W16"))
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W17_1"))			
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W17"))	
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W17_NM"))	
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
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W29"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W30"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W31"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W32"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W33"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W34"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W35"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W36"))			 
				iStrData = iStrData & Chr(11) & ""	 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W36_NM"))
				iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
				iStrData = iStrData & Chr(11) & Chr(12)
				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent										" & vbCr
			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_3 & ")        " & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
			Response.Write " Call .SetTotalLine                                  " & vbCr
			Response.Write " End With	                        " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If		
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
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W3, A.W4, A.W5, A.W6_1, W6_2, A.W89" & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_54H A WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "   A.SEQ_NO, A.W7, A.W8, A.W9, A.W10, A.W11, A.W12, A.W13 "
            lgStrSQL = lgStrSQL & "	, B.MINOR_NM W8_NM, C.MINOR_NM W9_NM " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_54D1 A WITH (NOLOCK) " & vbCrLf
            lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN ufn_TB_MINOR('W1082', '" & C_REVISION_YM & "') B ON A.W8 = B.MINOR_CD " & vbCrLf
            lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN ufn_TB_MINOR('W1083', '" & C_REVISION_YM & "') C ON A.W9 = C.MINOR_CD " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

      Case "RD2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "   A.W16, A.W17_1, A.W17, A.W18, A.W19, A.W20, A.W21, A.W22, A.W23, A.W24, A.W25 "
            lgStrSQL = lgStrSQL & " , A.W26, A.W27  , A.W28, A.W29, A.W30, A.W31, A.W32, A.W33, A.W34, A.W35, A.W36"
            lgStrSQL = lgStrSQL & "	, B.MINOR_NM W17_NM, C.MINOR_NM W36_NM " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_54D2 A WITH (NOLOCK) "
            lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN ufn_TB_MINOR('W1034', '" & C_REVISION_YM & "') B ON A.W17 = B.MINOR_CD " & vbCrLf
            lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN ufn_TB_MINOR('W1035', '" & C_REVISION_YM & "') C ON A.W36 = C.MINOR_CD " & vbCrLf
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
    
 	' 그리드 
	'PrintLog "txtSpread0 = " & Request("txtSpread" & CStr(TYPE_1))
			
    arrColVal = Split(Request("txtSpread" & CStr(TYPE_1)), gColSep)    
	
	'PrintLog "txtHeadMode=" & Request("txtHeadMode") & ";" & OPMD_CMODE
	
	If CDbl(Request("txtHeadMode")) = OPMD_CMODE Then
	    Call SubBizSaveMultiCreate(TYPE_1, arrColVal)                            '☜: Create
	Else
	    Call SubBizSaveMultiUpdate(TYPE_1, arrColVal)                            '☜: Update
	End If
				    
	' 그리드 
	For iType = TYPE_2_1 To TYPE_3
	
		'PrintLog "txtSpread" & CStr(iType) & " = " & Request("txtSpread" & CStr(iType))
				
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
		       lgErrorPos = lgErrorPos & arrColVal(iDx) & gColSep
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
	
			lgStrSQL = "INSERT INTO TB_54H WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W1, W2, W3, W4, W5, W6_1, W6_2, W89 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")     & "," 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")     & "," 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W4))),"null","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W5))),"null","S")     & "," 
			
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6_1))),"null","S")     & "," 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6_2))),"null","S")     & "," 
			
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W89), "0"),"0","D")     & "," & vbCrLf
			
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"

		Case TYPE_2_1, TYPE_2_2

			lgStrSQL = "INSERT INTO TB_54D1 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , SEQ_NO, W7, W8, W9, W10, W11, W12, W13 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")     & "," & vbCrLf
			If isDate(arrColVal(C_W7)) = False Then arrColVal(C_W7) = ""
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"null","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")  & "," & vbCrLf
	
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"

		Case TYPE_3

			lgStrSQL = "INSERT INTO TB_54D2 WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W16, W17_1, W17, W18, W19, W20, W21, W22, W23 " & vbCrLf
			lgStrSQL = lgStrSQL & " , W24, W25, W26, W27, W28, W29, W30, W31, W32, W33, W34, W35, W36 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")     & "," & vbCrLf

			If isNumeric(arrColVal(C_W17)) = False Then arrColVal(C_W17) = ""
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W17_1))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W19))),"''","S")     & "," & vbCrLf
	
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W21)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D")     & "," & vbCrLf
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
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W35)), "0"),"0","D")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W36))),"''","S")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"			
	End Select
	'PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
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
		
			lgStrSQL = "UPDATE  TB_54H WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(Trim(UCase(arrColVal(C_W4))),"NULL","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(Trim(UCase(arrColVal(C_W5))),"NULL","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6_1	= " &  FilterVar(Trim(UCase(arrColVal(C_W6_1))),"NULL","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W6_2))),"NULL","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W89		= " &  FilterVar(UNICDbl(arrColVal(C_W89), "0"),"0","D") & "," & vbCrLf

			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
		
		Case TYPE_2_1, TYPE_2_2
		
			lgStrSQL = "UPDATE  TB_54D1 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			If isDate(arrColVal(C_W7)) = False Then arrColVal(C_W7) = ""
			lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S") & "," & vbCrLf
		
			lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W11		= " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W12		= " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W13		= " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf

			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 

		Case TYPE_3
		    If isNumeric(arrColVal(C_W17)) = False Then arrColVal(C_W17) = ""
			lgStrSQL = "UPDATE  TB_54D2 WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W16		= " &  FilterVar(Trim(UCase(arrColVal(C_W16))),"null","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W17_1	= " &  FilterVar(Trim(UCase(arrColVal(C_W17_1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W17		= " &  FilterVar(Trim(UCase(arrColVal(C_W17))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W18		= " &  FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W19		= " &  FilterVar(Trim(UCase(arrColVal(C_W19))),"''","S") & "," & vbCrLf
			
			lgStrSQL = lgStrSQL & " W20		= " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W21		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W21)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W22		= " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W23		= " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W24		= " &  FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W25		= " &  FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W26		= " &  FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W27		= " &  FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W28		= " &  FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W29		= " &  FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W30		= " &  FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W31		= " &  FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W32		= " &  FilterVar(UNICDbl(arrColVal(C_W32), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W33		= " &  FilterVar(UNICDbl(arrColVal(C_W33), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W34		= " &  FilterVar(UNICDbl(arrColVal(C_W34), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W35		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W35)), "0"),"0","D") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W36		= " &  FilterVar(Trim(UCase(arrColVal(C_W36))),"''","S") & "," & vbCrLf

			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W16 = " & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") 	 & vbCrLf 			
									
	End Select
	
	'PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
	
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

		Case TYPE_3

			lgStrSQL =            "DELETE TB_54D2 WITH (ROWLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND W16 = " & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")  	 & vbCrLf 

	End Select
	
	'PrintLog "SubBizSaveMultiDelete = " & lgStrSQL 
	
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
