<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../wcm/incServeradodb.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
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

	Dim sFISC_YEAR, sREP_TYPE, sUsrID
	Dim lgStrPrevKey, lgCurrGrid
	Dim lgUserID, lgPassword, lgURL, lgW1, lgW2, lgW6
	
	Const C_ERP_uniERP2	= "1"
	Const C_ERP_uniERP1	= "2"
	Const C_ERP_Other	= "3"
	Const C_ERP_Excel	= "4"
	
	Const C_ERP_SP			= "1"
	Const C_ERP_WebService	= "2"
	
	
	Const C_MINOR_1 = "01"
	Const C_MINOR_2 = "02"
	Const C_MINOR_3 = "03"
	Const C_MINOR_4 = "04"
	Const C_MINOR_5 = "05"
	Const C_MINOR_6 = "06"
	Const C_MINOR_7 = "07"
	Const C_MINOR_8 = "08"
	Const C_MINOR_9 = "09"
	Const C_MINOR_10 = "10"
	Const C_MINOR_11 = "11"
	Const C_MINOR_12 = "12"
	Const C_MINOR_13 = "13"
	Const C_MINOR_14 = "14"
	
	Const C_INTERFACE_OK = "1"
	Const C_NO_DATA = "2"
	Const C_NO_ACCT = "3"
	Const C_ERROR = "4"
	
	Dim C_W1
	Dim C_W1_NM
	Dim C_W_CHK
	Dim C_W2
	Dim C_UPDT_USER
	Dim C_UPDT_DT
	Dim C_W3
	Dim C_W4

	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	lgStrPrevKey		= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	sUsrID		= FilterVar(gUsrID,"''", "S")		' 신고구분 

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

	C_W1		= 1	
	C_W1_NM		= 2
	C_W_CHK		= 3
	C_W2		= 4
	C_UPDT_USER	= 5
	C_UPDT_DT	= 6
	C_W3		= 7
	C_W4		= 8

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

End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax, sData, sData2
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 


	Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	     lgStrPrevKey = ""
	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		    
	Else

	    iDx = 1
		    
	    Do While Not lgObjRs.EOF
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("W1"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("W1_NM"))
			sData = sData & Chr(11) & ""
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("UPDT_USER_ID"))	
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("UPDT_DT"))			 
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("W3_NM"))
			sData = sData & Chr(11) & "" 'Replace(Server.HTMLEncode(ConvSPChars(lgObjRs("W4"))), vbCrLf, "")
			sData = sData & Chr(11) & iDx
			sData = sData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	    Loop 
		    
	    lgObjRs.Close
			
	End If
    
	Set lgObjRs = Nothing

	' -- 환경설정 로드 
	Call SubMakeSQLStatements("C",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	sData2 = ""
	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
'PrintRs lgObjRs
		sData2 =		"	.cboW1.value = """ & ConvSPChars(lgObjRs("W1")) & """" & vbCrLf
		sData2 = sData2 & "	Call parent.cboW1_onChange " & vbCrLf
		sData2 = sData2 & "	.cboW2.value = """ & ConvSPChars(lgObjRs("W2")) & """" & vbCrLf
		sData2 = sData2 & "	Call parent.cboW2_onChange " & vbCrLf
		sData2 = sData2 & "	.txtW3.value = """ & ConvSPChars(lgObjRs("W3")) & """" & vbCrLf
		sData2 = sData2 & "	.txtW4.value = """ & ConvSPChars(lgObjRs("W4")) & """" & vbCrLf
		sData2 = sData2 & "	.txtW5.value = """ & ConvSPChars(lgObjRs("W5")) & """" & vbCrLf
		sData2 = sData2 & "	.txtW6.value = """ & ConvSPChars(lgObjRs("W6")) & """" & vbCrLf
		lgObjRs.Close
		
	End If
    
	Set lgObjRs = Nothing
	
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " Sub window_onload() " & vbCr
	Response.Write " With parent.frm1                                   " & vbCr
    Response.Write "	parent.ggoSpread.Source = .vspdData              " & vbCr
    Response.Write "	parent.ggoSpread.SSShowData """ & sData       & """" & vbCr & vbCr
    
    If sData2 <> "" Then
		Response.Write "	parent.lgblnConfig = true" & vbCr
		Response.Write sData2       & vbCr
	End If
	
    Response.Write "	parent.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " End Sub " & vbCr
    Response.Write " </Script>                                          " & vbCr
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R" '-- Tab1 쿼리 

			lgStrSQL = " SELECT  "
            lgStrSQL = lgStrSQL & " A.MINOR_CD W1, A.MINOR_NM W1_NM, ISNULL(B.W2, 0) W2, B.W4, dbo.ufn_getCodeName('W1073', B.W3) W3_NM, B.UPDT_USER_ID, B.UPDT_DT  " & vbCrLf

            lgStrSQL = lgStrSQL & " FROM B_MINOR A" & vbCrLf
            lgStrSQL = lgStrSQL & " 	LEFT OUTER JOIN TB_ERP_INTERFACE B ON A.MINOR_CD = B.W1" & vbCrLf
			lgStrSQL = lgStrSQL & "			AND B.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND B.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND B.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "WHERE A.MAJOR_CD = 'W1055' "

      Case "C"	' -- Tab2 쿼리 

			lgStrSQL = " SELECT  "
            lgStrSQL = lgStrSQL & " W1, W2, W3, W4, W5, W6  " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_ERP_CONFIG A" & vbCrLf
			lgStrSQL = lgStrSQL & "	WHERE	CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
       
	End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub

Sub SubMakeSQLStatements2(pMode,pCode1, pCode2, pCode3, pCode4)
    Select Case pMode 
 
 	  Case "E3"	' -- 보조부 계정조회 
		
			lgStrSQL =  " SELECT ACCT_CD  "
            lgStrSQL = lgStrSQL & " FROM TB_ACCT_MATCH " & vbCrLf
            lgStrSQL = lgStrSQL & " WHERE	CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND MATCH_CD = " & FilterVar(GetMatchCd(pCode4),"''", "S") 	 & vbCrLf
	End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub

Function GetMatchCd(Byval pMinorCd)
	' W1015 코드 참조 
	
	Select Case pMinorCd
		Case C_MINOR_6 ' 접대비보조부 
			GetMatchCd = "10"	
		Case C_MINOR_7 ' 기부금보조부 
			GetMatchCd = "18"	
		Case C_MINOR_8 ' 선급법인세보조부 
			GetMatchCd = "20"	
		Case C_MINOR_9 ' 이자비용보조부 
			GetMatchCd = "13"	
		Case C_MINOR_10 ' 배당금수익보조부 
			GetMatchCd = "05"	
	End Select
End Function

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i, sData
	Dim iKey1, iKey2, iKey3
	
    On Error Resume Next
    Err.Clear 

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	' 환경정보를 읽어온다 
	Call SubMakeSQLStatements("C",iKey1, iKey2, iKey3)                              '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	     lgStrPrevKey = ""
	    Call Displaymsgbox("WB0001", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		Exit Sub   
	Else
		lgStrSQL = ""
		lgW1		= lgObjRs("W1")	' 연결유형 
		lgW2		= lgObjRs("W2")	' 연결방식 
		lgW6		= lgObjRs("W6")	' 원격지 DB
		
	    lgURL		= lgObjRs("W3")	' 원격지URL
		lgUserID	= lgObjRs("W4")	' 원격지ID
		lgPassword	= lgObjRs("W5")	' 원격지PWD
				    
	    lgObjRs.Close
		Set lgObjRs = Nothing
	End If
			
	' ERP추출을 호출한다.
	sData = Request("txtSpread")
	PrintLog "1번째 그리드.. : " & sData
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
			    
		    Select Case arrColVal(C_W1)
		        Case C_MINOR_1
					Call ERPGet1_1()	' 인사-급여 
		        Case C_MINOR_2
					Call ERPGet1_2()	' 인사-상시근로자수 
		        Case C_MINOR_3
					Call ERPGet1_3()	' 인사-임직원대여금 
		        Case C_MINOR_4
					Call ERPGet1_4()	' 인사-퇴직금 
		        Case C_MINOR_5		
		            Call ERPGet2()  ' 계정별잔액 
		        Case C_MINOR_6, C_MINOR_7, C_MINOR_8, C_MINOR_9, C_MINOR_10
		            Call ERPGet3(arrColVal(C_W1))   ' 보조부 
		        Case C_MINOR_12
					Call ERPGet6()  ' 계정마스터 
				Case C_MINOR_14 
					Call ERPGet8()	' 부가세과세표준 2007.05 이부분 작성 안되었음.
				Case C_MINOR_11 
					Call ERPGet5()	' 수입금액 2007.05 이부분 작성 안되었음.
		    End Select
			    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit Sub
		       
		    End If
		    
		    		    
		Next
		
		    IF lgErrorStatus = "NO" 	Then
    	        Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
		    END IF
	End If		

End Sub  

'============================================================================================================
' Name : ERPGet1_1
' Desc : 인사DB - 급여관련 
'============================================================================================================
Sub ERPGet1_1()

	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg, i
    'On Error Resume Next
    Err.Clear 

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	Select Case lgW2
		Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.

			Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
			sURL = lgURL & "GetERP1_1.xml?co_cd=" & wgCO_CD & "t_year=" & sFISC_YEAR 
			
			PrintLog "sURL = " & sURL
			objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
			objHttp.Send
	
			Set xmlDoc = objHttp.ResponseXML
			sServerResponseText = objHttp.ResponseText
	
			Set objHttp = Nothing
	
			lgStrSQL = ""
			If xmlDoc is Nothing Then
				sW3 = C_ERROR	
				
				lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			ElseIf xmlDoc.xml = "" Then
				sW3 = C_ERROR	
				
				lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			Else
				Set oNodeList = xmlDoc.selectNodes("//row")
				
				If oNodeList is Nothing Then
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_INTERFACE_OK
					iW2 = oNodeList.Length

					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_1 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_1 (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
							
						For Each oNode In oNodeList
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","

							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w1").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w2").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w3").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w4").text, "0"),"0","D")     & ","
							
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						Next
					
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
					End If
					
					Set oNode = Nothing
					
					
				End If

				Set oNodeList = Nothing
				
			End If
	
		Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.
			lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get1_1 " & sCoCd & ", " & sFiscYear & ", " & sRepType & ", " & FilterVar(lgW6,"''", "S")	' ERP연결 쿼리 
			PrintLog "lgstrSQL = " & lgstrSQL
			gCursorLocation = 3	' -- adUseClient
			If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				iW2 = lgObjRs.RecordCount
				sW3 = C_INTERFACE_OK
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_1 & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				
				' 기존 데이타를 삭제한다.
				lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_1 WITH (ROWLOCK) " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
		
				If iW2 > 0 Then
				
					lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_1 (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
						
					Do Until lgObjRs.EOF
						lgStrSQL = lgStrSQL & "SELECT "
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","

						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W1").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W2").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W3").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W4").value, "0"),"0","D")     & ","
						
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
						lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						
						lgObjRs.MoveNext
					Loop
				
					lgObjRs.Close
					Set lgObjRs = Nothing
					
					lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
				End If
							
			Else
				If lgErrorStatus    = "YES" Then
					sW3 = C_ERROR	
		
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_NO_DATA
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				End If
			End If
		    
	End Select		

	PrintLog " ERPGet1_1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : ERPGet1_2
' Desc : 인사DB - 상시근로자수 
'============================================================================================================
Sub ERPGet1_2()

	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg, i, iCol, sCol
    'On Error Resume Next
    Err.Clear 

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	Select Case lgW2
		Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.

			Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
			sURL = lgURL & "GetERP1_2.xml?co_cd=" & wgCO_CD & "t_year=" & sFISC_YEAR 
			
			PrintLog "sURL = " & sURL
			objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
			objHttp.Send
	
			Set xmlDoc = objHttp.ResponseXML
			sServerResponseText = objHttp.ResponseText
	
			Set objHttp = Nothing
	
			lgStrSQL = ""
			If xmlDoc is Nothing Then
				sW3 = C_ERROR	
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_2 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			ElseIf xmlDoc.xml = "" Then
				sW3 = C_ERROR	
				
				lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_1 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			Else	
				Set oNode = xmlDoc.selectSingleNode("//row")
				
				If oNode is Nothing Then
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_2 & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_INTERFACE_OK
					iW2 = 1

					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_2 & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_2 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_2 (CO_CD, FISC_YEAR, REP_TYPE, W_TYPE, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf

						For Each oNode In oNodeList
						
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","

							lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("w1").text),"''","S") & ","

							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w2").text, "0"),"0","D")     & ","

							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION" & vbCrLf   

						Next
						
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
					End If

				End If

				Set oNode = Nothing
				
			End If
	
		Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.
			lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get1_2 " & sCoCd & ", " & sFiscYear & ", " & sRepType & ", " & FilterVar(lgW6,"''", "S")	' ERP연결 쿼리 
			PrintLog "lgstrSQL = " & lgstrSQL
			gCursorLocation = 3	' -- adUseClient
			If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				iW2 = lgObjRs.RecordCount
				sW3 = C_INTERFACE_OK
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_2 & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				
				' 기존 데이타를 삭제한다.
				lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_2 WITH (ROWLOCK) " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
		
				If iW2 > 0 Then
				
					lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_2 (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
					
					Do Until lgObjRs.EOF
						lgStrSQL = lgStrSQL & "SELECT "
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
					
					
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W1").value)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W2").value, "0"),"0","D")     & ","
						
						
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
						lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
				
						lgObjRs.MoveNext
					Loop
					
					lgObjRs.Close
					Set lgObjRs = Nothing
					
					lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 

				End If
							
			Else
				If lgErrorStatus    = "YES" Then
					sW3 = C_ERROR	
		
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_2 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_NO_DATA
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_2 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				End If
			End If
		    
	End Select		

	PrintLog " ERPGet1_2 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : ERPGet1_3
' Desc : 인사DB - 임직원대여금 
'============================================================================================================
Sub ERPGet1_3()

	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg, i, iSeqNO
    'On Error Resume Next
    Err.Clear 

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	Select Case lgW2
		Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.

			Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
			sURL = lgURL & "GetERP1_3.xml?co_cd=" & wgCO_CD & "t_year=" & sFISC_YEAR 
			
			PrintLog "sURL = " & sURL
			objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
			objHttp.Send
	
			Set xmlDoc = objHttp.ResponseXML
			sServerResponseText = objHttp.ResponseText
	
			Set objHttp = Nothing
	
			lgStrSQL = ""
			If xmlDoc is Nothing Then
				sW3 = C_ERROR	
				
				lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_3 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			Else	
				Set oNodeList = xmlDoc.selectNodes("//row")
				
				If oNodeList is Nothing Then
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_3 & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_INTERFACE_OK
					iW2 = oNodeList.Length

					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_3 & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_3 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_2 (CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W9, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
						iSeqNO = 1
						For Each oNode In oNodeList
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
							
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(iSeqNO, "0"),"0","D") & ","
							
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("w1").text)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("w2").text)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("w3").text)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("w4").text)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("w5").text)),"''","S") & ","
							
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w6").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w7").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w8").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w9").text, "0"),"0","D")     & ","
							
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
							iSeqNO = iSeqNO + 1
						Next
					
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
					End If
					
					Set oNode = Nothing
					
					
				End If

				Set oNodeList = Nothing
				
			End If
	
		Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.
			lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get1_3 " & sCoCd & ", " & sFiscYear & ", " & sRepType & ", " & FilterVar(lgW6,"''", "S")	' ERP연결 쿼리 
			PrintLog "lgstrSQL = " & lgstrSQL
			gCursorLocation = 3	' -- adUseClient
			If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				iW2 = lgObjRs.RecordCount
				sW3 = C_INTERFACE_OK
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_3 & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				
				' 기존 데이타를 삭제한다.
				lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_3 WITH (ROWLOCK) " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
		
				If iW2 > 0 Then
				
					lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_3 (CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W9, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
					iSeqNo = 1
					Do Until lgObjRs.EOF
						lgStrSQL = lgStrSQL & "SELECT "
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
						
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(iSeqNO, "0"),"0","D") & ","
						
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W1").value)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W2").value)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W3").value)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W4").value)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W5").value)),"''","S") & ","
						
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W6").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W7").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W8").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W9").value, "0"),"0","D")     & ","
						
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
						lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						
						lgObjRs.MoveNext
						iSeqNO = iSeqNO + 1
					Loop
				
					lgObjRs.Close
					Set lgObjRs = Nothing
					
					lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
				End If
							
			Else
				If lgErrorStatus    = "YES" Then
					sW3 = C_ERROR	
		
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_3 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_NO_DATA
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_3 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				End If
			End If
		    
	End Select		

	PrintLog " ERPGet1_3 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : ERPGet1_4
' Desc : 인사DB - 퇴직금관련 
'============================================================================================================
Sub ERPGet1_4()

	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg, i, iSeqNO
    'On Error Resume Next
    Err.Clear 

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	Select Case lgW2
		Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.

			Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
			sURL = lgURL & "GetERP1_4.xml?co_cd=" & wgCO_CD & "t_year=" & sFISC_YEAR 
			
			PrintLog "sURL = " & sURL
			objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
			objHttp.Send
	
			Set xmlDoc = objHttp.ResponseXML
			sServerResponseText = objHttp.ResponseText
	
			Set objHttp = Nothing
	
			lgStrSQL = ""
			If xmlDoc is Nothing Then
				sW3 = C_ERROR	
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_4 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			Else	
				Set oNodeList = xmlDoc.selectNodes("//row")
				
				If oNodeList is Nothing Then
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_4 & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_INTERFACE_OK
					iW2 = oNodeList.Length

					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_4 & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_4 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_4 (CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
						iSeqNO = 1
						For Each oNode In oNodeList
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
							
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(iSeqNO, "0"),"0","D") & ","

							lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("w1").text),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("w2").text),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("w3").text),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("w4").text, "0"),"0","D")     & ","

							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
							iSeqNO = iSeqNO + 1
						Next
					
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
					End If
					
					Set oNode = Nothing
					
					
				End If

				Set oNodeList = Nothing
				
			End If
	
		Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.
			
			lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get1_4 " & sCoCd & ", " & sFiscYear & ", " & sRepType & ", " & FilterVar(lgW6,"''", "S")	' ERP연결 쿼리 
			PrintLog " lgstrSQL = " & lgStrSQL
			
			gCursorLocation = 3	' -- adUseClient
			If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				iW2 = lgObjRs.RecordCount
				sW3 = C_INTERFACE_OK
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_4 & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				
				' 기존 데이타를 삭제한다.
				lgStrSQL = lgStrSQL & "DELETE TB_WORK_1_4 WITH (ROWLOCK) " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

				If iW2 > 0 Then
				
					lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_1_4 (CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
					iSeqNo = 1
					Do Until lgObjRs.EOF
						lgStrSQL = lgStrSQL & "SELECT "
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
						
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(iSeqNO, "0"),"0","D") & ","
						
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W1").value)),"''","S")     & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W2").value)),"''","S")     & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("W3").value)),"''","S")    & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("W4").value, "0"),"0","D")     & ","
						
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
						lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						
						lgObjRs.MoveNext
						iSeqNO = iSeqNO + 1
					Loop
				
					lgObjRs.Close
					Set lgObjRs = Nothing
					
					lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
				End If
							
			Else
				If lgErrorStatus    = "YES" Then
					sW3 = C_ERROR	
		
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_4 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_NO_DATA
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_4 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				End If
			End If
		    
	End Select			

	PrintLog " ERPGet1_4 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 계정별잔액 
'============================================================================================================
Sub ERPGet2()
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    	
	PrintLog "ERPGet2 = Running"
	Select Case lgW2
		Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.
					
			Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
			sURL = lgURL & "GetERP2.xml?t_year=" & sFISC_YEAR & "&bs_pl_fg="
			PrintLog "sURL = " & sURL
			objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
			objHttp.Send
	
			Set xmlDoc = objHttp.ResponseXML
			sServerResponseText = objHttp.ResponseText
	
			Set objHttp = Nothing
	
			lgStrSQL = ""
			If xmlDoc is Nothing Then
				sW3 = C_ERROR	
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_5 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			Else	
				Set oNodeList = xmlDoc.selectNodes("//row")
				
				If oNodeList is Nothing Then
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_5 & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_INTERFACE_OK
					iW2 = oNodeList.Length

					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_5 & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_2 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_2 (CO_CD, FISC_YEAR, REP_TYPE, ACCT_CD, ACCT_NM, DEBIT_BASIC_AMT, CREDIT_BASIC_AMT, DEBIT_SUM_AMT, CREDIT_SUM_AMT, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
							
						For Each oNode In oNodeList
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("acct_cd").text),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("acct_nm").text),"''","S") & ","

							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("dr_b").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("cr_b").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("dr_sum").text, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("cr_sum").text, "0"),"0","D")     & ","
							
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						Next
					
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
					End If
					
					Set oNode = Nothing
					
					
				End If

				Set oNodeList = Nothing
				
			End If

		Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.
			
			lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get2 " & sCoCd & ", " & sFiscYear & ", " & sRepType & ", " & FilterVar(lgW6,"''", "S")	' ERP연결 쿼리 
			PrintLog " SP EXEC = " & lgStrSQL
			
			gCursorLocation = 3	' -- adUseClient
			If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				iW2 = lgObjRs.RecordCount
				sW3 = C_INTERFACE_OK
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_5 & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				
				' 기존 데이타를 삭제한다.
				lgStrSQL = lgStrSQL & "DELETE TB_WORK_2 WITH (ROWLOCK) " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

				If iW2 > 0 Then
				
					lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_2 (CO_CD, FISC_YEAR, REP_TYPE, ACCT_CD, ACCT_NM, DEBIT_BASIC_AMT, CREDIT_BASIC_AMT, DEBIT_SUM_AMT, CREDIT_SUM_AMT, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
					iSeqNo = 1
					Do Until lgObjRs.EOF
						lgStrSQL = lgStrSQL & "SELECT "
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
						
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("acct_cd").value)),"''","S")     & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("acct_nm").value)),"''","S")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("dr_b").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("cr_b").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("dr_sum").value, "0"),"0","D")     & ","
						lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("cr_sum").value, "0"),"0","D")     & ","
						
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
						lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						
						lgObjRs.MoveNext
						iSeqNO = iSeqNO + 1
					Loop
				
					lgObjRs.Close
					Set lgObjRs = Nothing
					
					lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
				End If
							
			Else
				If lgErrorStatus    = "YES" Then
					sW3 = C_ERROR	
		
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_5 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_5 & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				End If
			End If
		    
	End Select		
		
	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL & vbCrLf 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 보조부 
'============================================================================================================
Sub ERPGet3(Byval pMinorCd)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg, sAcctCd, iSeqNO

	Dim iKey1, iKey2, iKey3
	
    'On Error Resume Next
    Err.Clear 

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	' 계정정보를 읽어온다 
	Call SubMakeSQLStatements2("E3",sCoCd, sFiscYear, sRepType, pMinorCd)                              '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		lgStrPrevKey = ""
		lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & pMinorCd & "', 0, '" & C_NO_ACCT & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

	Else
		' 계정코드 생성 
		Do While Not lgObjRs.EOF
			sAcctCd = sAcctCd & "'" & lgObjRs("ACCT_CD") & "',"
		    lgObjRs.MoveNext
		Loop 
		sAcctCd = Left(sAcctCd, Len(sAcctCd)-1)
		 
		lgObjRs.Close
		Set lgObjRs = Nothing	


		Select Case lgW2
			Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.
	
				' 원격XML호출 
				    
				Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
				Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
				sURL = lgURL & "GetERP3.xml?t_year=" & sFISC_YEAR & "&acct_cd=" & sAcctCd
				PrintLog "sURL = " & sURL
				objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
				objHttp.Send
	
				Set xmlDoc = objHttp.ResponseXML
				sServerResponseText = objHttp.ResponseText
	
				Set objHttp = Nothing
	
				lgStrSQL = ""
				If xmlDoc is Nothing Then
					sW3 = C_ERROR	
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & pMinorCd & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else	
					Set oNodeList = xmlDoc.selectNodes("//row")
					PrintLog "oNodeList = " & (oNodeList.Length)
					If oNodeList.Length = 0 Then
						sW3 = C_NO_DATA
						
						lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & pMinorCd & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
					Else
						sW3 = C_INTERFACE_OK
						iW2 = oNodeList.Length

						lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & pMinorCd & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

						' 기존 데이타를 삭제한다.
						lgStrSQL = lgStrSQL & "DELETE TB_WORK_3 WITH (ROWLOCK) " & vbCrLf
						lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
						lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
						lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
						lgStrSQL = lgStrSQL & "		AND ACCT_CD IN ( " & sAcctCd & " ) " & vbCrLf  & vbCrLf 
	
						If iW2 > 0 Then
						
							lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_3 (CO_CD, FISC_YEAR, REP_TYPE,  ACCT_CD, ACCT_NM, DOC_DT, DOC_AMT, CREDIT_DEBIT, DOC_DESC, DOC_TYPE, DOC_TYPE2, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
							
							iSeqNo = 1
							For Each oNode In oNodeList
								lgStrSQL = lgStrSQL & "SELECT "
								lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
								'lgStrSQL = lgStrSQL & FilterVar(UNICDbl(iSeqNO, "0"),"0","D") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("acct_cd").text),"''","S") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("acct_nm").text),"''","S") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("doc_dt").text),"''","S") & ","

								lgStrSQL = lgStrSQL & FilterVar(UNICDbl(oNode.attributes.getNamedItem("doc_amt").text, "0"),"0","D")     & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("credit_debit").text),"''","S") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("doc_desc").text),"' '","S") & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("DOC_TYPE").text),"' '","S")     & ","
								lgStrSQL = lgStrSQL & FilterVar(Trim(oNode.attributes.getNamedItem("DOC_TYPE2").text),"' '","S")     & ","

													
								'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
								lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
								'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
								lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
								lgStrSQL = lgStrSQL & " UNION ALL" & vbCrLf   
								'iSeqNO = iSeqNO + 1
							Next
						
							lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-11)	' 마지막 UNION제거 
						End If
						
						Set oNode = Nothing
						
						
					End If

					Set oNodeList = Nothing
					
				End If

			Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.

				lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get3 " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & Replace(sAcctCd, "'" , "") & "', " & FilterVar(lgW6,"''", "S") & ", " & FilterVar(pMinorCd,"''", "S")	' ERP연결 쿼리 
				PrintLog " lgstrSQL = " & lgStrSQL

				gCursorLocation = 3	' -- adUseClient
				If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				
					iW2 = lgObjRs.RecordCount
					sW3 = C_INTERFACE_OK
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & pMinorCd & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
					
					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_3 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
					lgStrSQL = lgStrSQL & "		AND ACCT_CD IN ( " & sAcctCd & " ) " & vbCrLf  & vbCrLf 
					
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_3 (CO_CD, FISC_YEAR, REP_TYPE, ACCT_CD, ACCT_NM, DOC_DT, DOC_AMT, CREDIT_DEBIT, DOC_DESC, DOC_TYPE, DOC_TYPE2, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
						iSeqNo = 1

						Do Until lgObjRs.EOF
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","
							
							'lgStrSQL = lgStrSQL & FilterVar(UNICDbl(iSeqNO, "0"),"0","D") & ","

							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("acct_cd").value)),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("acct_nm").value)),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("doc_dt").value)),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(UNICDbl(lgObjRs("doc_amt").value, "0"),"0","D")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("credit_debit").value)),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("doc_desc").value)),"' '","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("DOC_TYPE").value)),"' '","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("DOC_TYPE2").value)),"' '","S")     & ","
							
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION ALL" & vbCrLf   
							
							lgObjRs.MoveNext
							'iSeqNO = iSeqNO + 1
						Loop
					PrintLog "ERR = " & ERR.Description  & vbCrLf 
						lgObjRs.Close
						Set lgObjRs = Nothing
						
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-11)	' 마지막 UNION제거 
					End If
								
				Else
					If lgErrorStatus    = "YES" Then
						sW3 = C_ERROR	
		
						lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & pMinorCd & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
					Else
						sW3 = C_NO_DATA
						
						lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & pMinorCd & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
					End If
				End If
			    
		End Select	
							
	End If
	
	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL & vbCrLf 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : ERPGet6
' Desc : 계정마스터 
'============================================================================================================
Sub ERPGet6()

	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
	Dim sCoCd, sFiscYear, sRepType
	Dim objHttp, sURL, xmlDoc, sServerResponseText, sW3, iW2, sW4
	Dim oNode, oNodeList, sStatusFlg, i
    'On Error Resume Next
    Err.Clear 

    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    sRepType	= FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
	Select Case lgW2
		Case C_ERP_WebService	' 웹서비스방식이면 ERP유형하고 별상관엄다.

			Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			Set xmlDoc = Server.CreateObject("Msxml2.DomDocument")
	
			sURL = lgURL & "GetERP6.xml?co_cd=" & wgCO_CD & "t_year=" & sFISC_YEAR 
			
			PrintLog "sURL = " & sURL
			objHttp.open "GET", sURL , false, lgUserID, lgPassword
	
			objHttp.Send
	
			Set xmlDoc = objHttp.ResponseXML
			sServerResponseText = objHttp.ResponseText
	
			Set objHttp = Nothing
	
			lgStrSQL = ""
			If xmlDoc is Nothing Then
				sW3 = C_ERROR	
				
				lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_12 & "', 0, '" & sW3 & "', '" & Replace(sServerResponseText, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
			Else	
				Set oNodeList = xmlDoc.selectNodes("//row")
				
				If oNodeList is Nothing Then
					sW3 = C_NO_DATA
					
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_12 & "', 0, '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_INTERFACE_OK
					iW2 = oNodeList.Length

					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_12 & "', " & CStr(iW2) & ", '" & sW3 & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf

					' 기존 데이타를 삭제한다.
					lgStrSQL = lgStrSQL & "DELETE TB_WORK_6 WITH (ROWLOCK) " & vbCrLf
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
					If iW2 > 0 Then
					
						lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_6 (CO_CD, FISC_YEAR, REP_TYPE, ACCT_CD, ACCT_NM, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
							
						For Each oNode In oNodeList
							lgStrSQL = lgStrSQL & "SELECT "
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","

							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("acct_cd").text)),"''","S")     & ","
							lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(oNode.attributes.getNamedItem("acct_nm").text)),"''","S")     & ","
							
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
							'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
							lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
							lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						Next
					
						lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
					End If
					
					Set oNode = Nothing
					
					
				End If

				Set oNodeList = Nothing
				
			End If
	
		Case C_ERP_SP	' SP방식이면 유형에 따라 함수가 다르다.
			lgstrSQL = "EXEC dbo.usp_ERP_TYPE" & lgW1 & "_Get6 " & sFiscYear & ", " & FilterVar(lgW6,"''", "S")	' ERP연결 쿼리 
			PrintLog "lgstrSQL = " & lgstrSQL
			gCursorLocation = 3	' -- adUseClient
			If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = True Then
				iW2 = lgObjRs.RecordCount
				sW3 = C_INTERFACE_OK
				
				lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & sRepType & ", '" & C_MINOR_12 & "', " & CStr(iW2) & ", '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				
				' 기존 데이타를 삭제한다.
				lgStrSQL = lgStrSQL & "DELETE TB_WORK_6 WITH (ROWLOCK) " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
		
				If iW2 > 0 Then
				
					lgStrSQL = lgStrSQL & "INSERT INTO TB_WORK_6 (CO_CD, FISC_YEAR, REP_TYPE, ACCT_CD, ACCT_NM, INSRT_USER_ID, UPDT_USER_ID)" & vbCrLf
						
					Do Until lgObjRs.EOF
						lgStrSQL = lgStrSQL & "SELECT "
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ","

						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("acct_cd").value)),"''","S")     & ","
						lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgObjRs("acct_nm").value)),"''","S")     & ","
						
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
						'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
						lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")      & vbCrLf
						lgStrSQL = lgStrSQL & " UNION" & vbCrLf   
						
						lgObjRs.MoveNext
					Loop
				
					lgObjRs.Close
					Set lgObjRs = Nothing
					
					lgStrSQL = LEFT(lgStrSQL, Len(lgStrSQL)-7)	' 마지막 UNION제거 
				End If
							
			Else
				If lgErrorStatus    = "YES" Then
					sW3 = C_ERROR	
		
					lgStrSQL =  "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & C_MINOR_12 & ", '" & pType & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				Else
					sW3 = C_NO_DATA
					
					lgStrSQL = "EXEC dbo.usp_TB_ERP_INTERFACE_Save " & sCoCd & "," & sFiscYear & "," & C_MINOR_12 & ", '" & pType & "', 0, '" & sW3 & "', '" & Replace(Err.Description, "'", "''") & "', '" & gUsrId & "'" & vbCrLf & vbCrLf
				End If
			End If
		    
	End Select		

	PrintLog " ERPGet6 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : ERPGet8
' Desc : 부가세과세표준 
'============================================================================================================
Sub ERPGet8()
	lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_14 & "', 0, '" & C_NO_ACCT & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
	PrintLog " ERPGet8 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : ERPGet5
' Desc : 부가세과세표준 
'============================================================================================================
Sub ERPGet5()
	lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_ERP_INTERFACE_Save '" & wgCO_CD & "'," & sFISC_YEAR & "," & sREP_TYPE & ", '" & C_MINOR_11 & "', 0, '" & C_NO_ACCT & "', '', '" & gUsrId & "'" & vbCrLf & vbCrLf
	PrintLog " ERPGet8 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

Function GetW6(Byval pW6)
	If pW6 <> "" Then 
		GetW6 = pW6 & "."
	Else
		GetW6 = ""
	End If
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