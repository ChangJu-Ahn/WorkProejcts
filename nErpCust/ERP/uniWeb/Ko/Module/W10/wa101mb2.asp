<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<%Option Explicit%> 

<% session.CodePage=949 %>


<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../wcm/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<!-- #Include file="../wcm/inc_HomeTaxFunc.asp" -->


<% ' -------  전자신고 Include ---------------------  %>
<!-- #Include file="WB101MA1_HTF.asp" -->
<!-- #Include file="W8107MA1_HTF.asp" -->
<!-- #Include file="W8101MA1_HTF.asp" -->
<!-- #Include file="W5103MA1_HTF.asp" -->
<!-- #Include file="W7109MA1_HTF.asp" -->
<!-- #Include file="W8113MA1_HTF.asp" -->
<!-- #Include file="W8111MA1_HTF.asp" -->
<!-- #Include file="W6125MA1_HTF.asp" -->
<!-- #Include file="W4105MA1_HTF.asp" -->
<!-- #Include file="W9101MA1_HTF.asp" -->
<!-- #Include file="W2107MA1_HTF.asp" -->
<!-- #Include file="W1101MA1_HTF.asp" -->
<!-- #Include file="W1105MA1_HTF.asp" -->
<!-- #Include file="W1107MA1_HTF.asp" -->
<!-- #Include file="W1109MA1_HTF.asp" -->
<!-- #Include file="W1111MA1_HTF.asp" -->
<!-- #Include file="W1113MA1_HTF.asp" -->
<!-- #Include file="W1115MA1_HTF.asp" -->
<!-- #Include file="W1117MA1_HTF.asp" -->
<!-- #Include file="W9107MA1_HTF.asp" -->
<!-- #Include file="W5109MA1_HTF.asp" -->
<!-- #Include file="W9123MA1_HTF.asp" -->
<!-- #Include file="W9109MA1_HTF.asp" -->
<!-- #Include file="W9111MA1_HTF.asp" -->
<!-- #Include file="W2105MA1_HTF.asp" -->
<!-- #Include file="W3129MA1_HTF.asp" -->
<!-- #Include file="W6101MA1_HTF.asp" -->
<!-- #Include file="W9113MA1_HTF.asp" -->
<!-- #Include file="W6127MA1_HTF.asp" -->
<!-- #Include file="W1119MA1_HTF.asp" -->
<!-- #Include file="W9103MA1_HTF.asp" -->
<!-- #Include file="W7105MA1_HTF.asp" -->
<!-- #Include file="W6121MA1_HTF.asp" -->
<!-- #Include file="W8109MA1_HTF.asp" -->
<!-- #Include file="WB107MA1_HTF.asp" -->
<!-- #Include file="W6119MA1_HTF.asp" -->
<!-- #Include file="W9121MA1_HTF.asp" -->
<!-- #Include file="W9119MA1_HTF.asp" -->
<!-- #Include file="W7101MA1_HTF.asp" -->
<!-- #Include file="W6124MA1_HTF.asp" -->
<!-- #Include file="W6103MA1_HTF.asp" -->
<!-- #Include file="W6105MA1_HTF.asp" -->
<!-- #Include file="W8105MA1_HTF.asp" -->
<!-- #Include file="W6113MA1_HTF.asp" -->
<!-- #Include file="W9115MA1_HTF.asp" -->
<!-- #Include file="W6109MA1_HTF.asp" -->
<!-- #Include file="W6111MA1_HTF.asp" -->
<!-- #Include file="W9117MA1_HTF.asp" -->

<% ' -------  전자신고는 아니지만 참조해야 하는 서식 Include  ---------------------  %>
<!-- #Include file="W7107MA1_HTF.asp" -->
<!-- #Include file="W4101MA1_HTF.asp" -->
<!-- #Include file="W4103MA1_HTF.asp" -->

<% ' -------  200603 개정판 추가 서식 Include  ---------------------  %>
<!-- #Include file="W9125MA1_HTF.asp" -->
<!-- #Include file="W9127MA1_HTF.asp" -->
<!-- #Include file="W9129MA1_HTF.asp" -->

<%	
dim strFilePath
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    'On Error Resume Next
    Err.Clear

	Dim lgsFISC_YEAR, lgsREP_TYPE
	Dim lgsPGM_ID		' -- 각 전자신고파일생성 Include파일에서 값을 넣어야 한다.
	Dim wgcCompanyInfo	' -- include에서 사용할 법인정보 클래스 
	Dim lgsHTFBody		' -- include에서 전자신고파일로 생성할 스트링 
	Dim lgsTAX_DOC_CD		' -- 현재 전자신고파일의 서식명 
	Dim lgsPGM_NM		' -- 현재 전자신고파일의 프로그램명 
	Dim lgarrRs			' -- 전자파일생성할 레코드셋 
	
	
	Dim C_PGM_ID
	Dim C_TAX_DOC_CD
	Dim C_PGM_NM


	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgsFISC_YEAR		= Request("txtFISC_YEAR")
    lgsREP_TYPE			= Request("cboREP_TYPE")
    
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    
    Call CheckVersion(lgsFISC_YEAR, lgsREP_TYPE)	' 2005-03-11 버전관리기능 추가 
   	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
         Case CStr(UID_M0003)                                                         '☜: Save,Update
             Call SubFileDownLoad()     
    End Select

    Call SubCloseDB(lgObjConn)
    
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 

	C_PGM_ID		= 0
	C_TAX_DOC_CD	= 1	
	C_PGM_NM		= 2

End Sub

' -- 서식 프로그램이 존재하는지 체크 
Function SearchTaxDocCd(Byval pDocCd)
	Dim i, iMaxRows
	iMaxRows = UBound(lgarrRs, 2)
	For i = 0 To iMaxRows - 1
		If lgarrRs(C_TAX_DOC_CD, i) = pDocCd Then
			SearchTaxDocCd = True
			Exit Function
		End If
	Next
	SearchTaxDocCd = False
End Function




'========================================================================================
Sub SubBizSave()
    Dim iKey1, iKey2, iKey3, i
    Dim iDx, iMaxRows
    Dim sExecData, sPGM_ID, blnError, oRs
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	sPGM_ID = "WA101MA1"' -- 전자신고 프로그램 
    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 
	blnError = False
	
	' 전자신고용 파일을 가져온다.
	Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3, Request("txtSpread"))                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn, oRs,lgStrSQL, "", "") = False Then
  
	     lgStrPrevKey = ""
	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		    
	Else		
		' ------------- 신고서 공통 ---------------------
		Call SubBizDelete		' -- 전자신고 에러테이블 초기화 
		
		lgarrRs = oRs.GetRows()
		iMaxRows = UBound(lgarrRs, 2)

		Call SubCloseRs(oRs)	
		
		' ------------- 각 서식별 -----------------------
		For i = 0 To iMaxRows 
           
			lgsTAX_DOC_CD	= lgarrRs(C_TAX_DOC_CD, i)
			lgsPGM_NM		= lgarrRs(C_PGM_NM, i)
		
			' -- include 한  함수를 동적으로 호출한다.
			sExecData = "	Call MakeHTF_" & lgarrRs(C_PGM_ID, i) & "()" & vbCrLf
			PrintLog "sExecData : " & sExecData
			
			Execute sExecData
 
			If Err.number > 0 Then
				PrintLog "Execute Error.. : " & Err.Description
				 Call PrintMesg(UNIGetMesg(TYPE_CHK_HTF_MODULE, lgsPGM_NM,""))
					
				Exit For
			End If
		
	    Next 
	     IF lgErrorStatus = "NO" 	Then
    	    Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
    	
		 END IF

		Call CloseFileSystem	' -- 파일생성은 법인정보공통(WB101MA1)에서 한다.
	End If
	
    
	PrintLog "SubBizSave .. : Success and blnError = " & blnError
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3, pWhere)
    Select Case pMode 
      Case "H"
			lgStrSQL = lgStrSQL & " SELECT  "
            lgStrSQL = lgStrSQL & " B.PGM_ID , B.TAX_DOC_CD, A.MNU_NM " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM V_MENU	A " & vbCrLf	' 사용자권한별 메뉴뷰 
            lgStrSQL = lgStrSQL & " INNER JOIN TB_TAX_DOC B ON A.CALLED_FRM_ID = B.PGM_ID AND B.HT_TYPE = 'Y' " & vbCrLf	' -- 전자신고문서 정의테이블 
            'lgStrSQL = lgStrSQL & "	LEFT OUTER JOIN TB_TAX_DOC_HTF D ON A.CALLED_FRM_ID = D.PGM_ID AND D.CO_CD=" & pCode1 & " AND D.FISC_YEAR=" & pCode2 & " AND D.REP_TYPE = " & pCode3  & vbCrLf	' 오류테이블 

            lgStrSQL = lgStrSQL & "WHERE A.USR_ID = '" & gUsrID & "' " & vbCrLf
            
            If pWhere <> "" Then
				lgStrSQL = lgStrSQL & " AND PGM_ID IN (" & pWhere & ")" & vbCrLf
			End If
			lgStrSQL = lgStrSQL & "ORDER BY TAX_DOC_CD, PGM_ID DESC" & vbCrLf
 
      Case "D"
			lgStrSQL = lgStrSQL & " SELECT  "
            lgStrSQL = lgStrSQL & " A.ERR_DOC, A.ERR_VAL" & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_TAX_DOC_HTF	A " & vbCrLf	
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.PGM_ID = " & lgPgmID 	 & vbCrLf
	End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub


Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_TAX_DOC_HTF WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(lgsFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(lgsREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SaveHTFError
' Desc : 
'============================================================================================================
Sub SaveHTFError(Byval pPGMID, Byval pErrVal, Byval pErrDoc)
	dim i
	On Error Resume Next 
	Err.Clear            
	                                                        '☜: Clear Error status
    Call SubCreateCommandObject(lgObjComm)
        
    With lgObjComm
        .CommandText = "usp_TB_TAX_DOC_HTF_SaveHTFError"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        'lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"    ,adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CO_CD"          ,adVarChar,adParamInput, 20, wgCO_CD)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@FISC_YEAR"      ,adVarChar,adParamInput, 4, lgsFISC_YEAR)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@REP_TYPE"       ,adVarChar,adParamInput, 10, lgsREP_TYPE)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PGM_ID"         ,adVarChar,adParamInput, 8, pPGMID)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ERR_DOC"        ,adVarChar,adParamInput, 4000, pErrDoc)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ERR_VAL"		,adVarChar,adParamInput, 1000, pErrVal)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@UPDT_USER_ID"   ,adVarChar,adParamInput, 20, gUsrID)

        lgObjComm.Execute ,, adExecuteNoRecords
    End With
   
	If Err.number > 0 Then
		PrintLog "SaveHTFError.. : " & Err.Description
		
	End If

	PrintLog "SaveHTFError.. : pPGMID=" & pPGMID & vbCrLf & "ERR_VAL=" & pErrVal & vbCrLf & "ERR_DOC=" & pErrDoc
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	Call SubCloseCommandObject(lgObjComm)

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


'-----------------------------------------------------------------------------------------------
'            File DownLoad(With B.A)
'-----------------------------------------------------------------------------------------------
Sub SubFileDownLoad()
	
		Err.Clear 

		Call HideStatusWnd

		strFilePath = "http://" & Request.ServerVariables("LOCAL_ADDR") & ":" _
					   & Request.ServerVariables("SERVER_PORT")
        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
        End If
		strFilePath = strFilePath  & "files/" & wgCO_CD &"/HomeTaxFile_" & wgCO_CD & ".A100"
		
		if wgCO_CD = "" Then
		   lgErrorStatus = "YES"
		End If   
		
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
             Set SF = CreateObject("uni2kCM.SaveFile")
      
				If SF.SaveTextFile("<%= strFilePath %>") = True Then
        
					Set SF = Nothing
					parent.subDiskOK("OK")
				Else
	
					Set SF = Nothing
					 parent.subDiskOK("FAIL")
				End If
		  Else
		       parent.subDiskOK("FAIL")		
          End If   
    End Select    
       
</Script>

