<%@ LANGUAGE=VBSCript%>
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

	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey, lgCurrGrid, lgPgmID, lgTaxDocCd
		
	Const TYPE_1 = 0
	Const TYPE_2 = 1

	Dim C_PGM_ID
	Dim C_TAX_DOC_CD
	Dim C_PGM_NM
	Dim C_ERR_TYPE

	Dim C_SEQ_NO
	Dim C_ERR_DOC
	Dim C_ERR_VAL

	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	lgStrPrevKey		= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPgmID				= FilterVar(Request("PGM_ID"),"''", "S")		' 프로그램ID
    lgCurrGrid			= CDbl(Request("txtCurrGrid"))
    
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    
    Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 버전관리기능 추가 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
			If lgCurrGrid = TYPE_1 Then
				Call SubBizQuery()
			Else
			
				Call SubBizQuery2()
			End If
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()

    End Select

    Call SubCloseDB(lgObjConn)
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 

	C_PGM_ID		= 1
	C_TAX_DOC_CD	= 2	
	C_PGM_NM		= 3
	C_ERR_TYPE		= 4
	
	C_SEQ_NO		= 1
	C_ERR_DOC		= 2
	C_ERR_VAL		= 3

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
    lgStrSQL =            "DELETE TB_TAX_DOC_HTF WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx
    Dim iLoopMax, sData
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 


	Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	     lgStrPrevKey = ""
	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		    
	Else

	    iDx = 1
		    
	    Do While Not lgObjRs.EOF
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("PGM_ID"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("CHK_FLG"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("TAX_DOC_CD"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("MNU_NM"))			
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("ERR_TYPE"))	
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("STATUS_FLG"))	
			sData = sData & Chr(11) & iDx
			sData = sData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	    Loop 
		    
	    lgObjRs.Close
			
	End If
    
	Set lgObjRs = Nothing

    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData0             " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & sData       & """" & vbCr

    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

' -- 전자신고 그리드2 조회 
'========================================================================================
Sub SubBizQuery2()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx
    Dim iLoopMax, sData
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 



	Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	     lgStrPrevKey = ""
	   ' Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		    
	Else

	    iDx = 1
		    
	    Do While Not lgObjRs.EOF

	    
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("ERR_DOC"))
			sData = sData & Chr(11) & ConvSPChars(replace(lgObjRs("ERR_VAL"),chr(13),""))
			sData = sData & Chr(11) & iDx
			sData = sData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	    Loop 
		     sData = replace(sData,chr(13),"")
		     sData = replace(sData,chr(10),"")
	    lgObjRs.Close
			
	End If
    
	Set lgObjRs = Nothing

    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData1             " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & sData       & """" & vbCr

    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "H"
			lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_TAX_DOC_CheckProgress " & pCode1 & "," & pCode2 & "," & pCode3 & "," & FilterVar(gUsrID,"''", "S") & vbCrLf & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT  "
            
            lgStrSQL = lgStrSQL & " A.CALLED_FRM_ID PGM_ID, B.TAX_DOC_CD, A.MNU_NM ,  CASE When D.STATUS_FLG  = 2 Then dbo.TB_TAX_DOC_GetHTFStatus(" & pCode1 & "," & pCode2 & "," & pCode3 & ", A.CALLED_FRM_ID) ELSE '' END ERR_TYPE, D.STATUS_FLG" & vbCrLf
            lgStrSQL = lgStrSQL & " , CASE D.STATUS_FLG " & vbCrLf
            lgStrSQL = lgStrSQL & "	  WHEN '2' THEN 1 " & vbCrLf
            lgStrSQL = lgStrSQL & "   ELSE 0  " & vbCrLf
            lgStrSQL = lgStrSQL & "   END CHK_FLG " & vbCrLf
            
            lgStrSQL = lgStrSQL & " FROM V_MENU	A " & vbCrLf	' 사용자권한별 메뉴뷰 
            lgStrSQL = lgStrSQL & " INNER JOIN TB_TAX_DOC B ON A.CALLED_FRM_ID = B.PGM_ID AND B.HT_TYPE = 'Y' " & vbCrLf	' -- 전자신고문서 정의테이블 
            'lgStrSQL = lgStrSQL & "	LEFT OUTER JOIN TB_TAX_DOC_HTF D ON A.CALLED_FRM_ID = D.PGM_ID AND D.CO_CD=" & pCode1 & " AND D.FISC_YEAR=" & pCode2 & " AND D.REP_TYPE = " & pCode3  & vbCrLf	' 오류테이블 
			lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN TB_TAX_DOC_DTL D ON A.CALLED_FRM_ID = D.PGM_ID AND D.CO_CD=" & pCode1 & " AND D.FISC_YEAR=" & pCode2 & " AND D.REP_TYPE = " & pCode3  & vbCrLf
			
            lgStrSQL = lgStrSQL & "WHERE A.USR_ID = '" & gUsrID & "' " & vbCrLf
			lgStrSQL = lgStrSQL & "ORDER BY TAX_DOC_CD, PGM_ID DESC" & vbCrLf
 
      Case "D"
			lgStrSQL = lgStrSQL & " SELECT  "
            lgStrSQL = lgStrSQL & " A.ERR_DOC, A.ERR_VAL" & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_TAX_DOC_HTF	A " & vbCrLf	' 사용자권한별 메뉴뷰 
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.PGM_ID = " & lgPgmID 	 & vbCrLf
	End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
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