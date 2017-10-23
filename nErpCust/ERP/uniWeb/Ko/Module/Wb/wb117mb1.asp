<%@ LANGUAGE=VBSCript%>
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

	Dim sFISC_YEAR, sREP_TYPE, sUsrID
	Dim lgStrPrevKey, lgCurrGrid
		
	Dim C_PGM_ID
	Dim C_GROUP_NM
	Dim C_MNU_NM
	Dim C_STATUS_FLG
	Dim C_CONFIRM_FLG
	Dim C_UPDT_USER_ID
	Dim C_UPDT_DT


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

	C_PGM_ID		= 1	
	C_GROUP_NM		= 2
	C_MNU_NM		= 3
	C_STATUS_FLG	= 4
	C_CONFIRM_FLG	= 5
	C_UPDT_USER_ID	= 6
	C_UPDT_DT		= 7

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
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx
    Dim iLoopMax, sData
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
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
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("PGM_ID"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("GROUP_NM"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("MNU_NM"))			
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("STATUS_NM"))	
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("CONFIRM_FLG"))	
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("LAST_UPDT_USER"))	
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("LAST_UPDT_DT"))			 
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
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & sData       & """" & vbCr

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
      Case "R"
			lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_TAX_DOC_CheckProgress " & pCode1 & "," & pCode2 & "," & pCode3 & "," & sUsrID & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT  "
            lgStrSQL = lgStrSQL & " D.PGM_ID, A.MNU_NM , CASE WHEN CONFIRM_FLG = '1' THEN '완료' ELSE dbo.ufn_GetCodeName('W1070', D.STATUS_FLG) END STATUS_NM, D.CONFIRM_FLG,  D.LAST_UPDT_USER, D.LAST_UPDT_DT, X.MNU_NM GROUP_NM " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM (  " & vbCrLf
            lgStrSQL = lgStrSQL & " 	SELECT MNU_ID, MNU_NM, MNU_SEQ UPPER_MNU_SEQ FROM V_MENU WHERE USR_ID= '" & gUsrID & "'" & vbCrLf
            lgStrSQL = lgStrSQL & "		) X  " & vbCrLf
            lgStrSQL = lgStrSQL & "		INNER JOIN V_MENU	A  ON X.MNU_ID = A.UPPER_MNU_ID AND A.USR_ID = '" & gUsrID & "' " & vbCrLf
            lgStrSQL = lgStrSQL & "		INNER JOIN TB_TAX_DOC_DTL D ON A.CALLED_FRM_ID = D.PGM_ID AND D.CO_CD=" & pCode1 & " AND D.FISC_YEAR=" & pCode2 & " AND D.REP_TYPE = " & pCode3  & vbCrLf
			lgStrSQL = lgStrSQL & "ORDER BY UPPER_MNU_SEQ, MNU_SEQ" & vbCrLf
	End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
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
	PrintLog "1번째 그리드.. : " & sData
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
		    Select Case arrColVal(0)
		        Case "U"
		                Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
		    End Select
		    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit Sub
		    End If
		    
		Next
	End If
	
End Sub  


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_TAX_DOC_DTL WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " CONFIRM_FLG	   = " &  FilterVar(Trim(UCase(arrColVal(C_CONFIRM_FLG))),"''","S") & ","
                
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND PGM_ID = " & FilterVar(Trim(UCase(arrColVal(C_PGM_ID))),"''","S")  

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
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