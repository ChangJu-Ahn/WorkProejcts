<%@ LANGUAGE=VBSCript Transaction=required %>
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
     lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
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

	
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

   Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To UBound(arrRowVal)
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
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

End Sub    


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    lgStrSQL = "INSERT INTO TB_ACCT_MATCH("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " MATCH_CD, "
    lgStrSQL = lgStrSQL & " ACCT_CD, "
    lgStrSQL = lgStrSQL & " ACCT_NM, "
    lgStrSQL = lgStrSQL & " ACCT_GP_CD, "
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"

    lgStrSQL = lgStrSQL & " VALUES( "
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(3))),"","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                          
    lgStrSQL = lgStrSQL & ")"


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	if	arrColVal(4)="YES" then
		arrColVal(4) = "2"
    Else
		arrColVal(4) = "1"
    End if
    
    if  arrColVal(5)="YES" then
        arrColVal(5) = "Y"
    Else
        arrColVal(5) = "N"
    End if
    
  
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE  TB_ACCT_MATCH"
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "       and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 
    lgStrSQL = lgStrSQL & "       and Match_CD =" &FilterVar(Trim(UCase(lgKeyStream(3))),"","S") 
    lgStrSQL = lgStrSQL & "       and ACCT_CD   = " &   FilterVar(Trim(UCase(arrColVal(2))),"","S") 


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("MINOR_CD"))
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
		
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

	
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_GP_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_GP_NM"))
          
            lgstrData = lgstrData & Chr(11) &  iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
	    Loop 
		    
	    lgObjRs.Close
			
	End If
    'Response.Write lgstrData
	Set lgObjRs = Nothing

    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData1             " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & lgstrData       & """" & vbCr
    Response.Write "	.DbQueryOk2                                      " & vbCr
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
			lgStrSQL = " SELECT MINOR_CD ,  MINOR_NM  "

            lgStrSQL = lgStrSQL & " FROM B_MINOR " & vbCrLf	'
            lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = 'W1015'"  & vbCrLf	' 
            lgStrSQL = lgStrSQL & " ORDER BY  MINOR_CD "
   
      Case "D"
		    lgStrSQL = "SELECT  ACCT_CD, ACCT_NM, ACCT_GP_CD, "
		    lgStrSQL = lgStrSQL & " (Case when match_cd = '10' then (select minor_nm from b_minor where  major_cd = 'W1056' and  minor_cd =a.ACCT_GP_CD) "   & vbCrLf
			lgStrSQL = lgStrSQL & "		  when match_cd = '06' then (select minor_nm from b_minor where  major_cd = 'W1084' and  minor_cd =a.ACCT_GP_CD)"    & vbCrLf
			lgStrSQL = lgStrSQL & "		  when match_cd = '07' then (select minor_nm from b_minor where  major_cd = 'W1085' and  minor_cd =a.ACCT_GP_CD)"    & vbCrLf
			lgStrSQL = lgStrSQL & "		  when match_cd = '34' then (select minor_nm from b_minor where  major_cd = 'W1086' and  minor_cd =a.ACCT_GP_CD)" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		 else '' end 	    )  ACCT_GP_NM  "
            lgStrSQL = lgStrSQL & " FROM  TB_ACCT_MATCH A"
         	lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.match_cd = '" & trim(request("sMAP_CD")) &"'"  & vbCrLf
			lgStrSQL = lgStrSQL & " ORDER BY ACCT_CD "  & vbCrLf
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