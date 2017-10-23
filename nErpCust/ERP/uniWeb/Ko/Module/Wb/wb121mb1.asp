<%@ Language=VBScript%>
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
	Dim C_CO_CD
	Dim C_C0_NM
	Dim C_FISC_YEAR
	Dim C_REP_TYPE
	Dim C_REP_TYPE_NM
	Dim C_FISC_START_DT
	Dim C_FISC_END_DT
	Dim C_REVISION_YM2

	Dim lgCO_CD, lgFISC_YEAR
	
	lgErrorStatus    = "NO"
	                                       
    lgCO_CD			= Request("txtCO_CD")
    lgFISC_YEAR		= Request("txtFISC_YEAR")
	lgOpModeCRUD    = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    Call InitSpreadPosVariables()
    
    Call SubOpenDB(lgObjConn) 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
    End Select

    Call SubCloseDB(lgObjConn)

'========================================================================================
sub InitSpreadPosVariables()
	
	C_CO_CD			= 1
	C_C0_NM			= 2
	C_FISC_YEAR		= 3
	C_REP_TYPE		= 4
	C_REP_TYPE_NM	= 5
	C_FISC_START_DT = 6
	C_FISC_END_DT	= 7
	C_REVISION_YM2	= 8
end sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iStrData3, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(lgFISC_YEAR,"''", "S")	' 사업연도 

    Call SubMakeSQLStatements("R1",iKey1, iKey2, "")                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        iStrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
        
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("CO_CD"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("CO_NM"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("FISC_YEAR"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("REP_TYPE"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("REP_TYPE_NM"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("FISC_START_DT"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("FISC_END_DT"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("REVISION_YM"))	
			iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
              
        Loop 
        
        lgObjRs.Close
        Set lgObjRs = Nothing
 		       
    End If
    
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
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
            lgStrSQL = lgStrSQL & " A.FISC_YEAR, A.CO_CD, A.CO_NM, A.REP_TYPE, dbo.ufn_TAX_GetCodeName('W1018', A.REP_TYPE, '" & C_REVISION_YM & "') REP_TYPE_NM, A.FISC_START_DT, A.FISC_END_DT, A.REVISION_YM"
            lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A WITH (NOLOCK) "
            
            If pCode1 <> "''" Or pCode2 <> "''" Then
				lgStrSQL = lgStrSQL & " WHERE "
            End If
            
            If pCode1 <> "''" Then
				lgStrSQL = lgStrSQL & " A.CO_CD = " & pCode1 & " AND"
			End If
			If pCode2 <> "''" Then
				lgStrSQL = lgStrSQL & "	A.FISC_YEAR = " & pCode2  & " AND"
			End If

            If pCode1 <> "''" Or pCode2 <> "''" Then
				lgStrSQL = Left(lgStrSQL, Len(lgStrSQL)-4) & vbcrlf
            End If
            lgStrSQL = lgStrSQL & " AND A.USE_FLG = 'Y' " & vbcrlf
            lgStrSQL = lgStrSQL & " ORDER BY  A.CO_CD, A.FISC_YEAR DESC" & vbcrlf

    End Select
	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

	PrintLog "SubBizSaveMulti.."
	
    'On Error Resume Next
    Err.Clear 
    
	PrintLog "1번째 그리드. .: " & Request("txtSpread") 
	' --- 1번째 그리드 
	arrColVal = Split(Request("txtSpread"), gColSep)                                 '☜: Split Row    data
	
	Call SetCompanyInfo(arrColVal(C_CO_CD), arrColVal(C_C0_NM), arrColVal(C_FISC_YEAR), arrColVal(C_REP_TYPE), arrColVal(C_REVISION_YM2))
	
	
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
