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
    
'   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
'   Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
   Dim lgDataExist
   Dim lgPageNo
   
   Dim lgCO_CD
   Dim lgFISC_YEAR
   Dim lgREP_TYPE
   Dim lgACCT_CD
   Dim lgCREDIT_DEBIT

    'On Error Resume Next
    Err.Clear


   lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
   lgMaxCount       = UNICInt(Trim(Request("lgMaxCount")),0)				'☜ : 한번에 가져올수 있는 데이타 건수 
   lgDataExist      = "No"

   lgCO_CD          = FilterVar(Trim(Request("txtCO_CD")),"''", "S")
   lgFISC_YEAR      = FilterVar(Trim(Request("txtFISC_YEAR")),"''", "S")
   lgREP_TYPE       = FilterVar(Trim(Request("cboREP_TYPE")),"''", "S")
   lgACCT_CD        = FilterVar(Trim(Request("txtACCT_CD")),"''", "S")
   lgCREDIT_DEBIT   = FilterVar(Trim(Request("cboCREDIT_DEBIT")),"''", "S")

    Call SubOpenDB(lgObjConn) 
    	
    Call  SubMakeSQLStatements()
    Call  QueryData()

    Call SubCloseDB(lgObjConn)


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       lgObjRs.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("DOC_DT"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("DOC_AMT"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("CREDIT_DEBIT"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("DOC_DESC"))
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        lgObjRs.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                     '☜: Check if next data exists
       lgPageNo = ""
    End If
    lgObjRs.Close															'☜: Close recordset object
    Set lgObjRs = Nothing													'☜: Release ADF

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

	lgStrSQL =			  " SELECT "
    lgStrSQL = lgStrSQL & " A.SEQ_NO, A.ACCT_CD, A.ACCT_NM, A.DOC_DT, A.DOC_AMT, dbo.ufn_GetCodeName('A1012', A.CREDIT_DEBIT) CREDIT_DEBIT, A.DOC_DESC "

    lgStrSQL = lgStrSQL & " FROM TB_WORK_3 A WITH (NOLOCK), TB_ACCT_MATCH B WITH (NOLOCK) "
	lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & lgCO_CD 	 & vbCrLf
	lgStrSQL = lgStrSQL & "	AND A.FISC_YEAR = " & lgFISC_YEAR 	 & vbCrLf
	lgStrSQL = lgStrSQL & "	AND A.REP_TYPE = " & lgREP_TYPE 	 & vbCrLf
	lgStrSQL = lgStrSQL & " AND A.CO_CD = B.CO_CD " & vbCrLf
	lgStrSQL = lgStrSQL & "	AND A.FISC_YEAR = B.FISC_YEAR " & vbCrLf
	lgStrSQL = lgStrSQL & "	AND A.REP_TYPE = B.REP_TYPE " & vbCrLf
	lgStrSQL = lgStrSQL & "	AND A.ACCT_CD = B.ACCT_CD " & vbCrLf
	lgStrSQL = lgStrSQL & "	AND B.MATCH_CD IN ('08','09') " & vbCrLf

	If lgACCT_CD <> "''" Then
	    lgStrSQL = lgStrSQL & "		AND A.ACCT_CD = " & lgACCT_CD 	 & vbCrLf
	End If

	If lgCREDIT_DEBIT <> "''" Then
	    lgStrSQL = lgStrSQL & "		AND A.CREDIT_DEBIT = " & lgCREDIT_DEBIT 	 & vbCrLf
	End If

    lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC, A.ACCT_CD ASC " & vbcrlf

'    lgStrSQL = "select  A.SEQ_NO, A.ACCT_CD, A.ACCT_NM, A.DOC_DT, A.DOC_AMT, dbo.ufn_GetCodeName('A1012', A.CREDIT_DEBIT) CREDIT_DEBIT, A.DOC_DESC from tb_work_3 A (nolock) "
'    lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC, A.ACCT_CD ASC " & vbcrlf
	PrintLog "SubBizDelete = " & lgStrSQL
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgDataExist = False

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
	    
	Else
		lgDataExist = True
	    
	    Call MakeSpreadSheetData()
	End If  
    
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
<Script Language=vbscript>    

    If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       'Show multi spreadsheet data from this line       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>"          '☜ : Display data
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
    End If   
</Script>	
