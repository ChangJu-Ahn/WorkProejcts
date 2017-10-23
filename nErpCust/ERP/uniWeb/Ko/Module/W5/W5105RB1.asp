<%@ Transaction=required  CODEPAGE=949 Language=VBScript%>
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
<%
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
'   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
'   Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
	Dim lgStrData_2
   Dim lgDataExist
   Dim lgPageNo
   
   Dim lgCO_CD
   Dim lgFISC_YEAR
   Dim lgREP_TYPE

    'On Error Resume Next
    Err.Clear


   lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
   lgMaxCount       = UNICInt(Trim(Request("lgMaxCount")),0)				'☜ : 한번에 가져올수 있는 데이타 건수 
   lgDataExist      = "No"

   lgCO_CD          = FilterVar(Trim(Request("txtCO_CD")),"''", "S")
   lgFISC_YEAR      = FilterVar(Trim(Request("txtFISC_YEAR")),"''", "S")
   lgREP_TYPE       = FilterVar(Trim(Request("cboREP_TYPE")),"''", "S")

    Call SubOpenDB(lgObjConn) 
    	
'    Call  SubMakeSQLStatements()
    Call  QueryData()

    Call SubCloseDB(lgObjConn)


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData(pMode)

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
  
'    If CLng(lgPageNo) > 0 Then
'       lgObjRs.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
'    End If
    
    iLoopCount = -1
	iRowStr = ""
    
	Do Until lgObjRs.EOF
        iLoopCount =  iLoopCount + 1
'        iRowStr = ""
        

		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W_TYPE"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W1"))			
		iRowStr = iRowStr & Chr(11)			
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W1_NM"))			
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W2"))	
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W3_NM"))	
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W3"))	
		iRowStr = iRowStr & Chr(11) & ConvSPChars(lgObjRs("W4"))		
		iRowStr = iRowStr & Chr(11) & iLoopCount
		iRowStr = iRowStr & Chr(11) & Chr(12)
 
'        If iLoopCount < lgMaxCount Then
'           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
'        Else
'           lgPageNo = lgPageNo + 1
'           Exit Do
'        End If
        
        lgObjRs.MoveNext
	Loop

'    If iLoopCount < lgMaxCount Then                                     '☜: Check if next data exists
'       lgPageNo = ""
'    End If

	Select Case pMode
		Case "R1"
			lgstrData = iRowStr
		Case "R2"
			lgstrData_2 = iRowStr
	End Select
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
Sub SubMakeSQLStatements(pMode)

	Select Case pMode
      Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.SEQ_NO, A.W1, B.ITEM_NM AS W1_NM, A.W2, dbo.ufn_GetCodeName('W1001', A.W3) AS W3_NM, A.W3, A.W4 "
            lgStrSQL = lgStrSQL & " FROM TB_15 A WITH (NOLOCK), TB_ADJUST_ITEM B WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & lgCO_CD 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & lgFISC_YEAR 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & lgREP_TYPE 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W_TYPE = '1' " & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W1 *= B.ITEM_CD " & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf

      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.SEQ_NO, A.W1, B.ITEM_NM AS W1_NM, A.W2, dbo.ufn_GetCodeName('W1002', A.W3) AS W3_NM, A.W3, A.W4 "
            lgStrSQL = lgStrSQL & " FROM TB_15 A WITH (NOLOCK), TB_ADJUST_ITEM B WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & lgCO_CD 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & lgFISC_YEAR 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & lgREP_TYPE 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W_TYPE = '2' " & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W1 *= B.ITEM_CD " & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf
                        
    End Select
  '  PrintLog "SubMakeSQLStatements = " & lgStrSQL
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgDataExist = "Yes"
    Dim blnNoData1, blnNoData2
    
    blnNoData1 = True : blnNoData2 = True
    
    Call  SubMakeSQLStatements("R1")

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		blnNoData1 = False
	Else
		Call MakeSpreadSheetData("R1")
	End If

	Call  SubMakeSQLStatements("R2")
	
	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	    blnNoData2 = False
	Else
		Call MakeSpreadSheetData("R2")
	End If


	If blnNoData1 = False And blnNoData2 = False Then
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
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
       parent.ggoSpread.SSShowData "<%=lgstrData%>"          '☜ : Display data
       parent.ggoSpread.Source  = parent.frm1.vspdData2
       parent.ggoSpread.SSShowData "<%=lgstrData_2%>"          '☜ : Display data
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
    End If   
</Script>	
