<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/IncSvrDate.inc" -->
<!-- #Include file="../inc/AdoVbs.inc" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/incServerAdoDB.asp" -->
<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
Call LoadBasisGlobalInf

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgMaxCount		  = 30		
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)



'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
	
Dim strCharCd
Dim strCharNm

Dim TmpBuffer
Dim iTotalStr

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

strCharCd = FilterVar(Trim(Request("txtCharCd")) ,"''", "S")
strCharNm = FilterVar(Trim(Request("txtCharNm")) ,"''", "S")

IF strCharNm = "''" Or (strCharCd <> "''" And strCharNm <> "''" ) Then
	Call SubBizQueryMulti("CHAR_CD")
Else		
    Call SubBizQueryMulti("CHAR_NM")
End If
    
Call SubCloseDB(lgObjConn)

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
		If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
	        Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
			Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	        Response.Write ".ggoSpread.SSShowDataByClip " & """" & ConvSPChars(iTotalStr) & """" & vbCrLf
        End If
		
		' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
		Response.Write "If .vspdData.MaxRows < .VisibleRowCnt(.vspdData,0)  And .lgStrPrevKeyIndex <> """" Then" & vbCrLf
			Response.Write ".DbQuery" & vbCrLf
		Response.Write "Else" & vbCrLf
			Response.Write ".DbQueryOk" & vbCrLf
		Response.Write "End If" & vbCrLf
		Response.Write ".vspdData.Focus" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pType)
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iDx		
	
	iDx = 0    
	
	If pType = "CHAR_CD" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		Call SubMakeSQLStatements("CD",strCharCd,strCharNm)           '☜ : Make sql statements
	Else	
		Call SubMakeSQLStatements("NM",strCharCd,strCharNm)           '☜ : Make sql statements
	End If
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
		Call DisplayMsgBox("122630", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
				 
		Response.End
	Else
		
		IntRetCD = 1
		ReDim TmpBuffer(0)
		
		Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex )
        Do While Not lgObjRs.EOF
			
			lgstrData = ""
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHAR_CD"))		'사양항목 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHAR_NM")			'사양항목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("CHAR_VALUE_DIGIT")	'사양값자리수 
			
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
	        
	        ReDim Preserve TmpBuffer(iDx)
	        
	        TmpBuffer(iDx) = lgstrData
	         
	        iDx =  iDx + 1
	            
	        If iDx >= lgMaxCount Then			'처음에 최상위품목row를 한줄 써주었으므로 
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	           Exit Do
	        End If   
	                  
        Loop 
        
        iTotalStr = Join(TmpBuffer, "")
        
		If iDx < lgMaxCount Then
		   lgStrPrevKeyIndex = ""
		End If   

		Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		Call SubCloseRs(lgObjRs)       
    End If
 
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "CD"
			lgStrSQL = "SELECT CHAR_CD, CHAR_NM, CHAR_VALUE_DIGIT"
			lgStrSQL = lgStrSQL & " FROM B_CHARACTERISTIC"
			lgStrSQL = lgStrSQL & " WHERE CHAR_CD >= " & pCode
			lgStrSQL = lgStrSQL & " AND CHAR_NM >=  " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY CHAR_CD, CHAR_NM" 
			        
		Case "NM"
			lgStrSQL = "SELECT CHAR_CD, CHAR_NM, CHAR_VALUE_DIGIT"
			lgStrSQL = lgStrSQL & " FROM B_CHARACTERISTIC"
			lgStrSQL = lgStrSQL & " WHERE CHAR_CD >= " & pCode
			lgStrSQL = lgStrSQL & " AND CHAR_NM >=  " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY CHAR_NM, CHAR_CD" 
			
   End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub
%>
