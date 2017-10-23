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
lgMaxCount        = 30						                                     '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Dim strCharCd
Dim strCharNm
Dim strCharValueCd
Dim strCharValueNm

Dim TmpBuffer
Dim iTotalStr

strCharCd = FilterVar(Trim(Request("txtCharCd")) ,"''", "S")
strCharValueCd = FilterVar(Trim(Request("txtCharValueCd")) ,"''", "S")
strCharValueNm = FilterVar(Trim(Request("txtCharValueNm")) ,"''", "S")

Call SubOpenDB(lgObjConn)

Call SubBizQuerySingle()

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "	parent.txtCharNm.value = """ & ConvSPChars(strCharNm) & """" & vbCrLf
Response.Write "</Script>" & vbCrLf

IF strCharValueNm = "''" Or (strCharValueCd <> "''" And strCharValueNm <> "''" ) Then
	Call SubBizQueryMulti("CHAR_VALUE_CD")
Else		
    Call SubBizQueryMulti("CHAR_VALUE_NM")
End If
    
Call SubCloseDB(lgObjConn)

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
        Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
		Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
        Response.Write ".ggoSpread.SSShowDataByClip " & """" & ConvSPChars(iTotalStr) & """" & vbCrLf
		
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
' Name : SubBizQuerySingle
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuerySingle()
	
	On Error Resume Next
    Err.Clear
    
	Call SubMakeSQLStatements("CHARCD",strCharCd,strCharValueCd,strCharValueNm)

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		Call DisplayMsgBox("122630", vbInformation, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "	parent.txtCharNm.value = """"" & vbCrLf
		Response.Write "	parent.txtCharCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Call SubCloseRs(lgObjRs)
		Call SubCloseDB(lgObjConn)
		Response.End
	Else
		strCharNm = lgObjRs("CHAR_NM")
		Call SubCloseRs(lgObjRs)
    End If

End Sub    

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

	If pType = "CHAR_VALUE_CD" Then
		Call SubMakeSQLStatements("CD",strCharCd,strCharValueCd,strCharValueNm)
	Else	
		Call SubMakeSQLStatements("NM",strCharCd,strCharValueCd,strCharValueNm)
	End If

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKeyIndex = ""    
		Call DisplayMsgBox("122640", vbInformation, "", "", I_MKSCRIPT)
		Call SubCloseRs(lgObjRs)
		Call SubCloseDB(lgObjConn)
		Response.End
	Else
		ReDim TmpBuffer(0)
		Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex )
        Do While Not lgObjRs.EOF
			
			lgstrData = "" 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHAR_VALUE_CD"))	'사양값 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHAR_VALUE_NM")			'사양값명        
			
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
	        
	        ReDim Preserve TmpBuffer(iDx)
	        
	        TmpBuffer(iDx) = lgstrData
	         
	        iDx =  iDx + 1
	            
	        If iDx >= lgMaxCount Then
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	           Exit Do
	        End If   
	                  
        Loop 
        
        iTotalStr = Join(TmpBuffer, "")
        
		If iDx < lgMaxCount Then
		   lgStrPrevKeyIndex = ""
		End If   

		Call SubCloseRs(lgObjRs)
    End If

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "CHARCD"
			lgStrSQL = "SELECT CHAR_NM"
			lgStrSQL = lgStrSQL & " FROM B_CHARACTERISTIC"
			lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & pCode

		Case "CD"
			lgStrSQL = "SELECT CHAR_VALUE_CD, CHAR_VALUE_NM"
			lgStrSQL = lgStrSQL & " FROM B_CHAR_VALUE"
			lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND CHAR_VALUE_CD >= " & pCode1
			lgStrSQL = lgStrSQL & " AND CHAR_VALUE_NM >= " & pCode2
			lgStrSQL = lgStrSQL & " ORDER BY CHAR_VALUE_CD, CHAR_VALUE_NM" 
         
		Case "NM"
			lgStrSQL = "SELECT CHAR_VALUE_CD, CHAR_VALUE_NM"
			lgStrSQL = lgStrSQL & " FROM B_CHAR_VALUE"
			lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND CHAR_VALUE_CD >= " & pCode1
			lgStrSQL = lgStrSQL & " AND CHAR_VALUE_NM >= " & pCode2
			lgStrSQL = lgStrSQL & " ORDER BY CHAR_VALUE_NM, CHAR_VALUE_CD" 
				
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

%>
