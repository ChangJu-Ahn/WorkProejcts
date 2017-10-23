<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q","A", "NOCOOKIE", "QB")                                                             
	
	Dim FromDate, ToDate, strAcctCd, strAcctNm, strDocCur			
	Dim strMgntCd1Fr, strMgntCd1To, strMgntCd2Fr, strMgntCd2To
	

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL
 
    
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus		= ""
    lgKeyStream			= Split(Request("txtKeyStream"),gColSep)
    lgStrPrevKeyIndex	= UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgPrevNext			= Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    
    FromDate		= lgKeyStream(0)
    ToDate			= lgKeyStream(1)
    strAcctCd		= lgKeyStream(2)
    strDocCur		= lgKeyStream(3)
	strMgntCd1Fr	= Request("txtMgntCd1Fr")
	strMgntCd1To	= Request("txtMgntCd1To")
	if strMgntCd1To="" then strMgntCd1To="zzzzzzzz"

	strMgntCd2Fr	= Request("txtMgntCd2Fr")
	strMgntCd2To	= Request("txtMgntCd2To")    
    if strMgntCd2To="" then strMgntCd2To="zzzzzzzz"

    'a5408mb1.asp?txtMode=1500&txtKeyStream=2002-11-012002-11-3011KRW&txtPrevNext=Q&lgStrPrevKeyIndex=&txtMaxRows=33&ProcessOption=1&QueryOption=1

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    	
    Select Case CStr(Request("txtMode"))                                             '☜: Read Operation Mode (CRUD)
        Case CStr(UID_M0001)
			                                                      '☜: Query
             Select Case lgPrevNext
                Case "N","P","Q" 
					Call SubBizQuery()                                'Next
					Call SubBizQueryMulti()
					
                Case "R"  
                    Call SubBizQueryMulti()
					
             End Select 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
'             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
    'Dim lgSchoolCD
    
	On Error Resume Next                                                             '☜: Protect System From Crashing
    Err.Clear  
    	
	If strAcctCD <> "" Then
	lgStrSQL = " Select Acct_Nm from A_Acct Where Acct_cd =  " & FilterVar(strAcctCD, "''", "S") & ""
		
		 If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
		 	If isNull(lgObjRs("Acct_Nm")) = False Then
				strAcctNm = Trim(lgObjRs("Acct_Nm"))
			End If
		 End if
'		 Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		 Call SubCloseRs(lgObjRs)   
	End If
    
    Call SubMakeSQLStatements("SX","x","x")                                              '☆: Make sql statements
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                'R(Read) X(CursorType) X(LockType) 

        Exit Sub
    Else
		lgStrData = ""	
		lgStrData = lgStrData & Chr(11) & "합계"
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt0"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt1"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt2"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt3"),ggAmtOfMoney.DecPoint, 0)
		'lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
		lgstrData = lgstrData & Chr(11) & Chr(12)
				
		'Response.Write "lgstrData : " & lgstrData'
		'	
		'If iDx <= C_SHEETMAXROWS_D Then
		'   lgStrPrevKeyIndex = ""
		'End If   

		If CheckSYSTEMError(Err,True) = True Then
		   ObjectContext.SetAbort
		   Exit Sub
		End If   '
	
		
		Response.Write  " <Script Language=vbscript> " & vbCr
		Response.Write  " With Parent                " & vbCr
		Response.Write  "    .ggoSpread.Source     = .frm1.vspdData1         " & vbCr
		Response.Write  "    .lgStrPrevKeyIndex    = """ & lgStrPrevKeyIndex & """" & vbCr
		Response.Write  "    .ggoSpread.SSShowData   """ & lgstrData         & """" & vbCr
		Response.Write  "    .DBQueryOk   " & vbCr      
		Response.Write  " End With                   " & vbCr        
		Response.Write  " </Script>                  " & vbCr
				
		Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
'		Call SubBizQueryMulti()
       
    End If
    
End Sub	

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iSelCount
    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    
    Dim strYear,strMonth,strDay, OneDayBefore
    
    Const C_SHEETMAXROWS_D  = 100														'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear																			'☜: Clear Error status
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    If strAcctCD <> "" Then
	lgStrSQL = " Select Acct_Nm from A_Acct(nolock) Where Acct_cd =  " & FilterVar(strAcctCD, "''", "S") & ""
		
		 If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
		 	If isNull(lgObjRs("Acct_Nm")) = False Then
				strAcctNm = Trim(lgObjRs("Acct_Nm"))
			End If
		 End if
'		 Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		 Call SubCloseRs(lgObjRs)   
	End If
	    
	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
	    
	lgStrSql =				"select "
	lgStrSql = lgStrSql &	"	mgnt_val1, mgnt_val2 ,gl_no,gl_seq, "  & vbcrlf
	lgStrSql = lgStrSql &	"	s_order, GL_DT, CLS_DT, GL_DESC, bas_amt, open_amt, cls_amt, bal_amt " & vbcrlf
	lgStrSql = lgStrSql &	"from	dbo.ufn_A5409MA1 (" 
	lgStrSql = lgStrSql &	FilterVar(FromDate, "''", "S")		& ","
	lgStrSql = lgStrSql &	FilterVar(ToDate, "''", "S")		& ","
	lgStrSql = lgStrSql &	FilterVar(strDocCur, "''", "S")		& ","
	    
	lgStrSql = lgStrSql &	FilterVar(strAcctCd, "''", "S")		& ","
	lgStrSql = lgStrSql &	FilterVar(strMgntCd1Fr, "''", "S")	& ","
	lgStrSql = lgStrSql &	FilterVar(strMgntCd1To, "''", "S")	& ","
	lgStrSql = lgStrSql &	FilterVar(strMgntCd2Fr, "''", "S")	& ","
	lgStrSql = lgStrSql &	FilterVar(strMgntCd2To, "''", "S")	& ","
	    
	' 권한관리 추가 
	If lgAuthBizAreaCd	= "" Then	lgAuthBizAreaCd	= "%"
	If lgInternalCd		= "" Then	lgInternalCd		= "%"
'	If lgSubInternalCd	= "" Then	lgSubInternalCd		= "%"
	If lgAuthUsrID		= "" Then	lgAuthUsrID			= "%"
	
	lgStrSql = lgStrSql &	FilterVar(lgAuthBizAreaCd, "''", "S")	& ","
	lgStrSql = lgStrSql &	FilterVar(lgInternalCd, "''", "S")	& ","
	lgStrSql = lgStrSql &	FilterVar(lgSubInternalCd & "%", "''", "S")	& ","
	lgStrSql = lgStrSql &	FilterVar(lgAuthUsrID, "''", "S")		& " "

	lgStrSql = lgStrSql &	")"      
	lgStrSQL = lgStrSQL &	" order by 1,2,3,4,5 "
		
	'Response.Write lgstrsql
		
	If 	FncOpenRs("U",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	    lgStrPrevKeyIndex = ""
	    Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
	    lgErrorStatus     = "YES"
	    Exit Sub
	Else 
       
		lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

		If CInt(lgStrPrevKeyIndex) > 0 Then
			lgObjRs.Move     = C_SHEETMAXROWS_D * UniCInt(lgStrPrevKeyIndex,0)
		End If
		          
		lgstrData = ""
		        
		iDx = 1
		   
		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_val1"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_val2"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))
					
			'Response.Write UNIDateClientFormat(lgObjRs("GLDT")) & "-"
			'UniConvYYYYMMDDToDate(gDateFormat,"2001","02","21") -> 02-21-2001
					
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("GL_DT"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("CLS_DT"))
					
			'lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(Left(lgObjRs("GLDT"),4) & "-" & mid(lgObjRs("GLDT"),5,2) & "-" & Right(lgObjRs("GLDT"),2))
			'lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(Left(lgObjRs("CLSDT"),4) & "-" & mid(lgObjRs("CLSDT"),5,2) & "-" & Right(lgObjRs("CLSDT"),2))
					
			'lgstrData = lgstrData & Chr(11) & lgObjRs("GLDT")
			'lgstrData = lgstrData & Chr(11) & lgObjRs("CLSDT")
				
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_DESC"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("bas_amt"),	ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("open_amt"),	ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("cls_amt"),	ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("bal_amt"),	ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext
					
			iDx =  iDx + 1

			If iDx > C_SHEETMAXROWS_D Then
				lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
				Exit Do
			End If   
		Loop 
      
    End If
    
    'Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If iDx <= C_SHEETMAXROWS_D Then
	   lgStrPrevKeyIndex = ""
    End If   
	
    'If CheckSYSTEMError(Err,True) = True Then
    '   ObjectContext.SetAbort
    '   Exit Sub
    'End If   
	If lgErrorStatus  = "" Then
	
		Response.Write  " <Script Language=vbscript> " & vbCr
		Response.Write  " With Parent                " & vbCr
		Response.Write  "   .Frm1.txtAcctNm.Value    = """ & strAcctNm & """" & vbCr
		Response.Write  " End With                   " & vbCr        
		Response.Write  " </Script>                  " & vbCr
		
    
       Response.Write  " <Script Language=vbscript>                                    " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData         " & vbCr
	   Response.Write  "    Parent.lgStrPrevKeyIndex    = """ & lgStrPrevKeyIndex & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData         & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pSchoolCD,arrColVal)
    Dim iSelCount
    
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                    Select Case  lgPrevNext 
                         Case "Q"
								
									
                         Case "P"
                                                                 
                         Case "N"
                         
                     End Select 
                     
               Case "X"
					Select Case  lgPrevNext 
						Case "Q"
							lgStrSQL =				"select	sum(bas_amt) 'amt0', sum(open_amt) 'amt1', sum(cls_amt) 'amt2', sum(bal_amt) 'amt3'  "
							lgStrSQL = lgStrSQL &	"from	dbo.ufn_A5409MA1 (" 
							lgStrSql = lgStrSql &	FilterVar(FromDate, "''", "S")		& ","
							lgStrSql = lgStrSql &	FilterVar(ToDate, "''", "S")		& ","
							lgStrSql = lgStrSql &	FilterVar(strDocCur, "''", "S")		& ","
    
							lgStrSql = lgStrSql &	FilterVar(strAcctCd, "''", "S")		& ","
							lgStrSql = lgStrSql &	FilterVar(strMgntCd1Fr, "''", "S")	& ","
							lgStrSql = lgStrSql &	FilterVar(strMgntCd1To, "''", "S")	& ","
							lgStrSql = lgStrSql &	FilterVar(strMgntCd2Fr, "''", "S")	& ","
							lgStrSql = lgStrSql &	FilterVar(strMgntCd2To, "''", "S")	& ","
							
							' 권한관리 추가 
							If lgAuthBizAreaCd	= "" Then	lgAuthBizAreaCd		= "%"
							If lgInternalCd		= "" Then	lgInternalCd		= "%"
'							If lgSubInternalCd	= "" Then	lgSubInternalCd		= "%"
							If lgAuthUsrID		= "" Then	lgAuthUsrID			= "%"

							lgStrSql = lgStrSql &	FilterVar(lgAuthBizAreaCd, "''", "S")	& ","
							lgStrSql = lgStrSql &	FilterVar(lgInternalCd, "''", "S")	& ","
							lgStrSql = lgStrSql &	FilterVar(lgSubInternalCd & "%", "''", "S")	& ","
							lgStrSql = lgStrSql &	FilterVar(lgAuthUsrID, "''", "S")		& " "
							
							lgStrSql = lgStrSql &	") where s_order = 3 "      
							
						Case "P"
						
						Case "N"
                         
                    End Select 
           End Select
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

