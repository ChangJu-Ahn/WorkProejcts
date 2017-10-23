<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<%

    On Error Resume Next 
    Err.Clear                                                                        '☜: Clear Error status                                                                '☜: Protect system from crashing

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q","A", "NOCOOKIE", "QB") 
                       
	
	Dim FromDate, ToDate, strAcctCd, strAcctNm, strDocCur			
	Dim strProcessOption, strQueryOption
	Dim strMgntCd1Fr, strMgntCd1To, strMgntCd2Fr, strMgntCd2To
    
	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = ""
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    FromDate = left(lgKeyStream(0),4) & mid(lgKeyStream(0),6,2) & right(lgKeyStream(0),2)
    ToDate = left(lgKeyStream(1),4) & mid(lgKeyStream(1),6,2) & right(lgKeyStream(1),2)

     
    
    strAcctCd = lgKeyStream(2)
    strDocCur = lgKeyStream(3)
    strProcessOption = Request("ProcessOption")
    strQueryOption = Request("QueryOption")

	strMgntCd1Fr = Request("txtMgntCd1Fr")
	strMgntCd1To = Request("txtMgntCd1To")
	strMgntCd2Fr = Request("txtMgntCd2Fr")
	strMgntCd2To = Request("txtMgntCd2To")
 
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
    
	On Error Resume Next                                                             '☜: Protect system from crashing
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

    Call SubMakeSQLStatements("SR","x","x")                                              '☆: Make sql statements
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                'R(Read) X(CursorType) X(LockType) 
       If lgPrevNext = "Q" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
          lgErrorStatus = "YES" 
          Exit Sub
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the starting data. 
          lgPrevNext = "Q"
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the ending data.
          lgPrevNext = "Q"
          Call SubBizQuery()
       End If
       
    Else
		lgStrData = ""
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & "합계"
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgStrData = lgStrData & Chr(11) & ""
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt1"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt2"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt3"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt4"),ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & Chr(12)
		

		If CheckSYSTEMError(Err,True) = True Then
		   ObjectContext.SetAbort
		   Exit Sub
		End If   '
		Response.Write  " <Script Language=vbscript> " & vbCr
		Response.Write  " With Parent                " & vbCr
		Response.Write  "   .Frm1.txtAcctNm.Value    = """ & strAcctNm & """" & vbCr
		Response.Write  " End With                   " & vbCr        
		Response.Write  " </Script>                  " & vbCr
		
		Response.Write  " <Script Language=vbscript> " & vbCr
		Response.Write  " With Parent                " & vbCr
		Response.Write  "    .ggoSpread.Source     = .frm1.vspdData1         " & vbCr
		Response.Write  "    .lgStrPrevKeyIndex    = """ & lgStrPrevKeyIndex & """" & vbCr
		Response.Write  "    .ggoSpread.SSShowData   """ & lgstrData         & """" & vbCr
		Response.Write  "    .DBQueryOk   " & vbCr      
		Response.Write  " End With                   " & vbCr        
		Response.Write  " </Script>                  " & vbCr

       		
       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
       Call SubBizQueryMulti()
       
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
    Dim iCtrl_cd1
    Dim iCtrl_val1
    Dim iCtrl_cd2
    Dim iCtrl_val2
    Dim OUT_DATA_COLM_NM1
    Dim OUT_DATA_COLM_NM2
    
    Dim strYear,strMonth,strDay, OneDayBefore
    
    Const C_SHEETMAXROWS_D  = 100														'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear																			'☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
    iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1

	If strQueryOption = "1" Then		'시점별 

		lgStrSQL =				"	Select	X.mgnt_cd1, X.mgnt_val1, '' 'mgnt_nm1', X.mgnt_cd2, X.mgnt_val2, '' 'mgnt_nm2', "
		lgStrSQL = lgStrSQL &	"	sum(X.amt1) 'amt1', sum(X.amt2) 'amt2', sum(X.amt3) 'amt3', sum(X.amt1+X.amt2-X.amt3) 'amt4' "
		lgStrSQL = lgStrSQL &	"	from	( "
		lgStrSQL = lgStrSQL &	"		select a.gl_no, a.gl_seq, a.mgnt_cd1, a.mgnt_val1, a.mgnt_cd2, a.mgnt_val2, a.open_amt-ISNULL(b.cls_amt,0) 'amt1', 0 'amt2', 0 'amt3' "
		lgStrSQL = lgStrSQL &	"		from a_open_acct a(nolock) "
		lgStrSQL = lgStrSQL &	"		left outer join (SELECT c.GL_NO, c.GL_SEQ, ISNULL(SUM(c.CLS_AMT),0) 'CLS_AMT' "
		lgStrSQL = lgStrSQL &	"			FROM A_CLS_ACCT_ITEM c(nolock) "
		lgStrSQL = lgStrSQL &	"			join A_CLS_ACCT d(nolock) on c.CLS_NO = d.CLS_NO "
		lgStrSQL = lgStrSQL &	"			join A_OPEN_ACCT e(nolock) on c.gl_no = e.gl_no and c.gl_seq = e.gl_seq "
		lgStrSQL = lgStrSQL &	"			WHERE d.CLS_DT <= " & FilterVar(FromDate , "''", "S") & " " 
		lgStrSQL = lgStrSQL &	"			AND c.acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"			AND e.doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"			GROUP BY c.GL_NO, c.GL_SEQ ) b on a.gl_no = b.gl_no and a.gl_seq = b.gl_seq "
		lgStrSQL = lgStrSQL &	"		where a.acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and a.gl_dt < " & FilterVar(FromDate, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and a.doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
		
		If strMgntCd1Fr <> "" Then
			lgStrSQL = lgStrSQL & "		and a.mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "
		End If
		If strMgntCd1To <> "" Then
			lgStrSQL = lgStrSQL & "		and a.mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "
		End If
		If strMgntCd2Fr <> "" Then
			lgStrSQL = lgStrSQL & "		and isnull(a.mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "
		End If
		If strMgntCd2To <> "" Then
			lgStrSQL = lgStrSQL & "		and isnull(a.mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "
		End If						
				
		lgStrSQL = lgStrSQL &	"		union all "

		lgStrSQL = lgStrSQL &	"		select gl_no, gl_seq, mgnt_cd1, mgnt_val1, mgnt_cd2, mgnt_val2, 0, open_amt, 0 "
		lgStrSQL = lgStrSQL &	"		from a_open_acct(nolock) "
		lgStrSQL = lgStrSQL &	"		where acct_cd= " & FilterVar(strAcctCd, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and gl_dt between " & FilterVar(FromDate, "''", "S") & " and " & FilterVar(ToDate, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and doc_cur = " & FilterVar(strDocCur, "''", "S") & " "

		If strMgntCd1Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
		End If
		If strMgntCd1To <> "" Then
			lgStrSQL = lgStrSQL &	"	and mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
		End If
		If strMgntCd2Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
		End If
		If strMgntCd2To <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
		End If						

		lgStrSQL = lgStrSQL &	"		union all "

		lgStrSQL = lgStrSQL &	"		select a.gl_no, a.gl_seq, a.mgnt_cd1, a.mgnt_val1, a.mgnt_cd2, a.mgnt_val2, 0,0,sum(b.cls_amt) "
		lgStrSQL = lgStrSQL &	"		from a_open_acct a(nolock) "
		lgStrSQL = lgStrSQL &	"		join a_cls_acct_item b(nolock) on a.gl_no = b.gl_no and a.gl_seq = b.gl_seq "
		lgStrSQL = lgStrSQL &	"		join a_cls_acct c(nolock) on b.cls_no = c.cls_no "
		lgStrSQL = lgStrSQL &	"		where a.acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and a.doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and c.cls_dt between " & FilterVar(FromDate, "''", "S") & " and " & FilterVar(ToDate, "''", "S") & " "
		If strMgntCd1Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and a.mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
		End If
		If strMgntCd1To <> "" Then
			lgStrSQL = lgStrSQL &	"	and a.mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
		End If
		If strMgntCd2Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(a.mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
		End If
		If strMgntCd2To <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(a.mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
		End If						

		lgStrSQL = lgStrSQL &	"		group by a.gl_no, a.gl_seq, a.mgnt_cd1, a.mgnt_val1, a.mgnt_cd2,a.mgnt_val2) X "
		lgStrSQL = lgStrSQL &	"	Join A_GL_ITEM Y(NOLOCK) on X.gl_no = Y.gl_no and X.gl_seq = Y.item_seq "
		lgStrSQL = lgStrSQL &	"	where 1=1 "
		
		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then			
			lgBizAreaAuthSQL		= " AND Y.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
		End If			
		If lgInternalCd <> "" Then			
			lgInternalCdAuthSQL		= " AND Y.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
		End If			
		If lgSubInternalCd <> "" Then	
			lgSubInternalCdAuthSQL	= " AND Y.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
		End If	
		If lgAuthUsrID <> "" Then	
			lgAuthUsrIDAuthSQL		= " AND Y.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
		End If	

		lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


		
		lgStrSQL = lgStrSQL &	"	group by X.mgnt_val1, X.mgnt_val2, X.mgnt_cd1, X.mgnt_cd2 "
		
		If strProcessOption = 2 Then
			lgStrSQL = lgStrSQL &	" Having Sum(X.amt1+X.amt2-X.amt3) <> 0 "
		End If

		lgStrSQL = lgStrSQL &	" order by X.mgnt_val1, X.mgnt_val2"
									
	Else	
		lgStrSQL =				"	Select X.mgnt_cd1, X.mgnt_val1, '' 'mgnt_nm1', X.mgnt_cd2, X.mgnt_val2, '' 'mgnt_nm2', "
		lgStrSQL = lgStrSQL &	"	sum(X.amt1) 'amt1',sum(X.amt2) 'amt2',sum(X.amt3) 'amt3',sum(X.amt1+X.amt2-X.amt3) 'amt4' "
		lgStrSQL = lgStrSQL &	"	from	( "	
		lgStrSQL = lgStrSQL &	"		select gl_no, gl_seq, mgnt_cd1, mgnt_val1, mgnt_cd2, mgnt_val2, open_amt 'amt1',0 'amt2', temp_amt+cls_amt 'amt3' "
		lgStrSQL = lgStrSQL &	"		from a_open_acct(nolock) "
		lgStrSQL = lgStrSQL &	"		where acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and gl_dt < " & FilterVar(FromDate, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
		If strMgntCd1Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
		End If
		If strMgntCd1To <> "" Then
			lgStrSQL = lgStrSQL &	"	and mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
		End If
		If strMgntCd2Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
		End If
		If strMgntCd2To <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
		End If						

		If strProcessOption = 2 Then
			lgStrSQL = lgStrSQL &	"	and open_amt-temp_amt-cls_amt <> 0 "
		End If
		
		lgStrSQL = lgStrSQL &	"		union all "

		lgStrSQL = lgStrSQL &	"		select gl_no, gl_seq, mgnt_cd1, mgnt_val1, mgnt_cd2, mgnt_val2, 0, open_amt, temp_amt+cls_amt "
		lgStrSQL = lgStrSQL &	"		from a_open_acct(nolock) "
		lgStrSQL = lgStrSQL &	"		where acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and gl_dt between " & FilterVar(FromDate, "''", "S") & " and  " & FilterVar(ToDate, "''", "S") & " "
		lgStrSQL = lgStrSQL &	"		and doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
		If strMgntCd1Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
		End If
		If strMgntCd1To <> "" Then
			lgStrSQL = lgStrSQL &	"	and mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
		End If
		If strMgntCd2Fr <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
		End If
		If strMgntCd2To <> "" Then
			lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
		End If						

		If strProcessOption = 2 Then
			lgStrSQL = lgStrSQL &	"	and open_amt-temp_amt-cls_amt <> 0 "
		End If

		lgStrSQL = lgStrSQL &	"	) X Join A_GL_ITEM Y(NOLOCK) on X.gl_no = Y.gl_no and X.gl_seq = Y.item_seq "
		lgStrSQL = lgStrSQL &	"	where 1=1 "

		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then			
			lgBizAreaAuthSQL		= " AND Y.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
		End If			
		If lgInternalCd <> "" Then			
			lgInternalCdAuthSQL		= " AND Y.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
		End If			
		If lgSubInternalCd <> "" Then	
			lgSubInternalCdAuthSQL	= " AND Y.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
		End If	
		If lgAuthUsrID <> "" Then	
			lgAuthUsrIDAuthSQL		= " AND Y.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
		End If	

		lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

		lgStrSQL = lgStrSQL &	"	group by X.mgnt_val1, X.mgnt_val2, X.mgnt_cd1, X.mgnt_cd2 "
		lgStrSQL = lgStrSQL &	"	order by X.mgnt_val1, X.mgnt_val2 "
	End If
	
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
      
    If Trim(lgObjRs("mgnt_cd1")) <> "" and Trim(lgObjRs("mgnt_val1")) <> "" Then
         
		Call SubCreateCommandObject(lgObjComm)

		iCtrl_cd1 = Trim(lgObjRs("mgnt_cd1"))
		iCtrl_val1 = Trim(lgObjRs("mgnt_val1"))
		
  		
		With lgObjComm
			.CommandText = "USP_A_MGNT_NAME"
			.CommandType = adCmdStoredProc
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_CD"  ,adVarWChar,adParamInput,Len(iCtrl_cd1),iCtrl_cd1)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_VAL" ,adVarWChar,adParamInput,Len(iCtrl_val1),iCtrl_val1)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OUT_DATA_COLM_NM",adVarWChar,adParamOutput ,128)
							    
			lgObjComm.Execute ,, adExecuteNoRecords
		End With
	  
   			If  Err.number = 0 Then
				OUT_DATA_COLM_NM1 = lgObjComm.Parameters("@OUT_DATA_COLM_NM").Value
			Else
				OUT_DATA_COLM_NM1 = ""
				Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)				
			End If
		
		Call SubCloseCommandObject(lgObjComm)
		
    Else 
         OUT_DATA_COLM_NM1 = ""
        
    End If
       				
	If Trim(lgObjRs("mgnt_cd2")) <> "" and Trim(lgObjRs("mgnt_val2")) <> ""Then	

		Call SubCreateCommandObject(lgObjComm)
		       
		iCtrl_cd2 = Trim(lgObjRs("mgnt_cd2"))
		iCtrl_val2 = Trim(lgObjRs("mgnt_val2"))
			
		With lgObjComm
			.CommandText = "USP_A_MGNT_NAME"
			.CommandType = adCmdStoredProc
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_CD"  ,adVarWChar,adParamInput,Len(iCtrl_cd2),iCtrl_cd2)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_VAL" ,adVarWChar,adParamInput,Len(iCtrl_val2),iCtrl_val2)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OUT_DATA_COLM_NM",adVarWChar,adParamOutput ,128)
									    
			lgObjComm.Execute ,, adExecuteNoRecords
		End With
				
		If  Err.number = 0 Then
			OUT_DATA_COLM_NM2 = lgObjComm.Parameters("@OUT_DATA_COLM_NM").Value
		Else
			OUT_DATA_COLM_NM2 = ""
          
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
		
		End If
					
		Call SubCloseCommandObject(lgObjComm)
	Else 
		OUT_DATA_COLM_NM2 = ""
	        
	End If	
			
		
		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_cd1"))
		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_val1"))
		lgstrData = lgstrData & Chr(11) & ConvSPChars(OUT_DATA_COLM_NM1)

		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_cd2"))
		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_val2"))
		lgstrData = lgstrData & Chr(11) & ConvSPChars(OUT_DATA_COLM_NM2)

		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt1")	,ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt2")	,ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt3")	,ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("amt4")	,ggAmtOfMoney.DecPoint, 0)
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
    
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If iDx <= C_SHEETMAXROWS_D Then
	
       lgStrPrevKeyIndex = ""
    End If   

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   

	
	
    If lgErrorStatus  = "" Then
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
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
'    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
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
								' FromDate, ToDate, strAcctCd, strDocCur, strProcessOption, strQueryOption
								
								If strQueryOption = "1" Then		'시점별 
								
									lgStrSQL =				"	Select	sum(X.amt1) 'amt1', sum(X.amt2) 'amt2', sum(X.amt3) 'amt3', sum(X.amt1+X.amt2-X.amt3) 'amt4' "
									lgStrSQL = lgStrSQL &	"	from	( "
									lgStrSQL = lgStrSQL &	"		select a.gl_no, a.gl_seq, a.mgnt_cd1, a.mgnt_val1, a.mgnt_cd2, a.mgnt_val2, a.open_amt-ISNULL(b.cls_amt,0) 'amt1', 0 'amt2', 0 'amt3' "
									lgStrSQL = lgStrSQL &	"		from a_open_acct a(nolock) "
									lgStrSQL = lgStrSQL &	"		left outer join (SELECT c.GL_NO, c.GL_SEQ, ISNULL(SUM(c.CLS_AMT),0) 'CLS_AMT' "
									lgStrSQL = lgStrSQL &	"			FROM A_CLS_ACCT_ITEM c(nolock) "
									lgStrSQL = lgStrSQL &	"			join A_CLS_ACCT d(nolock) on c.CLS_NO = d.CLS_NO "
									lgStrSQL = lgStrSQL &	"			join A_OPEN_ACCT e(nolock) on c.gl_no = e.gl_no and c.gl_seq = e.gl_seq "
									lgStrSQL = lgStrSQL &	"			WHERE d.CLS_DT <= " & FilterVar(FromDate , "''", "S") & " " 
									lgStrSQL = lgStrSQL &	"			AND c.acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"			AND e.doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"			GROUP BY c.GL_NO, c.GL_SEQ ) b on a.gl_no = b.gl_no and a.gl_seq = b.gl_seq "
									lgStrSQL = lgStrSQL &	"		where a.acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and a.gl_dt < " & FilterVar(FromDate, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and a.doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
		
									If strMgntCd1Fr <> "" Then
										lgStrSQL = lgStrSQL & "		and a.mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "
									End If
									If strMgntCd1To <> "" Then
										lgStrSQL = lgStrSQL & "		and a.mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "
									End If
									If strMgntCd2Fr <> "" Then
										lgStrSQL = lgStrSQL & "		and isnull(a.mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "
									End If
									If strMgntCd2To <> "" Then
										lgStrSQL = lgStrSQL & "		and isnull(a.mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "
									End If						
											
									lgStrSQL = lgStrSQL &	"		union all "

									lgStrSQL = lgStrSQL &	"		select gl_no, gl_seq, mgnt_cd1, mgnt_val1, mgnt_cd2, mgnt_val2, 0, open_amt, 0 "
									lgStrSQL = lgStrSQL &	"		from a_open_acct(nolock) "
									lgStrSQL = lgStrSQL &	"		where acct_cd= " & FilterVar(strAcctCd, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and gl_dt between " & FilterVar(FromDate, "''", "S") & " and " & FilterVar(ToDate, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and doc_cur = " & FilterVar(strDocCur, "''", "S") & " "

									If strMgntCd1Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
									End If
									If strMgntCd1To <> "" Then
										lgStrSQL = lgStrSQL &	"	and mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
									End If
									If strMgntCd2Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
									End If
									If strMgntCd2To <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
									End If						

									lgStrSQL = lgStrSQL &	"		union all "

									lgStrSQL = lgStrSQL &	"		select a.gl_no, a.gl_seq, a.mgnt_cd1, a.mgnt_val1, a.mgnt_cd2, a.mgnt_val2, 0,0,sum(b.cls_amt) "
									lgStrSQL = lgStrSQL &	"		from a_open_acct a(nolock) "
									lgStrSQL = lgStrSQL &	"		join a_cls_acct_item b(nolock) on a.gl_no = b.gl_no and a.gl_seq = b.gl_seq "
									lgStrSQL = lgStrSQL &	"		join a_cls_acct c(nolock) on b.cls_no = c.cls_no "
									lgStrSQL = lgStrSQL &	"		where a.acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and a.doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and c.cls_dt between " & FilterVar(FromDate, "''", "S") & " and " & FilterVar(ToDate, "''", "S") & " "
									If strMgntCd1Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and a.mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
									End If
									If strMgntCd1To <> "" Then
										lgStrSQL = lgStrSQL &	"	and a.mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
									End If
									If strMgntCd2Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(a.mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
									End If
									If strMgntCd2To <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(a.mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
									End If						

									lgStrSQL = lgStrSQL &	"		group by a.gl_no, a.gl_seq, a.mgnt_cd1, a.mgnt_val1, a.mgnt_cd2,a.mgnt_val2) X "
									lgStrSQL = lgStrSQL &	"	Join A_GL_ITEM Y(NOLOCK) on X.gl_no = Y.gl_no and X.gl_seq = Y.item_seq "
									lgStrSQL = lgStrSQL &	"	where 1=1 "
		
									' 권한관리 추가 
									If lgAuthBizAreaCd <> "" Then			
										lgBizAreaAuthSQL		= " AND Y.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
									End If			
									If lgInternalCd <> "" Then			
										lgInternalCdAuthSQL		= " AND Y.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
									End If			
									If lgSubInternalCd <> "" Then	
										lgSubInternalCdAuthSQL	= " AND Y.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
									End If	
									If lgAuthUsrID <> "" Then	
										lgAuthUsrIDAuthSQL		= " AND Y.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
									End If	

									lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


		
									If strProcessOption = 2 Then
										lgStrSQL = lgStrSQL &	" Having Sum(X.amt1+X.amt2-X.amt3) <> 0 "
									End If
	
																									
								Else								'현재시점 

									lgStrSQL =				"	Select sum(X.amt1) 'amt1',sum(X.amt2) 'amt2',sum(X.amt3) 'amt3',sum(X.amt1+X.amt2-X.amt3) 'amt4' "
									lgStrSQL = lgStrSQL &	"	from	( "	
									lgStrSQL = lgStrSQL &	"		select gl_no, gl_seq, mgnt_cd1, mgnt_val1, mgnt_cd2, mgnt_val2, open_amt 'amt1',0 'amt2', temp_amt+cls_amt 'amt3' "
									lgStrSQL = lgStrSQL &	"		from a_open_acct(nolock) "
									lgStrSQL = lgStrSQL &	"		where acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and gl_dt < " & FilterVar(FromDate, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
									If strMgntCd1Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
									End If
									If strMgntCd1To <> "" Then
										lgStrSQL = lgStrSQL &	"	and mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
									End If
									If strMgntCd2Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
									End If
									If strMgntCd2To <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
									End If						

									If strProcessOption = 2 Then
										lgStrSQL = lgStrSQL &	"	and open_amt-temp_amt-cls_amt <> 0 "
									End If
		
									lgStrSQL = lgStrSQL &	"		union all "

									lgStrSQL = lgStrSQL &	"		select gl_no, gl_seq, mgnt_cd1, mgnt_val1, mgnt_cd2, mgnt_val2, 0, open_amt, temp_amt+cls_amt "
									lgStrSQL = lgStrSQL &	"		from a_open_acct(nolock) "
									lgStrSQL = lgStrSQL &	"		where acct_cd = " & FilterVar(strAcctCd, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and gl_dt between " & FilterVar(FromDate, "''", "S") & " and  " & FilterVar(ToDate, "''", "S") & " "
									lgStrSQL = lgStrSQL &	"		and doc_cur = " & FilterVar(strDocCur, "''", "S") & " "
									If strMgntCd1Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and mgnt_val1 >= " & FilterVar(strMgntCd1Fr, "''", "S") & " "		
									End If
									If strMgntCd1To <> "" Then
										lgStrSQL = lgStrSQL &	"	and mgnt_val1 <= " & FilterVar(strMgntCd1To, "''", "S") & " "		
									End If
									If strMgntCd2Fr <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'ZZZZZZZZZZ') >= " & FilterVar(strMgntCd2Fr, "''", "S") & " "		
									End If
									If strMgntCd2To <> "" Then
										lgStrSQL = lgStrSQL &	"	and isnull(mgnt_val2,'') <= " & FilterVar(strMgntCd2To, "''", "S") & " "		
									End If						

									If strProcessOption = 2 Then
										lgStrSQL = lgStrSQL &	"	and open_amt-temp_amt-cls_amt <> 0 "
									End If

									lgStrSQL = lgStrSQL &	"	) X Join A_GL_ITEM Y(NOLOCK) on X.gl_no = Y.gl_no and X.gl_seq = Y.item_seq "
									lgStrSQL = lgStrSQL &	"	where 1=1 "

									' 권한관리 추가 
									If lgAuthBizAreaCd <> "" Then			
										lgBizAreaAuthSQL		= " AND Y.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
									End If			
									If lgInternalCd <> "" Then			
										lgInternalCdAuthSQL		= " AND Y.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
									End If			
									If lgSubInternalCd <> "" Then	
										lgSubInternalCdAuthSQL	= " AND Y.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
									End If	
									If lgAuthUsrID <> "" Then	
										lgAuthUsrIDAuthSQL		= " AND Y.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
									End If	

									lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

								End If
								                                   
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
    lgErrorStatus    = "YES"
End Sub
%>

