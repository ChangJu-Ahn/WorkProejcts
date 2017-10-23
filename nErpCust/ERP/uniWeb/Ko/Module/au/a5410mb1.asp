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
	Dim lgStrPrevKey
	Dim	FromDate,ToDate
	    
	Dim strEmpNo
	 
	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


	On Error Resume Next															'☜: Protect prorgram from crashing
	Err.Clear																		'☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q","A", "NOCOOKIE", "MB")        
	    
	Call HideStatusWnd																'☜: Hide Processing message

	lgErrorStatus	= ""
	lgKeyStream		= Split(Request("txtKeyStream"),gColSep) 
	lgStrPrevKey	= Trim(Request("lgStrPrevKey"))									'☜: Next Key
	strEmpNo		= Trim(Request("txtEmpNo"))
	FromDate		= lgKeyStream(0)
	ToDate			= lgKeyStream(1)

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

	Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	    
	Select Case CStr(Request("txtMode"))
	    Case CStr(UID_M0001)                                                         '☜: Query
	         Call SubBizQuery()
	    Case CStr(UID_M0002)                                                         '☜: Save,Update
'			Call SubBizSaveMulti()
	    Case CStr(UID_M0003)                                                         '☜: Delete
'			Call SubBizDelete()
	End Select

	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
    Dim lgStrSQL

	Dim strSumAmt
	Dim strEmpNM


	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear  
                                                                          '☜: Clear Error status
	If lgStrPrevKey = "" Then
		strEmpNM = ""
		If strEmpNo <> "" Then
			lgStrSQL = " select isnull(b.name,a.use_user_id)	as	EmpNM	"
			lgStrSQL = lgStrSQL & " from    haa010t b(nolock), b_credit_card a(nolock) "
			lgStrSQL = lgStrSQL & " where b.emp_no=*a.use_user_id "
			lgStrSQL = lgStrSQL & " and   isnull(a.use_user_id,'')	<>	'' "
			lgStrSQL = lgStrSQL & " and   a.use_user_id =  " & FilterVar(strEmpNo, "''", "S")

			 If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
			 	If isNull(lgObjRs("EmpNM")) = False Then
					strEmpNM = Trim(lgObjRs("EmpNM"))
				End If
			 End if
'			 Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
			 Call SubCloseRs(lgObjRs)   
		End If		    

		lgStrSQL = lgStrSQL &	"Select	isnull(sum(a.OPEN_AMT),0) 'OpenAmt'	"
		lgStrSQL = lgStrSQL &	"From	A_OPEN_ACCT a(nolock), A_ACCT b(nolock), B_CREDIT_CARD c(nolock), A_GL_ITEM d(nolock)	"
		lgStrSQL = lgStrSQL &	"where	a.acct_cd	= b.acct_cd		"
		lgStrSQL = lgStrSQL &	"and	a.mgnt_val1	= c.Credit_No	"
		lgStrSQL = lgStrSQL &	"and	a.gl_no		= d.gl_no		"
		lgStrSQL = lgStrSQL &	"and	a.gl_seq	= d.item_seq	"
		lgStrSQL = lgStrSQL &	"and	a.gl_dt between	" & FilterVar(FromDate , "''", "S") & " and " & FilterVar(ToDate, "''", "S") & "	"
		lgStrSQL = lgStrSQL &	"and	b.mgnt_type	= " & FilterVar("6", "''", "S") & "	"

		if	Trim(strEmpNo) <> "" and isnull(strEmpNo) = false then
			lgStrSQL = lgStrSQL & " and	c.use_user_id =	" & FilterVar(strEmpNo , "''", "S") & "	"
		end if	

		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then			
			lgBizAreaAuthSQL		= " AND d.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
		End If			
		If lgInternalCd <> "" Then			
			lgInternalCdAuthSQL		= " AND d.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
		End If			
		If lgSubInternalCd <> "" Then	
			lgSubInternalCdAuthSQL	= " AND d.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
		End If	
		If lgAuthUsrID <> "" Then	
			lgAuthUsrIDAuthSQL		= " AND d.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
		End If	

		lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


	'	Response.Write lgStrSQL
		
       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = false Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'☜: No data is found.
                            
	   else

	
			If CheckSYSTEMError(Err,True) = True Then
			   ObjectContext.SetAbort
			   Exit Sub
			End If   
		
			strSumAmt = UNINumClientFormat(lgObjRs("OpenAmt"),ggAmtOfMoney.DecPoint, 0)
			
		'	Response.Write lgstrdata
		'	Response.Write strSumAmt
			
			Response.Write  " <Script Language=vbscript> " & vbCr
			Response.Write  " With Parent                " & vbCr
			Response.Write  "    .lgStrPrevKey			= """ & lgStrPrevKey	& """" & vbCr
			Response.Write  "    .frm1.txtSumAmt.text	= """ & strSumAmt		& """" & vbCr
			Response.Write	"	 .frm1.txtEmpNm.value	= """ & strEmpNM		& """" & vbCr
			Response.Write  "    .DBQueryOk   " & vbCr
			Response.Write  " End With                   " & vbCr        
			Response.Write  " </Script>                  " & vbCr

       'lgSchoolCD = lgObjRs("SchoolCD")
		
			Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
			Call SubBizQueryMulti()
       end if
	else
	   Call SubBizQueryMulti()
    End If
    
End Sub	




'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iSelCount
    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    Dim strCur
    Dim strYear,strMonth,strDay, OneDayBefore
    Dim strDataFormat
    
    Const C_SHEETMAXROWS_D  = 100														'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear																			'☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'Response.Write pSchoolCD
    
    lgStrSQL = ""
   

	lgStrSQL = lgStrSQL &	"Select	a.gl_dt, a.mgnt_val1, "
	lgStrSQL = lgStrSQL &	"	(select isnull(s.name,c.use_user_id) from haa010t s(nolock) "
	lgStrSQL = lgStrSQL &	"	where s.emp_no =* c.use_user_id and isnull(c.use_user_id,'') <> '' ) 'EmpNM',	"
	lgStrSQL = lgStrSQL &	"	c.credit_nm, a.gl_desc, a.open_amt	" 
	lgStrSQL = lgStrSQL &	"From	A_OPEN_ACCT a(nolock), A_ACCT b(nolock), B_CREDIT_CARD c(nolock), A_GL_ITEM d(nolock)	"
	lgStrSQL = lgStrSQL &	"where	a.acct_cd	= b.acct_cd		"
	lgStrSQL = lgStrSQL &	"and	a.mgnt_val1	= c.Credit_No	"
	lgStrSQL = lgStrSQL &	"and	a.gl_no		= d.gl_no		"
	lgStrSQL = lgStrSQL &	"and	a.gl_seq	= d.item_seq	"
	lgStrSQL = lgStrSQL &	"and	a.gl_dt between	" & FilterVar(FromDate , "''", "S") & " and " & FilterVar(ToDate, "''", "S") & "	"
	lgStrSQL = lgStrSQL &	"and	b.mgnt_type	= " & FilterVar("6", "''", "S") & "	"


	if	Trim(strEmpNo) <> "" and isnull(strEmpNo) = false then
		lgStrSQL = lgStrSQL & " and	c.use_user_id =	" & FilterVar(strEmpNo , "''", "S") & "	"
	end if	


	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND d.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			
	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND d.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			
	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND d.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	
	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND d.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

	lgStrSQL = lgStrSQL &	" order by 1,2,3 "


 '   Response.Write "------- Multi :" & lgStrSQL
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
    Else 

		lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

		If UniCInt(lgStrPrevKey,0) > 0 Then
			lgObjRs.Move     = C_SHEETMAXROWS_D * UniCInt(lgStrPrevKey,0)
		End If

		lgstrData = ""
		        
		iDx = 1

		Do While Not lgObjRs.EOF
			strDataFormat = replace(gServerDateFormat, gServerDateType, "")
			call ExtractDateFrom(lgObjRs("gl_dt"),strDataFormat,gServerDateType,strYear,strMonth,strDay)
			lgstrData = lgstrData & Chr(11) & MakeDateTo("YMD",gDateFormat,gComDateType ,strYear,strMonth,strDay)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mgnt_val1"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EmpNM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("credit_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gl_desc"))        
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("open_amt"),ggAmtOfMoney.DecPoint, 0)       
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext

			iDx =  iDx + 1

			If iDx > C_SHEETMAXROWS_D Then
				lgStrPrevKey = UniCInt(lgStrPrevKey,0) + 1
				Exit Do
			End If   
		Loop 
    End If
	  
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If iDx <= C_SHEETMAXROWS_D Then

       lgStrPrevKey = ""
    End If   

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   

'	Response.werite  "test_test"
    If lgErrorStatus  = "" Then
	   Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.lgStrPrevKey    = """ & lgStrPrevKey & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr
       Response.Write  " </Script>             " & vbCr
    End If
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

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


