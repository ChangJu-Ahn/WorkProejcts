<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->

<%

    Dim lgStrPrevKey
    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")

    lgErrorStatus  = ""
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep)
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))                   '☜: NextKey
    lgPrevNext     = Request("txtPrevNext")                          '☜: "P"(Prev search) "N"(Next search)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim hBizAreaCd
Dim hInernalCd
Dim hCostCd
Dim aChangeOrgId

aChangeOrgId = Trim(request("horgchangeid"))

Dim strAutoCardNo
Dim strDate, strYear, strMonth, strDay

Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인 

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case CStr(Request("txtMode"))                                             '☜: Read Operation Mode (CRUD)
        Case CStr(UID_M0001)                                                         '☜: Query
             Select Case lgPrevNext
                Case "N","P","Q" : Call SubBizQuery()                                'Next										
                Case "R"         :   Call SubBizQueryMulti(lgKeyStream(0))             '
             End Select 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim lgCardNo
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
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
       Response.Write  " <Script Language=vbscript>	" & vbCr
       Response.Write  " With Parent							" & vbCr             
       Response.Write  "   .Frm1.txtCardNoQry.Value    = """ & ConvSPChars(lgObjRs("NOTE_NO"))				& """" & vbCr
'       Response.Write  "   .Frm1.txtCardNo.Value			= """ & ConvSPChars(lgObjRs("NOTE_NO"))			& """" & vbCr
       Response.Write  "   .Frm1.htxtInternalCd.Value	= """ & ConvSPChars(lgObjRs("INTERNAL_CD"))		& """" & vbCr
       Response.Write  "   .Frm1.htxtBizAreaCd.Value	= """ & ConvSPChars(lgObjRs("BIZ_AREA_CD"))		& """" & vbCr
       Response.Write  "   .Frm1.horgchangeid.Value	= """ & ConvSPChars(lgObjRs("ORG_CHANGE_ID"))		& """" & vbCr
       Response.Write  "   .Frm1.htxtCostCd.Value		= """ & ConvSPChars(lgObjRs("COST_CD"))			& """" & vbCr       
       
       Response.Write  "   .Frm1.txtIssueDt.Text			= """ & UniDateClientFormat(lgObjRs("ISSUE_DT"))	& """" & vbCr
       Response.Write  "   .Frm1.txtDueDt.Text			= """ & UniDateClientFormat(lgObjRs("DUE_DT"))	& """" & vbCr
       Response.Write  "   .Frm1.txtDeptCD.Value			= """ & ConvSPChars(lgObjRs("DEPT_CD"))				& """" & vbCr
       Response.Write  "   .Frm1.txtDeptNm.Value		= """ & ConvSPChars(lgObjRs("DEPT_NM"))				& """" & vbCr
       Response.Write  "   .Frm1.txtBpCd.Value			= """ & ConvSPChars(lgObjRs("BP_CD"))					& """" & vbCr
       Response.Write  "   .Frm1.txtBpNM.Value			= """ & ConvSPChars(lgObjRs("BP_NM"))				& """" & vbCr
       '2003-02-18 카드사/은행 정보 모두 QUERY
       If  Trim(ConvSPChars(lgObjRs("CARD_CO_CD")))  <> ""  Then 		       
			Response.Write  "   .Frm1.txtCardCoCd.Value		= """ & ConvSPChars(lgObjRs("CARD_CO_CD"))		& """" & vbCr
			Response.Write  "   .Frm1.txtCardCoNm.Value	= """ & ConvSPChars(lgObjRs("CARD_CO_NM"))		& """" & vbCr
	   Else 
			Response.Write  "   .Frm1.txtCardCoCd.Value		= """ & ConvSPChars(lgObjRs("BANK_CD"))		& """" & vbCr
			Response.Write  "   .Frm1.txtCardCoNm.Value	= """ & ConvSPChars(lgObjRs("BANK_NM"))		& """" & vbCr
	   End If 
	   
       Response.Write  "   .Frm1.txtCardAmt.Value       = """ & UNINumClientFormat(lgObjRs("NOTE_AMT"), ggAmtOfMoney.DecPoint, 0)   & """" & vbCr
       Response.Write  "   .Frm1.txtSttlAmt.Value			= """ & UNINumClientFormat(lgObjRs("STTL_AMT"), ggAmtOfMoney.DecPoint, 0)           & """" & vbCr
       Response.Write  "   .Frm1.txtCardDesc.Value      = """ & ConvSPChars(lgObjRs("NOTE_DESC"))			& """" & vbCr
       Response.Write  "   .Frm1.htxtCardNo.Value		= """ & ConvSPChars(lgObjRs("NOTE_NO"))			& """" & vbCr       
       Response.Write  " End With                   " & vbCr                
       Response.Write  " </Script>                  " & vbCr

       lgCardNo = lgObjRs("NOTE_NO")

	   Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet	    
       Call SubBizQueryMulti(lgCardNo)
     
    End If
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    Dim lgIntFlgMode

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    ' -- 권한체크 
    If ChkAuth() = False Then Exit Sub
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    
    Select Case lgIntFlgMode
        Case  OPMD_CMODE  : Call SubBizSaveSingleCreate()                            '☜ : Create
        Case  OPMD_UMODE  : Call SubBizSaveSingleUpdate()                            '☜ : Update
    End Select

End Sub	


' -- 권한관리 
Function ChkAuth()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    ' 권한관리용 변수 define
    Dim iVarOutPut1 , iStrSQL
    Dim objAChkDataAuth 

    Dim L1_a_data_auth_cud_char 
    
    Dim L2_a_pgm_value
    Const L2_a_pgm_value_dept_cd = 0
    Const L2_a_pgm_value_internal_cd = 1
    Const L2_a_pgm_value_biz_area_cd = 2
    Const L2_a_pgm_value_updt_user_id = 3

	' -- 권한관리추가 
	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(L2_a_pgm_value_dept_cd)			= Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(L2_a_pgm_value_internal_cd)		= Trim(Request("txthhInternalCd"))
	I1_a_data_auth(L2_a_pgm_value_biz_area_cd)		= Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(L2_a_pgm_value_updt_user_id)     = Trim(Request("txthAuthUsrID"))

	ReDim L2_a_pgm_value(3) 

	ChkAuth = False

    ' -- 권한관리 추가 2006-08-01 JYK start .. CHOE0TAE 커스터마이징한것 
    ' -- 권한 DLL 에 넘기는 명령어 
    Select Case CInt(Request("txtFlgMode"))
        Case  OPMD_CMODE  : L1_a_data_auth_cud_char = "C"                           '☜ : Create
        Case  OPMD_UMODE  : L1_a_data_auth_cud_char = "U"                           '☜ : Update
        Case  Else		  : L1_a_data_auth_cud_char = "D"	
    End Select
    
    ' -- 조회 SQL
    iStrSQL = ""
    iStrSQL = iStrSQL & "DECLARE @DEPT_CD    CHAR(10) " & vbCrLf & _
                        ",   @INTERNAL_CD    CHAR(30) " & vbCrLf & _
                        ",   @BIZ_AREA_CD    VARCHAR(10) " & vbCrLf & _
                        ",   @UPDT_USER_ID   VARCHAR(13) " & vbCrLf
        
    If UCase(Trim(L1_a_data_auth_cud_char)) = "C" Or UCase(Trim(L1_a_data_auth_cud_char)) = "U" Then
        'asp에서 입력받은 부서정보 
        'asp에서 입력받은 내부부서코드/입력받은 조직개편아이디(org_change_id)와 부서코드(dept_cd)로 펫치한 내부부서코드(internal_cd)
        'asp에서 입력받은 사업장코드 / 입력받은 조직개편아이디와 부서코드로 cost_cd를 찾고 b_cost_center에서 cost_cd로 펫치한 사업장코드(biz_area_cd)

        iStrSQL = iStrSQL & "SELECT  @INTERNAL_CD = A.INTERNAL_CD " & vbCrLf & _
                            ",   @BIZ_AREA_CD = B.BIZ_AREA_CD " & vbCrLf & _
                            ",   @DEPT_CD = " & FilterVar(Request("txtDeptCD"), "''", "S") & vbCrLf & _
                            "FROM    B_ACCT_DEPT  A " & vbCrLf & _
                            "    INNER JOIN B_COST_CENTER B ON A.COST_CD = B.COST_CD " & vbCrLf & _
                            "WHERE   A.ORG_CHANGE_ID = " & FilterVar(aChangeOrgId, "''", "S") & vbCrLf & _
                            "AND A.DEPT_CD = " & FilterVar(Request("txtDeptCD"), "''", "S") & vbCrLf
    End If
        
    If UCase(Trim(L1_a_data_auth_cud_char)) = "U" Then
         '수정대상 데이터의 수정자ID
             
        iStrSQL = iStrSQL & "SELECT  @UPDT_USER_ID = UPDT_USER_ID " & vbCrLf & _
                            "FROM    F_NOTE " & vbCrLf & _
                            "WHERE   NOTE_NO = " & FilterVar(Trim(Request("txtCardNoQry")), "''", "S") & vbCrLf

    ElseIf UCase(Trim(L1_a_data_auth_cud_char)) = "D" Then
        '삭제대상 데이터의 부서코드 
        '삭제대상 데이터의 내부부서코드 
        '삭제대상 데이터의 사업장코드 
        '삭제대상 데이터의 수정자ID
        
        iStrSQL = iStrSQL & "SELECT  @DEPT_CD = DEPT_CD " & vbCrLf & _
                            ",   @INTERNAL_CD = INTERNAL_CD " & vbCrLf & _
                            ",   @BIZ_AREA_CD = BIZ_AREA_CD " & vbCrLf & _
                            ",   @UPDT_USER_ID = UPDT_USER_ID " & vbCrLf & _
                            "FROM    F_NOTE " & vbCrLf & _
                            "WHERE   NOTE_NO = " & FilterVar(Trim(Request("txtCardNoQry")), "''", "S") & vbCrLf
    End If
        
    ' -- 데이타 리턴 
    iStrSQL = iStrSQL & "SELECT @DEPT_CD DEPT_CD, @INTERNAL_CD INTERNAL_CD, @BIZ_AREA_CD BIZ_AREA_CD, @UPDT_USER_ID UPDT_USER_ID"
        
    If 	FncOpenRs("R", lgObjConn, lgObjRs, iStrSQL, "X", "X") = True Then
    
        L2_a_pgm_value(L2_a_pgm_value_dept_cd) = lgObjRs(0)
        L2_a_pgm_value(L2_a_pgm_value_internal_cd) = lgObjRs(1)
        L2_a_pgm_value(L2_a_pgm_value_biz_area_cd) = lgObjRs(2)
        L2_a_pgm_value(L2_a_pgm_value_updt_user_id) = lgObjRs(3)
    End If
            
    ' -- 권한관리 호출 
    Set objAChkDataAuth = Server.CreateObject("PA0CG07.cAChkDataAuthSvr")
            
    Call objAChkDataAuth.A_CHECK_DATA_AUTH_SVR(gStrGlobalCollection, L1_a_data_auth_cud_char, I1_a_data_auth, L2_a_pgm_value)
            
    Set objAChkDataAuth = Nothing

    If CheckSYSTEMError(Err,True) = True Then
		lgErrorStatus = "YES"
		ObjectContext.SetAbort
		Exit Function
    End If
                
    ' -- 권한관리 추가 2006-08-01 JYK end
    
    ChkAuth = True
End Function
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record    

    Call CommonQueryRs("count(*)", "f_note_item a, f_note b", _ 
									"a.note_no = b.note_no and a.note_no= " & FilterVar(Request("txtCardNoQry"), "''", "S") & " ", _
									lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	If Cint(lgF0) <> 0 Then
		Call DisplayMsgBox("141229", vbInformation, "", "", I_MKSCRIPT)               '☜ : Note Delete or Update : F_NOTE_ITEM AE 
		lgErrorStatus = "YES" 
		Exit Sub
	End If
	
	lgStrSQL = "DELETE  F_NOTE"
    lgStrSQL = lgStrSQL & " WHERE NOTE_NO   = " & FilterVar(Request("txtCardNoQry"), "''", "S")    

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       Response.Write  " <Script Language=vbscript>	" & vbCr
       Response.Write  "       Parent.DbDeleteOk			" & vbCr
       Response.Write  " </Script>							" & vbCr
    End If   

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pCardCD)

    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
           
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status		
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		lgStrSQL = "SELECT " 
		lgStrSQL = lgStrSQL & " B.STS_DT,					B.SEQ,   			B.AMT,		B.NOTE_ACCT_CD, " 
		lgStrSQL = lgStrSQL & " C.ACCT_NM,				B.TEMP_GL_NO,	B.GL_NO" 		
		lgStrSQL = lgStrSQL & " FROM " 
		lgStrSQL = lgStrSQL & " F_NOTE A	INNER JOIN   F_NOTE_ITEM B "
		lgStrSQL = lgStrSQL & "					ON (A.NOTE_NO = B.NOTE_NO  AND A.NOTE_FG = B.NOTE_FG ) "
		lgStrSQL = lgStrSQL & "					INNER JOIN   A_ACCT C "
		lgStrSQL = lgStrSQL & "					ON B.NOTE_ACCT_CD = C.ACCT_CD	 "		
		lgStrSQL = lgStrSQL & " WHERE A.NOTE_FG = " & FilterVar("CR", "''", "S") & "  "      
		lgStrSQL = lgStrSQL & " AND A.NOTE_NO = " & FilterVar(pCardCD   , "''", "S")  
		lgStrSQL = lgStrSQL & " AND B.NOTE_STS = ''  "		
		lgStrSQL = lgStrSQL & " AND (B.TEMP_GL_NO <>'' OR B.GL_NO <> '' )   "  		
		lgStrSQL = lgStrSQL & " ORDER BY B.SEQ"	


	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
'	    Call DisplayMsgBox("141300", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
'	    lgStrPrevKey  = ""
'	    lgErrorStatus = "YES"
'	    Exit Sub 
	Else    
       iDx = 1
       lgstrData = ""
       lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
				
       Do While Not lgObjRs.EOF 
		  lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("STS_DT"))	
		  lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ"))          
		  lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AMT"),ggAmtOfMoney.DecPoint, 0)	
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NOTE_ACCT_CD"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
          lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_NO"))                                                  
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)
  
           lgObjRs.MoveNext

          iDx =  iDx + 1
          If iDx > C_SHEETMAXROWS_D Then
             Exit Do
         End If   
      Loop 
	End If
	
    If Not lgObjRs.EOF Then
       lgStrPrevKey = lgObjRs("NOTE_NO")
    Else
       lgStrPrevKey = ""
    End If       
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   
    
'     If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "    Parent.ggoSpread.Source = Parent.frm1.vspdData       " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DBQueryOk " & vbCr      
       Response.Write  " </Script>				" & vbCr
'    End If
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim lgStrSQL
    Dim tmpDate
	Dim strDeptCd
	Dim hOrgChangeId
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status	
	
	Call ExtractDateFrom(Request("txtIssueDt"), gDateFormat, gComDateType, strYear, strMonth, strDay)
		
	If Trim(Request("txtCardNoQry")) = "" Then 
		Call CardPayAutoNum()		
	Else 
	
		If UCase(Left(Trim(Request("txtCardNoQry")),2)) = "CS" Then
          Call DisplayMsgBox("120714", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
          lgErrorStatus = "YES" 
          Exit Sub
		End If
		
		strAutoCardNo = strYear & strMonth & strDay & "-" & Trim(Request("txtCardNoQry"))	
	%>
	<Script Language="VBScript">
		Parent.frm1.txtCardNoQry.value =  "<%=strAutoCardNo%>"
		Parent.frm1.htxtCardNo.value =  "<%=strAutoCardNo%>"
	</script>
	<%			
	End If		
	
		
	If Trim(Request("htxtBizAreaCd")) = "" or Trim(Request("htxtInternalCd")) = "" or Trim(Request("htxtCostCd"))  = "" Then
		Call CommonQueryRs("a.dept_cd, a.dept_nm, a.cost_cd, a.internal_cd, c.biz_area_cd", "b_acct_dept a, b_cost_center b, b_biz_area c", _ 
									"a.org_change_id =  " & FilterVar(aChangeOrgId, "''", "S") & " AND a.dept_cd= " & FilterVar(Request("txtDeptCD"), "''", "S") & " AND a.cost_cd = b.cost_cd AND b.biz_area_cd = c.biz_area_cd ", _
									lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		hBizAreaCd	= Trim(Replace(lgF4,Chr(11),""))
		hInernalCd	= Trim(Replace(lgF3,Chr(11),""))
		hCostCd		= Trim(Replace(lgF2,Chr(11),""))	
	Else 
		hBizAreaCd	=  Request("htxtBizAreaCd") 
		hInernalCd	=  Request("htxtInternalCd") 
		hCostCd		=  Request("htxtCostCd") 
	End If

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------  
    lgStrSQL = "INSERT INTO F_NOTE"
    lgStrSQL = lgStrSQL & " ( NOTE_NO      , "
    lgStrSQL = lgStrSQL & "   NOTE_FG      , "
    lgStrSQL = lgStrSQL & "   NOTE_AMT , "
    lgStrSQL = lgStrSQL & "   STTL_AMT       , "
    lgStrSQL = lgStrSQL & "   ISSUE_DT         , "
    lgStrSQL = lgStrSQL & "   DUE_DT  , "
    lgStrSQL = lgStrSQL & "   PLACE     , "
    lgStrSQL = lgStrSQL & "   RCPT_FG      , "
    lgStrSQL = lgStrSQL & "   PUBLISHER , "
    lgStrSQL = lgStrSQL & "   NOTE_STS       , "
    lgStrSQL = lgStrSQL & "   NOTE_DESC         , "
    lgStrSQL = lgStrSQL & "   BP_CD  , "   
    lgStrSQL = lgStrSQL & "   BANK_CD     , "
    lgStrSQL = lgStrSQL & "   BIZ_AREA_CD      , "
    lgStrSQL = lgStrSQL & "   ORG_CHANGE_ID , "
    lgStrSQL = lgStrSQL & "   DEPT_CD       , "
    lgStrSQL = lgStrSQL & "   INTERNAL_CD         , "
    lgStrSQL = lgStrSQL & "   COST_CD  , "   
    lgStrSQL = lgStrSQL & "   ENDORSE_FG , "
    lgStrSQL = lgStrSQL & "   BP_ENDORSE_CD       , "
    lgStrSQL = lgStrSQL & "   BP_ORG_CD         , "
    lgStrSQL = lgStrSQL & "   USED_FG  , "    
    lgStrSQL = lgStrSQL & "   INSRT_USER_ID       , "
    lgStrSQL = lgStrSQL & "   INSRT_DT         , "
    lgStrSQL = lgStrSQL & "   UPDT_USER_ID       , "
    lgStrSQL = lgStrSQL & "   UPDT_DT         , "
    lgStrSQL = lgStrSQL & "   CASH_RATE  , "
    lgStrSQL = lgStrSQL & "   CASH_AMT  , "                        
    lgStrSQL = lgStrSQL & "   CARD_CO_CD      ) "
    lgStrSQL = lgStrSQL & "  VALUES(" & FilterVar(Trim(strAutoCardNo)	, "''", "S")  & ","
    lgStrSQL = lgStrSQL &					"" & FilterVar("CR", "''", "S") & " " & ","
    lgStrSQL = lgStrSQL &					UNIConvNum(Request("txtCardAmt"),0)              & ","
    lgStrSQL = lgStrSQL &					0              & ","
    lgStrSQL = lgStrSQL &					FilterVar(UniConvDate(Request("txtIssueDt"))     , "''", "S") & ","
    lgStrSQL = lgStrSQL &					FilterVar(UniConvDate(Request("txtDueDt"))     , "''", "S") & ","
    lgStrSQL = lgStrSQL &					"''" & ","
    lgStrSQL = lgStrSQL &					"''" & ","
    lgStrSQL = lgStrSQL &					"''" & ","
    lgStrSQL = lgStrSQL &					"" & FilterVar("BG", "''", "S") & " " & ","
    lgStrSQL = lgStrSQL &					FilterVar(            Request("txtCardDesc"), "''", "S") & ","
    lgStrSQL = lgStrSQL &					FilterVar(            Request("txtBpCd"), "''", "S") & ","
    lgStrSQL = lgStrSQL &					"NULL"	& ","
    lgStrSQL = lgStrSQL &					FilterVar(            hBizAreaCd , "''", "S") & ","
    lgStrSQL = lgStrSQL &					FilterVar(			aChangeOrgId				, "''", "S")			& ", "
    lgStrSQL = lgStrSQL &					FilterVar(            Request("txtDeptCD"), "''", "S") & ","
    lgStrSQL = lgStrSQL &					FilterVar(            hInernalCd , "''", "S") & ","
    lgStrSQL = lgStrSQL &					FilterVar(            hCostCd , "''", "S") & ","   
    lgStrSQL = lgStrSQL &					"" & FilterVar("CR", "''", "S") & " " & ","
    lgStrSQL = lgStrSQL &					"''" & ","
    lgStrSQL = lgStrSQL &					"''" & ","
    lgStrSQL = lgStrSQL &					"" & FilterVar("N", "''", "S") & " " & ","
    lgStrSQL = lgStrSQL &					FilterVar(		gUsrId				, "''", "S")					& ", "
	lgStrSQL = lgStrSQL &					FilterVar(		GetSvrDateTime		,NULL,"S")			& ", " 					
	lgStrSQL = lgStrSQL &					FilterVar(		gUsrId				, "''", "S")					& ", " 	
	lgStrSQL = lgStrSQL &					FilterVar(		GetSvrDateTime		,NULL,"S")			& ", " 					
	lgStrSQL = lgStrSQL &					0 & ","
	lgStrSQL = lgStrSQL &					0 & ","		
	lgStrSQL = lgStrSQL &					FilterVar(           UCase(Request("txtCardCoCd")), "''", "S") & ")"

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else       
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Parent.DBSaveOk     " & vbCr       
       Response.Write  " </Script>                  " & vbCr
    End If   

End Sub
'============================================================================================================
' Name : CardPayAutoNum
' Desc : Batch		'autonumbering
'============================================================================================================
Sub CardPayAutoNum()
    Dim IntRetCD
    Dim strYYYYMMDD
    Dim strYYYY, strMM, strDD
    CONST CALLSPNAME = "usp_a_tempgl_no_auto_gen"
    CONST AUTONUMPREFIX = "CS"    

    Call SubCreateCommandObject(lgObjComm)

    Call ExtractDateFrom(Request("txtIssueDt"), gDateFormat, gComDateType, strYYYY, strMM, strDD)        
    strYYYYMMDD = strYYYY & strMM & strDD
    

    With lgObjComm
        .CommandText = CALLSPNAME			'CALLSPNAME
        .CommandType = adCmdStoredProc
		.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
		.Parameters.Append lgObjComm.CreateParameter("@type",adVarWChar,adParamInput,Len(Trim(AUTONUMPREFIX)),AUTONUMPREFIX)
		.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date", adVarWChar,adParamInput,Len(Trim(strYYYYMMDD)), strYYYYMMDD)
		.Parameters.Append lgObjComm.CreateParameter("@usr_id", adVarWChar,adParamInput,Len(Trim(gUsrID)), gUsrID)
		.Parameters.Append lgObjComm.CreateParameter("@last_auto_no", adVarWChar, adParamOutput, 18)
		.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD = 1 then
            strAutoCardNo = lgObjComm.Parameters("@last_auto_no").Value
            strAutoCardNo = strYear & strMonth & strDay & "-" & strAutoCardNo
%>
<Script Language="VBScript">
	Parent.frm1.txtCardNoQry.value =  "<%=strAutoCardNo%>"
	Parent.frm1.htxtCardNo.value =  "<%=strAutoCardNo%>"
</script>
<%
'			Call SubMakeSQLStatements("SC")                                              '☜ : Make sql statements , SC : Single Create
        end if
    Else 
        lgErrorStatus     = "YES"                                                         '☜: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End if

    Call SubCloseCommandObject(lgObjComm)

    If lgErrorStatus    = "YES" Then
       lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
    End If
End Sub
'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
   '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    Call CommonQueryRs("count(*)", "f_note_item a, f_note b", _ 
									"a.note_no = b.note_no and a.note_no= " & FilterVar(Request("txtCardNoQry"), "''", "S") & " ", _
									lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	If Cint(lgF0) <> 0 Then
		Call DisplayMsgBox("141229", vbInformation, "", "", I_MKSCRIPT)              '☜ : Note Delete or Update : F_NOTE_ITEM AE 
		lgErrorStatus = "YES" 
		Exit Sub
	End If	
        
    lgStrSQL = "UPDATE  F_NOTE"
    lgStrSQL = lgStrSQL & " SET "  
    lgStrSQL = lgStrSQL & "   NOTE_AMT	= " & UNIConvNum(Request("txtCardAmt"),0)  & ","
    lgStrSQL = lgStrSQL & "   ISSUE_DT		= " & FilterVar(UniConvDate(Request("txtIssueDt"))     , "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   DUE_DT		= " & FilterVar(UniConvDate(Request("txtDueDt"))     , "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   NOTE_DESC	= " & FilterVar(            Request("txtCardDesc"), "''", "S")   & ","
    lgStrSQL = lgStrSQL & "   BP_CD			= " & FilterVar(            Request("txtBpCd"), "''", "S")   & ","
    lgStrSQL = lgStrSQL & "   BIZ_AREA_CD		= " & FilterVar(            Request("htxtBizAreaCd"), "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   ORG_CHANGE_ID	= " & FilterVar(			aChangeOrgId				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   DEPT_CD				= " & FilterVar(            Request("txtDeptCD"), "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   INTERNAL_CD		= " & FilterVar(            Request("htxtInternalCd"), "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   COST_CD				= " & FilterVar(            Request("htxtCostCd"), "''", "S")  & ","
    lgStrSQL = lgStrSQL & "   UPDT_USER_ID		= " & FilterVar(			gUsrId				, "''", "S")  & ","                
    lgStrSQL = lgStrSQL & "   UPDT_DT				= " & FilterVar(			GetSvrDateTime		,NULL,"S")  & ","                      
    lgStrSQL = lgStrSQL & "   CARD_CO_CD		= " & FilterVar(            Request("txtCardCoCd"), "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE NOTE_NO		= " & FilterVar(           UCase(Request("txtCardNoQry"))     , "''", "S")                                                

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Parent.DBSaveOk      " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If   


End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pSchoolCD,arrColVal)
    Dim iSelCount, strCond

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If  

    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case "Q"
									lgStrSQL = "SELECT 	A.NOTE_NO, "
									lgStrSQL = lgStrSQL & " A.NOTE_FG, "
									lgStrSQL = lgStrSQL & " A.NOTE_AMT, "
									lgStrSQL = lgStrSQL & " A.STTL_AMT, "
									lgStrSQL = lgStrSQL & " A.ISSUE_DT, " 
									lgStrSQL = lgStrSQL & " A.DUE_DT,"
									lgStrSQL = lgStrSQL & " A.NOTE_STS,"
									lgStrSQL = lgStrSQL & " A.NOTE_DESC,"
									lgStrSQL = lgStrSQL & " F.BP_CD,"
									lgStrSQL = lgStrSQL & " F.BP_NM,"
									lgStrSQL = lgStrSQL & " D.BIZ_AREA_CD,"
									lgStrSQL = lgStrSQL & " B.ORG_CHANGE_ID, "
									lgStrSQL = lgStrSQL & " B.DEPT_CD,"
									lgStrSQL = lgStrSQL & " B.DEPT_NM,"
									lgStrSQL = lgStrSQL & " B.INTERNAL_CD,"
									lgStrSQL = lgStrSQL & " C.COST_CD,"
									lgStrSQL = lgStrSQL & " E.CARD_CO_CD,"
									lgStrSQL = lgStrSQL & " E.CARD_CO_NM,	"
									lgStrSQL = lgStrSQL & " G.BANK_CD,"
									lgStrSQL = lgStrSQL & " G.BANK_NM	"		
									lgStrSQL = lgStrSQL & " FROM 	F_NOTE		A,	"
									lgStrSQL = lgStrSQL & " B_ACCT_DEPT	B,			"
									lgStrSQL = lgStrSQL & " B_COST_CENTER		C,	"
									lgStrSQL = lgStrSQL & " B_BIZ_AREA		D,		"
									lgStrSQL = lgStrSQL & " B_CARD_CO		E,		"
									lgStrSQL = lgStrSQL & " B_BIZ_PARTNER		F,	"
									lgStrSQL = lgStrSQL & " B_BANK		G	"
									lgStrSQL = lgStrSQL & " WHERE A.DEPT_CD = B.DEPT_CD		"	
									lgStrSQL = lgStrSQL & " AND A.ORG_CHANGE_ID = B.ORG_CHANGE_ID	"
									lgStrSQL = lgStrSQL & " AND B.COST_CD = C.COST_CD						"
									lgStrSQL = lgStrSQL & " AND C.BIZ_AREA_CD = D.BIZ_AREA_CD			"
									lgStrSQL = lgStrSQL & " AND A.CARD_CO_CD *= E.CARD_CO_CD			"
									lgStrSQL = lgStrSQL & " AND A.BANK_Cd *= G.BANK_Cd						"
									lgStrSQL = lgStrSQL & " AND A.BP_CD = F.BP_CD								"
									lgStrSQL = lgStrSQL & " AND A.NOTE_NO = " & FilterVar(lgKeyStream(0), "''", "S")  
									
									lgStrSQL = lgStrSQL  & vbCrLf & strCond & vbCrLf
                       
                             Case "P"
									lgStrSQL = "SELECT 	A.NOTE_NO, "
									lgStrSQL = lgStrSQL & " A.NOTE_FG, "
									lgStrSQL = lgStrSQL & " A.NOTE_AMT, "
									lgStrSQL = lgStrSQL & " A.STTL_AMT, "
									lgStrSQL = lgStrSQL & " A.ISSUE_DT, " 
									lgStrSQL = lgStrSQL & " A.DUE_DT,"
									lgStrSQL = lgStrSQL & " A.NOTE_STS,"
									lgStrSQL = lgStrSQL & " A.NOTE_DESC,"
									lgStrSQL = lgStrSQL & " F.BP_CD,"
									lgStrSQL = lgStrSQL & " D.BIZ_AREA_CD,"
									lgStrSQL = lgStrSQL & " B.ORG_CHANGE_ID, "
									lgStrSQL = lgStrSQL & " B.DEPT_CD,"
									lgStrSQL = lgStrSQL & " B.DEPT_NM,"
									lgStrSQL = lgStrSQL & " B.INTERNAL_CD,"
									lgStrSQL = lgStrSQL & " C.COST_CD,"
									lgStrSQL = lgStrSQL & " E.CARD_CO_CD,"
									lgStrSQL = lgStrSQL & " E.CARD_CO_NM,	"
									lgStrSQL = lgStrSQL & " G.BANK_CD,"
									lgStrSQL = lgStrSQL & " G.BANK_NM	"		
									lgStrSQL = lgStrSQL & " FROM 	F_NOTE		A,	"
									lgStrSQL = lgStrSQL & " B_ACCT_DEPT	B,			"
									lgStrSQL = lgStrSQL & " B_COST_CENTER		C,	"
									lgStrSQL = lgStrSQL & " B_BIZ_AREA		D,		"
									lgStrSQL = lgStrSQL & " B_CARD_CO		E,		"
									lgStrSQL = lgStrSQL & " B_BIZ_PARTNER		F,	"
									lgStrSQL = lgStrSQL & " B_BANK		G	"
									lgStrSQL = lgStrSQL & " WHERE A.NOTE_FG = " & " " & FilterVar("CR", "''", "S") & "  "
									lgStrSQL = lgStrSQL & " AND A.DEPT_CD = B.DEPT_CD		"		
									lgStrSQL = lgStrSQL & " AND A.ORG_CHANGE_ID = B.ORG_CHANGE_ID	"
									lgStrSQL = lgStrSQL & " AND B.COST_CD = C.COST_CD						"
									lgStrSQL = lgStrSQL & " AND C.BIZ_AREA_CD = D.BIZ_AREA_CD			"
									lgStrSQL = lgStrSQL & " AND A.CARD_CO_CD *= E.CARD_CO_CD			"
									lgStrSQL = lgStrSQL & " AND A.BANK_Cd *= G.BANK_Cd						"
									lgStrSQL = lgStrSQL & " AND A.BP_CD = F.BP_CD								"
									lgStrSQL = lgStrSQL & " AND A.NOTE_NO < " & FilterVar(lgKeyStream(0), "''", "S")  
									
									lgStrSQL = lgStrSQL  & vbCrLf & strCond & vbCrLf
									
									lgStrSQL = lgStrSQL & " ORDER BY A.NOTE_NO DESC "

                             Case "N"
									lgStrSQL = "SELECT 	A.NOTE_NO, "
									lgStrSQL = lgStrSQL & " A.NOTE_FG, "
									lgStrSQL = lgStrSQL & " A.NOTE_AMT, "
									lgStrSQL = lgStrSQL & " A.STTL_AMT, "
									lgStrSQL = lgStrSQL & " A.ISSUE_DT, " 
									lgStrSQL = lgStrSQL & " A.DUE_DT,"
									lgStrSQL = lgStrSQL & " A.NOTE_STS,"
									lgStrSQL = lgStrSQL & " A.NOTE_DESC,"
									lgStrSQL = lgStrSQL & " F.BP_CD,"
									lgStrSQL = lgStrSQL & " D.BIZ_AREA_CD,"
									lgStrSQL = lgStrSQL & " B.ORG_CHANGE_ID, "
									lgStrSQL = lgStrSQL & " B.DEPT_CD,"
									lgStrSQL = lgStrSQL & " B.DEPT_NM,"
									lgStrSQL = lgStrSQL & " B.INTERNAL_CD,"
									lgStrSQL = lgStrSQL & " C.COST_CD,"
									lgStrSQL = lgStrSQL & " E.CARD_CO_CD,"
									lgStrSQL = lgStrSQL & " E.CARD_CO_NM,	"
									lgStrSQL = lgStrSQL & " G.BANK_CD,"
									lgStrSQL = lgStrSQL & " G.BANK_NM	"		
									lgStrSQL = lgStrSQL & " FROM 	F_NOTE		A,	"
									lgStrSQL = lgStrSQL & " B_ACCT_DEPT	B,			"
									lgStrSQL = lgStrSQL & " B_COST_CENTER		C,	"
									lgStrSQL = lgStrSQL & " B_BIZ_AREA		D,		"
									lgStrSQL = lgStrSQL & " B_CARD_CO		E,		"
									lgStrSQL = lgStrSQL & " B_BIZ_PARTNER		F,	"
									lgStrSQL = lgStrSQL & " B_BANK		G	"
									lgStrSQL = lgStrSQL & " WHERE A.NOTE_FG = " & " " & FilterVar("CR", "''", "S") & "  "
									lgStrSQL = lgStrSQL & " AND A.DEPT_CD = B.DEPT_CD		"		
									lgStrSQL = lgStrSQL & " AND A.ORG_CHANGE_ID = B.ORG_CHANGE_ID	"
									lgStrSQL = lgStrSQL & " AND B.COST_CD = C.COST_CD						"
									lgStrSQL = lgStrSQL & " AND C.BIZ_AREA_CD = D.BIZ_AREA_CD			"
									lgStrSQL = lgStrSQL & " AND A.CARD_CO_CD *= E.CARD_CO_CD			"
									lgStrSQL = lgStrSQL & " AND A.BANK_Cd *= G.BANK_Cd						"
									lgStrSQL = lgStrSQL & " AND A.BP_CD = F.BP_CD								"
									lgStrSQL = lgStrSQL & " AND A.NOTE_NO > " & FilterVar(lgKeyStream(0), "''", "S")  
									
									lgStrSQL = lgStrSQL  & vbCrLf & strCond & vbCrLf
									
									lgStrSQL = lgStrSQL & " ORDER BY A.NOTE_NO ASC "							                                 
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

