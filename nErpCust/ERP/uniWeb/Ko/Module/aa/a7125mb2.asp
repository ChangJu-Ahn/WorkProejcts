
<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    Dim lgStrPrevKey

    On Error Resume Next                                                   '☜: Protect prorgram from crashing
    Call LoadBasisGlobalInf()    
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB") 
    gChangeOrgId = GetGlobalInf("gChangeOrgId")

    Err.Clear                                                              '☜: Clear Error status
    
    Call HideStatusWnd                                                     '☜: Hide Processing message

    lgErrorStatus  = "NO"
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep) 
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))                   '☜: Next Key

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim lgAcqNo
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("SR","x","x")                                              '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
        lgStrPrevKey  = ""
        lgErrorStatus = "YES"
        Exit Sub 
       
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " With Parent                " & vbCr             
       Response.Write  "   .Frm1.htxtAcqNo.Value     = """ & ConvSPChars(lgObjRs("Acq_no"))          & """" & vbCr

       Response.Write  "   .Frm1.txtAcqNo.Value     = """ & ConvSPChars(lgObjRs("Acq_no"))          & """" & vbCr
       Response.Write  "   .Frm1.txtDocCur.Value     = """ & ConvSPChars(lgObjRs("Doc_cur"))          & """" & vbCr

       Response.Write  "   .Frm1.txtAcqDt.Text     = """ & UNIDateClientFormat(lgObjRs("Acq_dt"))          & """" & vbCr
       Response.Write  "   .Frm1.txtBpCd.Value     = """ & ConvSPChars(lgObjRs("Bp_cd"))          & """" & vbCr
       Response.Write  "   .Frm1.txtBpNm.Value      = """ & ConvSPChars(lgObjRs("Bp_nm"))      & """" & vbCr
	   Response.Write  "   .Frm1.txtXchrate.text     = """ & UNINumClientFormat(lgObjRs("Xch_rate"),   ggExchRate.DecPoint, 0) &					"""" & vbCr                             '환율         

       Response.Write  " End With                   " & vbCr        
       Response.Write  " </Script>                  " & vbCr

       lgAcqNo =  ConvSPChars(lgObjRs("Acq_no"))
       
	   Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
       Call SubBizQueryMulti(lgAcqNo)
    End If
End Sub	


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pAcqNo)

    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    Dim cnt
    Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    lgStrSQL = "Select TOP " & C_SHEETMAXROWS_D + 1 & "	a.Dept_cd,	b.dept_nm, 	a.Acct_cd,	c.acct_nm,"
    lgStrSQL = lgStrSQL & " 		a.Asst_no,		a.Asst_nm,	a.Acq_amt,		a.Acq_loc_amt,"  
    lgStrSQL = lgStrSQL & " 		a.Acq_qty,		a.Res_amt,	a.Ref_no,		a.Asset_desc"  
    lgStrSQL = lgStrSQL & " From	a_asset_master a, b_acct_dept b, a_acct c" 
    lgStrSQL = lgStrSQL & " WHERE	a.Acq_no = " & FilterVar(pAcqNo, "''", "S")  
    lgStrSQL = lgStrSQL & " AND		a.asst_no >= " & FilterVar(lgStrPrevKey, "''", "S")
    
    lgStrSQL = lgStrSQL & " And		a.dept_cd =  b.dept_cd"
    lgStrSQL = lgStrSQL & " And		b.org_change_id = " & FilterVar(gChangeOrgId, "''", "S")
    lgStrSQL = lgStrSQL & " And		a.acct_cd =  c.acct_cd"
    lgStrSQL = lgStrSQL & " ORDER BY	a.asst_no"



    If 	FncOpenRs("U",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
    Else 
     
       iDx = 1
       lgstrData = ""
       lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
       Response.Write cnt
       Do While Not lgObjRs.EOF

		  cnt=cnt+1 	
          Response.Write cnt & VBCRLf

          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Dept_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Acct_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Asst_no"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Asst_nm"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Acq_amt"),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Acq_loc_amt"),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Acq_qty"),ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Res_amt"),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Ref_no"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Asset_desc"))
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
		lgStrPrevKey = lgObjRs("asst_no")
    Else
		lgStrPrevKey = ""
    End If
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "NO" Then
       Response.Write  " <Script Language=vbscript>                            " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey   & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData      & """" & vbCr
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

    lgStrSQL = "Select a.acq_no, a.Doc_cur, a.Acq_dt, a.bp_cd, b.bp_nm, a.xch_rate"
    lgStrSQL = lgStrSQL & " From  a_asset_acq a, b_biz_partner b "
    lgStrSQL = lgStrSQL & " WHERE a.bp_cd *= b.bp_cd "
    lgStrSQL = lgStrSQL & " AND   a.acq_no = " & FilterVar(lgKeyStream(0), "''", "S")
    
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

