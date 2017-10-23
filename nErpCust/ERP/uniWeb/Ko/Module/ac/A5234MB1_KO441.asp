<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
    Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
    Call LoadBasisGlobalInf()

    Dim lgStrPrevKey
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	Dim StrNo
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
	DIM StrId
    lgErrorStatus     = ""
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '¢Ð: Fetch count at a time for VspdData
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubCreateCommandObject(lgObjComm)
             Call SubBizQueryMulti()
             Call SubCloseCommandObject(lgObjComm)
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim iSelCount,StrNo,StrItrm,StrFdt,StrTdt
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear                                                                            '¢Ð: Clear Error status

	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    
    lgStrSQL =            "    SELECT TOP " & iSelCount  & "	(CASE WHEN GROUPING(acct_cd) = 1 THEN 'ÃÑ    °è' WHEN GROUPING(a.gl_dt) = 1 THEN '°èÁ¤¼Ò°è' ELSE a.acct_cd END ) AS acct_cd," & vbCrlf 
    lgStrSQL = lgStrSQL & "           a.gl_dt, a.gl_no, a.dept_cd, dr=sum(dr), cr= sum(cr), dept_nm=dbo.ufn_GetDeptName(a.dept_cd, a.gl_dt), " & vbCrlf 
    lgStrSQL = lgStrSQL & "           minor_nm, a.item_desc, acct_nm=dbo.ufn_x_getcodename('a_acct',acct_cd,''), " & vbCrlf 
    lgStrSQL = lgStrSQL & "           (select dbo.ufn_a_ctrl_val(d.ctrl_cd,d.ctrl_val) from a_gl_dtl d where d.gl_no = a.gl_no and d.item_seq = a.item_seq and d.dtl_seq = 1) as ctrl_val1, " & vbCrlf 
    lgStrSQL = lgStrSQL & "           (select dbo.ufn_a_ctrl_val(e.ctrl_cd,e.ctrl_val) from a_gl_dtl e where e.gl_no = a.gl_no and e.item_seq = a.item_seq and e.dtl_seq = 2) as ctrl_val2, " & vbCrlf 
    lgStrSQL = lgStrSQL & "           (select dbo.ufn_a_ctrl_val(f.ctrl_cd,f.ctrl_val) from a_gl_dtl f where f.gl_no = a.gl_no and f.item_seq = a.item_seq and f.dtl_seq = 3) as ctrl_val3, " & vbCrlf 
    lgStrSQL = lgStrSQL & "           (select dbo.ufn_a_ctrl_val(g.ctrl_cd,g.ctrl_val) from a_gl_dtl g where g.gl_no = a.gl_no and g.item_seq = a.item_seq and g.dtl_seq = 4) as ctrl_val4, " & vbCrlf 
    lgStrSQL = lgStrSQL & "           (select dbo.ufn_a_ctrl_val(h.ctrl_cd,h.ctrl_val) from a_gl_dtl h where h.gl_no = a.gl_no and h.item_seq = a.item_seq and h.dtl_seq = 5) as ctrl_val5, " & vbCrlf 
    lgStrSQL = lgStrSQL & "           (select dbo.ufn_a_ctrl_val(i.ctrl_cd,i.ctrl_val) from a_gl_dtl i where i.gl_no = a.gl_no and i.item_seq = a.item_seq and i.dtl_seq = 6) as ctrl_val6, " & vbCrlf 

'   IF lgKeyStream(7) <> "" THEN
'      lgStrSQL = lgStrSQL & "           , ctrl_val=dbo.ufn_x_getcodename('b_biz_partner',j.ctrl_val,'')  " & vbCrlf 
'   End If

    lgStrSQL = lgStrSQL & "           Max(a.DOC_CUR) as DOC_CUR, sum(a.ITEM_AMT) as ITEM_AMT, Max(a.XCH_RATE) as XCH_RATE " & vbCrlf 

    lgStrSQL = lgStrSQL & "      from ( " & vbCrlf 
    lgStrSQL = lgStrSQL & "           select a.gl_dt, a.gl_no, a.item_seq, a.dept_cd,  a.acct_cd, c.minor_nm, a.item_desc, "  & vbCrlf 
    lgStrSQL = lgStrSQL & "            DR=case when  a.dr_cr_fg = 'DR' then a.item_loc_amt else 0 end, CR=case when  a.dr_cr_fg = 'CR' then a.item_loc_amt  else 0 end, " & vbCrlf 
    lgStrSQL = lgStrSQL & "            a.DOC_CUR, a.XCH_RATE, a.ITEM_AMT " & vbCrlf 
    lgStrSQL = lgStrSQL & "           from a_gl_item a " & vbCrlf 
    lgStrSQL = lgStrSQL & "           join a_gl b on b.gl_no = a.gl_no " & vbCrlf 
    lgStrSQL = lgStrSQL & "           join b_minor c on c.major_cd = 'a1001' and c.minor_cd = b.gl_input_type " & vbCrlf 

    IF lgKeyStream(7) <> "" THEN
		IF lgKeyStream(8) = "" THEN
		   lgKeyStream(8) = "%"
		End If   
		lgStrSQL = lgStrSQL & "      join a_gl_dtl j on j.gl_no = a.gl_no and j.item_seq = a.item_seq and j.ctrl_cd = " &  FilterVar(lgKeyStream(7),"","S")  & " and j.ctrl_val LIKE " &  FilterVar(lgKeyStream(8),"","S") & vbCrlf 
    End If
    lgStrSQL = lgStrSQL & "     WHERE a.gl_dt BETWEEN   " &  FilterVar(lgKeyStream(0),"","S")  & " AND " &  FilterVar(lgKeyStream(1),"","S") & vbCrlf 
    IF lgKeyStream(2) <> "" THEN
    lgStrSQL = lgStrSQL & "       AND a.acct_cd >= 	" &  FilterVar(lgKeyStream(2),"","S") & vbCrlf 
    End If
    IF lgKeyStream(3) <> "" THEN
    lgStrSQL = lgStrSQL & "       AND a.acct_cd <= 	" &  FilterVar(lgKeyStream(3),"","S") & vbCrlf 
    End If
    IF lgKeyStream(4) <> "" THEN
    lgStrSQL = lgStrSQL & "       AND a.BIZ_AREA_CD = 	" &  FilterVar(lgKeyStream(4),"","S") & vbCrlf 
    End If
    IF lgKeyStream(5) <> "" THEN
    lgStrSQL = lgStrSQL & "       AND a.dr_cr_fg = 	" &  FilterVar(lgKeyStream(5),"","S") & vbCrlf 
    End If
    IF lgKeyStream(6) <> "" THEN
    lgStrSQL = lgStrSQL & "       AND a.DEPT_CD = 	" &  FilterVar(lgKeyStream(6),"","S") & vbCrlf 
    End If
    lgStrSQL = lgStrSQL & "      ) a " & vbCrlf 
    lgStrSQL = lgStrSQL & " group by acct_cd,a.gl_dt, a.gl_no,a.item_seq, a.dept_cd,minor_nm, a.item_desc, a.DOC_CUR with rollup " & vbCrlf 
    lgStrSQL = lgStrSQL & " HAVING  GROUPING(a.acct_cd) = 1 or GROUPING(a.gl_dt) = 1 OR GROUPING(a.DOC_CUR) = 0  " & vbCrlf 
    lgStrSQL = lgStrSQL & " order by a.acct_cd, gl_dt  " & vbCrlf 


'    caLL SVRMSGBOX(lgStrSQL, 0, 1)
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '¢Ð: No data is found. 
        lgStrPrevKey  = ""
        lgErrorStatus = "YES"
        Response.Write  " <Script Language=vbscript>                                  " & vbCr
        'Response.Write  "    Parent.DBQueryfalse   " & vbCr      
        Response.Write  " </Script>             " & vbCr
        Exit Sub 
    Else    
		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKey)

        lgstrData = ""        
        iDx       = 1 
       Do While Not lgObjRs.EOF
			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gl_dt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gl_no"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_cd"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DR"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CR"))

			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_CUR"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("XCH_RATE"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_AMT"))

			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_desc"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val1"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val2"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val3"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val4"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val5"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val6"))
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
          
			lgObjRs.MoveNext
	
			iDx =  iDx + 1
            If iDx > lgMaxCount Then
				lgStrPrevKey = lgStrPrevKey + 1
				Exit Do
            End If   
                
      Loop 
    End If
    
    If iDx <= lgMaxCount Then
       lgStrPrevKey = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)  
    
    If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If

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


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            