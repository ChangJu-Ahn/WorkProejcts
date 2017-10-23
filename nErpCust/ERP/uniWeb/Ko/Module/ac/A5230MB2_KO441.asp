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
             Call SubBizBatch()
             Call SubCloseCommandObject(lgObjComm)
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

Sub SubBizBatch()

    Dim strMsg_cd ,IntRetCD
    Dim StrFrDt,StrToDt,StrAcct,StrAcct2, StrFrBiz,StrToBiz,StrAcctVal,StrAcctVal2,StrAcctVal3,StrAcctVal4
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    StrFrDt		= lgKeyStream(0)
    StrToDt		= lgKeyStream(1)
    StrAcct		= lgKeyStream(2)
    StrAcct2            = lgKeyStream(3)
    StrFrBiz	= lgKeyStream(4)
    StrToBiz	= lgKeyStream(5)
    StrAcctVal	= lgKeyStream(6)
    StrAcctVal2	= lgKeyStream(7)
    StrAcctVal3	= ""
    StrAcctVal4	= ""
    'IF lgKeyStream(2) <> "" THEN
    
    'ELSE
    'StrAcct = "''"
    'END IF
    
         
  ' Call SvrMsgBox(StrFrDt &"_"& StrToDt &"_"& StrAcct &"_"& StrAcct2 &"_"& StrFrBiz &"_"& StrToBiz &"_"& StrAcctVal &"_"& StrAcctVal2 &"_"& StrAcctVal3 &"_"&StrAcctVal4, vbInformation, I_MKSCRIPT)
    With lgObjComm
    
        .CommandText = "x_usp_balance_amt_of_acct_cd_by_ctrl_cd_ko441"
        .CommandType = adCmdStoredProc       

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_mnth"      ,adVarXChar,adParamInput,10, StrFrDt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_mnth"		,adVarXChar,adParamInput,10, StrToDt)
	    	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_biz"			,adVarXChar,adParamInput,20, StrFrBiz)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_biz"			,adVarXChar,adParamInput,10, StrToBiz)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@acct_cd"		,adVarXChar,adParamInput,11, StrAcct)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@acct_cd2"		,adVarXChar,adParamInput,11, StrAcct2)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ctrl_cd"		,adVarXChar,adParamInput,10, StrAcctVal)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ctrl_cd2"		,adVarXChar,adParamInput,10, StrAcctVal3)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ctrl_val"		,adVarXChar,adParamInput,30, StrAcctVal2)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ctrl_val2"		,adVarXChar,adParamInput,30, StrAcctVal4)

	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"			,adVarXChar,adParamoutput,08)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id"			,adVarXChar,adParamoutput,13)

	   'lgObjComm.Parameters.Append lgObjComm.CreateParameter("@USER_ID"    ,adVarChar,adParamInput,13, gUsrId)
	   
        
	    .Execute ,, adExecuteNoRecords
    
    End With

	If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD = -1 Then
			strMsg_Cd = lgObjComm.Parameters("@msg_cd").Value
  
			Call SetErrorStatus
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
		Else
			StrId = lgObjComm.Parameters("@sp_id").Value
			'CALL SVRMSGBOX(StrId,0,1) 

			Call SubBizQueryMulti()

        End If
    Else 
 	'Call SvrMsgBox("aa" ,vbInformation, I_MKSCRIPT) 
		Call SetErrorStatus
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End If
	
End Sub	


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL
    Dim lgstrData,S,T
    Dim iDx
    Dim iSelCount,StrNo,StrItrm,StrFdt,StrTdt

    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear                                                                            '¢Ð: Clear Error status

	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
  'caLL SVRMSGBOX(lgKeyStream(5), 0, 1)
  'caLL SVRMSGBOX(lgKeyStream(6), 0, 1)

   ' if lgKeyStream(5) <> "" and lgKeyStream(6) <> "" then
     
    lgStrSQL =            "    SELECT TOP " & iSelCount  & "  (CASE WHEN GROUPING(A.ACCT_NM) = 1 THEN '¼Ò°è' ELSE B.ACCT_CD END )ACCT_CD, "
    lgStrSQL = lgStrSQL & "           (CASE WHEN GROUPING(A.ACCT_NM) = 1 THEN '' ELSE A.ACCT_NM END )ACCT_NM, "
    lgStrSQL = lgStrSQL & "           B.CTRL_VAL,B.CTRL_VAL_NM, B.CTRL_VAL_NM2,"
    lgStrSQL = lgStrSQL & "           SUM(B.L_YR_BAL_LOC_AMT)L_YR_BAL_LOC_AMT, "
    lgStrSQL = lgStrSQL & "           SUM(B.T_PRD_INC_LOC_AMT)T_PRD_INC_LOC_AMT, "
    lgStrSQL = lgStrSQL & "           SUM(B.T_PRD_DEC_LOC_AMT)T_PRD_DEC_LOC_AMT, "
    lgStrSQL = lgStrSQL & "           SUM(B.T_PRD_BAL_LOC_AMT)T_PRD_BAL_LOC_AMT "
    lgStrSQL = lgStrSQL & "         FROM x_balance_amt_of_acct_cd_by_ctrl_cd B, A_ACCT A "
    lgStrSQL = lgStrSQL & "     WHERE b.sp_id = '" & StrId & "' "
    lgStrSQL = lgStrSQL & "           AND A.ACCT_CD = B.ACCT_CD"
    lgStrSQL = lgStrSQL & "                              and b.ctrl_val_nm <> ''  "
    lgStrSQL = lgStrSQL & "       group by b.ACCT_CD,A.ACCT_NM,B.CTRL_VAL,B.CTRL_VAL_NM, B.CTRL_VAL_NM2  with rollup    "
    lgStrSQL = lgStrSQL & "      HAVING  GROUPING(b.ACCT_CD) <> 1 and GROUPING(A.ACCT_NM) = 1 or GROUPING(CTRL_VAL_NM) = 0 and GROUPING(CTRL_VAL_NM2) = 0 "
'    else
 '   lgStrSQL =            "    SELECT TOP " & iSelCount  & " b.acct_cd,a.acct_nm,b.ctrl_val,b.ctrl_val_nm, "
 '   lgStrSQL = lgStrSQL & "   b.ctrl_val2,b.ctrl_val_nm2,b.l_yr_bal_loc_amt,b.t_prd_inc_loc_amt,b.t_prd_dec_loc_amt,b.t_prd_bal_loc_amt "
 '   lgStrSQL = lgStrSQL & "      FROM x_balance_amt_of_acct_cd_by_ctrl_cd b(NOLOCK), a_acct a "
 '   lgStrSQL = lgStrSQL & "     WHERE a.acct_cd = b.acct_cd and sp_id = '" & StrId & "' "

 '   lgStrSQL = lgStrSQL & "  ORDER BY b.ACCT_CD,b.ctrl_val "
  '  End if
    
  '  caLL SVRMSGBOX(lgStrSQL, 0, 1)  
  
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
	   
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_cd"))   
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_nm"))  			   			
			IF ConvSPChars(lgObjRs("ctrl_val")) <> "ZZZZZZZZ" THEN
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val"))
			S = ""
			ELSE
			lgstrData = lgstrData & Chr(11) & ""
			S = "S"
			END IF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val_NM"))
			'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val2"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val_NM2"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("l_yr_bal_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("t_prd_inc_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("t_prd_dec_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("t_prd_bal_loc_amt"))
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
          
			lgObjRs.MoveNext
	
			iDx =  iDx + 1
            
            If iDx  > lgMaxCount  Then
				lgStrPrevKey = lgStrPrevKey + 1
				Exit Do
            End If   
			
			if S = "S" then 
			S = "" 
			Exit Do
			End if
			
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
       Response.Write  "    Parent.Frm1.txtID.value  = """ & StrID & """" & vbCr   
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


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            