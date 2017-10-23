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
    Dim StrFrDt,StrToDt,StrAcct,StrFrBiz,StrToBiz,StrAcctVal
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    StrFrDt		= REPLACE(lgKeyStream(0),"-","")
    StrToDt		= REPLACE(lgKeyStream(1),"-","")
    StrAcct		= lgKeyStream(2)
    StrFrBiz	= lgKeyStream(3)
    StrToBiz	= lgKeyStream(4)
    StrAcctVal	= lgKeyStream(5)
    'IF lgKeyStream(2) <> "" THEN
    
    'ELSE
    'StrAcct = "''"
    'END IF

         
    'Call SvrMsgBox(StrFrDt &"_"& StrToDt &"_"& StrAcct &"_"& StrFrBiz &"_"& StrToBiz &"_"& StrAcctVal, vbInformation, I_MKSCRIPT)
    With lgObjComm
    
        .CommandText = "x_usp_balance_amt_of_acct_cd"
        .CommandType = adCmdStoredProc       

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_mnth_dd"      ,adVarChar,adParamInput,08, StrFrDt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_mnth_dd"		,adVarChar,adParamInput,08, StrToDt)
	    
'	    if StrFrBiz = "" then
		'   StrFrBiz = "ZZZZZZZZ"
		'endif
		
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_biz"			,adVarChar,adParamInput,20, StrFrBiz)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_biz"			,adVarChar,adParamInput,10, StrToBiz)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_acct_cd"		,adVarChar,adParamInput,10, StrAcct)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_acct_cd"		,adVarChar,adParamInput,10, StrAcctVal)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"			,adVarChar,adParamoutput,08)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id"			,adVarChar,adParamoutput,13)

	   ' lgObjComm.Parameters.Append lgObjComm.CreateParameter("@USER_ID"    ,adVarChar,adParamInput,13, gUsrId)
	   
        
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
			
			Call SubBizQueryMulti()
        End If
    Else 
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
    Dim lgstrData
    Dim iDx
    Dim iSelCount,StrNo,StrItrm,StrFdt,StrTdt
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear                                                                            '¢Ð: Clear Error status

	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
    lgStrSQL =            "    SELECT TOP " & iSelCount  & " * "
    lgStrSQL = lgStrSQL & "      FROM x_balance_amt_of_acct_cd(NOLOCK) "
    lgStrSQL = lgStrSQL & "      WHERE sp_id = '" & StrId & "' "

    lgStrSQL = lgStrSQL & "      ORDER BY ACCT_CD "
    ',ctrl_val "
    
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
			
			IF ConvSPChars(lgObjRs("acct_cd")) <> "ZZZZZZZZ" THEN
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
			ELSE
			lgstrData = lgstrData & Chr(11) & ""
			END IF
			
			'caLL SVRMSGBOX(ConvSPChars(lgObjRs("sp_id")), 0, 1)
			
			'IF ConvSPChars(lgObjRs("ctrl_val")) <> "ZZZZZZZZ" THEN
			'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val"))
			'ELSE
			'lgstrData = lgstrData & Chr(11) & ""
			'END IF
			'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ctrl_val_NM"))
			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("l_yr_bal_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("l_prd_bal_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("t_prd_inc_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("t_prd_dec_loc_amt"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("t_prd_bal_loc_amt"))
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


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            