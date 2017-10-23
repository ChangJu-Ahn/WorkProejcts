<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<% 
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*","NOCOOKIE", "MB")

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	Dim lgErrorStatus
	Dim lgErrorPos
	Dim lgOpModeCRUD
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iPB2SA10
    Dim iStrData
    Dim ExportData(1)
    Dim ExportReturn
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrBank_cd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
	Dim itxtbank_cd
    
    Const C_SHEETMAXROWS_D  = 100
    Const C_QueryConut		= 0
    Const C_MaxQueryReCord = 1
    Const C_TaxOffice_Cd = 2

    Const A0286_b_bank_bank_cd1 = 0
    Const A0286_b_bank_bank_nm1 = 1
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	iStrPrevKey		= Request("lgStrPrevKey")
	itxtbank_cd		= Request("txtbank_cd")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If Trim(iStrPrevKey) = "" Then
		iStrBank_cd	= itxtbank_cd
	Else
		iStrBank_cd	= iStrPrevKey
    End If
    
    If Len(Trim(iIntQueryCount))  Then                                        '¢Ð : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = 0
    End If
        
    ReDim importArray(2)        
    importArray(C_QueryConut)		= iIntQueryCount
    importArray(C_MaxQueryReCord)	= C_SHEETMAXROWS_D
    importArray(C_TaxOffice_Cd)		= iStrBank_cd
    
	Set iPB2SA10 = Server.CreateObject("PB2SA10.cBListBankLoanSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
	
	Call iPB2SA10.B_LIST_BANK_LOAN_SVR(gStrGloBalCollection, importArray, itxtbank_cd, ExportData, ExportReturn)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPB2SA10 = Nothing
       Exit Sub
    End If    

    Set iPB2SA10 = Nothing

	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(ExportReturn, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(ExportReturn(iLngRow, 0))
			iStrData = iStrData & Chr(11) & ConvSPChars(ExportReturn(iLngRow, 1)) 
			iStrData = iStrData & Chr(11) & "" 
			iStrData = iStrData & Chr(11) & UNINumClientFormat((ExportReturn(iLngRow, 2)), ggAmtOfMoney.DecPoint, 0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat((ExportReturn(iLngRow, 3)), ggAmtOfMoney.DecPoint, 0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat((ExportReturn(iLngRow, 4)), ggAmtOfMoney.DecPoint, 0)

			iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)


	    Else
			iStrPrevKey = ExportReturn(UBound(ExportReturn, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.frm1.txtbank_nm.value = """ & ConvSPChars(ExportData(A0286_b_bank_bank_nm1))  & """" & vbCr
	Response.Write "	.frm1.hBankCd.value = """ & ConvSPChars(ExportData(A0286_b_bank_bank_cd1))  & """" & vbCr
    Response.Write "	.lgPageNo = """ & iIntQueryCount		   & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)		   & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    Dim iPB2SA10
    Dim import_Group
    Dim import_String
    Dim import_GroupString
    Dim iErrorPosition
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    import_Group = Trim(Request("txtbank_cd"))
    import_GroupString = replace(Trim(Request("txtSpread")),",","")

    Set iPB2SA10 = Server.CreateObject("PB2SA10.cBMngBankLoanSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPB2SA10 = Nothing
       Exit Sub
    End If    
    
    Call iPB2SA10.B_MANAGE_BANK_LOAN_SVR(gStrGloBalCollection, import_Group, import_GroupString,iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then				
       Set iPB2SA10 = Nothing
       Exit Sub
    End If    
    
    Set iPB2SA10 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode)
    On Error Resume Next
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status


End Sub

%>
