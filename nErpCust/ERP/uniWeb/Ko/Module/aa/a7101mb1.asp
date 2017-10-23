<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : as001mb1
'*  4. Program Name         : 고정자산 계정정보등록 
'*  5. Program Desc         : 고정자산별 계정정보를 등록,수정,삭제,조회 
'*  6. Comproxy List        : +As0011ManageSvr
'                             +As0018ListSvr
'*  7. Modified date(First) : 2000/03/24
'*  8. Modified date(Last)  : 2000/09/14
'*  9. Modifier (First)     : 조익성 
'* 10. Modifier (Last)      : hersheys
'* 11. Comment              :
'**********************************************************************************************

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
                                                                '☜: Clear Error status
                                                                
	Dim lgOpModeCRUD
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    
    Call LoadBasisGlobalInf()

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
'        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iPAAG005
    Dim iStrData
    Dim exportData
    Dim exportReturn
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrAcctcd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
    
    Const C_SHEETMAXROWS  = 100
    Const C_QueryConut		= 0
    Const C_MaxQueryReCord = 1
    Const C_Acctcd = 2

    Const C_Acct_ACCT_CD = 0
    Const C_Acct_ACCT_NM = 1
    Const C_Acct_DEPR_MTHD = 2
    Const C_Acct_DUR_YRS = 3
    Const C_Acct_TEMP_FG1 = 4
    Const C_Acct_TEMP_FG2 = 5
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If iStrPrevKey = "" Then
		iStrAcctcd	= Request("txtAcctCd")
	Else
		iStrAcctcd	= iStrPrevKey
    End If
    
    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = 0
    End If
        
    ReDim importArray(2)        
    importArray(C_QueryConut)		= iIntQueryCount
    importArray(C_MaxQueryReCord)	= C_SHEETMAXROWS
    importArray(C_Acctcd)		    = iStrAcctcd
    
	Set iPAAG005 = Server.CreateObject("PAAG005.cAAS0018Listvr")
	
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End
       Exit Sub
    End If    

	Call iPAAG005.AS0018_LIST_SVR(gStrGlobalCollection, importArray, exportData, exportReturn)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG005 = Nothing
       Response.End
       Exit Sub
    End If    

    Set iPAAG005 = Nothing

	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportReturn, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
	    
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_Acct_ACCT_CD))'0 계정코드 
			iStrData = iStrData & Chr(11) & ""'C_AcctCdPopup
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_Acct_ACCT_NM))'C_AcctNm
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_Acct_DEPR_MTHD))'3C_DeprMthd
			iStrData = iStrData & Chr(11) & ""'ConvSPChars(exportReturn(iLngRow, 2))'C_DeprMthdNm
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_Acct_DUR_YRS))'5C_DurYrs
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_Acct_TEMP_FG1))'6C_AcctFg
			iStrData = iStrData & Chr(11) & ""'ConvSPChars(exportReturn(iLngRow, 4))'7C_AcctFgNm
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_Acct_TEMP_FG2))'8C_DeprFg
			iStrData = iStrData & Chr(11) & ""'ConvSPChars(exportReturn(iLngRow, 5))'9C_DeprFgNm
			iStrData = iStrData & Chr(11) & iLngRow+1                                
            iStrData = iStrData & Chr(11) & Chr(12) 
			    
	    Else
			iStrPrevKey = exportReturn(UBound(exportReturn, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If


	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData               " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData        & """" & vbCr
    
    if isarray(exportData) then
		Response.Write "	.frm1.txtAcctNm.value = """ & exportData(0) & """" & vbCr
	end if
    
    Response.Write "	.lgPageNo = """ & iIntQueryCount		    & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & iStrPrevKey		    & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
	Response.End

End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    Dim iPAAG005
    Dim import_String
    Dim import_GroupString
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    'import_Group = Trim(Request("txtAcctCd"))
    import_GroupString = replace(Trim(Request("txtSpread")),",","")
    
    Set iPAAG005 = Server.CreateObject("PAAG005.cAAS0011MngSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    
    'Call iPAAG005.AS0011_MANAGE_SVR(gStrGlobalCollection, import_Group, import_GroupString)
    Call iPAAG005.AS0011_MANAGE_SVR(gStrGlobalCollection, import_GroupString)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG005 = Nothing
       Response.End
       Exit Sub
    End If    
    
    Set iPAAG005 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

<Script Language="VBScript">
	parent.DbSaveOk																		'☜: 화면 처리 ASP 를 지칭함 
</Script>	
<%					

    Set pAS0011 = Nothing                                                   '☜: Unload Comproxy

	Response.End
%>
