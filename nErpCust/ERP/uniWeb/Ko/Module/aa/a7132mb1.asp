<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<% Call LoadBasisGlobalInf() 
   Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")


'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7132mb1
'*  4. Program Name         : 감가상각방법등록 
'*  5. Program Desc         : 감가상각방법을 등록, 삭제,조회 
'*  6. Modified date(First) : 2003/09/19
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, Joon-Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'**********************************************************************************************

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
                                                                '☜: Clear Error status
                                                                
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow
	
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    
    Call LoadBasisGlobalInf()

    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    

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

    Dim iPAAG095
    Dim iStrData
    Dim exportData
    Dim exportReturn
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iDeprcd
    Dim iIntMaxRows
    Dim iIntLoopCount
    Dim importArray
	Dim lgMaxCount
    
    Const C_SHEETMAXROWS_D  = 100
    
	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
	Const C_DeprCd     = 2
          
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
     
	'Key 값을 읽어온다 
	iDeprcd     = Trim(Request("txtDeprCd"))
	iStrPrevKey = Trim(Request("lgStrPrevKey"))
	
  	  	
    'Component 입력변수        
    ReDim importArray(2)
   
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	importArray(C_DeprCd)		= iDeprcd
    
    Const C_DEPR_MTH_CD = 0
    Const C_DEPR_MTH_NM = 1
    Const C_DEPR_FG = 2
    Const C_DEPR_UNIT = 3
    Const C_DEPR_TYPE = 4
    Const C_DEPR_TERM = 5
    Const C_DEPR_INC  = 6
    Const C_DEPR_SOLD = 7
    Const C_RES_RATE  = 8
    Const C_DEPR_CLOSE_TYPE = 9
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

   
	Set iPAAG095 = Server.CreateObject("PAAG095.cAListDeprMethSvr")
	
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End
       Exit Sub
    End If    

	Call iPAAG095.A_LIST_DEPR_METH_SVR(gStrGlobalCollection, importArray, exportData, exportReturn)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG095 = Nothing
        Response.Write " <Script Language=vbscript>	                         " & vbCr
        Response.Write " parent.frm1.txtDeprNm.value = """ & ConvSPChars(exportData)		& """" & vbCr
        Response.Write "</Script>  " & vbCr
        Call SetErrorStatus
       Exit Sub
    End If    
        
    Set iPAAG095 = Nothing


	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportReturn, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		If  iIntLoopCount < (lgMaxCount + 1) Then
	    
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_MTH_CD))
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_MTH_NM))
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_FG))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_UNIT))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_TYPE))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_TERM))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_INC))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_SOLD))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportReturn(iLngRow, C_RES_RATE),ggExchRate.DecPoint,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, C_DEPR_CLOSE_TYPE))
			iStrData = iStrData & Chr(11) & iLngRow+1                                
            iStrData = iStrData & Chr(11) & Chr(12) 
			    
	    Else
			iStrPrevKey = exportReturn(UBound(exportReturn, 1), 0)
			iIntLoopCount = iIntLoopCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
	End If


	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData               " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData        & """" & vbCr
	Response.Write "	.frm1.txtDeprNm.value = """ & ConvSPChars(exportData(0)) & """" & vbCr
    Response.Write "	.frm1.hDeprCd.value = """ & iDeprCd        & """    " & vbCr
    Response.Write "	.lgStrPrevKey = """ & iStrPrevKey		    & """" & vbCr
'   Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr

End Sub	


Sub SubBizSave()

    Dim iPAAG095
    Dim import_String
    Dim import_GroupString
    Dim iErrPosition
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    import_GroupString = Trim(Request("txtSpread"))
    
    Set iPAAG095 = Server.CreateObject("PAAG095.cAMngDeprMethSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    
    Call iPAAG095.A_MANAGE_DEPR_METH_SVR(gStrGlobalCollection, import_GroupString, iErrPosition)

	If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
		Call SetErrorStatus
		Set iPAAG095 = Nothing
		Exit Sub
    End If    
    
    Set iPAAG095 = Nothing
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
        lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus() 
    lgErrorStatus     = "YES"                                                         '☜: Set error status   
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
</Script>	







	

