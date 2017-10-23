<% Option Explicit %>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Call LoadBasisGlobalInf()
	
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow
    Dim lgObjConn
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             'Call SubBizSave()
             'Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             'Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iPC1G045Q1
    Dim iStrData			'Spread2
    Dim iStrData1			'Spread3
 
    Dim exportData
    Dim iLngRow,iLngCol
    Dim importArray
    
    Const C_VerCd = 0
    Const C_CostCd = 1
    
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear     
                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    ReDim importArray(1)
            
    importArray(C_VerCd)   = Trim(Request("txtVerCd"))
    importArray(C_CostCd)  = Trim(Request("txtCostCd"))

	Set iPC1G045Q1 = Server.CreateObject("PC1G045.cCListRcvRlByCcSvr")

	
    If CheckSYSTEMError(Err, True) = True Then
		Call SetErrorStatus
		Exit Sub
    End If
    
	Call iPC1G045Q1.C_LIST_RECV_RULE_BY_CC_SVR(gStrGloBalCollection,  importArray, exportData)
	
    If CheckSYSTEMError(Err, True) = True Then					
		Call SetErrorStatus         
         Set iPC1G045Q1 = Nothing
       Exit Sub
       
    End If    

    
    Set iPC1G045Q1 = Nothing
	iStrData = ""
	iStrData1 = ""
	For iLngRow = 0 To UBound(exportData, 1) 		
		
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(Trim(Request("txtCostCd")))
			
		For iLngCol = 0 To UBound(exportData, 2)
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, iLngCol)))
				iStrData1 = iStrData1 & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, iLngCol)))
		Next
		iStrData = iStrData & Chr(11) & iLngRow
		iStrData1 = iStrData1 & Chr(11) & iLngRow
		iStrData = iStrData & Chr(11) & Chr(12)
		iStrData1 = iStrData1 & Chr(11) & Chr(12)
	Next


	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	' Spread2
	
	   Response.Write "	.ggoSpread.Source = .frm1.vspdData2             " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    ' Spread3(Append)
    Response.Write "	.ggoSpread.Source = .frm1.vspdData3             " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData1       & """" & vbCr
    
    Response.Write "	.frm1.htxtVerCd.value = """ & Trim(Request("txtVerCd"))    & """" & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr                                                        '¢Ð: Release RecordSSet
	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()


End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

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


%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
            .DBQueryOk2()
	         End with
	      Else
             parent.Frm1.vspdData.Focus 
          End If   
    End Select    
</Script>	
