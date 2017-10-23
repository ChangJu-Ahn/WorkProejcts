<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%
     
    Call LoadBasisGlobalInf() 
    Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow
    Dim txtVersion
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    txtVersion		  = Request("txtVersion")
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim iPC4G012D
		
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
 
    Set iPC4G012D = Server.CreateObject("PC4G012.cCMngMovTypeConfgSvr")


	
    If CheckSYSTEMError(Err, True) = True Then
	   SetErrorStatus()						
       Exit Sub
    End If    

	
    Call iPC4G012D.C_MANAGE_MOV_TYPE_CONFG_SVR(gStrGloBalCollection,"D",txtVersion)		
		
    
    If CheckSYSTEMError(Err, True) = True Then					
        Set iPC4G012D = Nothing
		Call SetErrorStatus
       Exit Sub
       
    End If    

    
    Set iPC4G012D = Nothing
	    
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPC4G012Q
    Dim iStrData
    
    Dim exportData
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iIntLoopCount
    
   	Dim  arrTemp

 	Set iPC4G012Q = Server.CreateObject("PC4G012.cCListMovTypeConfgSvr")


	
    If CheckSYSTEMError(Err, True) = True Then
		SetErrorStaus()
		Exit Sub
    End If    
	
	Call iPC4G012Q.C_LIST_MOV_TYPE_CONFG_SVR(gStrGloBalCollection, txtVersion, exportData)
	
    
    
    If CheckSYSTEMError(Err, True) = True Then					
         Set iPC4G012Q = Nothing
		Call SetErrorStatus
       Exit Sub
       
    End If    

    
    Set iPC4G012Q = Nothing
	
	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportData, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
		For iLngCol = 0 To UBound(exportData, 2)
		    IF iLngCol = 8 or iLngCol = 9 or iLngCol = 10 or iLngCol = 11 Then

				IF UCase(Trim(exportData(iLngRow, iLngCol))) = "Y" Then
					iStrData = iStrData & Chr(11) & 1
				ELSE
					iStrData = iStrData & Chr(11) & 0
				END IF
		    ELSEIF iLngCol = 0Then
				IF UCase(Trim(exportData(iLngRow, iLngCol))) = "D" Then
					iStrData = iStrData & Chr(11) & "입고"
				ELSEIF UCase(Trim(exportData(iLngRow, iLngCol))) = "C" AND UCase(Trim(exportData(iLngRow, iLngCol + 1))) <> "ST" Then
					iStrData = iStrData & Chr(11) & "출고"
				ELSE
					iStrData = iStrData & Chr(11) & "이동"
				END IF
			ELSEIF iLngCol = 5 Then
				iStrData = iStrData &  Chr(11) & ConvSPChars(Trim(exportData(iLngRow, iLngCol))) & Chr(11) & "" 
			ELSE
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, iLngCol)))
			END IF	
		Next
		iStrData = iStrData & Chr(11) & iLngRow
		iStrData = iStrData & Chr(11) & Chr(12)
	Next
		

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    'Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr 
    '---------- Developer Coding part (End)   ---------------------------------------------------------------
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iPC4G012S
	Dim iErrPosition
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   

    Set iPC4G012S = Server.CreateObject("PC4G012.cCMngMovTypeConfgSvr")

    If CheckSYSTEMError(Err, True) = True Then
	   SetErrorStatus()						
       Exit Sub
    End If    

	
    Call iPC4G012S.C_MANAGE_MOV_TYPE_CONFG_SVR(gStrGloBalCollection,"U",txtVersion,Trim(Request("txtSpread")),iErrPosition)		
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
       Call SetErrorStatus
       Set iPC4G012S = Nothing
       Exit Sub
    End If    
    
    Set iPC4G012S = Nothing
	

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

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
    lgErrorStatus    = "YES"
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
                .frm1.hVersion.value = "<%=ConvSPChars(txtVersion)%>"        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
</Script>	
