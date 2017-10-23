<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%
	Call LoadBasisGlobalInf()

    Dim iStrData 
    Dim exportData
	Dim iCboCostType
	Dim iPlantCd
    Dim iStrPrevKey
    Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow,  lgMaxCount
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'   lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    
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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
			

End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPC1G001Q
    
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim importArray
    Dim iIntLoopCount
   	Dim  arrTemp																
    
    Const C_CboCostType = 0
    Const C_PlantCd = 1
    
  	Const C_SHEETMAXROWS_D  = 100 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time

	iStrPrevKey = Trim(Request("lgStrPrevKey"))         '☜: Next Key Value
	
	arrTemp = Split(Request("txtKeyStream"), gColSep)									
	
	'Key 값을 읽어온다 
	iCboCostType= arrTemp(0)
	iPlantCd	= arrTemp(1)	

    'Component 입력변수        
    ReDim importArray(1)
            
    importArray(C_CboCostType)	= iCboCostType
    importArray(C_PlantCd)		= iPlantCd
    
	Set iPC1G001Q = Server.CreateObject("PC1G000.cCListCoConfgSvr")

    If CheckSYSTEMError(Err, True) = True Then
    	SetErrorStatus()		
		Exit Sub
    End If    
	
	Call iPC1G001Q.C_LIST_COST_CONFG_SVR(gStrGloBalCollection, lgMaxCount,importArray, exportData, exportData1,iStrPrevKey)
    
    If CheckSYSTEMError(Err, True) = True Then					
         Set iPC1G001Q = Nothing
         Call SetErrorStatus
       Exit Sub
    End If    
    
    Set iPC1G001Q = Nothing
	
	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (lgMaxCount + 1) Then
			For iLngCol = 0 To UBound(exportData1, 2)
			    IF iLngCol = 1 or iLngCol = 3 Then
					iStrData = iStrData & Chr(11) & "" & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, iLngCol)))
			    ELSE
					iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, iLngCol))
				END IF
			Next
				iStrData = iStrData & Chr(11) & iLngRow
			    iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), 0)
			Exit For
			  
		End If
	Next
		
	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
	End If
	
   '---------- Developer Coding part (End)   ---------------------------------------------------------------
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iPC1G001S
	Dim importArray
	Dim txtSpread

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    ReDim importArray(1)
    
    importArray(0) = Trim(Request("CboCostType"))
    importArray(1) = Trim(Request("txtPlantCd"))
    txtSpread     = Trim(Request("txtSpread"))
    
    Set iPC1G001S = Server.CreateObject("PC1G000.cCMngCoConfgSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       SetErrorStatus()
       Exit Sub
    End If    

    Call iPC1G001S.C_MANAGE_COST_CONFG_SVR(gStrGloBalCollection, importArray, txtSpread)		
		
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPC1G001S = Nothing
       SetErrorStatus()
       Exit Sub
    End If    
    
    Set iPC1G001S = Nothing
  
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
          Parent.frm1.txtPlantNm.value = "<%=ConvSPChars(exportData)%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
				.ggoSpread.Source = .frm1.vspdData
				.ggoSpread.SSShowData "<%=iStrData%>"
                .frm1.hPlantCd.value = "<%=iPlantCd%>"
                .frm1.hCboCostType.value = "<%=iCboCostType%>"
                .lgStrPrevKey = "<%=ConvSPChars(iStrPrevKey)%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
</Script>	
