<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Distribution Factor
'*  3. Program ID           : c1701mb1
'*  4. Program Name         : 실제원가 배부요소 정보 등록 
'*  5. Program Desc         : 실제원가 배부요소 정보 
'*  6. Modified date(First) : 2000/11/08
'*  7. Modified date(Last)  : 2002/06/18
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : cho Ig Sung/ Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Call LoadBasisGlobalInf()
	
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow
	
	On Error Resume Next								'☜: 
	Err.Clear

	Call HideStatusWnd

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
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	Dim PC1G035Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim importArray
    Dim iIntLoopCount
	Dim iDstbFctr
	Dim lgMaxCount
	

	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
	Const C_DstbFctr   = 2
          
    Const C_SHEETMAXROWS_D  = 100                   '☜: Max fetched data at a time
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D) 
     
	'Key 값을 읽어온다 
	iDstbFctr   = Trim(Request("txtDstbFctr"))
	iStrPrevKey = Trim(Request("lgStrPrevKey"))
	
  	  	
    'Component 입력변수        
    ReDim importArray(2)
   
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	importArray(C_DstbFctr)		= iDstbFctr
   

    Set PC1G035Data = Server.CreateObject("PC1G035.cCListDstFctrSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
	   Call SetErrorStatus										
       Exit Sub
    End If    
    
    Call PC1G035Data.C_LIST_DSTB_FCTR_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PC1G035Data = Nothing
		Response.Write " <Script Language=vbscript>	                         " & vbCr
		Response.Write " parent.frm1.txtDstbFctrNm.value = """ & ConvSPChars(exportData)   	& """" & vbCr
		Response.Write "</Script>  " & vbCr
		Call SetErrorStatus					  
       Exit Sub
    End If    
  
    Set PC1G035Data = nothing 
    
    Const E_DstbFctr = 0
	Const E_DstbFctrNm = 1
	Const E_GenFlag = 2	
	Const E_WeightAppFlag = 3	
'	Const E_GenFlagNm = 3
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    If  iIntLoopCount < (lgMaxCount + 1) Then
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DstbFctr)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DstbFctrNm)))
			iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, E_GenFlag))
			iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, E_GenFlag))
			IF UCase(Trim(exportData1(iLngRow, E_WeightAppFlag))) = "Y" Then
				iStrData = iStrData & Chr(11) & 1
			ELSE
				iStrData = iStrData & Chr(11) & 0
			END IF			
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_DstbFctr)
			Exit For
		End If
 	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey  = ""
	End If

	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.txtDstbFctrNm.value = """ & ConvSPChars(exportData)   	& """" & vbCr
    Response.Write " .frm1.hDstbFctr.value = """ & iDstbFctr        & """" & vbCr
    Response.Write " .lgStrPrevKey          = """ & ConvSPChars(iStrPrevKey)		& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr  

End Sub    	 



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
  
    Dim PC1G035Data
    Dim txtSpread
    Dim iErrPosition 
    
    txtSpread    = Trim(Request("txtSpread"))
    
    Set PC1G035Data = Server.CreateObject("PC1G035.cCMngDstFctrSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC1G035Data.C_MANAGE_DSTB_FCTR_SVR(gStrGlobalCollection, txtSpread, iErrPosition)			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then												
       Call SetErrorStatus
       Set PC1G035Data = Nothing
       Exit Sub
    End If    
    
    Set PC1G035Data = Nothing
	
   
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
