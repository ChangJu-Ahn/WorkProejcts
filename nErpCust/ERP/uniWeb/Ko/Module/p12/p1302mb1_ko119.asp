<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1302mb1_ko119.asp
'*  4. Program Name         : 작업지시를 위한 생산라인 등록
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2006/04/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
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
'   lgMaxCount        = CInt(Request("lgMaxCount"))
   
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


	Dim PP1G204_ko119_Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrPrevKey1
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd
	Dim iLineGroup   
	Dim lgMaxCount
	Dim lgStrPrevKeyPlantCd	' 이전 값 
    Dim lgStrPrevKeyWorkLine' 이전 값 
 		

	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_NextKey1   = 2
    Const C_PlantCd    = 3
    Const C_LineGroup  = 4
  
	Const C_SHEETMAXROWS_D  = 100 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
	'Key 값을 읽어온다 
	iPlantCd      = Trim(Request("txtPlantCd"))
	iStrPrevKey   = Trim(Request("lgStrPrevKeyPlantCd"))           '☜: Next Key Value
	iStrPrevKey1  = Trim(Request("lgStrPrevKeyWorkLine"))
	iLineGroup    = Trim(Request("cboLineGrp"))


    'Component 입력변수        
    ReDim importArray(4)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
    importArray(C_NextKey)		= iStrPrevKey
    importArray(C_NextKey1)		= iStrPrevKey1
    importArray(C_PlantCd)		= iPlantCd
	importArray(C_LineGroup)	= iLineGroup   
 
   
    Set PP1G204_ko119_Data = Server.CreateObject("PP1G204_KO119.cPListProLine")
    
	If CheckSYSTEMError(Err, True) = True Then
	   Call SetErrorStatus					
       Exit Sub
    End If    
   
    Call PP1G204_ko119_Data.C_LIST_PRO_LINE_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PP1G204_ko119_Data = Nothing
        Response.Write " <Script Language=vbscript>	                         " & vbCr
        Response.Write " parent.frm1.txtPlantNm.value = """ & ConvSPChars(exportData)			& """" & vbCr
        Response.Write "</Script>  " & vbCr 
        Call SetErrorStatus
       Exit Sub
    End If    
        
    Set PP1G204_ko119_Data = nothing    
	
	Const E_PlantCd = 0		
	Const E_WorkLine = 1
	Const E_WorkLineDesc = 2
	Const E_LineGroup = 3
	Const E_Remark = 4
	
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_PlantCd)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_LineGroup)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_LineGroup)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_WorkLine)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_WorkLineDesc)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_Remark)))
'			iStrData = iStrData & Chr(11) & ""
'			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_CostElmtNm)))
'			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_RelCostElmtCd)))
'			iStrData = iStrData & Chr(11) & ""
'			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_RelCostElmtNm)))
            iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_PlantCd)
			IStrPrevKey1 = exportData1(UBound(exportData1, 1), E_WorkLine)
			Exit For
		End If
	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
		iStrPrevKey1 = ""
	End If

	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData)			& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & iPlantCd    & """" & vbCr
    Response.Write " .lgStrPrevKeyPlantCd = """ & ConvSPChars(iStrPrevKey)		& """" & vbCr
    Response.Write " .lgStrPrevKeyWorkLine = """ & ConvSPChars(iStrPrevKey1)    	& """" & vbCr   
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr

End Sub    	 



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PP1G204_ko119_Data
    Dim importString
    Dim txtSpread
    Dim iErrPosition 

   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                        '☜: Clear Error status

    
   importString = Trim(Request("txtPlantCd"))
   txtSpread    = Trim(Request("txtSpread"))

    Set PP1G204_ko119_Data = Server.CreateObject("PP1G204_KO119.cPMngProLine")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PP1G204_ko119_Data.P_MANAGE_PRODUCTION_LINE(gStrGlobalCollection, importString, txtSpread, iErrPosition)			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
       Call SetErrorStatus
       Set PP1G204_ko119_Data = Nothing
       Exit Sub
    End If    
    
    Set PP1G204_ko119_Data = Nothing
	
    
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
