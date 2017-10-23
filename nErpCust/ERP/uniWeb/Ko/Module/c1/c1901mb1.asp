<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : COST
'*  2. Function Name        : P_ROUTING_DETAIL
'*  3. Program ID           : c1901mb1.asp
'*  4. Program Name         : 완성품 환산율 등록 
'*  5. Program Desc         : 완성품 환산율 조회, 수정 
'*  6. Modified date(First) : 2001/01/19
'*  7. Modified date(Last)  : 2002/06/20
'*  8. Modifier (First)     : Cho Ig Sung
'*  9. Modifier (Last)      : Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	
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

	Dim PC1G065Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrPrevKey1
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd, iItemCd
	Dim lgMaxCount
	Dim iIntMaxRows
    Dim lgStrRoutNoPrevKey	' 이전 값 
	Dim lgStrOprNoPrevKey	' 이전 값 
    

	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_NextKey1   = 2
    Const C_PlantCd    = 3
    Const C_ItemCd     = 4
  
	Const C_SHEETMAXROWS_D  = 100 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
	'Key 값을 읽어온다 
	iPlantCd      = Request("txtPlantCd")
	iItemCd       = Request("txtItemCd")
	iStrPrevKey   = Request("lgStrRoutNoPrevKey")          '☜: Next Key Value
	iStrPrevKey1  = Request("lgStrOprNoPrevKey")


    'Component 입력변수        
    ReDim importArray(4)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
    importArray(C_NextKey)		= iStrPrevKey
    importArray(C_NextKey1)		= iStrPrevKey1
    importArray(C_PlantCd)		= iPlantCd
    importArray(C_ItemCd)       = iItemCd
    
    ReDim exportData(1)
   
    Set PC1G065Data = Server.CreateObject("PC1G065.cCListRoutingDtlSvr")
    
	If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
   
    Call PC1G065Data.C_LIST_ROUTING_DETAIL_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PC1G065Data = Nothing
        Response.Write " <Script Language=vbscript>	                         " & vbCr
	    Response.Write " With parent                                         " & vbCr
        Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))		& """" & vbCr
        Response.Write " .frm1.txtItemNm.value = """ & ConvSPChars(exportData(1))		& """" & vbCr
        Response.Write "End With   " & vbCr
        Response.Write "</Script>  " & vbCr
       Exit Sub
    End If    
        
    Set PC1G065Data = nothing    
	
    iStrData = ""
    iIntLoopCount = 0	
    
    Const E_RoutNo = 0
    Const E_OprNo = 1
    Const E_RoutOrder = 2
    Const E_WcCd = 3
    Const E_WcNm = 4
    Const E_ProdRate = 5

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		 If  iIntLoopCount < (lgMaxCount + 1) Then
			iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, E_RoutNo))
            iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, E_OprNo))
			iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, E_RoutOrder))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_WcCd)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_WcNm)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_ProdRate),ggExchRate.DecPoint,0)			
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
         Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_RoutNo)
			iStrPrevKey1 = exportData1(UBound(exportData1, 1), E_OprNo)
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
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))		& """" & vbCr
    Response.Write " .frm1.txtItemNm.value = """ & ConvSPChars(exportData(1))		& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & iPlantCd    & """" & vbCr
    Response.Write " .frm1.hItemCd.value = """ & ihItemCd    & """" & vbCr
    Response.Write " .lgStrRoutNoPrevKey = """ & ConvSPChars(iStrPrevKey)			& """" & vbCr
    Response.Write " .lgStrOprNoPrevKey = """ & ConvSPChars(iStrPrevKey1)    	& """" & vbCr   
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
    Err.Clear                                                '☜: Clear Error status

    Dim PC1G065Data
    Dim importString 
	Dim importString1
	Dim txtSpread
    Dim iErrPosition                                                      
   
   importString  = Trim(Request("txtPlantCd"))
   importString1 = Trim(Request("txtItemCd"))
   txtSpread     = Trim(Request("txtSpread"))

    Set PC1G065Data = Server.CreateObject("PC1G065.cCMngRoutingDtlSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC1G065Data.C_MANAGE_ROUTING_DETAIL_SVR(gStrGlobalCollection, importString, importString1,txtSpread,iErrPosition)			
                                  
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then								
       Call SetErrorStatus
       Set PC1G065Data = Nothing
       Exit Sub
    End If    
    
    Set PC1G065Data = Nothing
	
  
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