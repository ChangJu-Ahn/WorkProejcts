<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : Costing
'*  2. Function Name        : 공장별 우선순위 등록 
'*  3. Program ID           : c1904mb1
'*  4. Program Name         : 공장별 우선순위 등록 
'*  5. Program Desc         : 공장별 우선순위 등록 
'*  6. Modified date(First) : 2004/03/22
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Tae Soo
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadBasisGlobalInf() 
   Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

   Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD
   Dim lgLngMaxRow
   Dim lgItemAcctPrevKey,lgPlantPrevKey,lgItemGroupPrevKey,lgItemPrevKey,lgCostElmtPRevKey
   Dim txtItemAcctCd,txtPlantCd,txtItemGroupCd,txtItemCd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            


  Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
	
	
	
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    
  
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


	Dim PC1G085Data		
    Dim iStrData
    Dim exportData1 
    Dim exportData
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd
    Dim lgMaxCount
	


	Const C_SHEETMAXROWS_D  = 100 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
	lgItemAcctPrevKey		= Request("lgItemAcctPrevKey")
	lgPlantPrevKey			= Request("lgPlantPrevKey")
	lgItemGroupPrevKey		= Request("lgItemGroupPrevKey")
	lgItemPrevKey			= Request("lgItemPrevKey")
	lgCostElmtPrevKey		= Request("lgCostElmtPrevKey")
	txtItemAcctCd			= Request("txtItemAcctCd")
	txtPlantCd				= Request("txtPlantCd")
	txtItemGroupCd			= Request("txtItemGroupCd")
	txtItemCd				= Request("txtItemCd")
		

    'Component 입력변수        
    ReDim importArray(9)
            
    importArray(0)	= lgMaxCount
    importArray(1)	= lgItemAcctPrevKey
	importArray(2)	= lgPlantPrevKey
	importArray(3)	= lgItemGroupPrevKey
	importArray(4)	= lgItemPrevKey
	importArray(5)	= lgCostElmtPrevKey
	importArray(6)	= txtItemAcctCd
	importArray(7)	= txtPlantCd
	importArray(8)	= txtItemGroupCd
	importArray(9)	= txtItemCd


		            
    Set PC1G085Data = Server.CreateObject("PC1G085.cCListStdRatebyCESvr")
	
	If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus				
       Exit Sub
    End If    

    Call PC1G085Data.C_LIST_STD_RATE_BY_CE_SVR(gStrGlobalCollection,importArray, exportData,exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PC1G085Data = Nothing
       Call SetErrorStatus
       Exit Sub
    End If    
        
   

    iStrData = ""
    iIntLoopCount = 0

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (lgMaxCount + 1) Then
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 0)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 1)))
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 2)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 3)))
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 4)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 5)))
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 6)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 7)))
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 8)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 9)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, 10),ggExchRate.DecPoint,0)			
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
        Else
			lgItemAcctPrevKey = exportData1(UBound(exportData1, 1), 0)
			lgPlantPrevKey = exportData1(UBound(exportData1, 1), 2)
			lgItemGroupPrevKey = exportData1(UBound(exportData1, 1), 4)
			lgItemPrevKey = exportData1(UBound(exportData1, 1), 6)
			lgCostElmtPrevKey =  exportData1(UBound(exportData1, 1), 8)
			Exit For
			  
		End If
	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
			lgItemAcctPrevKey = ""
			lgPlantPrevKey = ""
			lgItemGroupPrevKey = ""
			lgItemPrevKey = ""
			lgCostElmtPrevKey = ""
	End If
	
	Set PC1G085Data = nothing  
	
	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .lgItemAcctPrevKey = """ & ConvSPChars(lgItemAcctPrevKey)			& """" & vbCr
    Response.Write " .lgPlantPrevKey = """ & ConvSPChars(lgPlantPrevKey)			& """" & vbCr
    Response.Write " .lgItemGroupPrevKey = """ & ConvSPChars(lgItemGroupPrevKey)			& """" & vbCr
    Response.Write " .lgItemPrevKey = """ & ConvSPChars(lgItemPrevKey)			& """" & vbCr
    Response.Write " .lgCostElmtPrevKey = """ & ConvSPChars(lgCostElmtPrevKey)			& """" & vbCr
    Response.Write " .frm1.hItemAcctCd.value = """ & ConvSPChars(txtItemAcctCd)			& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & ConvSPChars(txtPlantCd)			& """" & vbCr
    Response.Write " .frm1.hItemGroupCd.value = """ & ConvSPChars(txtItemGroupCd)			& """" & vbCr
    Response.Write " .frm1.hItemCd.value = """ & ConvSPChars(txtItemCd)			& """" & vbCr
    Response.Write " .frm1.txtItemAcctNm.value = """ & ConvSPChars(exportData(0))			& """" & vbCr
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(1))			& """" & vbCr
    Response.Write " .frm1.txtItemGroupNm.value = """ & ConvSPChars(exportData(2))			& """" & vbCr
    Response.Write " .frm1.txtItemNm.value = """ & ConvSPChars(exportData(3))			& """" & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr
    
End Sub    	 


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PC1G085Data
    Dim importString 
    Dim txtSpread 
    Dim iErrPosition

   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                        '☜: Clear Error status


   txtSpread     = Trim(Request("txtSpread"))
   
    Set PC1G085Data = Server.CreateObject("PC1G085.cCMngStdRateByCESvr")

    If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus				
       Exit Sub
    End If    
	
    Call PC1G085Data.C_MANAGE_STD_RATE_BY_CE_SVR(gStrGloBalCollection, txtSpread, iErrPosition)		
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
       Call SetErrorStatus
       Set PC1G085Data = Nothing
       Exit Sub
    End If    
    
    Set PC1G085Data = Nothing

    
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
