<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : COST
'*  2. Function Name        : C_COST_BASIC_DATA_FOR_COSTING
'*  3. Program ID           : c1801mb1.asp
'*  4. Program Name         : CC별 배부근거 등록 
'*  5. Program Desc         : CC별 배부근거 조회, 등록, 수정, 삭제 
'*  6. Modified date(First) : 2000/11/06
'*  7. Modified date(Last)  : 2002/08/09
'*  8. Modifier (First)     : Cho Ig Sung / Park, Joon-Won
'*  9. Modifier (Last)      : Joo, Sungho
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

	Call HideStatusWnd								'☜: 

  
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
        Case CStr(UID_M0004)                                                         '☜: copy
             Call SubBizBulkCopy()
    End Select

'============================================================================================================
' Name : SubBizBulkCopy
' Desc : Copy Data from Previous Month
'============================================================================================================

Sub SubBizBulkCopy()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PC1G060Data
    Dim txtYYYYMM
    Dim txtDstbFctrCd

    txtYYYYMM = Request("txtYYYYMM")
      txtDstbFctrCd = Request("txtDstbFctrCd")

    Set PC1G060Data = Server.CreateObject("PC1G060.cCMngDstBsisByCcSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC1G060Data.C_MANAGE_DSTB_BASIS_BY_CC_SVR(gStrGlobalCollection, txtYYYYMM, txtDstbFctrCd , , ,"X")			
		
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Set PC1G060Data = Nothing
       Exit Sub
    End If    
    
    Set PC1G060Data = Nothing    
        
End Sub



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


	Dim PC1G060Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim importArray
    Dim iIntLoopCount
	Dim iYyyymm, i
	Dim lgMaxCount
	Dim iDstbFctrCd, iCostCd

	
	Const C_SHEETMAXROWS_D  = 500                      

	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_Yyyymm     = 2
    Const C_DstbFctrCd = 3
	Const C_CostCd     = 4
          
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
     
	'Key 값을 읽어온다 
	iYyyymm     = Request("txtYyyymm")
	iDstbFctrCd = Request("txtDstbFctrCd")
	iCostCd     = Request("txtCostCd")
	iStrPrevKey = Request("lgStrPrevKey")
	
  	  	
    'Component 입력변수        
    ReDim importArray(4)
   
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	importArray(C_Yyyymm)		= iYyyymm
	importArray(C_DstbFctrCd)	= iDstbFctrCd
	importArray(C_CostCd)       = iCostCd
	
	Redim exportData(2)
   
    Set PC1G060Data = Server.CreateObject("PC1G060.cCListDstBsisByCcSvr")
    
	If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus														
       Exit Sub
    End If    
    
    Call PC1G060Data.C_LIST_DSTB_BASIS_BY_CC_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then	
       Set PC1G060Data = Nothing
        Response.Write " <Script Language=vbscript>	                         " & vbCr
	    Response.Write " With parent                                         " & vbCr
        Response.Write " .frm1.txtDstbFctrNm.value = """ & ConvSPChars(exportData(0))   & """    " & vbcr
        Response.Write " .frm1.txtCostNm.value = """ & ConvSPChars(exportData(2))   & """    " & vbcr
		Response.Write " .frm1.hDataExists.value = """ & "N"  & """    " & vbCr        
		Response.Write " .DBQueryOk   " & vbCr        
        Response.Write "End With   " & vbCr
        Response.Write "</Script>  " & vbCr 
        Call SetErrorStatus										 
       Exit Sub
    End If    
  
    Set PC1G060Data = nothing    
	
    iStrData = ""
    iIntLoopCount = 0
    
    Const E_CostCd = 0 
    Const E_CostNM = 1
    Const E_ProdBasicQty = 2
    Const E_ProdBasicAmt = 3
    Const E_AdjRate = 4	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
   	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_CostCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_CostNM)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_ProdBasicQty),ggQty.DecPoint,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_ProdBasicAmt),ggAmtOfMoney.DecPoint,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_AdjRate),ggExchRate.DecPoint,0)						
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
        Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_CostCd)
			Exit For
		End If
 	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey  = ""
	End If

	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData		& """    " & vbCr
    Response.Write " .frm1.txtDstbFctrNm.value = """ & ConvSPChars(exportData(0))   & """    " & vbcr
    Response.Write " .frm1.txtCostNm.value = """ & ConvSPChars(exportData(2))   & """    " & vbcr
    Response.Write " .frm1.hYyyymm.value = """ & UNIDateClientFormat(iYyyymm)  & """    " & vbCr
    Response.Write " .frm1.hDstbFctrCd.value = """ & iDstbFctrCd        & """    " & vbCr 
    Response.Write " .frm1.hCostCd.value = """ & iCostCd        & """    " & vbCr 
    Response.Write " .frm1.hGenFlag.value = """ & ConvSPChars(exportData(1))        & """    " & vbCr 
    Response.Write " .lgStrPrevKey          = """ & ConvSPChars(iStrPrevKey) & """    " & vbCr
  
    Response.Write " .frm1.hDataExists.value = """ & "Y"  & """    " & vbCr

	
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

  
    Dim PC1G060Data
    Dim importString 
    Dim importString1
    Dim txtSpread
    Dim iErrPosition  
    
    importString  = Trim(Request("hYYYYMM"))
    importString1 = Trim(Request("txtDstbFctrCd"))
    txtSpread     = Trim(Request("txtSpread"))
    
    Set PC1G060Data = Server.CreateObject("PC1G060.cCMngDstBsisByCcSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC1G060Data.C_MANAGE_DSTB_BASIS_BY_CC_SVR(gStrGlobalCollection,importString,importString1,txtSpread, iErrPosition, "a")			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then				
       Set PC1G060Data = Nothing
       Call SetErrorStatus
       Exit Sub
    End If    
    
    Set PC1G060Data = Nothing
	

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
       Case "<%=UID_M0004%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOK
          End If  
    End Select    
</Script>