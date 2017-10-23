<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Costing
'*  2. Function Name        : 원가계산 기준정보 
'*  3. Program ID           : c1210mb1
'*  4. Program Name         : 간접비 배부율 등록 
'*  5. Program Desc         : 공장별 표준계산시 간접비에 대한 배부율을 등록한다.
'*  6. Modified date(First) : 2000/09/02
'*  7. Modified date(Last)  : 2002/08/09
'*  8. Modifier (First)     : Ig Sung, Cho
'*  9. Modifier (Last)      : Park, Joon-Won
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
             Call SubBizSaveMulti()
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

	Dim iPC4G019
    Dim iStrData
    Dim exportData
    Dim exportGroupData
    Dim iLngRow,iLngCol
    Dim iStrPrevKey, iStrPrevKey1
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd
    Dim lgMaxCount
	

	
	Const C_PlantCd		= 0
	Const C_ItemAcct	= 1
	Const C_ProcurType	= 2
	Const C_ItemCd		= 3
	Const C_NextKey		= 4
    Const C_NextKey1	= 5
    Const C_MaxFetchRc	= 6
    

	Const C_SHEETMAXROWS_D  = 100 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
    iStrPrevKey   = Trim(Request("lgPlantPrevKey"))           '☜: Next Key Value
	iStrPrevKey1  = Trim(Request("lgItemPrevKey"))

	

    'Component 입력변수        
    ReDim importArray(6)
            
    importArray(C_PlantCd)		= Trim(Request("txtPlantCd"))
    importArray(C_ItemAcct)		= Trim(Request("txtItemAcct"))
    importArray(C_ProcurType)	= Trim(Request("txtProcurType"))
    importArray(C_ItemCd)		= Trim(Request("txtItemCd"))
    importArray(C_NextKey)		= iStrPrevKey
    importArray(C_NextKey1)		= iStrPrevKey1
    importArray(C_MaxFetchRc)	= lgMaxCount
	            
    Set iPC4G019 = Server.CreateObject("PC4G019.cCListAdjRateSvr")

	If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus														
       Exit Sub
    End If    

		   
    Call iPC4G019.C_LIST_ADJ_RATE_BY_ITEM_SVR(gStrGlobalCollection,importArray, exportData, exportGroupData)
	
	
	
	If CheckSYSTEMError(Err, True) = True Then					
	    Set iPC4G019 = Nothing
        IF not IsEmpty(exportData) Then
			Response.Write " <Script Language=vbscript>	                         " & vbCr
			Response.Write " parent.frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))		& """" & vbCr
			Response.Write " parent.frm1.txtItemAcctNm.value = """ & ConvSPChars(exportData(1))		& """" & vbCr
			Response.Write " parent.frm1.txtProcurTypeNm.value = """ & ConvSPChars(exportData(2))		& """" & vbCr
			Response.Write " parent.frm1.txtItemNm.value = """ & ConvSPChars(exportData(3))		& """" & vbCr
			Response.Write "</Script>  " & vbCr
		End IF
		Call SetErrorStatus										
       Exit Sub
    End If    
        
 
   
    Const E_PlantCd		= 0
	Const E_PlantNm		= 1
	Const E_ItemCd		= 2  
	Const E_ItemNm		= 3  
	Const E_ItemAcct	= 4
	Const E_ItemAcctNm	= 5
	Const E_ProcurType	= 6  
	Const E_ProcurTypeNm =7  
	Const E_AdjRate		= 8  
	
	
    iStrData = ""
    iIntLoopCount = 0	
    
	For iLngRow = 0 To UBound(exportGroupData, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (lgMaxCount + 1) Then
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, E_PlantCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, E_PlantNm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, E_ItemCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, E_ItemNm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, E_ItemAcctNm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, E_ProcurTypeNm)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportGroupData(iLngRow, E_AdjRate),ggExchRate.DecPoint,0)			
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
        Else
			iStrPrevKey = exportGroupData(UBound(exportGroupData, 1), E_PlantCd)
			iStrPrevKey1 = exportGroupData(UBound(exportGroupData, 1), E_ItemCd)
			Exit For
			  
		End If
	Next
	
	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
		iStrPrevKey1 = ""
	End If

    Set iPC4G019 = nothing  

	
	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))		& """" & vbCr
    Response.Write " .frm1.txtItemAcctNm.value = """ & ConvSPChars(exportData(1))    & """" & vbCr
    Response.Write " .frm1.txtProcurTypeNm.value = """ & ConvSPChars(exportData(2))    & """" & vbCr
    Response.Write " .frm1.txtItemNm.value = """ & ConvSPChars(exportData(3))    & """" & vbCr
    Response.Write " .lgPlantPrevKey = """ & ConvSPChars(iStrPrevKey)			& """" & vbCr
    Response.Write " .lgItemPrevKey = """ & ConvSPChars(iStrPrevKey1)			& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr
    
End Sub    	 


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iPC4G019
    Dim txtSpread 
    Dim iErrPosition

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    txtSpread     = Trim(Request("txtSpread"))


   
    Set iPC4G019 = Server.CreateObject("PC4G019.cCMngAdjRateSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
	
    Call iPC4G019.C_MANAGE_ADJ_RATE_BY_ITEM_SVR(gStrGloBalCollection, txtSpread, iErrPosition)		
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
       Call SetErrorStatus
       Set iPC4G019 = Nothing
       Exit Sub
    End If    
    
    Set iPC4G019 = Nothing

    
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
          End If   
    End Select    
</Script>	
