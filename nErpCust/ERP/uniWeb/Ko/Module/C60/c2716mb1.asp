<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 표준원가관리 
'*  3. Program ID           : c2716mb1
'*  4. Program Name         : 품목별 가공비 정보 등록 
'*  5. Program Desc         : 품목별 인건비/제조경비 금액 및 원가요소 정보를 설정한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/08/24
'*  8. Modified date(Last)  : 2002/06/21
'*  9. Modifier (First)     : Chang Goo, Kang
'* 10. Modifier (Last)      : Cho Ig Sung / Park, Joon-Won
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True	
							'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow

	'On Error Resume Next								'☜: 
	'Err.Clear	
	
	Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
 '  lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    
  
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

	Dim PC2G035Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd
    Dim lgMaxCount
	
	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_PlantCd    = 2

	Const C_SHEETMAXROWS_D  = 100 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
	iStrPrevKey = Trim(Request("lgStrPrevKey"))           '☜: Next Key Value

	
	'Key 값을 읽어온다 
	iPlantCd = Trim(Request("txtPlantCd"))

    'Component 입력변수        
    ReDim importArray(2)
            
    importArray(C_MaxFetchRc)	= lgMaxCount
    importArray(C_NextKey)		= iStrPrevKey
    importArray(C_PlantCd)		= iPlantCd
	            
    Set PC2G035Data = Server.CreateObject("PC2G035.cCListPrcsCoByItmSvr")
	
	If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus																		
       Exit Sub
    End If    
   
    Call PC2G035Data.C_LIST_PRC_COST_BY_ITEM_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PC2G035Data = Nothing
		Response.Write " <Script Language=vbscript>	                         " & vbCr
		Response.Write " parent.frm1.txtPlantNm.value = """ & ConvSPChars(exportData)		& """" & vbCr
		Response.Write "</Script>  " & vbCr
		Call SetErrorStatus														
       Exit Sub
    End If    
        
    Set PC2G035Data = nothing    
	
	Const E_ItemCd = 0
	Const E_ItemNm = 1
	Const E_LaborCost = 2
	Const E_LaborCostElmtCd = 3													'☆: Spread Sheet의 Column별 상수 
	Const E_LaborCostElmtNm = 4
	Const E_Expense = 5	
	Const E_ExpenseCostElmtCd = 6
    Const E_ExpenseCostElmtNm = 7
	
    iStrData = ""
    iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (lgMaxCount + 1) Then
            iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ItemCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ItemNm)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_LaborCost),ggUnitCost.DecPoint,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_LaborCostElmtCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_LaborCostElmtNm)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_Expense),ggUnitCost.DecPoint,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ExpenseCostElmtCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ExpenseCostElmtNm)))
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
        Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_ItemCd)
			Exit For
			  
		End If
	Next
	
	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
	End If
	
	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData)		& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & iPlantCd    & """" & vbCr
    Response.Write " .lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)			& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr
    
End Sub    	 


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PC2G035Data
    Dim importString
    Dim txtSpread 
    Dim iErrPosition   

   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
   importString = Trim(Request("txtPlantCd"))
   txtSpread    = Trim(Request("txtSpread"))

    Set PC2G035Data = Server.CreateObject("PC2G035.cCMngPrcsCoByItmSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    
	
    Call PC2G035Data.C_MANAGE_PRC_COST_BY_ITEM_SVR(gStrGloBalCollection, importString, txtSpread, iErrPosition)		
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then								
	   Call SetErrorStatus
       Set PC2G035Data = Nothing
       Exit Sub
    End If    
    
    Set PC2G035Data = Nothing
	

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