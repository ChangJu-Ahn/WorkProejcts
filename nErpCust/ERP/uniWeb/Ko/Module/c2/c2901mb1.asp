<%
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 표준원가반영 
'*  3. Program ID           : C2901MB1.asp
'*  4. Program Name         : 표준원가반영 
'*  5. Program Desc         : 표준원가반영 BIZ Logic
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/10/31
'*  8. Modified date(Last)  : 2002/08/09
'*  9. Modifier (First)     : Lee Tae Soo 
'* 10. Modifier (Last)      : Park, Joon-Won
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = TRUE								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")


	On Error Resume Next								'☜: 
	Err.Clear
	
		
	Call HideStatusWnd

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""
    lgstrData		  = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)


    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Fetechd Count
 '   lgStrPrevKey      = Request("lgStrPrevKey")                                      '☜: Next Key
     
	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
   

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizBatch()
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


	Dim PC2G045Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd, iItemAccntCd, iItemCd
	Dim lgMaxCount

	
	Const C_MaxFetchRc  = 0
    Const C_NextKey     = 1
	Const C_PlantCd     = 2
	Const C_ItemAccntCd = 3 
	Const C_ItemCd      = 4	
    
	Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)
    
	'Key 값을 읽어온다 
	iPlantCd     = Trim(Request("txtPlantCd"))
	iItemAccntCd = Trim(Request("txtItemAccntCd"))
	iItemCd      = Trim(Request("txtItemCd"))
	iStrPrevKey	 = Trim(Request("lgStrPrevKey"))         '☜: Next Key Value

	
    'Component 입력변수        
    ReDim importArray(4)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
    importArray(C_PlantCd)		= iPlantCd
	importArray(C_ItemAccntCd)  = iItemAccntCd 
	importArray(C_ItemCd)       = iItemCd 
	
   
	ReDim exportData(2)
	
    Set PC2G045Data = Server.CreateObject("PC2G045.cCListStCoRefltSvr")
    
	If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    
   
    Call PC2G045Data.C_LIST_STD_COST_REFLECTION_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Set PC2G045Data = Nothing
		Response.Write " <Script Language=vbscript>	                         " & vbCr
		Response.Write " With parent                                         " & vbCr
        Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))	& """" & vbCr
		Response.Write " .frm1.txtItemAccntNM.value = """ & ConvSPChars(exportData(1))	& """" & vbCr
	    Response.Write " .frm1.txtItemNm.value = """ & ConvSPChars(exportData(2))	& """" & vbCr	
		Response.Write "End With   " & vbCr
        Response.Write "</Script>  " & vbCr
       Exit Sub
    End If    
        
    Set PC2G045Data = nothing    
	
	Const E_ItemCd = 0										'Spread Sheet의 Column별 상수 
	Const E_ItemNm = 1
	Const E_Basicunit = 2
	Const E_ItemSpec = 3
	Const E_StdPrc = 4
	Const E_StockStdPrc = 5
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
			iStrData = iStrData & Chr(11) & "0"
			iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, E_ItemCd))
			iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, E_ItemNm))
			iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, E_Basicunit))
			iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, E_ItemSpec))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_StdPrc),	ggUnitCost.DecPoint,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_StockStdPrc),	ggUnitCost.DecPoint,0)
			iStrData = iStrData & Chr(11)  
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
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))	& """" & vbCr
    Response.Write " .frm1.txtItemAccntNM.value = """ & ConvSPChars(exportData(1))	& """" & vbCr
    Response.Write " .frm1.txtItemNm.value = """ & ConvSPChars(exportData(2))	& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & iPlantCd    & """" & vbCr
    Response.Write " .frm1.hItemAccntCd.value = """ & iItemAccntCd  & """" & vbCr
    Response.Write " .frm1.hItemCd.value = """ & iItemCd  & """" & vbCr
    Response.Write " .lgStrPrevKey        = """ & ConvSPChars(iStrPrevKey)			& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr

End Sub    	 



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PC2G045Data
    Dim importString
    Dim importString1
	Dim txtSpread
	Dim iErrPosition   
  
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
    importString = Trim(Request("txtPlantCd"))
    importString1 = Trim(Request("txtItemAccntCd"))
 '  importString2 = Trim(Request("hChecked"))
    txtSpread    = Trim(Request("txtSpread"))

    Set PC2G045Data = Server.CreateObject("PC2G045.cCUpdStCoRefltSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC2G045Data.C_UPDATE_STD_COST_REFLECTION_SVR(gStrGlobalCollection, "S",importString, importString1, txtSpread, iErrPosition)
                     			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then								
       Call SetErrorStatus
       Set PC2G045Data = Nothing
       Exit Sub
    End If    
            

    Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
	
    
    Set PC2G045Data = Nothing
	
    
End Sub 

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizBatch()

    Dim PC2G045Data
    Dim importString
    Dim importString1
	Dim txtSpread
	Dim iErrPosition   
  
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
    importString = Trim(Request("txtPlantCd"))
    importString1 = Trim(Request("txtItemAccntCd"))
 '  importString2 = Trim(Request("hChecked"))
    

    Set PC2G045Data = Server.CreateObject("PC2G045.cCUpdStCoRefltSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC2G045Data.C_UPDATE_STD_COST_REFLECTION_SVR(gStrGlobalCollection,"A", importString, importString1)
                     			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then								
       Call SetErrorStatus
       Set PC2G045Data = Nothing
       Exit Sub
    End If    
            

    Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
	
    
    Set PC2G045Data = Nothing
	
    
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
       Case "<%=UID_M0002%>"
           If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.FncBtnOk
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
 
       Case "<%=UID_M0003%>"                                                         '☜ : Save                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.FncBtnOk
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
  
    End Select    
</Script>	

