<% Option Explicit 
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3900mb1
'*  4. Program Name         : 평가금액반영 
'*  5. Program Desc         : 원가품목정보 조회, 재고금액평가반영 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/01/09
'*  8. Modified date(Last)  : 2002/06/21
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")     
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim lgStrPlantPrevKey
	Dim lgStrTrnsPrevKey
	Dim lgStrMovPrevKey
	Dim lgStrCostPrevKey
	Dim lgStrItemPrevKey
	Dim lgStrTrnsPlantPrevKey
	Dim lgStrTrnsSlPrevKey
	Dim lgStrTrnsItemPrevKey

	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD
    Dim lgLngMaxRow,   lgMaxCount

	

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
             'Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPC3G045Q
    Dim iStrData
    
    Dim exportData2		'Tot diff. Amt
    Dim exportGroupData	'vspd data
    
    
    Dim iLngRow,iLngCol
    Dim importArray1
    Dim importArray2
    Dim iIntLoopCount
    
   	Const C_SHEETMAXROWS_D  = 100 
   lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time    
    
   
    Const C_TotDiffamt = 0
    
    'Condition
    Const C203_I1_yyyymm = 0
    Const C203_I1_procure_type = 1
    Const C203_I1_plant_cd = 2
    Const C203_I1_trns_type = 3
    Const C203_I1_mov_type = 4
    Const C203_I1_item_acct = 5
    Const C203_I1_Item_cd = 6
    Const C203_I1_Cost_Cd = 7
    
    'Next Key
	Const C203_I2_PLANT_CD = 0
    Const C203_I2_TRNS_TYPE = 1
    Const C203_I2_MOV_TYPE = 2
    Const C203_I2_COST_CD = 3
    Const C203_I2_ITEM_CD = 4
    Const C203_I2_TRNS_PLANT_CD = 5
    Const C203_I2_TRNS_SL_CD = 6
    Const C203_I2_TRNS_ITEM_CD = 7


	lgStrPlantPrevKey		= Trim(Request("lgStrPlantPrevKey"))         '☜: Next Key Value
	lgStrTrnsPrevKey		= Trim(Request("lgStrTrnsPrevKey"))
	lgStrMovPrevKey			= Trim(Request("lgStrMovPrevKey"))         '☜: Next Key Value
	lgStrCostPrevKey		= Trim(Request("lgStrCostPrevKey"))
	lgStrItemPrevKey		= Trim(Request("lgStrItemPrevKey"))         '☜: Next Key Value
	lgStrTrnsPlantPrevKey	= Trim(Request("lgStrTrnsPlantPrevKey"))
	lgStrTrnsSlPrevKey		= Trim(Request("lgStrTrnsSlPrevKey"))         '☜: Next Key Value
	lgStrTrnsItemPrevKey	= Trim(Request("lgStrTrnsItemPrevKey"))

    'Component 입력변수        
    ReDim importArray1(7)
    ReDim importArray2(7)

    importArray1(C203_I1_yyyymm)		= Trim(Request("txtYyyymm")) 
    importArray1(C203_I1_procure_type)	= Trim(Request("cboProcurType")) 
    
    
    importArray1(C203_I1_plant_cd)		= Trim(Request("txtPlantCd"))    
	IF importArray1(C203_I1_plant_cd) = "" Then
		importArray1(C203_I1_plant_cd)		= "%"
	END IF
	         
    importArray1(C203_I1_trns_type)		= Trim(Request("txtTrnsTypeCd"))         
	IF importArray1(C203_I1_trns_type) = "" Then
		importArray1(C203_I1_trns_type)		= "%"
	END IF         

    importArray1(C203_I1_mov_type)		= Trim(Request("txtMovTypeCd"))         
	IF importArray1(C203_I1_mov_type) = "" Then
		importArray1(C203_I1_mov_type)		= "%"
	END IF         

    importArray1(C203_I1_item_acct)		= Trim(Request("txtItemAcctCd"))         
	IF importArray1(C203_I1_item_acct) = "" Then
		importArray1(C203_I1_item_acct)		= "%"
	END IF         

    importArray1(C203_I1_Item_cd)		= Trim(Request("txtItemCd"))         
	IF importArray1(C203_I1_Item_cd) = "" Then
		importArray1(C203_I1_Item_cd)		= "%"
	END IF         

    importArray1(C203_I1_Cost_Cd)		= Trim(Request("txtCostCd"))
	IF importArray1(C203_I1_Cost_Cd) = "" Then
		importArray1(C203_I1_Cost_Cd)		= "%"
	END IF         

	
   
    importArray2(C203_I2_PLANT_CD)		= lgStrPlantPrevKey
	importArray2(C203_I2_TRNS_TYPE)		= lgStrTrnsPrevKey
    importArray2(C203_I2_MOV_TYPE)		= lgStrMovPrevKey
	importArray2(C203_I2_COST_CD)		= lgStrCostPrevKey
	importArray2(C203_I2_ITEM_CD)		= lgStrItemPrevKey
	importArray2(C203_I2_TRNS_PLANT_CD)	= lgStrTrnsPlantPrevKey
	importArray2(C203_I2_TRNS_SL_CD)	= lgStrTrnsSlPrevKey
	importArray2(C203_I2_TRNS_ITEM_CD)	= lgStrTrnsItemPrevKey
		

	Set iPC3G045Q = Server.CreateObject("PC3G045.cClistCItmByPlantSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Call SetErrorStatus
		Exit Sub
    End If  


	
	Call iPC3G045Q.C_LIST_C_ITEM_BY_PLANT_SVR(gStrGloBalCollection, lgMaxCount, importArray1,exportData2,exportGroupData,importArray2)
	
	
	
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Set iPC3G045Q = Nothing
       Exit Sub
    End If    


    
    Set iPC3G045Q = Nothing
	
	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportGroupData, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (lgMaxCount + 1) Then
			For iLngCol = 0 To UBound(exportGroupData, 2)
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportGroupData(iLngRow, iLngCol)))
			Next
				iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
				iStrData = iStrData & Chr(11) & Chr(12)
	    Else

		lgStrPlantPrevKey		= exportGroupData(UBound(exportGroupData, 1), 0)         '☜: Next Key Value
		lgStrTrnsPrevKey		= exportGroupData(UBound(exportGroupData, 1), 1)
		lgStrMovPrevKey			= exportGroupData(UBound(exportGroupData, 1), 3)         '☜: Next Key Value
		lgStrCostPrevKey		= exportGroupData(UBound(exportGroupData, 1), 5)
		lgStrItemPrevKey		= exportGroupData(UBound(exportGroupData, 1), 7)        '☜: Next Key Value
		lgStrTrnsPlantPrevKey	= exportGroupData(UBound(exportGroupData, 1), 12)
		lgStrTrnsSlPrevKey		= exportGroupData(UBound(exportGroupData, 1), 13)         '☜: Next Key Value
		lgStrTrnsItemPrevKey	= exportGroupData(UBound(exportGroupData, 1), 14)
			Exit For
			  
		End If
	Next

		
	If  iIntLoopCount < (lgMaxCount + 1) Then
		lgStrPlantPrevKey = ""
	End If

		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.lgStrPlantPrevKey = """ & lgStrPlantPrevKey    & """" & vbCr
    Response.Write "	.lgStrTrnsPrevKey = """ & lgStrTrnsPrevKey    & """" & vbCr
    Response.Write "	.lgStrMovPrevKey = """ & lgStrMovPrevKey    & """" & vbCr
    Response.Write "	.lgStrCostPrevKey = """ & lgStrCostPrevKey    & """" & vbCr
    Response.Write "	.lgStrItemPrevKey = """ & lgStrItemPrevKey    & """" & vbCr
    Response.Write "	.lgStrTrnsPlantPrevKey = """ & lgStrTrnsPlantPrevKey    & """" & vbCr
    Response.Write "	.lgStrTrnsSlPrevKey = """ & lgStrTrnsSlPrevKey    & """" & vbCr
    Response.Write "	.lgStrTrnsItemPrevKey = """ & lgStrTrnsItemPrevKey    & """" & vbCr
    Response.Write "	.frm1.hYyyymm.value = """ & Trim(Request("txtYyyymm"))   & """" & vbCr
    Response.Write "	.frm1.hProcurType.value = """ & Trim(Request("cboProcurType"))   & """" & vbCr
    Response.Write "	.frm1.hPlantCd.value		= """ & Trim(Request("txtPlantCd"))   & """" & vbCr
    Response.Write "	.frm1.hCostCd.value			= """ & Trim(Request("txtCostCd"))   & """" & vbCr
    Response.Write "	.frm1.hTrnsTypeCd.value		= """ & Trim(Request("txtTrnsTypeCd"))     & """" & vbCr
    Response.Write "	.frm1.hMovTypeCd.value		= """ & Trim(Request("txtMovTypeCd"))  & """" & vbCr
    Response.Write "	.frm1.hItemAcctCd.value		= """ & Trim(Request("txtItemAcctCd"))   & """" & vbCr       
    Response.Write "	.frm1.hItemCd.value			= """ & Trim(Request("txtItemCd"))    & """" & vbCr    

        
    Response.Write "	.frm1.txtSum.text = """    & UNINumClientFormat(exportData2(C_TotDiffamt),ggAmtOfMoney.Decpoint,0)   & """" & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr 
    '---------- Developer Coding part (End)   ---------------------------------------------------------------
End Sub    



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------


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
          With Parent
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				
                .DBQueryOk        
	         
			End If   
          End with
    End Select    
    
       
</Script>	
