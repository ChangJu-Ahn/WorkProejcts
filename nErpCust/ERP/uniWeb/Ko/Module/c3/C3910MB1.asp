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
'*  9. Modifier (First)     : Ig Sung, Cho
'* 10. Modifier (Last)      : Lee Tae Soo
'* 11. Comment              :
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
	Server.ScriptTimeOut = 10000
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")     
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim lgStrPlantPrevKey
	Dim lgStrTrnsPrevKey
	Dim lgStrMovPrevKey
	Dim lgStrCostPrevKey
	Dim lgStrItemPrevKey
    
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
        Case "ExeReflect"                                                         '☜: Save,Update
             Call ExeReflect()
             'Call SubBizSaveMulti()
        Case "ExeCancel"                                                        '☜: Delete
             Call ExeCancel()
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
    Dim iPC3G050Q
    Dim iStrData
    
   	Const C_SHEETMAXROWS_D  = 100 
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time    
    
    Dim exportData2		'Tot diff. Amt
    Dim exportGroupData	'vspd data
    
    
    Dim iLngRow,iLngCol
    Dim importArray1
    Dim importArray2
    Dim iIntLoopCount
    
   	Dim  arrTemp																
    
	Const C208_I1_yyyymm = 0
    Const C208_I1_plant_cd = 1
    Const C208_I1_trns_type = 2
    Const C208_I1_mov_type = 3
    Const C208_I1_item_acct = 4
    Const C208_I1_Item_cd = 5
    Const C208_I1_Cost_Cd = 6
    
    'Next Key
	Const C208_I2_PLANT_CD = 0
    Const C208_I2_TRNS_TYPE = 1
    Const C208_I2_MOV_TYPE = 2
    Const C208_I2_COST_CD = 3
    Const C208_I2_ITEM_CD = 4
    
    Const C_TotDiffamt = 0


	lgStrPlantPrevKey		= Trim(Request("lgStrPlantPrevKey"))         '☜: Next Key Value
	lgStrTrnsPrevKey		= Trim(Request("lgStrTrnsPrevKey"))
	lgStrMovPrevKey			= Trim(Request("lgStrMovPrevKey"))         '☜: Next Key Value
	lgStrCostPrevKey		= Trim(Request("lgStrCostPrevKey"))
	lgStrItemPrevKey		= Trim(Request("lgStrItemPrevKey"))         '☜: Next Key Value
	
    'Component 입력변수        
    ReDim importArray1(6)
    ReDim importArray2(4)
    
    importArray2(0) = lgStrPlantPrevKey
    importArray2(1) = lgStrTrnsPrevKey
    importArray2(2) = lgStrMovPrevKey
    importArray2(3) = lgStrCostPrevKey
    importArray2(4) = lgStrItemPrevKey

    importArray1(C208_I1_yyyymm)		= Trim(Request("txtYyyymm")) 
   
    
    importArray1(C208_I1_plant_cd)		= Trim(Request("txtPlantCd"))    
	IF importArray1(C208_I1_plant_cd) = "" Then
		importArray1(C208_I1_plant_cd)		= "%"
	END IF

         
    importArray1(C208_I1_trns_type)		= Trim(Request("txtTrnsTypeCd"))         
	IF importArray1(C208_I1_trns_type) = "" Then
		importArray1(C208_I1_trns_type)		= "%"
	END IF         
	

    importArray1(C208_I1_mov_type)		= Trim(Request("txtMovTypeCd"))         
	IF importArray1(C208_I1_mov_type) = "" Then
		importArray1(C208_I1_mov_type)		= "%"
	END IF         


    importArray1(C208_I1_item_acct)		= Trim(Request("txtItemAcctCd"))         
	IF importArray1(C208_I1_item_acct) = "" Then
		importArray1(C208_I1_item_acct)		= "%"
	END IF         

    importArray1(C208_I1_Item_cd)		= Trim(Request("txtItemCd"))         
	IF importArray1(C208_I1_Item_cd) = "" Then
		importArray1(C208_I1_Item_cd)		= "%"
	END IF         

    importArray1(C208_I1_Cost_Cd)		= Trim(Request("txtCostCd"))
	IF importArray1(C208_I1_Cost_Cd) = "" Then
		importArray1(C208_I1_Cost_Cd)		= "%"
	END IF         
	

	Set iPC3G050Q = Server.CreateObject("PC3G050.cCListCRcDifSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If    
	
	Call iPC3G050Q.C_LIST_C_RCPT_DIFF_SVR(gStrGloBalCollection, lgMaxCount, importArray1, exportGroupData,exportData2,importArray2)
	
    
    If CheckSYSTEMError(Err, True) = True Then					
         Set iPC3G050Q = Nothing
       Exit Sub
       
    End If    

    
    Set iPC3G050Q = Nothing
	
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
    
    Response.Write "	.frm1.hYyyymm.value = """ & Trim(Request("txtYyyymm"))   & """" & vbCr
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
Sub ExeReflect()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim iPC3G050S

	Const C_YyyyMm	= 0
	
	
	Dim importArray1
	
	Redim importArray1(0)

	importArray1(C_YyyyMm) = Trim(Request("txtYyyymm"))
	
    Set iPC3G050S = Server.CreateObject("PC3G050.cCRefltOfRcDifSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

	
    Call iPC3G050S.C_REFLECTION_OF_RCPT_DIFF_SVR(gStrGloBalCollection,importArray1)		
		
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Set iPC3G050S = Nothing
       Exit Sub
    End If    
    
    Set iPC3G050S = Nothing
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub ExeCancel()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim iPC3G050S

	Const C_YyyyMm	= 0
	

	Dim importArray1
	
	Redim importArray1(0)
    

	importArray1(C_YyyyMm) = Trim(Request("txtYyyymm"))
	
    Set iPC3G050S = Server.CreateObject("PC3G050.cCCnclRefltRcDifSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

	
    Call iPC3G050S.C_CANCEL_OF_REFLT_RCPT_DIFF_SVR(gStrGloBalCollection,importArray1)		
		
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Set iPC3G050S = Nothing
       Exit Sub
    End If    
    
    Set iPC3G050S = Nothing   
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
                Parent.DBQueryOk        
          End If
		Case "<%="ExeReflect"%>"
		   If Trim("<%=lgErrorStatus%>") = "NO" Then
		      Parent.ExeReflectOk
		   End If   
		Case "<%="ExeCancel"%>"
		   If Trim("<%=lgErrorStatus%>") = "NO" Then
		      Parent.ExeCancelOk
		   End If      
    End Select    
    
       
</Script>	
