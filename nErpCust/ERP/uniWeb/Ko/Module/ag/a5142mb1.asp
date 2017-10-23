<% Option Explicit %>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear
    
    Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I","A", "NOCOOKIE", "MB")  
                                                                           '¢Ð: Clear Error status
	
	Dim lgStrPrevKeyTempGlDt	
	Dim lgStrPrevKeyBatchNo
    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)

    'Multi SpreadSheet

    lgLngMaxRow       = CInt(Request("txtMaxRows"))                                        '¢Ð: Read Operation Mode (CRUD)
    lgMaxCount        = Request("lgMaxCount")                                  '¢Ð: Fetch count at a time for VspdData
    lgStrPrevKeyTempGlDt = Trim(Request("lgStrPrevKeyTempGlDt"))
    lgStrPrevKeyBatchNo = Trim(Request("lgStrPrevKeyBatchNo"))
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             'Call SubBizDelete()
        Case CStr(UID_M0004)                                                         '¢Ð: 
             Call SubBizUpperQuery()   
        Case CStr(UID_M0005)                                                         '¢Ð: 
             Call SubBizDownQuery()
        End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizUpperQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizUpperQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizUpperQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizUpperQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDownQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizDownQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Const C_SHEETMAXROWS	= 100
	
	Const A894_I2_gl_dt_from = 0
    Const A894_I2_gl_dt_to = 1    
    
    Const A894_I3_gl_dt_previous = 0
    Const A894_I3_batch_no_previous = 1
    Const A894_I3_gl_trans_type = 2
    Const A894_I3_auto_trans_fg = 3
    
    
    Dim PAGG116_cAListOneBtchSvr
        
    Dim I1_b_biz_area
    Dim I2_a_batch
    Dim I3_a_batch
    Dim I4_a_vat_type
    Dim E1_a_batch
    Dim E2_a_batch
    Dim E3_a_batch
    Dim EG1_export_group
	
    Dim iStrData			
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iIntMaxRows
    Dim iIntMaxCount
    Dim iIntLoopCount
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear     
                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
  
    
    
    ReDim I2_a_batch(1)
    ReDim I3_a_batch(3)
                
    I1_b_biz_area							= Trim(Request("txtBizCd"))
        
    I2_a_batch(A894_I2_gl_dt_from)			= UNIConvDate(Request("txtFromReqDt"))
    I2_a_batch(A894_I2_gl_dt_to)			= UNIConvDate(Request("txtToReqDt"))
    
    I3_a_batch(A894_I3_gl_dt_previous)		= UNIConvDate(Request("lgStrPrevKeyTempGlDt"))
    I3_a_batch(A894_I3_batch_no_previous)	= Trim(Request("lgStrPrevKeyBatchNo"))
    I3_a_batch(A894_I3_gl_trans_type)		= Trim(Request("txtTransType"))
    
    I4_a_vat_type                           = Trim(Request("txtVatType"))
    
    
	Set PAGG116_cAListOneBtchSvr = Server.CreateObject("PAGG116.cAListOneBtchSvr")	
	  
	
    If CheckSYSTEMError(Err, True) = True Then
		Call SetErrorStatus()
		Exit Sub
    End If    


	
	Call PAGG116_cAListOneBtchSvr.A_LIST_ONE_BATCH_SVR(gStrGloBalCollection, _
												C_SHEETMAXROWS, _
												I1_b_biz_area, _
												I2_a_batch, _
												I3_a_batch, _
												I4_a_vat_type, _
												E1_a_batch, _
												E2_a_batch, _
												E3_a_batch, _
												EG1_export_group)
	
'	if err.number <> 0 then
'		Response.Write "xx" & "::" & err.source & "::" & err.description
'		Response.End
'	end if
	
    If CheckSYSTEMError(Err, True) = True Then					
         Call SetErrorStatus()
         Set PAGG116_cAListOneBtchSvr = Nothing
       'Exit Sub
    End If 
    
    If lgErrorStatus <> "YES" Then
    
		Set PAGG116_cAListOneBtchSvr = Nothing
		iStrData = ""
		iIntLoopCount = 0
		
	 Const A894_EG1_E4_gl_dt = 0
     Const A894_EG1_E4_ref_no = 1
     Const A894_EG1_E1_biz_area_cd = 2
     Const A894_EG1_E1_biz_area_nm = 3
     Const A894_EG1_E4_gl_trans_type = 4
     Const A894_EG1_E4_gl_trnas_type_nm = 5
     Const A894_EG1_E4_batch_no = 6
     Const A894_EG1_E4_item_amt = 7
     Const A894_EG1_E4_item_loc_amt = 8
     Const A894_EG1_E4_desc = 9		
		
		IF isempty(EG1_export_group) = FALSE Then
			For iLngRow = 0 To UBound(EG1_export_group, 1) 
				
				iIntLoopCount = iIntLoopCount + 1
			    
			    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					
					iStrData = iStrData & Chr(11) & "0"
								
					For iLngCol = 0 To UBound(EG1_export_group, 2)
'						IF iLngCol = 0 or iLngCol = 7 Then 
'							iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, iLngCol))
'						ELSE
'							iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, iLngCol)
'						END IF
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A894_EG1_E4_gl_dt))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E4_ref_no)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E1_biz_area_cd)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E1_biz_area_nm)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E4_gl_trans_type)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E4_gl_trnas_type_nm)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E4_batch_no)))
                        iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, A894_EG1_E4_item_amt), ggAmtOfMoney.DecPoint, 0)
						iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, A894_EG1_E4_item_loc_amt), ggAmtOfMoney.DecPoint, 0)
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A894_EG1_E4_desc)))
					Next
						iStrData = iStrData & Chr(11) & Chr(12)
			    Else
					lgStrPrevKeyTempGlDt = EG1_export_group(UBound(EG1_export_group, 1), 0)
					lgStrPrevKeyBatchNo = EG1_export_group(UBound(EG1_export_group, 1), 1)
					Exit For
					  
				End If
			Next
			
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				lgStrPrevKeyTempGlDt = ""
				lgStrPrevKeyBatchNo = ""
			End If
		END IF
	
	End If
	
		
	
	
		Response.Write " <Script Language=vbscript>			" & vbCr
		Response.Write " With parent						" & vbCr
	
	IF not isEmpty(E1_a_batch) Then
		Response.Write "	.frm1.txtBizCd.value = """ & E1_a_batch(0)			& """" & vbCr 			 
		Response.Write "	.frm1.txtBizNm.value = """ & E1_a_batch(1)			& """" & vbCr			 
	ENd IF
	
	IF not isEmpty(E2_a_batch) Then
		Response.Write "	.frm1.txtTransType.value = """ & E2_a_batch(0)	& """" & vbCr 
		Response.Write "	.frm1.txtTransTypeNm.value = """ & E2_a_batch(1)	& """" & vbCr 		 
	END IF
	
	IF not isEmpty(E3_a_batch) Then
		Response.Write "	.frm1.txtVatType.value = """ & E3_a_batch(0)	& """" & vbCr 
		Response.Write "	.frm1.txtVatTypeNm.value = """ & E3_a_batch(1)	& """" & vbCr 		 
	END IF
	
		Response.Write "	.ggoSpread.Source = .frm1.vspdData						 " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & ConvSPChars(iStrData)	& """" & vbCr	
		Response.Write "	.lgStrPrevKeyTempGlDt = """ & lgStrPrevKeyTempGlDt  & """" & vbCr
		Response.Write "	.lgStrPrevKeyBatchNo = """ & lgStrPrevKeyBatchNo    & """" & vbCr
		Response.Write "	.frm1.hFromReqDt.Text = """ & Trim(Request("FromReqDt"))	& """" & vbCr
		Response.Write "	.frm1.hToReqDt.Text = """ & Trim(Request("ToReqDt"))		& """" & vbCr
		Response.Write "	.frm1.hGlTransType.value = """ & Trim(Request("txtTransType"))    & """" & vbCr
		Response.Write "	.frm1.hBizCd.value = """ & Trim(Request("txtBizCd"))				& """" & vbCr
		Response.Write "	.frm1.htxtVatType.value = """ & Trim(Request("txtVatType"))				& """" & vbCr
	If lgErrorStatus <> "NO" Then
'		Response.Write "		.frm1.txtBizCd.value			= """"" & vbCr'
'		Response.Write "		.frm1.txtBizNm.value			= """"" & vbCr	
'		Response.Write "		.frm1.txtTransType.value		= """"" & vbCr
'		Response.Write "		.frm1.txtTransTypeNm.value	= """"" & vbCr
	End If
		Response.Write " .DbQueryOk   " & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr                                                        '¢Ð: Release RecordSSet
		
End Sub    


'============================================================================================================
' Name : SubBizUpperQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizUpperQueryMulti()

	Const C_SHEETMAXROWS	= 100
	
	Const A896_I2_gl_dt_from = 0
    Const A896_I2_gl_dt_to = 1    
    
    Const A896_I3_gl_dt_previous = 0
    Const A896_I3_batch_no_previous = 1
    Const A896_I3_gl_trans_type = 2
    Const A896_I3_auto_trans_fg = 3
    
    
    Dim PAGG116_cAListUpperBtchSvr
        
    Dim I1_b_biz_area
    Dim I2_a_batch
    Dim I3_a_batch
    Dim I4_a_vat_type
    Dim E1_a_batch
    Dim E2_a_batch
    Dim E3_a_batch
    Dim EG1_export_group
	
    Dim iStrData			
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iIntMaxRows
    Dim iIntMaxCount
    Dim iIntLoopCount
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear     
                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
  
    
    
    ReDim I2_a_batch(1)
    ReDim I3_a_batch(3)

    I1_b_biz_area							= Trim(Request("txtBizCd1"))
        
    I2_a_batch(A896_I2_gl_dt_from)			= UNIConvDate(Request("txtFromReqDt1"))
    I2_a_batch(A896_I2_gl_dt_to)			= UNIConvDate(Request("txtToReqDt1"))


    I3_a_batch(A896_I3_gl_dt_previous)		= UNIConvDate(Request("lgStrPrevKeyTempGlDt"))
    I3_a_batch(A896_I3_batch_no_previous)	= Trim(Request("lgStrPrevKeyBatchNo"))
    I3_a_batch(A896_I3_gl_trans_type)		= Trim(Request("txtTransType1"))
    
    I4_a_vat_type                           = Trim(Request("txtVatType1"))
   

	Set PAGG116_cAListUpperBtchSvr = Server.CreateObject("PAGG116.cAListUpperBtchSvr")	
	  
	
    If CheckSYSTEMError(Err, True) = True Then
		Call SetErrorStatus()
		Exit Sub
    End If    


	
	Call PAGG116_cAListUpperBtchSvr.A_LIST_UPPER_BATCH_SVR(gStrGloBalCollection, _
												C_SHEETMAXROWS, _
												I1_b_biz_area, _
												I2_a_batch, _
												I3_a_batch, _
												I4_a_vat_type, _
												E1_a_batch, _
												E2_a_batch, _
												E3_a_batch, _
												EG1_export_group)
	
'	if err.number <> 0 then
'		Response.Write "xx" & "::" & err.source & "::" & err.description
'		Response.End
'	end if
	
    If CheckSYSTEMError(Err, True) = True Then					
         Call SetErrorStatus()
         Set PAGG116_cAListUpperBtchSvr = Nothing
       'Exit Sub
    End If 
    
    If lgErrorStatus <> "YES" Then
    
		Set PAGG116_cAListUpperBtchSvr = Nothing
		iStrData = ""
		iIntLoopCount = 0
		
	 Const A896_EG1_E4_gl_dt = 0
     Const A896_EG1_E4_chain_no = 1
     Const A896_EG1_E4_ref_no = 2
     Const A896_EG1_E1_biz_area_cd = 3
     Const A896_EG1_E1_biz_area_nm = 4
     Const A896_EG1_E4_gl_trans_type = 5
     Const A896_EG1_E4_gl_trans_type_nm = 6
     Const A896_EG1_E4_batch_no = 7
     Const A896_EG1_E2_issued_dt = 8
     Const A896_EG1_E2_gl_no = 9
     Const A896_EG1_E4_item_amt = 10
     Const A896_EG1_E4_item_loc_amt = 11
     Const A896_EG1_E4_desc = 12
     Const A896_EG1_E5_select_char = 13		
		
		IF isempty(EG1_export_group) = FALSE Then
			For iLngRow = 0 To UBound(EG1_export_group, 1) 
				
				iIntLoopCount = iIntLoopCount + 1
			    
			    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					
					iStrData = iStrData & Chr(11) & "0"
								
					For iLngCol = 0 To UBound(EG1_export_group, 2)
'						IF iLngCol = 0 or iLngCol = 8 Then 
'							iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, iLngCol))
'						ELSE
'							iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, iLngCol)
'						END IF
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A896_EG1_E4_gl_dt))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E4_chain_no)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E4_ref_no)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E1_biz_area_cd)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E1_biz_area_nm)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E4_gl_trans_type)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E4_gl_trans_type_nm)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E4_batch_no)))
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A896_EG1_E2_issued_dt))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E2_gl_no)))
                        iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, A896_EG1_E4_item_amt), ggAmtOfMoney.DecPoint, 0)
						iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, A896_EG1_E4_item_loc_amt), ggAmtOfMoney.DecPoint, 0)
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E4_desc)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A896_EG1_E5_select_char)))
					Next
						iStrData = iStrData & Chr(11) & Chr(12)
			    Else
					lgStrPrevKeyTempGlDt = EG1_export_group(UBound(EG1_export_group, 1), 0)
					lgStrPrevKeyBatchNo = EG1_export_group(UBound(EG1_export_group, 1), 2)
					Exit For
					  
				End If
			Next
			
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				lgStrPrevKeyTempGlDt = ""
				lgStrPrevKeyBatchNo = ""
			End If
		END IF
	
	End If 

		Response.Write " <Script Language=vbscript>			" & vbCr
		Response.Write " With parent						" & vbCr
	
	IF not isEmpty(E1_a_batch) Then
		Response.Write "	.frm1.txtBizCd1.value = """ & E1_a_batch(0)			& """" & vbCr 			 
		Response.Write "	.frm1.txtBizNm1.value = """ & E1_a_batch(1)			& """" & vbCr			 
	ENd IF
	
	IF not isEmpty(E2_a_batch) Then
		Response.Write "	.frm1.txtTransType1.value = """ & E2_a_batch(0)	& """" & vbCr 
		Response.Write "	.frm1.txtTransTypeNm1.value = """ & E2_a_batch(1)	& """" & vbCr 		 
	END IF
	
	IF not isEmpty(E3_a_batch) Then
		Response.Write "	.frm1.txtVatType1.value = """ & E3_a_batch(0)	& """" & vbCr 
		Response.Write "	.frm1.txtVatTypeNm1.value = """ & E3_a_batch(1)	& """" & vbCr 		 
	END IF
	
		Response.Write "	.ggoSpread.Source = .frm1.vspdData1						 " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & ConvSPChars(iStrData)	& """" & vbCr	
		Response.Write "	.lgStrPrevKeyTempGlDt = """ & lgStrPrevKeyTempGlDt  & """" & vbCr
		Response.Write "	.lgStrPrevKeyBatchNo = """ & lgStrPrevKeyBatchNo    & """" & vbCr
		Response.Write "	.frm1.hFromReqDt.Text = """ & Trim(Request("FromReqDt")) & """" & vbCr
		Response.Write "	.frm1.hToReqDt.Text = """ & Trim(Request("ToReqDt"))	& """" & vbCr
		Response.Write "	.frm1.hGlTransType.value = """ & Trim(Request("txtTransType"))    & """" & vbCr
		Response.Write "	.frm1.hBizCd.value = """ & Trim(Request("txtBizCd"))				& """" & vbCr
		Response.Write "	.frm1.htxtVatType.value = """ & Trim(Request("txtVatType1"))				& """" & vbCr
	If lgErrorStatus <> "NO" Then
'		Response.Write "		.frm1.txtBizCd1.value			= """"" & vbCr'
'		Response.Write "		.frm1.txtBizNm1.value			= """"" & vbCr	
'		Response.Write "		.frm1.txtTransType1.value		= """"" & vbCr
'		Response.Write "		.frm1.txtTransTypeNm1.value	= """"" & vbCr
	End If
		Response.Write " .DbQueryOk   " & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr                                                        '¢Ð: Release RecordSSet
		
End Sub    

'============================================================================================================
' Name : SubBizDownQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDownQueryMulti()

	Const C_SHEETMAXROWS	= 100
	
	Const A898_I2_gl_chain_no = 0
    Const A898_I2_gl_batch_no = 1    
    
    
    Dim PAGG116_cAListDownBtchSvr
        
    Dim I2_a_batch
    Dim EG1_export_group
	
    Dim iStrData			
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iIntMaxRows
    Dim iIntMaxCount
    Dim iIntLoopCount
    
    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear     
                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
  
    
    
    ReDim I2_a_batch(1)

    I2_a_batch(A898_I2_gl_chain_no)			= Trim(Request("UpperChainNo"))
    I2_a_batch(A898_I2_gl_batch_no)			= Trim(Request("UpperBatchNo"))

	Set PAGG116_cAListDownBtchSvr = Server.CreateObject("PAGG116.cAListDownBtchSvr")	
	  
	
    If CheckSYSTEMError(Err, True) = True Then
		Call SetErrorStatus()
		Exit Sub
    End If    


	
	Call PAGG116_cAListDownBtchSvr.A_LIST_DOWN_BATCH_SVR(gStrGloBalCollection, _
												I2_a_batch, _
												EG1_export_group)
	
	
    If CheckSYSTEMError(Err, True) = True Then					
         Call SetErrorStatus()
         Set PAGG116_cAListDownBtchSvr = Nothing
       'Exit Sub
    End If 
    
    If lgErrorStatus <> "YES" Then
    
		Set PAGG116_cAListDownBtchSvr = Nothing
		iStrData = ""
		iIntLoopCount = 0
		
	 Const A898_EG1_E4_gl_dt = 0
     Const A898_EG1_E4_chain_no = 1
     Const A898_EG1_E4_ref_no = 2
     Const A898_EG1_E1_biz_area_cd = 3
     Const A898_EG1_E1_biz_area_nm = 4
     Const A898_EG1_E4_gl_trans_type = 5
     Const A898_EG1_E3_gl_trans_type_nm = 6
     Const A898_EG1_E4_batch_no = 7
     Const A898_EG1_E2_issued_dt = 8
     Const A898_EG1_E2_gl_no = 9
     Const A898_EG1_E5_select_char = 10		
		
		IF isempty(EG1_export_group) = FALSE Then
			For iLngRow = 0 To UBound(EG1_export_group, 1) 
				
				iIntLoopCount = iIntLoopCount + 1
			    
			    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					
					iStrData = iStrData & Chr(11) & "0"
								
					For iLngCol = 0 To UBound(EG1_export_group, 2)
'						IF iLngCol = 0 or iLngCol = 8 Then 
'							iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, iLngCol))
'						ELSE
'							iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, iLngCol)
'						END IF
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A898_EG1_E4_gl_dt))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E4_chain_no)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E4_ref_no)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E1_biz_area_cd)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E1_biz_area_nm)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E4_gl_trans_type)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E3_gl_trans_type_nm)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E4_batch_no)))
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A898_EG1_E2_issued_dt))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E2_gl_no)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, A898_EG1_E5_select_char)))
					Next
						iStrData = iStrData & Chr(11) & Chr(12)
			    Else
					lgStrPrevKeyTempGlDt = EG1_export_group(UBound(EG1_export_group, 1), 0)
					lgStrPrevKeyBatchNo = EG1_export_group(UBound(EG1_export_group, 1), 2)
					Exit For
					  
				End If
			Next
		END IF
	
	End If
	
		
		Response.Write " <Script Language=vbscript>			" & vbCr
		Response.Write " With parent						" & vbCr
	
		Response.Write "	.ggoSpread.Source = .frm1.vspdData2 					 " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & ConvSPChars(iStrData)	& """" & vbCr	
		Response.Write " .DbQueryOk2   " & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr                                                        '¢Ð: Release RecordSSet
		
End Sub    
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)


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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


%>
