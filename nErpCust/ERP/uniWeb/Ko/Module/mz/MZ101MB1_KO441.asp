<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!--
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	
	Const C_SHEETMAXROWS_D = 100
	Dim lgIntFlgMode
	Dim lgOpModeCRUD
	Dim ls_msg
	On Error Resume Next
																								'☜: Protect system from crashing
	Err.Clear 
																								'☜: Clear Error status
	Call HideStatusWnd
	lgOpModeCRUD	=	Request("txtMode")
																								'☜: Read Operation Mode (CRUD)
	
	Select Case lgOpModeCRUD		
	   Case CStr(UID_M0001)																		'☜: Query
	      Call SubBizQuery()
	   Case CStr("DbQuery2")																		'☜: Query
	      Call SubBizQuery2()	      
	   Case CStr(UID_M0002)																		'☜: Save
	      Call SubBizSave()
	   Case CStr(UID_M0003)																		'☜: Delete
	      Call SubBizDelete()
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Err.Clear                                                               '☜: Protect system from crashing
	On Error Resume Next
		Dim LngMaxRow
		Dim iLngRow
		Dim strData
		Dim iStrNextKey
   	Dim PMZG100_KO441
		Dim I1_CONDITION_VALS
		Const I1_PLANT_CD = 0
		Const I1_FR_DT = 1
		Const I1_TO_DT = 2
		Const I1_ITEM_CD = 3
		Const I1_BAL_QTY_FLG = 4
		
		ReDim I1_CONDITION_VALS(I1_BAL_QTY_FLG)
		Dim I2_NEXT_KEY
		Dim EG1_export_group
		Const EG1_RCPT_NO 		= 0
		Const EG1_RCPT_DT 		= 1
		Const EG1_PLANT_CD 		= 2
		Const EG1_PLANT_NM 		= 3
		Const EG1_ITEM_CD 		= 4
		Const EG1_ITEM_NM 		= 5
		Const EG1_SPEC 				= 6
		Const EG1_UNIT 				= 7
		Const EG1_ISSUE_QTY 	= 8
		Const EG1_RCPT_QTY 		= 9
		Const EG1_PRICE 			= 10
		Const EG1_AMT 				= 11
		Const EG1_RCPT_DOC_NO = 12
		Const EG1_CLOSE_YN 		= 13
		Const EG1_BAL_QTY 		= 14
		
		Set PMZG100_KO441 = Server.CreateObject("PMZG100_KO441.cILstImpGoodsMvmt")    

		If CheckSYSTEMError(Err,True) = true then 		
			Exit Sub
		End if

		I1_CONDITION_VALS(I1_PLANT_CD) 		= Request("txtPlantCd")
		I1_CONDITION_VALS(I1_FR_DT) 			= Request("txtFrDt")
		I1_CONDITION_VALS(I1_TO_DT) 			= Request("txtToDt")
		I1_CONDITION_VALS(I1_ITEM_CD) 		= Request("txtItemCd")
		I1_CONDITION_VALS(I1_BAL_QTY_FLG) = Request("rdoBal")
				         				         
     Call PMZG100_KO441.I_LIST_IMP_GOODS_MVMT_HDR(gStrGlobalCollection, _
									  C_SHEETMAXROWS_D, _
									  I1_CONDITION_VALS, _
									  Request("lgStrPrevKey"), _
									  EG1_export_group)
    										      
			If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
				Set PMZG100_KO441 = Nothing												'☜: ComProxy Unload
				Exit Sub															'☜: 비지니스 로직 처리를 종료함 
			End If

	   Set PMZG100_KO441 = Nothing																	'☜: ComProxy UnLoad

    LngMaxRow = CLng(Request("txtMaxRows"))										
	For iLngRow = 0 To UBound(EG1_export_group,1)
	    If  iLngRow < C_SHEETMAXROWS_D  Then
			Else
		   iStrNextKey = ConvSPChars(EG1_export_group(iLngRow, EG1_RCPT_NO)) 
       Exit For
      End If  
      
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_RCPT_NO))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_RCPT_DT))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_PLANT_CD))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ITEM_CD))
        strData = strData & Chr(11)
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ITEM_NM))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_SPEC))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_UNIT))
        strData = strData & Chr(11)
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_RCPT_QTY))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_PRICE))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_AMT))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ISSUE_QTY))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_BAL_QTY))        
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_RCPT_DOC_NO))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_CLOSE_YN))
        strData = strData & Chr(11) & LngMaxRow + iLngRow
        strData = strData & Chr(11) & Chr(12)
    
    Next            
    
    Response.Write "<Script language=vbs> " & vbCr           
    Response.Write " Parent.frm1.vspdData.ReDraw = False " & vbCr
    Response.Write " Parent.ggoSpread.Source	= Parent.frm1.vspdData " &	 	  vbCr
    Response.Write " Parent.ggoSpread.SSShowData      """ & strData	 & """" & vbCr    
    Response.Write " Parent.frm1.txtPlantCd.Value=""" & ConvSPChars(EG1_export_group(0, EG1_PLANT_CD))	 & """" & vbCr    
    Response.Write " Parent.frm1.txtPlantNm.Value=""" & ConvSPChars(EG1_export_group(0, EG1_PLANT_NM))	 & """" & vbCr            
    Response.Write " Parent.lgStrPrevKey              = """ & iStrNextKey						& """" & vbCr  
    Response.Write " Parent.DbQueryOk "															& vbCr   
    Response.Write "</Script> "																	& vbCr      
	
End Sub																				'☜: Process End

'============================================================================================================
' Name : SubBizQuery2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery2()
	Err.Clear                                                               '☜: Protect system from crashing
	On Error Resume Next
		Dim LngMaxRow
		Dim iLngRow
		Dim strData
		Dim iStrNextKey
		Dim EG1_export_group
		Const EG1_SEQ = 0
		Const EG1_ISSUE_DT = 1
		Const EG1_ISSUE_TYPE = 2
		Const EG1_ISSUE_TYPE_NM = 3
		Const EG1_ISSUE_QTY = 4
		Const EG1_ISSUE_DOC_NO = 5
		Dim PMZG100_KO441
		
		Set PMZG100_KO441 = Server.CreateObject("PMZG100_KO441.cILstImpGoodsMvmt")    

		If CheckSYSTEMError(Err,True) = true then 		
			Exit Sub
		End if
				         
     Call PMZG100_KO441.I_LIST_IMP_GOODS_MVMT_DTL(gStrGlobalCollection, _
									  C_SHEETMAXROWS_D, _
									  Request("txtRcptNo"), _
									  Request("lgStrPrevKey2"), _
									  EG1_export_group)
    										      
			If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
				Set PMZG100_KO441 = Nothing												'☜: ComProxy Unload
				Exit Sub															'☜: 비지니스 로직 처리를 종료함 
			End If

	   Set PMZG100_KO441 = Nothing																	'☜: ComProxy UnLoad

	if isEmpty(EG1_export_group) Then
			Exit Sub
	End If
	
    LngMaxRow = CLng(Request("txtMaxRows"))										
	For iLngRow = 0 To UBound(EG1_export_group,1)
	    If  iLngRow < C_SHEETMAXROWS_D  Then
			Else
		   iStrNextKey = ConvSPChars(EG1_export_group(iLngRow, EG1_SEQ)) 
       Exit For
      End If  
      
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_SEQ))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ISSUE_DT))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ISSUE_TYPE))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ISSUE_TYPE_NM))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ISSUE_QTY))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_ISSUE_DOC_NO))
        strData = strData & Chr(11) & LngMaxRow + iLngRow
        strData = strData & Chr(11) & Chr(12)
    
    Next            
    
    Response.Write "<Script language=vbs> " & vbCr           
    Response.Write " Parent.frm1.vspdData1.ReDraw = False " & vbCr
    Response.Write " Parent.ggoSpread.Source	= Parent.frm1.vspdData1 " &	 	  vbCr
    Response.Write " Parent.ggoSpread.SSShowData      """ & strData	 & """" & vbCr    
    Response.Write " Parent.frm1.vspdData1.ReDraw = True " & vbCr
    Response.Write " Parent.lgStrPrevKey2              = """ & iStrNextKey						& """" & vbCr  
    Response.Write " Parent.SetSpreadLock(""B"")" & vbCr  
    Response.Write "</Script> "																	& vbCr      
	
End Sub																				'☜: Process End

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSave()																	'☜: 저장 요청을 받음 
	On Error Resume Next								
	Err.Clear																		'☜: Protect system from crashing
	Dim prErrorPosition
	Dim iStrCommandSent
	Dim PMZG100_KO441
	Dim M_IMP_GOODS_MVMT_HDR
	Const IMP_I1_RCPT_NO = 0
	Const IMP_I1_RCPT_DT = 1
	Const IMP_I1_PLANT_CD = 2
	Const IMP_I1_ITEM_CD = 3
	Const IMP_I1_UNIT = 4
	Const IMP_I1_ISSUE_QTY = 5
	Const IMP_I1_RCPT_QTY = 6
	Const IMP_I1_PRICE = 7
	Const IMP_I1_AMT = 8
	Const IMP_I1_RCPT_DOC_NO = 9
	Const IMP_I1_CLOSE_YN = 10
	ReDim M_IMP_GOODS_MVMT_HDR(IMP_I1_CLOSE_YN)	
	
	
	M_IMP_GOODS_MVMT_HDR(IMP_I1_RCPT_NO) 		= Request("txtRCPT_NO")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_RCPT_DT) 		= Request("txtRCPT_DT")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_PLANT_CD) 	= Request("txtPLANT_CD")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_ITEM_CD) 		= Request("txtITEM_CD")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_UNIT) 			= Request("txtUNIT")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_RCPT_QTY) 	= Request("txtRCPT_QTY")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_PRICE) 			= Request("txtPRICE")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_AMT) 				= Request("txtAMT")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_RCPT_DOC_NO)= Request("txtRCPT_DOC_NO")
	M_IMP_GOODS_MVMT_HDR(IMP_I1_CLOSE_YN) 	= Request("txtCLOSE_YN")

	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

	'-----------------------
	'Data manipulate area
	'-----------------------        
	If Request("txtHDcmd") = "입력" Then
		iStrCommandSent 							= "CREATE"
	ElseIf Request("txtHDcmd") = "수정" Then
		iStrCommandSent 							= "UPDATE"
	ElseIf Request("txtHDcmd") = "삭제" Then
		iStrCommandSent 							= "DELETE"
	End If
	
	'-----------------------
	'Com Action Area
	'-----------------------	
	'⊙: Lookup Pad 동작후 정상적인 데이타 이면, 저장 로직 시작 
    Set PMZG100_KO441 = Server.CreateObject("PMZG100_KO441.cIMntImpGoodsMvmt")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set PMZG100_KO441 = Nothing 		
			Exit Sub
	End If

	 Call PMZG100_KO441.I_MAINT_IMP_GOODS_MVMT(gStrGlobalCollection, _
											  iStrCommandSent, _
											  M_IMP_GOODS_MVMT_HDR, _
											  Request("txtSpread"), _
											  prErrorPosition)

		If CheckSYSTEMError2(Err,True,"",prErrorPosition,"","","") = true then 		
			Set PMZG100_KO441 = Nothing												'☜: ComProxy Unload
			Exit Sub															'☜: 비지니스 로직 처리를 종료함 
		End If

    Set PMZG100_KO441 = Nothing															'☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>"  								& vbCr
	Response.Write "With parent"	  											& vbCr	
	Response.Write "	.DbSaveOk" 	& vbCr
	Response.Write "End With" 		& vbCr
	Response.Write "</Script>"		& vbCr
				
    Set PMZG100_KO441 = Nothing															'☜: Unload Comproxy
	
End Sub																				'☜: Process End

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()																			'☜: 삭제 요청 
	
	On Error Resume Next
    Err.Clear																				'☜: Protect system from crashing
	
	Dim M31111
	Dim I5_m_pur_ord_hdr
	Const M193_I2_po_no = 0										

	Redim I5_m_pur_ord_hdr(76)

	I5_m_pur_ord_hdr(M193_I2_po_no)				= Trim(Request("txtPoNo"))

    Set M31111 = Server.CreateObject("PM3G111.cMMaintPurOrdHdrS")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set M31111 = Nothing 		
			Exit Sub
	End If
					   
	Call M31111.M_MAINT_PUR_ORD_HDR_SVR("F",gStrGlobalCollection, _
									  "DELETE", _
									  "", _
									  "", _
									  "", _
									  "", _
									  I5_m_pur_ord_hdr)

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set M31111 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

    Set M31111 = Nothing															'☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Call parent.DbDeleteOk()" & vbCr
	Response.Write "</Script>" & vbCr
		
    Set M31111 = Nothing																	'☜: Unload Comproxy
												
End Sub

	
%>
