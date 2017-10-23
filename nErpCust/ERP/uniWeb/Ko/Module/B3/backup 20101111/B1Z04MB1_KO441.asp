<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%												

On Error Resume Next
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iLngRow
Dim intGroupCount
Dim PB1G103_KO441		

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 
lgStrPrevKey = Request("lgStrPrevKey")
'Response.Write lgstrprevkey & "::"

Dim iErrorPosition		 
Dim I2_B_CDN_REQ_HDR
Const C_I2_HDR_ITEM_CD = 0
Const C_I2_HDR_CBM_DESCRIPTION = 1
Const C_I2_HDR_ITEM_NM = 2
Const C_I2_HDR_BASIC_UNIT = 3
Const C_I2_HDR_ITEM_ACCT = 4
Const C_I2_HDR_VALID_FROM_DT = 5
Const C_I2_HDR_SPEC = 6
Const C_I2_HDR_VAT_TYPE = 7
Const C_I2_HDR_VAT_RATE = 8
Const C_I2_HDR_PHANTOM_FLG = 9
Const C_I2_HDR_NOTE_DT = 10
Const C_I2_HDR_DEV_PROD_GB = 11

    
Select Case strMode
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	Err.Clear											
  Dim E1_B_CDN_REQ_HDR_KO441
  							
	Set PB1G103_KO441 = Server.CreateObject("PB1G103_KO441.cILstCDNReq")

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
   End If   

	Call PB1G103_KO441.I_LIST_REQ_HDR(gStrGlobalCollection, Request("txtItemCd"), E1_B_CDN_REQ_HDR_KO441)
	  			                                           	
	If CheckSYSTEMError(Err,True) = True Then
       Set PB1G103_KO441 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
    End If   

Const C_EG_ITEM_CD 				= 0
Const C_EG_CBM_DESCRIPTION= 1
Const C_EG_ITEM_NM 				= 2
Const C_EG_BASIC_UNIT 		= 3
Const C_EG_ITEM_ACCT 			= 4
Const C_EG_ITEM_ACCT_NM 	= 5
Const C_EG_VALID_FROM_DT 	= 6
Const C_EG_SPEC 					= 7
Const C_EG_VAT_TYPE 			= 8
Const C_EG_VAT_TYPE_NM 		= 9
Const C_EG_REFERENCE 			= 10
Const C_EG_PHANTOM_FLG 		= 11
Const C_EG_NOTE_DT 				= 12
Const C_EG_DEV_PROD_GB 		= 13
Const C_EG_CONFIRM_FLG 		= 14
Const C_EG_USR_NM 				= 15


	Response.Write "<Script language=vbs>  " & vbCr   			    
	'Response.Write " Parent.CurFormatNumericOCX  " & vbCr   		
	Response.Write " Parent.frm1.txtItemCd2.value					= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_ITEM_CD))			& """" & vbcr
	Response.Write " Parent.frm1.txtItemNm.value					= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_ITEM_NM))			& """" & vbcr
	Response.Write " Parent.frm1.txtCBMdescription.value	= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_CBM_DESCRIPTION))			& """" & vbcr
	Response.Write " Parent.frm1.txtItemNm1.value					= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_ITEM_NM))			& """" & vbcr
	Response.Write " Parent.frm1.txtUnit.value						= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_BASIC_UNIT))			& """" & vbcr
	Response.Write " Parent.frm1.cboItemAcct.value				= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_ITEM_ACCT))			& """" & vbcr
	Response.Write " Parent.frm1.txtValidDt.text					= """ & UNIDateClientFormat(E1_B_CDN_REQ_HDR_KO441(C_EG_VALID_FROM_DT ))         & """" & vbcr	
	Response.Write " Parent.frm1.txtItemSpec.value				= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_SPEC))			& """" & vbcr
	Response.Write " Parent.frm1.txtVatType.value					= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_VAT_TYPE))			& """" & vbcr
	Response.Write " Parent.frm1.txtVatTypeNm.value				= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_VAT_TYPE_NM))			& """" & vbcr
	Response.Write " Parent.frm1.txtVatRate.text					= """ & UNINumClientFormat(E1_B_CDN_REQ_HDR_KO441(C_EG_REFERENCE), ggQty.DecPoint, 0)    & """" & vbcr		
	If E1_B_CDN_REQ_HDR_KO441(C_EG_PHANTOM_FLG) = "Y" Then
		Response.Write " Parent.frm1.rdoPhantomType1.checked=true" & vbcr
	Else
		Response.Write " Parent.frm1.rdoPhantomType2.checked=true" & vbcr
	End If
	Response.Write " Parent.frm1.txtInsUser.value				= """ & ConvSPChars(E1_B_CDN_REQ_HDR_KO441(C_EG_USR_NM))			& """" & vbcr
	Response.Write " Parent.frm1.txtNoteDt.text					= """ & UNIDateClientFormat(E1_B_CDN_REQ_HDR_KO441(C_EG_NOTE_DT ))         & """" & vbcr	
	If E1_B_CDN_REQ_HDR_KO441(C_EG_DEV_PROD_GB) = "Y" Then
		Response.Write " Parent.frm1.rdoDP1.checked=true" & vbcr
	Else
		Response.Write " Parent.frm1.rdoDP2.checked=true" & vbcr
	End If
	Response.Write " Call parent.DbQueryOk() " & vbCr   
	Response.Write "</Script>      " & vbCr      
'----------------------------------detail query------------------------------------------
    
    Dim I2_s_cc_hdr
    Dim I1_s_cc_dtl
         
    Const C_SHEETMAXROWS_D  = 100
        
    Dim EG1_export_group
    
    Const EG1_E1_SEQ 					= 0
    Const EG1_E1_TODO_DOC 		= 1
    Const EG1_E1_COMBO_YN 		= 2
    Const EG1_E1_UD_MAJOR_CD 	= 3
    Const EG1_E1_UD_MINOR_CD 	= 4
    Const EG1_E1_DATA_TEXT 		= 5
    Const EG1_E1_PROCESS_TYPE = 6
    Const EG1_E1_MES_USE_YN 	= 7
    Const EG1_E1_CDN_BIZ 			= 8
    Const EG1_E1_CDN_BMP 			= 9
    Const EG1_E1_CDN_PKG 			= 10
    Const EG1_E1_CDN_PRD 			= 11
    Const EG1_E1_CDN_TQC 			= 12
    Const EG1_E1_REMARK 			= 13
            
    Dim LngLastRow      
    Dim LngMaxRow       
        
    Dim strTemp
    Dim strData
    Dim iStrNextKey
    Dim PS6G215_KO441
           
	Call PB1G103_KO441.I_LIST_REQ_DTL(gStrGlobalCollection, C_SHEETMAXROWS_D, Request("txtItemCd"),Request("lgStrPrevKey"),EG1_export_group)
	                                           	
	If CheckSYSTEMError(Err,True) = True Then
       Set PB1G103_KO441 = Nothing	
       Response.Write "<Script language=vbs>  " & vbCr   
       Response.Write "   Parent.frm1.txtItemCd.focus " & vbCr    
       Response.Write "</Script>      " & vbCr
       Response.End
       'Exit Sub
    End If   

    Set PB1G103_KO441 = Nothing   
                
    LngMaxRow = CLng(Request("txtMaxRows"))										

	For iLngRow = 0 To UBound(EG1_export_group,1)
	    If  iLngRow < C_SHEETMAXROWS_D  Then
			Else
		   iStrNextKey = ConvSPChars(EG1_export_group(iLngRow, EG1_E1_SEQ)) 
       Exit For
      End If  
      
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_SEQ))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_TODO_DOC))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_COMBO_YN))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_UD_MAJOR_CD))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_UD_MINOR_CD))
        strData = strData & Chr(11)
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_DATA_TEXT))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_PROCESS_TYPE))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_MES_USE_YN))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_CDN_BIZ))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_CDN_BMP))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_CDN_PKG))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_CDN_PRD))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_CDN_TQC))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, EG1_E1_REMARK))        
        strData = strData & Chr(11) & LngMaxRow + iLngRow
        strData = strData & Chr(11) & Chr(12)
    
    Next            
    
    Response.Write "<Script language=vbs> " & vbCr           
    Response.Write " Parent.frm1.vspdData.ReDraw = False " & vbCr
    Response.Write " Parent.ggoSpread.Source	= Parent.frm1.vspdData " &	 	  vbCr
    Response.Write " Parent.ggoSpread.SSShowData      """ & strData	 & """" & vbCr
    
    Response.Write " Parent.lgStrPrevKey              = """ & iStrNextKey						& """" & vbCr  
    Response.Write " Parent.DbQueryOk "															& vbCr   
    Response.Write "</Script> "																	& vbCr      
    
Case CStr(UID_M0002)														'☜: 현재 Save 요청을 받음	
	ReDim I2_B_CDN_REQ_HDR(C_I2_HDR_DEV_PROD_GB)

	I2_B_CDN_REQ_HDR(C_I2_HDR_ITEM_CD) 					= Request("txtItemCd2")
	I2_B_CDN_REQ_HDR(C_I2_HDR_CBM_DESCRIPTION) 	= Request("txtCBMdescription")
	I2_B_CDN_REQ_HDR(C_I2_HDR_ITEM_NM) 					= Request("txtItemNm1")
	I2_B_CDN_REQ_HDR(C_I2_HDR_BASIC_UNIT) 			= Request("txtUnit")
	I2_B_CDN_REQ_HDR(C_I2_HDR_ITEM_ACCT) 				= Request("cboItemAcct")
	I2_B_CDN_REQ_HDR(C_I2_HDR_VALID_FROM_DT) 		= Request("txtValidDt")
	I2_B_CDN_REQ_HDR(C_I2_HDR_SPEC) 						= Request("txtItemSpec")
	I2_B_CDN_REQ_HDR(C_I2_HDR_VAT_TYPE) 				= Request("txtVatType")
	I2_B_CDN_REQ_HDR(C_I2_HDR_VAT_RATE) 				= Request("txtVatRate")
	I2_B_CDN_REQ_HDR(C_I2_HDR_PHANTOM_FLG) 			= Request("rdoPhantomType")
	I2_B_CDN_REQ_HDR(C_I2_HDR_NOTE_DT)	 				= Request("txtNoteDt")
	I2_B_CDN_REQ_HDR(C_I2_HDR_DEV_PROD_GB) 			= Request("rdoDP")
		 
    Set PB1G103_KO441 = Server.CreateObject("PB1G103_KO441.cIMntCDNReq")       
   
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
    End If   
    
	Call PB1G103_KO441.I_MAINT_CDN_REQ(gStrGlobalCollection, Request("txtFlgMode"), I2_B_CDN_REQ_HDR, Request("txtSpread"), _
	                                        iErrorPosition)                                                   
	                                           	
	If CheckSYSTEMError2(Err, True, iErrorPosition ,"","","","") = True Then
       Set PB1G103_KO441 = Nothing
       Response.End
	End If  

    Set PB1G103_KO441 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr      
  Response.Write " Parent.frm1.txtItemCd.value      = """ & Request("txtItemCd2")						& """" & vbCr  
	Response.Write " Parent.DBSaveOk "		& vbCr   
	Response.Write "</Script> "				& vbCr      													

Case CStr(UID_M0003)														'☜: 현재 Save 요청을 받음	

	ReDim I2_B_CDN_REQ_HDR(C_I2_HDR_DEV_PROD_GB)

	I2_B_CDN_REQ_HDR(C_I2_HDR_ITEM_CD) 					= Request("txtItemCd")
		 
    Set PB1G103_KO441 = Server.CreateObject("PB1G103_KO441.cIMntCDNReq")       
   
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
    End If   
    
	Call PB1G103_KO441.I_MAINT_CDN_REQ(gStrGlobalCollection, UID_M0003, I2_B_CDN_REQ_HDR, "", _
	                                        iErrorPosition)                                                   
	                                           	
	If CheckSYSTEMError2(Err, True, iErrorPosition ,"","","","") = True Then
       Set PB1G103_KO441 = Nothing
       Response.End
	End If  

    Set PB1G103_KO441 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr      
  Response.Write " Parent.frm1.txtItemCd.value      = """ & Request("txtItemCd2")						& """" & vbCr  
	Response.Write " Parent.DbDeleteOk "		& vbCr   
	Response.Write "</Script> "				& vbCr      													

Case Else
	Response.End
End Select
%>
