<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%        

On Error Resume Next                                                             '☜: Protect system from crashing 

Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

Call HideStatusWnd                                                               '☜: Hide Processing message
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query			
         Call SubBizQuery()
    Case CStr(UID_M0002), CStr(UID_M0003)                                       '☜: Save,Update, Delete
         Call SubBizSaveMulti()
End Select

Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubBizQueryMulti()
End Sub    

'============================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub


'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQueryMulti()
	
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey, iStrNextKey
	Dim iArrCols, iArrRows
	Dim iLngSheetMaxRows

	Dim iArrDateInfo 	
	Const S536_I3_from_date = 0   
	Const S536_I3_to_date = 1

    Dim iStrBillType
	Dim iStrPostFlag 
	Dim iStrTaxBizArea 
	Dim iStrBillToParty 
	Dim iStrSoldToParty 
	Dim iStrSalesGrp
	Dim iStrSalesOrg
	
	Dim iArrWhere
	Const S536_E1_bp_cd22 = 0			' 발행처 
	Const S536_E1_bp_nm22 = 1			' 발행처명 
	Const S536_E1_bp_cd33 = 2			' 주문처 
	Const S536_E1_bp_nm33 = 3			' 주문처명 
	Const S536_E1_bill_type44 = 4		' 매출채권유형 
	Const S536_E1_bill_type_nm44 = 5	' 매출채권유형명 
	Const S536_E1_biz_area_cd55 = 6		' 세금신고사업장 
	Const S536_E1_biz_area_nm55 = 7		' 세금신고사업장명 
	Const S536_E1_sales_grp_cd66 = 8		' 영업그룹 
	Const S536_E1_sales_grp_nm66 = 9		' 영업그룹명 
	Const S536_E1_sales_org_cd77 = 10		' 영업조직 
	Const S536_E1_sales_org_nm77 = 11		' 영업조직명 


	Dim iArrRsOut 	
	Const S536_EG1_E2_bp_cd = 0   
	Const S536_EG1_E2_bp_nm = 1
	Const S536_EG1_E5_bill_no = 2 
	Const S536_EG1_E5_bill_dt = 3
	Const S536_EG1_E5_cur = 4
	Const S536_EG1_E5_bill_amt = 5
	Const S536_EG1_E5_vat_amt = 6
	Const S536_EG1_E5_tax_biz_area = 7
	Const S536_EG1_E5_collect_amt = 8
	Const S536_EG1_E5_post_flag = 9
	Const S536_EG1_E5_ext1_qty = 10
	Const S536_EG1_E5_ext2_qty = 11
	Const S536_EG1_E5_ext3_qty = 12
	Const S536_EG1_E5_ext1_amt = 13
	Const S536_EG1_E5_ext2_amt = 14
	Const S536_EG1_E5_ext3_amt = 15
	Const S536_EG1_E5_ext1_cd = 16
	Const S536_EG1_E5_ext2_cd = 17
	Const S536_EG1_E5_ext3_cd = 18
	Const S536_EG1_E3_bp_cd = 19
	Const S536_EG1_E3_bp_nm = 20
	Const S536_EG1_E1_bill_type = 21
	Const S536_EG1_E1_bill_type_nm = 22
	Const S536_EG1_E4_biz_area_nm = 23
	Const S536_EG1_E3_sales_grp_cd = 24
	Const S536_EG1_E3_sales_grp_nm = 25

    Dim PS7G159 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status
    
    ReDim iArrDateInfo(1)     

	Dim C_SHEETMAXROWS_D				' 한번에 Query할 Row수 

	If Request("txtBatchQuery") = "Y" Then
		C_SHEETMAXROWS_D = -1			' 조회조건에 해당되는 모든 Row를 반환한다.
	Else
		C_SHEETMAXROWS_D = 100
	End If

	iArrDateInfo(S536_I3_from_date) = UNIConvDate(Request("txtReqDateFrom"))
	iArrDateInfo(S536_I3_to_date) = UNIConvDate(Request("txtReqDateTo"))
	iStrBillType = Trim(Request("txtBillTypeCd"))
	iStrTaxBizArea = Trim(Request("txtTaxBizAreaCd"))
	iStrBillToParty = Trim(Request("txtBillToPartyCd"))
	iStrSoldToParty = Trim(Request("txtSoldToPartyCd"))
	iStrPostFlag = Trim(Request("txtPostFlag"))
	iStrSalesGrp = Trim(Request("txtSalesGrpCd"))
	iStrSalesOrg = Trim(Request("txtSalesOrgCd"))
        
    iStrPrevKey = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	    
  	
	Set PS7G159 = Server.CreateObject("PS7G159.CSLtBlHdrForBatPost")	
	
	if CheckSYSTEMError(Err,True) = True Then 
        Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Call Parent.SetFocusToDocument(""M"") " & vbCr    
        Response.Write "   Parent.frm1.txtReqDateFrom.focus " & vbCr    
        Response.Write "</Script>      " & vbCr
		Exit Sub
	end if
               
	Call PS7G159.S_LIST_BILL_HDR_FOR_BATCH_POST(gStrGlobalCollection, C_SHEETMAXROWS_D, iStrBillType, _
												iStrPrevKey, iArrDateInfo, iStrPostFlag, iStrTaxBizArea, _												
												iStrBillToParty, iStrSoldToParty, iStrSalesGrp, iStrSalesOrg, iArrWhere , iArrRsOut)
        
    If CheckSYSTEMError(Err,True) = True Then 		    
		Response.Write "<Script language=vbs> " & vbCr       
		Response.Write " Parent.frm1.txtBillTypeNm.value    = """ & ConvSPChars(iArrWhere(S536_E1_bill_type_nm44))     & """" & vbCr        
		Response.Write " Parent.frm1.txtSoldToPartyNm.value = """ & ConvSPChars(iArrWhere(S536_E1_bp_nm33))            & """" & vbCr 
		Response.Write " Parent.frm1.txtBillToPartyNm.value = """ & ConvSPChars(iArrWhere(S536_E1_bp_nm22))            & """" & vbCr    
		Response.Write " Parent.frm1.txtTaxBizAreaNm.value  = """ & ConvSPChars(iArrWhere(S536_E1_biz_area_nm55))      & """" & vbCr 
		Response.Write " Parent.frm1.txtSalesGrpNm.value  = """ & ConvSPChars(iArrWhere(S536_E1_sales_grp_nm66))      & """" & vbCr    
		Response.Write " Parent.frm1.txtSalesOrgNm.value  = """ & ConvSPChars(iArrWhere(S536_E1_sales_org_nm77))      & """" & vbCr		   
		Response.Write " Call Parent.SetFocusToDocument(""M"") " & vbCr    
		Response.Write "   Parent.frm1.txtReqDateFrom.focus " & vbCr    
		Response.Write "</Script> "				& vbCr          		
		Set PS7G159 = Nothing
		Exit Sub
	End if
	
	Set PS7G159 = Nothing	    

	' Set Next key
	If C_SHEETMAXROWS_D > 0 And Ubound(iArrRsOut,1) = C_SHEETMAXROWS_D Then
		'출고번호 
		iStrNextKey = iArrRsOut(C_SHEETMAXROWS_D, S536_EG1_E5_bill_no)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(iArrRsOut,1)
	End If

	ReDim iArrCols(19)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 
	
	iLngMaxRow  = CLng(Request("txtMaxRows")) + 1					'☜: Fetechd Count        
			
	iArrCols(0) = ""
   	iArrCols(1) = "0"

	For iLngRow = 0 To iLngSheetMaxRows
   		iArrCols(2) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_post_flag))			' 확정여부 
   		iArrCols(3) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_bill_no))			' 매출채권번호 
   		iArrCols(4) = UNIDateClientFormat(iArrRsOut(iLngRow,S536_EG1_E5_bill_dt)) 	' 매출채권일 
   		iArrCols(5) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E3_bp_cd ))			' 주문처 
   		iArrCols(6) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E3_bp_nm ))			' 주문처명 
   		iArrCols(7) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_cur ))				' 화폐단위 
   		iArrCols(8) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_bill_amt))			' 매출채권금액 
   		iArrCols(9) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_vat_amt))			' VAT금액 
   		iArrCols(10) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_collect_amt))		' 수금액 
   		iArrCols(11) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E2_bp_cd))			' 발행처 
   		iArrCols(12) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E2_bp_nm))			' 발행처명 
   		iArrCols(13) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E1_bill_type))		' 매출채권유형 
   		iArrCols(14) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E1_bill_type_nm))		' 매출채권유형명 
   		iArrCols(15) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E5_tax_biz_area))		' 세금신고사업장 
   		iArrCols(16) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E4_biz_area_nm ))		' 세금신고사업장명 
   		iArrCols(17) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E3_sales_grp_cd))			' 영업그룹 
   		iArrCols(18) = ConvSPChars(iArrRsOut(iLngRow,S536_EG1_E3_sales_grp_nm))			' 영업그룹명 
   		iArrCols(19) = iLngMaxRow + iLngRow 
   		
   		iArrRows(iLngRow) = Join(iArrCols, gColSep)
    Next        
    
    Response.Write "<Script language=vbs> " & vbCr       
    ' 조회조건명 설정 
    Response.Write " Parent.frm1.txtBillTypeNm.value    = """ & ConvSPChars(iArrWhere(S536_E1_bill_type_nm44))     & """" & vbCr        
    Response.Write " Parent.frm1.txtSoldToPartyNm.value = """ & ConvSPChars(iArrWhere(S536_E1_bp_nm33))            & """" & vbCr 
    Response.Write " Parent.frm1.txtBillToPartyNm.value = """ & ConvSPChars(iArrWhere(S536_E1_bp_nm22))            & """" & vbCr    
    Response.Write " Parent.frm1.txtTaxBizAreaNm.value  = """ & ConvSPChars(iArrWhere(S536_E1_biz_area_nm55))      & """" & vbCr
    Response.Write " Parent.frm1.txtSalesGrpNm.value  = """ & ConvSPChars(iArrWhere(S536_E1_sales_grp_nm66))      & """" & vbCr    
    Response.Write " Parent.frm1.txtSalesOrgNm.value  = """ & ConvSPChars(iArrWhere(S536_E1_sales_org_nm77))      & """" & vbCr     
    
    ' scroll bar의 조회를 위해 조회조건 hidden값으로 설정 
    If iStrNextKey <> "" Then
		Response.Write " Parent.frm1.HBillToParty.value  = """ & Trim(Request("txtBillToPartyCd")) & """" & vbCr
		Response.Write " Parent.frm1.HSoldToParty.value  = """ & Trim(Request("txtSoldToPartyCd")) & """" & vbCr
		Response.Write " Parent.frm1.HReqDateFrom.value  = """ & iArrDateInfo(S536_I3_from_date)   & """" & vbCr
		Response.Write " Parent.frm1.HReqDateTo.value    = """ & iArrDateInfo(S536_I3_to_date)	   & """" & vbCr
		Response.Write " Parent.frm1.HBillTypeCd.value   = """ & Trim(Request("txtBillTypeCd"))    & """" & vbCr
		Response.Write " Parent.frm1.HTaxBizAreaCd.value = """ & Trim(Request("txtTaxBizAreaCd"))  & """" & vbCr
		Response.Write " Parent.frm1.HSalesGrpCd.value = """ & Trim(Request("txtSalesGrpCd"))  & """" & vbCr
		Response.Write " Parent.frm1.HSalesOrgCd.value = """ & Trim(Request("txtSalesOrgCd"))  & """" & vbCr
		Response.Write " Parent.frm1.HPostFlag.value     = """ & Trim(Request("txtPostFlag"))	   & """" & vbCr
	End If
    
    Response.Write " Parent.ggoSpread.Source         = Parent.frm1.vspdData									      " & vbCr
 
    Response.Write  "Parent.frm1.vspdData.Redraw = False   "                     & vbCr      
    Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write " Parent.lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr  
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write  "Parent.frm1.vspdData.Redraw = True " & vbCr      
    Response.Write "</Script> "																							& vbCr      
    
End Sub    

'============================================
' Name : SubBizSave
' Desc : Save Data 
'============================================
Sub SubBizSaveMulti()

	Dim PS7G116
	Dim iCommandSent 
	Dim iErrorPosition	

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	iCommandSent = "SAVE"
	
	Set PS7G116 = Server.CreateObject("PS7G116.cSBatchArProcessSvr")	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If 
    
    Call PS7G116.S_BATCH_AR_PROCESS_SVR(gStrGlobalCollection, iCommandSent , _
										cstr(Trim(Request("txtSpread"))),iErrorPosition )					
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set PS7G116 = Nothing
	    Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " Call Parent.SubSetErrPos(" & iErrorPosition & ")" & vbCr
	    Response.Write "</Script> "               
		Exit Sub
	End If	
	
    Set PS7G116 = Nothing    
        
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "               
    
End Sub
%>
