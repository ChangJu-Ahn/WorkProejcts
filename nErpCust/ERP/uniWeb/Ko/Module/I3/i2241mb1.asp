<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Monthly Inventory
'*  3. Program ID           : I2241mb1.asp
'*  4. Program Name         : 이월 재고현황 조회 
'*  5. Program Desc         : 이월재고 정보의 자료를 조회한다.
'*  6. Comproxy List        :                            
'                             +I22218ListMonthlyInvSvr
'*  7. Modified date(First) : 2000/04/30
'*  8. Modified date(Last)  : 2005/12/01
'*  9. Modifier (First)     : Soon Ho Kweon
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "B","NOCOOKIE","MB")   

'Hide Processin Bar
Call HideStatusWnd 

On Error Resume Next
Err.Clear

Dim i22218             
Dim strData                                            
Dim strMode
Dim PvArr          

Dim lgStrPrevToKey  
Dim LngMaxRow 
Dim LngRow
Dim GroupCount
'Dim intGroupCount
	Const C_SHEETMAXROWS_D = 500

    '-----------------------
    'IMPORTS View
    '-----------------------
	'Dim I4_to_b_item_cd
	
	Dim I5_fr_b_item
		Const I314_I5_item_cd = 0    
		Const I314_I5_item_acct = 1
		Const I314_I5_item_class = 2
		Const I314_I5_tracking_no = 3
	ReDim I5_fr_b_item(I314_I5_tracking_no)
	
	Dim I7_fr_i_monthly_inventory
		Const I314_I7_mnth_inv_year = 0    
		Const I314_I7_mnth_inv_month = 1
	ReDim I7_fr_i_monthly_inventory(I314_I7_mnth_inv_month)
	
	Dim I8_b_plant_cd
 '-----------------------
 'EXPORTS View
 '-----------------------
	Dim E3_i_monthly_inventory
		Const I314_E3_mnth_inv_year = 0    
		Const I314_E3_mnth_inv_month = 1
	
	'Dim E4_next_b_item_cd
	'	Const I314_E4_item_cd = 0
	'	Const I314_E4_tracking_no = 1
	'Redim E4_next_b_item_cd(I314_E4_tracking_no)
	
	Dim EG1_export_group
		Const I314_EG1_E1_item_cd = 0
		Const I314_EG1_E1_item_nm = 1
		Const I314_EG1_E1_spec = 2
		Const I314_EG1_E1_item_acct = 3
		Const I314_EG1_E1_item_class = 4
		Const I314_EG1_E1_basic_unit = 5
		Const I314_EG1_E2_mnth_inv_year = 6
		Const I314_EG1_E2_mnth_inv_month = 7
		Const I314_EG1_E2_bas_inv_qty = 8
		Const I314_EG1_E2_bas_inv_amt = 9
		Const I314_EG1_E2_rcpt_qty = 10
		Const I314_EG1_E2_rcpt_amt = 11
		Const I314_EG1_E2_issue_qty = 12
		Const I314_EG1_E2_issue_amt = 13
		Const I314_EG1_E2_inv_qty = 14
		Const I314_EG1_E2_inv_amt = 15
		Const I314_EG1_E2_prc_ctrl_indctr = 16
		Const I314_EG1_E2_moving_avg_prc = 17
		Const I314_EG1_E2_std_prc = 18
		Const I314_EG1_E2_tracking_no = 19

	'StrNextKey = Request("lgStrPrevKey")
	'StrNextKey2 = Request("lgStrPrevKey2")
	
	LngMaxRow = CLng(Request("txtMaxRows"))
 
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I8_b_plant_cd = Trim(Request("txtPlantCd"))
	I5_fr_b_item(I314_I5_item_cd)		= Trim(Request("txtItemCd"))
	I5_fr_b_item(I314_I5_item_acct)		= Trim(Request("txtItemAcct"))
	I5_fr_b_item(I314_I5_tracking_no)	= Trim(Request("txtTrackingNo"))
	
	lgStrPrevToKey   = cint(Request("lgStrPrevToKey"))
	 
	I7_fr_i_monthly_inventory(I314_I7_mnth_inv_year)  = Request("txtYyyy")
	I7_fr_i_monthly_inventory(I314_I7_mnth_inv_month) = Request("txtMm")
	    
	'If StrNextKey <> "" then I5_fr_b_item(I314_I5_item_cd) = StrNextKey
	
	'If StrNextKey2 <> "*" then
	'	If I5_fr_b_item(I314_I5_tracking_no) = "" Then 
	'		I5_fr_b_item(I314_I5_tracking_no) = StrNextKey2
	'	End If
	'End If
	                                 
	Set i22218 = Server.CreateObject("PI3G050.cIListMonthlyInvSvr")
	    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End            
	End If    
 
	Call i22218.I_LIST_MONTHLY_INV_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
							I5_fr_b_item, _
							I7_fr_i_monthly_inventory, _
							I8_b_plant_cd, _
							E3_i_monthly_inventory, _
							EG1_export_group, _
							lgStrPrevToKey)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set i22218 = Nothing           
		Response.End              
	End If

	Set i22218 = Nothing
	
	Call GridShow(EG1_export_group)
	
	'============================================================================================================
	'GridShow
	'============================================================================================================
	Sub GridShow(pArr)
		Dim i,j,strData
		Dim LngRow
		
		j = 0
		
		For LngRow = (lgStrPrevToKey - 1) * C_SHEETMAXROWS_D to UBound(pArr,1)
			strData = strData &	Chr(11) & ConvSPChars(EG1_export_group(LngRow, I314_EG1_E1_item_cd)) & _
								Chr(11) & ConvSPChars(EG1_export_group(LngRow, I314_EG1_E1_item_nm)) & _
								Chr(11) & ConvSPChars(EG1_export_group(LngRow, I314_EG1_E1_spec)) & _
								Chr(11) & ConvSPChars(EG1_export_group(LngRow, I314_EG1_E2_tracking_no)) & _
								Chr(11) & ConvSPChars(EG1_export_group(LngRow, I314_EG1_E1_basic_unit)) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_bas_inv_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_bas_inv_amt),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_rcpt_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_rcpt_amt),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_issue_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_issue_amt),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_inv_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
								Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I314_EG1_E2_inv_amt),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0) & _
								Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
							
			j = j+1
		Next
			
		Response.Write "<Script Language=vbscript> "                & vbcr
		Response.Write "With parent "                               & vbcr
		Response.Write " .ggoSpread.Source = .frm1.vspdData  "      & vbcr
		Response.Write "    .frm1.vspdData.Redraw = False   "       & vbCr
		Response.Write " .ggoSpread.SSShowData """ & strData & """" & vbcr
		Response.Write "  .DbQueryOk "								& vbcr
		Response.Write "    .frm1.vspdData.Redraw = True "			& vbCr
		Response.Write "  .frm1.hPlantCd.value  = """ & Request("txtPlantCd")                     & """" & vbcr
		Response.Write "  .frm1.hYyyy.value     = """ & Request("txtYyyy")                        & """" & vbcr
		Response.Write "  .frm1.hMm.value       = """ & Request("txtMm")                          & """" & vbcr
		Response.Write "  .frm1.hItemAcct.value = """ & Request("txtItemAcct")                    & """" & vbcr
		Response.Write "  .frm1.hTrackingNo.value = "" " & ConvSPChars(Trim(Request("txtTrackingNo")))     & " "" " & vbcr 
		If j < C_SHEETMAXROWS_D then 
			Response.Write "    .DbChkFlag = true "	& vbCr  
		Else
			Response.Write "    .DbChkFlag = false " & vbCr  
		End if
	
		Response.Write "    .lgStrPrevToKey = "&lgStrPrevToKey+1&" " & vbCr
		Response.Write "End with "      & vbcr
		Response.Write "</Script> "     & vbcr
		
	End Sub
	
	 

%>


