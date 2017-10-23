<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Material valuation
'*  3. Program ID           : I2311mb1.asp
'*  4. Program Name         : Material Valuation list
'*  5. Program Desc         : 공장별 품목재고정보 자료를 조회한다.
'*  6. Comproxy List        :                            
'                             +i31118ListMonthlyInvSvr
'*  7. Modified date(First) : 2000/05/03
'*  8. Modified date(Last)  : 2005/02/17
'*  9. Modifier (First)     : Soon Ho Kweon
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : tracking no addition
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
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

Dim i31118             
Dim strData
Dim strMode            

Dim LngMaxRow  
Dim LngRow
Dim GroupCount
Dim lgStrPrevToKey
Dim PvArr           
	Const C_SHEETMAXROWS_D = 500     
    '-----------------------
    'IMPORTS View
    '-----------------------

Dim I4_fr_b_item
  Const I305_I4_item_cd = 0
  Const I305_I4_item_acct = 1
  Const I305_I4_item_class = 2
  Const I305_I4_tracking_no = 3
ReDim I4_fr_b_item(I305_I4_tracking_no)

Dim I5_b_plant_cd
Dim I6_qty_flag
 '-----------------------
 'EXPORTS View
 '-----------------------
 Dim EG1_export_group
  Const I305_EG1_E1_location = 0
  Const I305_EG1_E2_item_cd = 1
  Const I305_EG1_E2_item_nm = 2
  Const I305_EG1_E2_spec = 3
  Const I305_EG1_E2_item_acct = 4
  Const I305_EG1_E2_item_class = 5
  Const I305_EG1_E2_basic_unit = 6
  Const I305_EG1_E3_prc_ctrl_indctr = 7
  Const I305_EG1_E3_moving_avg_prc = 8
  Const I305_EG1_E3_std_prc = 9
  Const I305_EG1_E3_tot_stk_qty = 10
  Const I305_EG1_E3_tot_stk_val = 11
  Const I305_EG1_E3_prev_prc_ctrl = 12
  Const I305_EG1_E3_prev_moving_avg_prc = 13
  Const I305_EG1_E3_prev_std_prc = 14
  Const I305_EG1_E3_prev_tot_stk_qty = 15
  Const I305_EG1_E3_prev_tot_stk_val = 16
  Const I305_EG1_E3_curr_yr = 17
  Const I305_EG1_E3_curr_nmth = 18
  Const I305_EG1_E3_tracking_no = 19

 '-----------------------
 'Data manipulate  area(import view match)
 '-----------------------
    I5_b_plant_cd                   = Trim(Request("txtPlantCd"))    
    I4_fr_b_item(I305_I4_item_cd)  = Trim(Request("txtItemCd"))
    I4_fr_b_item(I305_I4_item_acct) = Trim(Request("txtAccntCd"))
    I4_fr_b_item(I305_I4_tracking_no) = Trim(Request("txtTrackingNo"))
    lgStrPrevToKey   = cint(Request("lgStrPrevToKey")  )
    I6_qty_flag = Request("txtFlag")

    
 Set i31118 = Server.CreateObject("PI3G030.cIListMaterialVal")
    
 If CheckSYSTEMError(Err, True) = True Then
	Response.End            
 End If    
 
 Call i31118.I_LIST_MATERIAL_VALUATION_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
          I4_fr_b_item, _
          I5_b_plant_cd, _
          I6_qty_flag, _
		  EG1_export_group, _ 
          lgStrPrevToKey )
          
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err, True) = True Then
	Set i31118 = Nothing          
	Response.End             
End If

Set i31118 = Nothing

call GridShow(EG1_export_group)
    
    
'============================================================================================================
'GridShow
'============================================================================================================
Sub GridShow(pArr)
	Dim i,j,strData
	Dim LngRow
	j=0
        For i=(lgStrPrevToKey - 1) * C_SHEETMAXROWS_D to uBound(pArr,1) 
		  
			strData =	strData & Chr(11) & ConvSPChars(pArr(i, I305_EG1_E2_item_cd)) 
			strData =	strData & Chr(11) & ConvSPChars(pArr(i, I305_EG1_E2_item_nm)) 
			strData =	strData & Chr(11) & ConvSPChars(pArr(i, I305_EG1_E2_spec)) 
			strData =	strData & Chr(11) & ConvSPChars(pArr(i, I305_EG1_E3_tracking_no)) 
			strData =	strData & Chr(11) & ConvSPChars(pArr(i, I305_EG1_E2_basic_unit)) 
			strData =	strData & Chr(11) & ConvSPChars(pArr(i, I305_EG1_E1_location)) 
			strData =	strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_tot_stk_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) 
			strData =	strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_tot_stk_val),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0)
			   
			IF EG1_export_group(i, I305_EG1_E3_prc_ctrl_indctr) = "M" THEN
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_moving_avg_prc),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)
			 	strData = strData & Chr(11) & "이동"
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_prev_tot_stk_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) 
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_prev_tot_stk_val),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0)
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_prev_moving_avg_prc),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)
			     strData = strData & Chr(11) & "이동"
									
		    Elseif EG1_export_group(i, I305_EG1_E3_prc_ctrl_indctr) = "S" THEN
		       
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_std_prc),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)
			 	strData = strData & Chr(11) & "표준"
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_prev_tot_stk_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) 
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_prev_tot_stk_val),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0)
			 	strData = strData & Chr(11) & UniConvNumberDBToCompany(pArr(i, I305_EG1_E3_prev_std_prc),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)
			 	strData = strData & Chr(11) & "표준"
		      
		    End if
		 
			strData =  strData & Chr(11) & i &  Chr(11) & Chr(12) 
			j=j+1
		Next 
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
		Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr   
		Response.Write "	.ggoSpread.SSShowData     """ & strData & """" & ",""F""" & vbCr
		Response.Write "	.DbQueryOk " & vbCr 
		Response.Write "    .frm1.vspdData.Redraw = True " & vbCr 
		Response.Write "    .frm1.hPlantCd.value       = "" " & ConvSPChars(Trim(Request("txtPlantCd")))        & " "" " & vbcr
		Response.Write "    .frm1.hItemCd.value        = "" " & ConvSPChars(Trim(Request("txtItemCd")))         & " "" " & vbcr
		Response.Write "    .frm1.hAccntCd.value       = "" " & ConvSPChars(Trim(Request("txtAccntCd")))        & " "" " & vbcr
		Response.Write "    .frm1.hTrackingNo.value    = "" " & ConvSPChars(Trim(Request("txtTrackingNo")))     & " "" " & vbcr 
		if j < C_SHEETMAXROWS_D then 
			Response.Write "    .DbChkFlag =true " & vbCr  
		else
			Response.Write "    .DbChkFlag =false " & vbCr  
		end if
		
		Response.Write "    .lgStrPrevToKey = "&lgStrPrevToKey+1&" " & vbCr  
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr    
    
 End Sub   
 
 
    
%>

