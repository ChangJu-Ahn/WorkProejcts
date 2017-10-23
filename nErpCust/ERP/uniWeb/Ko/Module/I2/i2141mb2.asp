<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : 실사선별posting 등록 asp
'*  2. Function Name        : 
'*  3. Program ID           : i2141mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : PI2G060 I_POST_PHY_INV

'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2003/05/07
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Mr  Ahn Jung Je
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%										
Call LoadBasisGlobalInf()

On Error Resume Next											
Err.Clear

Call HideStatusWnd

	Dim pPI2G060															
	
	Dim I1_i_physical_inventory_header_phy_inv_no
    Dim I2_b_cost_center_cost_cd
	Dim iErrorPosition
	
    Dim E1_good_mvmt_workset      
		Const I216_E1_item_document_no = 0   
		Const I216_E1_year = 1
	
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim i
    
    Dim arrRowVal					
	Dim arrColVal		
    Dim LngRow
    Dim iMaxRow
    Dim istrStatus
	
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For i = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(i)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")
    	
	I1_i_physical_inventory_header_phy_inv_no		= Request("txtCondPhyInvNo")
	I2_b_cost_center_cost_cd					    = UCase(Request("txtCostCd"))

	If itxtSpread <> "" Then
		arrRowVal = Split(itxtSpread, gRowSep)
		iMaxRow = UBound(arrRowVal) - 1
		
		For LngRow = 0 To iMaxRow
			arrColVal = Split(arrRowVal(LngRow), gColSep)
			istrStatus = arrColVal(0)
		Next
	End If

	Set pPI2G060 = Server.CreateObject("PI2G060.cIPostPhyInv")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
		Response.End
	End If
	
	Select Case istrStatus	
		Case "U"	
			Call pPI2G060.I_POST_PHY_INV(gStrGlobalCollection, _
										I1_i_physical_inventory_header_phy_inv_no, _
										I2_b_cost_center_cost_cd, _
										itxtSpread, _
										E1_good_mvmt_workset, _
										iErrorPosition)
									
			If CheckSYSTEMError(Err,True) = True Then
				Set pPI2G060 = Nothing													
				Response.Write "<Script Language=vbscript> " & vbCr  
				Response.Write " Call Parent.RemovedivTextArea	"	& vbCr
				If iErrorPosition(0) <> "" Then
					Response.Write " Call Parent.SheetFocus(" & iErrorPosition(0) & ", 2)  " & vbCr  
				End If
				Response.Write "</Script>	"	& vbCr
				Response.End
			End If
					
			Set pPI2G060 = Nothing
		Case "D"
			Call pPI2G060.I_POST_PHY_INV_DEL(gStrGlobalCollection, _
										I1_i_physical_inventory_header_phy_inv_no, _
										itxtSpread, _
										iErrorPosition)
									
			If CheckSYSTEMError(Err,True) = True Then
				Set pPI2G060 = Nothing													
				Response.Write "<Script Language=vbscript> " & vbCr  
				Response.Write " Call Parent.RemovedivTextArea	"	& vbCr
				
				If iErrorPosition(0) <> "" Then
					Response.Write " Call Parent.SheetFocus(" & iErrorPosition(0) & ", 2)  " & vbCr  
				End If
				
				Response.Write "</Script>	"	& vbCr
				Response.End
			End If
					
			Set pPI2G060 = Nothing
	End Select												
				
	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	Response.Write "	.frm1.txthItemDocumentNo.value = """ & ConvSPChars(E1_good_mvmt_workset(I216_E1_item_document_no)) & """" & vbCr  	   	  
  	Response.Write "    .RemovedivTextArea	"	& vbCr
  	Response.Write "    .DbSaveOk "				& vbCr
	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End 			


%>
