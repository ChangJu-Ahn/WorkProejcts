<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Create physical inventory Posting in batch 
'*  3. Program ID           : I2151mb2.asp
'*  4. Program Name         : 실사조정Batch등록(취소)
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2006/10/18
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : LEE SEUNG WOOK
'* 10. Modifier (Last)      : 
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%									
Call LoadBasisGlobalInf()

On Error Resume Next
Err.Clear
Call HideStatusWnd

Dim pPI2G070												

Dim strMode

Dim I1_i_physical_inventory_header_phy_inv_no

Dim prErrorPosition
	Const Err_item_cd = 0
    Const Err_tracking_no = 1
    Const Err_lot_no = 2
    Const Err_lot_sub_no = 3
 
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_i_physical_inventory_header_phy_inv_no       = Request("txtPhyinvNo")
 
 
    Set pPI2G070 = Server.CreateObject("PI2G070.cIPostPhyInvBatch")    

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

    '-----------------------
    'Com action area
    '-----------------------
	Call pPI2G070.I_POST_PHY_INV_BATCH_DEL(gStrGlobalCollection, _
										   I1_i_physical_inventory_header_phy_inv_no, _
										   prErrorPosition)
    
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		If prErrorPosition(0) <> "" Then
			Call ServerMesgBox("상세정보" & vbcrlf & vbcrlf & vbcrlf & _
							   "품목 : " & prErrorPosition(Err_item_cd) & vbtab & vbtab & vbcrlf & vbcrlf & _
							   "Tracking No : " & prErrorPosition(Err_tracking_no) & vbtab & vbcrlf &vbcrlf & _
							   "LOT NO : " & prErrorPosition(Err_lot_no) & vbcrlf & vbtab & vbcrlf & _
							   "Lot No.순번 : " & prErrorPosition(Err_lot_sub_no), _
							   	vbCritical, I_MKSCRIPT)  
		End If

		Set pPI2G070 = Nothing														
		Response.End
	End If

    Set pPI2G070 = Nothing	

	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	'Response.Write "	.frm1.hItemDocumentNo.value = """ & ConvSPChars(E1_good_mvmt_workset(E1_item_document_no)) & """" & vbCr  	   	  
  	Response.Write "    .DbSaveOk2 "				& vbCr
	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End 
%>