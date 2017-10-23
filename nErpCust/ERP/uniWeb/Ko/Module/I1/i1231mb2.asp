<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Good Movement Header/detail
'*  3. Program ID           : I1131mb2.asp
'*  4. Program Name         : 기타입고수불등록
'*  5. Program Desc         : 기타입고수불정보/상세정보를 등록한다.
'*  7. Modified date(First) : 2002/05/30
'*  8. Modified date(Last)  : 2003/04/29
'*  9. Modifier (First)     : HAN SUNG GYU
'* 10. Modifier (Last)      : AHN JUNG JE
'* 11. Comment              : VB CONVERSION시 반영
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
    On Error Resume Next                                                      
    Err.Clear        
    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")                                 
    Call HideStatusWnd                                                            
        
    Dim arrRowVal					
	Dim arrColVal		
	Dim iPI0C161
    Dim LngRow
    Dim iMaxRow
    Dim istrStatus
    
    Dim EG1_export_group
    Dim iErrorPosition
    
    Dim E1_good_mvmt_workset 
		Const I127_E3_item_document_no = 0
		Const I127_E3_year				= 1
    
    Dim I1_good_mvmt_workset 
		Const I127_I1_item_document_no = 0
		Const I127_I1_trns_type		= 1
		Const I127_I1_mov_type		= 2
		Const I127_I1_document_dt	= 3
		Const I127_I1_pos_dt		= 4
		Const I127_I1_document_text = 5
		Const I127_I1_plant			= 6
		Const I127_I1_cost_cd		= 7
    ReDim I1_good_mvmt_workset(I127_I1_cost_cd) 
    	
    Dim IG1_import_group
		Const I127_IG1_lot_no			= 0
		Const I127_IG1_lot_sub_no		= 1
		Const I127_IG1_item_status		= 2
		Const I127_IG1_entry_qty		= 3
		Const I127_IG1_entry_unit		= 4
		Const I127_IG1_auto_crtd_flag	= 5
		Const I127_IG1_amount			= 6
		Const I127_IG1_wc_cd			= 7
		Const I127_IG1_tracking_no		= 8
		Const I127_IG1_prodt_order_no	= 9
		Const I127_IG1_cost_cd			= 11
		Const I127_IG1_plant_cd			= 13
		Const I127_IG1_item_cd			= 14
		Const I127_IG1_sl_cd			= 15
		Const I127_IG1_po_no			= 16
		Const I127_IG1_po_seq			= 17
		Const I127_IG1_req_no			= 18
    
    Dim DI1_good_mvmt_workset
		Const I128_I1_item_document_no	= 0
		Const I128_I1_year				= 1
    ReDim DI1_good_mvmt_workset(I128_I1_year) 

    Dim DIG1_import_group
		Const I128_IG1_seq_no			= 0
		Const I128_IG1_sub_seq_no		= 1

    ReDim iErrorPosition(0)
 	
 	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim i   

    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For i = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(i)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")

	
	I1_good_mvmt_workset(I127_I1_item_document_no) = Trim(Request("txtDocumentNo2"))
	I1_good_mvmt_workset(I127_I1_trns_type)        = "OI"
	I1_good_mvmt_workset(I127_I1_mov_type)         = Request("txtMovType")
	I1_good_mvmt_workset(I127_I1_document_dt)	    = UNIConvDate(Request("txtDocumentDt"))
	I1_good_mvmt_workset(I127_I1_pos_dt)           = UNIConvDate(Request("txtPostingDt"))
	I1_good_mvmt_workset(I127_I1_document_text)    = Request("txtDocumentText")
	I1_good_mvmt_workset(I127_I1_plant)            = Request("txtPlantCd")
	I1_good_mvmt_workset(I127_I1_cost_cd)          = Request("txtCostCd")
	
	DI1_good_mvmt_workset(I128_I1_item_document_no) = Trim(Request("txtDocumentNo1"))
	DI1_good_mvmt_workset(I128_I1_year)             = Request("hYear")
	

	If itxtSpread <> "" Then
	
		arrRowVal = Split(itxtSpread, gRowSep)
	    iMaxRow = UBound(arrRowVal) - 1
	    
	    Redim IG1_import_group(iMaxRow,I127_IG1_req_no)  
        Redim DIG1_import_group(iMaxRow,I128_IG1_sub_seq_no)  
    	
    	For LngRow = 0 To iMaxRow
		    
			arrColVal = Split(arrRowVal(LngRow), gColSep)
			istrStatus = arrColVal(0)												
	         
       	     Select Case istrStatus
				Case "C"					
					IG1_import_group(LngRow,I127_IG1_item_cd)        = arrColVal(2)
					
					If arrColVal(3) <> "" Then
					   IG1_import_group(LngRow,I127_IG1_tracking_no) = arrColVal(3)	
					Else
					   IG1_import_group(LngRow,I127_IG1_tracking_no) = "*"
					End If
					
					IG1_import_group(LngRow,I127_IG1_lot_no)         = Trim(arrColVal(4))
					
					If arrColVal(5) = "" Then
						IG1_import_group(LngRow,I127_IG1_lot_sub_no)	= 0
					Else
						IG1_import_group(LngRow,I127_IG1_lot_sub_no)	= CInt(arrColVal(5))
					End If					
				    
					IG1_import_group(LngRow,I127_IG1_entry_qty)       = UNIConvNum(arrColVal(6), 0)	
					IG1_import_group(LngRow,I127_IG1_amount)		  = UNIConvNum(arrColVal(7), 0)
					IG1_import_group(LngRow,I127_IG1_entry_unit)      = Trim(arrColVal(8))
				    IG1_import_group(LngRow,I127_IG1_plant_cd)	       = Request("txtPlantCd")
                    IG1_import_group(LngRow,I127_IG1_sl_cd)           = Request("txtSLCd")
                    IG1_import_group(LngRow,I127_IG1_prodt_order_no)  = Trim(arrColVal(9))
                    IG1_import_group(LngRow,I127_IG1_wc_cd)           = Request("txtWCCd")
                    IG1_import_group(LngRow,I127_IG1_cost_cd)		  = ""
                    IG1_import_group(LngRow,I127_IG1_req_no)		  = Trim(arrColVal(10))
                    
            	case "D"
					DIG1_import_group(LngRow,I128_IG1_seq_no)	         = CInt(arrColVal(2))
					DIG1_import_group(LngRow,I128_IG1_sub_seq_no)      = CInt(arrColVal(3))
			End Select
		Next
	End If

    Set iPI0C161 = Server.CreateObject("PI0C161.cIGoodIssueForOthers")
    
     If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
		Response.End
     End If

	Select Case istrStatus
		Case "C"			
   
             Call iPI0C161.I_CREATE_GOOD_ISSUE_FOR_OTHERS(gStrGlobalCollection, _
											             I1_good_mvmt_workset, _
											             IG1_import_group, _
											             EG1_export_group, _
											             E1_good_mvmt_workset, _
											             iErrorPosition) 
   
             If CheckSYSTEMError(Err,True) = True Then
                Set iPI0C161 = Nothing
                If iErrorPosition(0) <> 0 Then
		     		Response.Write "<Script Language=VBScript>" & vbCrLF
					Response.Write " Parent.RemovedivTextArea	"	& vbCr
			        Response.Write "Call parent.SheetFocus(" & iErrorPosition(0) & ", 1)" & vbCrLF
			        Response.Write "</Script>" & vbCrLF
		        End If
				Response.End
             End If
    
             Set iPI0C161 = Nothing
                    
        Case "D"			

             Call iPI0C161.I_DELETE_GOOD_ISSUE_FOR_OTHERS(gStrGlobalCollection, _
														DI1_good_mvmt_workset, _
														DIG1_import_group) 
   
             If CheckSYSTEMError(Err,True) = True Then
                Set iPI0C161 = Nothing
				Response.Write "<Script Language=vbscript> " & vbCr   
				Response.Write " Parent.RemovedivTextArea	"	& vbCr
				Response.Write "</Script>	"	& vbCr
				Response.End
             End If
             Set iPI0C161 = Nothing
    End select                      
    
	If istrStatus  = "C" Then	
       Response.Write "<Script Language=vbscript> " & vbCrlf
	   Response.Write "With parent.frm1 " & vbCrlf
       Response.Write ".txtDocumentNo1.Value 	= """ & ConvSPChars(E1_good_mvmt_workset(I127_E3_item_document_no)) & """" & vbCrlf
       Response.Write ".txtYear.Text 		    = """ & E1_good_mvmt_workset(I127_E3_year) & """" & vbCrlf
       Response.Write ".txtDocumentNo2.Value	= """ & ConvSPChars(E1_good_mvmt_workset(I127_E3_item_document_no)) & """" & vbCrlf
       Response.Write "End With" & vbCrlf
       Response.Write "</Script>" & vbCrlf
	End If 
		
	Response.Write "<Script Language=vbscript> " & vbCrlf
	Response.Write " Parent.RemovedivTextArea	"	& vbCr
    Response.Write "Parent.DbSaveOk " & vbCrlf
    Response.Write "</Script>" & vbCrlf

	Response.End
%>

