<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 출고등록 
'*  3. Program ID           : I1523MB2
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Choi Sung Jae
'* 10. Modifier (Last)      : Ahn JUng Je
'* 11. Comment              : VB Conversion
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
    
	Dim iPI5G140
	Dim itxtFlgMode

    Dim I1_I_VMI_Goods_Mvmt_Hdr
		Const I507_I1_item_document_no = 0
		Const I507_I1_document_year = 1
		Const I507_I1_plant_cd = 2
		Const I507_I1_trns_type = 3
		Const I507_I1_document_dt = 4
		Const I507_I1_sl_cd = 5
		Const I507_I1_bp_cd = 6
		Const I507_I1_document_text = 7

    Dim E1_Good_Mvmt_Worket
		Const I507_E1_item_document_no = 0
		Const I507_E1_document_year = 1    

	Const I508_I1_item_document_no = 0
	Const I508_I1_document_year = 1

	Dim txtSpread
	Dim iErrorPosition

	itxtFlgMode = Request("txtpvCommandMode")
	If itxtFlgMode = "C" Then		
		ReDim I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_document_text)
		
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_item_document_no) = Request("txtItemDocumentNo2")
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_document_year)    = Request("hDocumentYear")
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_plant_cd)         = Request("txtPlantCd")
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_document_dt)      = UNIConvDate(Request("txtDocumentDt"))
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_sl_cd)            = Request("txtSlCd")
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_bp_cd)            = Request("txtBpCd")
		I1_I_VMI_Goods_Mvmt_Hdr(I507_I1_document_text)    = Request("txtDocumentText")
		txtSpread = Request("txtSpread")

	ElseIf itxtFlgMode = "U" Then
		'해당사항없음.
		
	ElseIf itxtFlgMode = "D" Then	
		ReDim I1_I_VMI_Goods_Mvmt_Hdr(I508_I1_document_year)

		I1_I_VMI_Goods_Mvmt_Hdr(I508_I1_item_document_no) = Request("txtItemDocumentNo2")
		I1_I_VMI_Goods_Mvmt_Hdr(I508_I1_document_year)    = Request("txtDocumentYear")
		txtSpread = Request("txtSpread")
	End If

    Set iPI5G140 = Server.CreateObject("PI5G140.cIVMIGoodsIssue")
    
    If CheckSYSTEMError(Err,True) = True Then
    	Response.End	
    End If

	If itxtFlgMode = "C" Then		
		Call iPI5G140.I_VMI_CREATE_GOODS_ISSUE(gStrGlobalCollection, _
											I1_I_VMI_Goods_Mvmt_Hdr, _
											txtSpread, _
											E1_Good_Mvmt_Worket, _
											iErrorPosition)  

	ElseIf itxtFlgMode = "U" Then	
		'해당사항없음.
		
	ElseIf itxtFlgMode = "D" Then	
		Call iPI5G140.I_VMI_DELETE_GOODS_ISSUE(gStrGlobalCollection, _
											I1_I_VMI_Goods_Mvmt_Hdr, _
											txtSpread, _
											iErrorPosition)		                   
   	End If

    If CheckSYSTEMError(Err,True) = True Then
		Set iPI5G140 = Nothing
		If iErrorPosition <> 0 Then
			Call SheetFocus(iErrorPosition, 1)
		End If
    	Response.End	
    End If
   
    Set iPI5G140 = Nothing

	'-------------------------------------------------------------------------------------------------------------
	If itxtFlgMode = "C" Then		
		Response.Write " <Script Language=vbscript> " & vbCrlf
		Response.Write " With parent.frm1 " & vbCrlf
		Response.Write "     .txtItemDocumentNo.Value  = """ & ConvSPChars(E1_Good_Mvmt_Worket(I507_E1_item_document_no)) & """" & vbCrlf
		Response.Write "     .txtDocumentYear.Text     = """ & E1_Good_Mvmt_Worket(I507_E1_document_year) & """" & vbCrlf
		Response.Write " End With " & vbCrlf
		Response.Write " Parent.DbSaveOk " & vbCrlf
		Response.Write " </Script>" & vbCrlf
	ElseIf itxtFlgMode = "U" Then	
		'해당사항없음.
		
	ElseIf itxtFlgMode = "D" Then	
		Response.Write " <Script Language=vbscript> " & vbCrlf
		Response.Write " Parent.DbSaveOk " & vbCrlf
		Response.Write " </Script>" & vbCrlf
   	End If

   	Response.End	
    

Sub SheetFocus(ByVal lRow, ByVal lCol)
	Response.Write " <Script Language=VBScript> "                    & vbCrLF
	Response.Write " With parent.frm1 "                              & vbCrlf
	Response.Write "	.vspdData.focus "                           & vbCrlf
	Response.Write "	.vspdData.Row = " & lRow                    & vbCrlf
	Response.Write "	.vspdData.Col = " & lCol                    & vbCrlf
	Response.Write "	.vspdData.Action = 0 "                      & vbCrlf
	Response.Write "	.vspdData.SelStart = 0 "                    & vbCrlf
	Response.Write "	.vspdData.SelLength = len(.vspdData.Text) " & vbCrlf
	Response.Write " End With"                                       & vbCrlf
	Response.Write " </Script>"                                      & vbCrLF
End Sub


%>


