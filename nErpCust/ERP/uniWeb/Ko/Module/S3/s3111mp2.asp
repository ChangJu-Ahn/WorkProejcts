<%

'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3111mb2.asp	
'*  4. Program Name         : 단가확정 
'*  5. Program Desc         : 단가확정 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------

Call SubBizQuery()        

'============================================================================================================
Sub SubBizQuery()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	Dim tmpRow
	
    Dim iObjPS3G142       
    
    '-----------------------------------------------
    ' Declare User Variable
    '-----------------------------------------------
    '  단가확정여부/단가적용규칙/단가적용기준일   
    Dim i3_ief_supplied1    
    Dim   i3_price_flag1 
    Dim   i3_base_date1
  
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	' 단가확정여부	
    '-----------------------------------------------
    i3_ief_supplied1 = Trim(Request("txtPostFlag"))
	'-----------------------------------------------
    ' 단가적용규칙여부	
    '-----------------------------------------------
    i3_price_flag1 = Trim(Request("txtPriceFlag"))
    '-----------------------------------------------
    ' 단가규칙적용기준일 
    '-----------------------------------------------
    i3_base_date1 = UNIConvDate(Request("txtBaseDate"))    
      
	set iObjPS3G142 = CREATEOBJECT("PS3G142.cCsLcHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
                  
	Call iObjPS3G142.S_LIST_FIXING_PRICE_POPUP (gStrGlobalCollection, _
		                                   Trim(Request("txtSpread")),exp_grp, I3_price_flag1, I3_base_date1)    
								      

	If CheckSYSTEMError(Err,TRUE) = True Then
        Response.Write "<Script language=vbs>  " & vbCr   
        Response.Write "</Script>      " & vbCr
       Set iObjPS3G142 = Nothing		         
       Exit Sub
    End If   
    
    Set iObjPS3G142 = Nothing	       
    
    Dim tmpSpreadData,tmpSpreadData2
    Dim tmpRow1, tmpRow2

        
    For iLngRow = 0 To UBound(exp_grp,1)-1    
		tmpRow1=split(exp_grp(iLngRow, c_exp_s_item_sales_price_item_price),":::")
		'진단가 
		'istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_s_item_sales_price_item_price), 0)
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(tmpRow1(0), 0)     
		tmpRow2=tmpRow2 & chr(11) & tmpRow1(1)             
    Next
	tmpSpreadData=split(istrdata,chr(11))
	tmpSpreadData2=split(tmpRow2,chr(11))
	
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " Parent.frm1.txtHBaseDate.value   = """ & Trim(Request("txtBaseDate")) & """" & vbCr            
    
    for iLngRow =1 to ubound(tmpSpreadData)
		
	    Response.Write "	Parent.ggoSpread.Source          = Parent.frm1.vspdData									      " & vbCr
	    Response.Write "	Parent.frm1.vspdData.Row= " & tmpSpreadData2(iLngRow)  & vbcr			
	    
	    Response.Write "	Parent.frm1.vspdData.Col = Parent.C_PriceFlagN " &	vbcr
	    Response.Write "	Parent.frm1.vspdData.text = """ & tmpSpreadData(iLngRow) & """" & vbcr 
	    Response.write "	Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.C_Currency,Parent.C_PriceFlagN,	""C"" ,""I"",""X"",""X"")" & vbCr
	    

   next
    Response.Write "</Script> "																							& vbCr      
    
    Exit sub 
                   	
End Sub    
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

