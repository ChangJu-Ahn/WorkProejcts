<%@ LANGUAGE=VBSCript%>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1111MB2
'*  4. Program Name         : 품목단가등록 
'*  5. Program Desc         : 품목단가등록 
'*  6. Comproxy List        : PS1G101.dll, PS1G102.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2005/05/03
'*  9. Modifier (First)     : sonbumyeol
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/20 : Grid성능 적용, Kang Jun Gu
'*				            : 2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->

<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	Call LoadBasisGlobalInf()

    Call HideStatusWnd                                                     '☜: Hide Processing message

    '---------------------------------------Common-----------------------------------------------------------
 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
Sub SubBizQueryMulti()

 On Error Resume Next                                                             '☜: Protect system from crashing
 Err.Clear                                                                        '☜: Clear Error status
 
 Dim PS1G102
 Dim StrNextKey       ' 다음 값 
 Dim iLngMaxRow       ' 현재 그리드의 최대Row
 Dim iLngRow
 Dim intGroupCount               '☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
 Dim StrNext
 Dim StrSeq
 Dim arrValue
 Dim lgStrPrevKey
 Dim istrData

 Dim imp_s_item_sales_price
 Dim imp_b_item
 Dim imp_next_b_item
 Dim imp_next_s_item_sales_price
 
 Redim imp_s_item_sales_price(4)
 Redim imp_next_s_item_sales_price(4)
 
 Dim exp_b_item
 Dim exp_b_unit_of_measure
 Dim exp_deal_type_b_minor
 Dim exp_pay_meth_b_minor
 Dim exp_b_currency
 Dim exp_next_b_item
 Dim exp_next_s_item_sales_price

 Dim exp_grp 
 Dim prGroupView
 
 Const imp_valid_from_dt = 0
 Const imp_deal_type = 1
 Const imp_pay_meth = 2
 Const imp_sales_unit = 3
 Const imp_currency = 4
 
 Const item_cd = 0
 
 Const next_item_cd = 0
 
 Const next_deal_type = 0
 Const next_pay_meth = 1
 Const next_from_dt = 2
 Const next_sales_unit = 3
 Const next_currency = 4
 
 Const exp_item_cd = 0
 Const exp_item_nm = 1
 Const exp_deal_type = 2
 Const exp_pay_meth = 3
 Const exp_sales_unit = 4
 Const exp_currency = 5
 Const exp_valid_from_dt = 6
 Const exp_item_price = 7
 Const exp_ext1_qty = 8
 Const exp_ext2_qty = 9
 Const exp_ext1_amt = 10
 Const exp_ext2_amt = 11
 Const exp_ext1_cd = 12
 Const exp_ext2_cd = 13
 Const exp_deal_type__nm = 14
 Const exp_pay_meth_nm = 15
 Const exp_spec = 16
 Const exp_price_Flag =17
 Const exp_remrk =18
 
 Const C_SHEETMAXROWS_D  = 100
 '-----------------------
    ' 수주헤더를 읽어온다.
    '-----------------------
 
 imp_s_item_sales_price(imp_valid_from_dt) = UNIConvDate(Request("txtconValid_from_dt"))
' imp_s_item_sales_price(imp_deal_type) = FilterVar(Trim(Request("txtconDeal_type")), "" ,  "SNM")
 imp_s_item_sales_price(imp_deal_type) = Trim(Request("txtconDeal_type")) 
 imp_s_item_sales_price(imp_pay_meth) = Trim(Request("txtconPay_terms")) 
 imp_s_item_sales_price(imp_sales_unit) = Trim(Request("txtconSales_unit"))
 imp_s_item_sales_price(imp_currency) = Trim(Request("txtconCurrency"))
 
 imp_b_item = FilterVar(Trim(Request("txtconItem_cd")), "" ,  "SNM") 
  
 lgStrPrevKey = Trim(Request("lgStrPrevKey"))
 
 If lgStrPrevKey <> "" then

  arrValue = Split(lgStrPrevKey, gColSep)

  imp_next_b_item = FilterVar(Trim(arrValue(0)), "" ,  "SNM") 
    
  imp_next_s_item_sales_price(next_deal_type) = Trim(arrValue(1))
  imp_next_s_item_sales_price(next_pay_meth) = Trim(arrValue(2))
  If Len(Trim(arrValue(3))) Then imp_next_s_item_sales_price(next_from_dt) = Trim(arrValue(3)) 
  imp_next_s_item_sales_price(next_sales_unit) = Trim(arrValue(4)) 
  imp_next_s_item_sales_price(next_currency) = Trim(arrValue(5))
  
 Else
  imp_next_b_item = ""
  
  imp_next_s_item_sales_price(next_deal_type) = ""
  imp_next_s_item_sales_price(next_pay_meth) = ""
  imp_next_s_item_sales_price(next_from_dt) = ""
  imp_next_s_item_sales_price(next_sales_unit) = ""
  imp_next_s_item_sales_price(next_currency) = ""
 End If 
     
    Set PS1G102 = Server.CreateObject("PS1G102.cSListItemPriceSvr")
 
 Call PS1G102.S_LIST_ITEM_PRICE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
 imp_next_b_item,imp_next_s_item_sales_price,imp_s_item_sales_price, _
 imp_b_item, exp_b_item, exp_b_unit_of_measure, exp_deal_type_b_minor, _
 exp_pay_meth_b_minor,exp_b_currency,exp_next_b_item,exp_next_s_item_sales_price, _
 exp_grp ,prGroupView)
  
 If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "122600" then    
 '품목 
  If CheckSYSTEMError(Err,True) = True Then
      prGroupView = -1
   Set PS1G102 = Nothing
   
   Response.Write "<Script language=vbs>  " & vbCr   
   Response.Write " With Parent        " & vbCr
   Response.Write " .frm1.txtconItem_cd.focus" & vbCr
   Response.Write " .frm1.txtconItem_nm.value  = """ & "" & """" & vbCr    
   Response.Write " .frm1.txtconDeal_type_nm.value = """ & ConvSPChars(exp_deal_type_b_minor(1)) & """" & vbCr    
   Response.Write " .frm1.txtconPay_terms_nm.value = """ & ConvSPChars(exp_pay_meth_b_minor(1)) & """" & vbCr    
   Response.Write "End With       " & vbCr                    
   Response.Write "</Script>      " & vbCr      
 
   Response.End
  End If   
 
 Elseif cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "200071" then
 '판매유형 
  If CheckSYSTEMError(Err,True) = True Then
   prGroupView = -1
   Set PS1G102 = Nothing
   
   Response.Write "<Script language=vbs>  " & vbCr   
   Response.Write " With Parent        " & vbCr
   Response.Write " .frm1.txtconDeal_type.focus" & vbCr
   Response.Write " .frm1.txtconDeal_type_nm.value  = """ & "" & """" & vbCr    
   Response.Write " .frm1.txtconItem_nm.value  = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
   Response.Write " .frm1.txtconPay_terms_nm.value = """ & ConvSPChars(exp_pay_meth_b_minor(1)) & """" & vbCr       
   Response.Write "End With       " & vbCr                    
   Response.Write "</Script>      " & vbCr       
  
   Response.End
  End If   
 
 Elseif cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "200054" then
 '결재방법 
  If CheckSYSTEMError(Err,True) = True Then
   prGroupView = -1
   Set PS1G102 = Nothing
   
   Response.Write "<Script language=vbs>  " & vbCr   
   Response.Write " With Parent        " & vbCr
   Response.Write " .frm1.txtconPay_terms.focus" & vbCr
   Response.Write " .frm1.txtconPay_terms_nm.value  = """ & "" & """" & vbCr    
   Response.Write " .frm1.txtconItem_nm.value  = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
   Response.Write " .frm1.txtconDeal_type_nm.value = """ & ConvSPChars(exp_deal_type_b_minor(1)) & """" & vbCr    
   Response.Write "End With       " & vbCr                    
   Response.Write "</Script>      " & vbCr        
  
   Response.End
  End If   
 
 Elseif cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "124000" then
 '단위 
  If CheckSYSTEMError(Err,True) = True Then
   prGroupView = -1
   Set PS1G102 = Nothing
   
   Response.Write "<Script language=vbs>  " & vbCr   
   Response.Write " With Parent        " & vbCr
   Response.Write " .frm1.txtconSales_unit.focus" & vbCr
   Response.Write " .frm1.txtconDeal_type_nm.value  = """ & "" & """" & vbCr    
   Response.Write " .frm1.txtconItem_nm.value  = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
   Response.Write " .frm1.txtconPay_terms_nm.value = """ & ConvSPChars(exp_pay_meth_b_minor(1)) & """" & vbCr       
   Response.Write "End With       " & vbCr                    
   Response.Write "</Script>      " & vbCr       
  
   Response.End
  End If   
 
 Elseif cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "121400" then
 '화폐 
  If CheckSYSTEMError(Err,True) = True Then
   prGroupView = -1
   Set PS1G102 = Nothing
   
   Response.Write "<Script language=vbs>  " & vbCr   
   Response.Write " With Parent        " & vbCr
   Response.Write " .frm1.txtconCurrency.focus" & vbCr
   Response.Write " .frm1.txtconPay_terms_nm.value  = """ & "" & """" & vbCr    
   Response.Write " .frm1.txtconItem_nm.value  = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
   Response.Write " .frm1.txtconDeal_type_nm.value = """ & ConvSPChars(exp_deal_type_b_minor(1)) & """" & vbCr    
   Response.Write "End With       " & vbCr                    
   Response.Write "</Script>      " & vbCr        
  
   Response.End
  End If   
 
 Else
  
  If CheckSYSTEMError(Err,True) = True Then
   prGroupView = -1
   Set PS1G102 = Nothing

   Response.Write "<Script language=vbs>  " & vbCr   
   Response.Write " With Parent        " & vbCr
   Response.Write " .frm1.txtconItem_cd.focus" & vbCr
   Response.Write " .frm1.txtconItem_nm.value  = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
   Response.Write " .frm1.txtconDeal_type_nm.value = """ & ConvSPChars(exp_deal_type_b_minor(1)) & """" & vbCr    
   Response.Write " .frm1.txtconPay_terms_nm.value = """ & ConvSPChars(exp_pay_meth_b_minor(1)) & """" & vbCr    
   Response.Write "End With       " & vbCr                    
   Response.Write "</Script>      " & vbCr      
  
   Response.End
  End If   
 
 End if 
 
    
 intGroupCount = prGroupView
    
 StrNext = exp_next_b_item(0)
 StrNext = StrNext & gColSep & exp_next_s_item_sales_price(0)
 StrNext = StrNext & gColSep & exp_next_s_item_sales_price(1)
 StrNext = StrNext & gColSep & exp_next_s_item_sales_price(2)
 StrNext = StrNext & gColSep & exp_next_s_item_sales_price(3)
 StrNext = StrNext & gColSep & exp_next_s_item_sales_price(4)
  
 StrSeq = exp_grp(intGroupCount,0)
 StrSeq = StrSeq & gColSep & exp_grp(intGroupCount,2)
 StrSeq = StrSeq & gColSep & exp_grp(intGroupCount,3)
 StrSeq = StrSeq & gColSep & exp_grp(intGroupCount,6)
 StrSeq = StrSeq & gColSep & exp_grp(intGroupCount,4)
 StrSeq = StrSeq & gColSep & exp_grp(intGroupCount,5)
    


  
    iLngMaxRow  = CLng(Request("txtMaxRows"))           '☜: Fetechd Count      
    
	For iLngRow = 0 To intGroupCount

		If  iLngRow < C_SHEETMAXROWS_D  Then
     	Else
	       StrNextKey = StrNext 		   		   
		   Exit For
        End If  

		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_item_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_item_nm))
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_spec))
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_deal_type))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_deal_type__nm))
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_pay_meth))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_pay_meth_nm))
        istrData = istrData & Chr(11) & UNIDateClientFormat(exp_grp(iLngRow,exp_valid_from_dt))
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_sales_unit))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_currency))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) &		UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow,exp_item_price), 0)
        '''''''''''''''''''''''''''''''
        istrData = istrData & Chr(11) &		ConvSPChars(exp_grp(iLngRow,exp_price_Flag))
        If ConvSPChars(exp_grp(iLngRow,exp_price_Flag))="T" then
			istrData = istrData & Chr(11) &		ConvSPChars("진단가")
        Else
			istrData = istrData & Chr(11) &		ConvSPChars("가단가")
        End If
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,exp_remrk))
                
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)     
    Next    
    
'    Response.Write istrData
 '   Response.Write chr(11) & chr(13)
  '  
   ' Response.End
    
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent        " & vbCr

    Response.Write "   .frm1.txtconItem_nm.value  = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
    Response.Write "   .frm1.txtconDeal_type_nm.value = """ & ConvSPChars(exp_deal_type_b_minor(1)) & """" & vbCr    
    Response.Write "   .frm1.txtconPay_terms_nm.value = """ & ConvSPChars(exp_pay_meth_b_minor(1)) & """" & vbCr
    
    Response.Write "   .frm1.txtHconItem_cd.value  = """ & ConvSPChars(Request("txtconItem_cd")) & """" & vbCr   
    Response.Write "   .frm1.txtHconDeal_type.value  = """ & ConvSPChars(Request("txtconDeal_type")) & """" & vbCr   
    Response.Write "   .frm1.txtHconPay_terms.value  = """ & ConvSPChars(Request("txtconPay_terms")) & """" & vbCr   
    Response.Write "   .frm1.txtHconValid_from_dt.value = """ & ConvSPChars(Request("txtconValid_from_dt")) & """" & vbCr   
    Response.Write "   .frm1.txtHconSales_unit.value = """ & ConvSPChars(Request("txtconSales_unit")) & """" & vbCr   
    Response.Write "   .frm1.txtHconCurrency.value  = """ & ConvSPChars(Request("txtconCurrency")) & """" & vbCr   
    
    Response.Write "   .ggoSpread.Source          = .frm1.vspdData          " & vbCr

    Response.Write "   .frm1.vspdData.Redraw = False   "                     & vbCr      
    Response.Write "   .ggoSpread.SSShowDataByClip   """ & istrData & """ ,""F""" & vbCr
    Response.Write "   Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1  & ",.C_Cur,.C_Item_Price,""C"" ,""I"",""X"",""X"")" & vbCr
    
    Response.Write "   .SetSpreadColor1 -1                                " & vbCr
    
    Response.Write "   .lgStrPrevKey              = """ & StrNextKey    & """" & vbCr  
    Response.Write "   .DbQueryOk()  " & vbCr   

    Response.Write  "    .frm1.vspdData.Redraw = True   "                      & vbCr      

    Response.Write "End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr     
     
            
End Sub    

'============================================================================================================
Sub SubBizSaveMulti()   
 
 On Error Resume Next                                                                 '☜: Protect system from crashing
 Err.Clear   
                                                                      
 Dim PS1G101 
 Dim iErrorPosition
  
 
    Dim reqtxtSpread
    reqtxtSpread = Request("txtSpread")

	Set PS1G101 = Server.CreateObject("PS1G101.cSItemPrcMultiSvr")  
    Call PS1G101.S_MAINT_ITEM_PRC_MULTI_SVR(gStrGlobalCollection, reqtxtSpread,iErrorPosition)
                  
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Set PS1G101 = Nothing
		Exit Sub
	End If 
  
    Set PS1G101 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk() "    & vbCr   
    Response.Write "</Script> "             & vbCr   
              
End Sub    

%>

