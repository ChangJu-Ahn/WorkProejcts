<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        :  
'*  3. Program ID           : S1611MB1
'*  4. Program Name         : 적립금정보등록 
'*  5. Program Desc         : 적립금정보등록 
'*  6. Comproxy List        : PS1G119.dll, PS1G120.dll
'*  7. Modified date(First) : 2002/05/28
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/15 : UI성능 적용 
'*                            2002/11/23 : Grid성능 적용, Kang Jun Gu
'*                            2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call LoadBasisGlobalInf()

Dim lgOpModeCRUD

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
	
lgOpModeCRUD  = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
    Dim LngRow	
	Dim LngMaxRow
	
	Dim lgstrData
	
	
	Dim StrPrevKey
	
    Dim iPS1G120    
    
    Dim intGroupCount
    
    Dim StrNext
    Dim StrNextKey  	
    Dim arrValue
    Const C_SHEETMAXROWS_D  = 100
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	LngMaxRow       = CLng(Request("txtMaxRows"))                                  '☜: Fetechd Count
	StrPrevKey      =Trim(Request("lgStrPrevKey"))                                 '☜: Next Key
    '-----------------------
    ' 적립금정보를 읽어온다.
    '-----------------------
   
    'View Name : imp s_reserve_price
    Const C_imp_s_reserve_price_bp_cd = 0
    Const C_imp_s_reserve_price_item_cd = 1
    Const C_imp_s_reserve_price_unit = 2
    Const C_imp_s_reserve_price_cur = 3
    Const C_imp_s_reserve_price_valid_from_dt = 4

   
    Const C_exp_grp_exp_item_b_biz_partner_bp_cd = 0        '거래처 
    Const C_exp_grp_exp_item_b_biz_partner_bp_nm = 1        '거래처명 
    Const C_exp_grp_exp_item_b_item_item_cd = 2             '품목 
    Const C_exp_grp_exp_item_b_item_item_nm = 3             '품목명 
    Const C_exp_grp_exp_item_s_reserve_price_unit = 4       '단위 
    Const C_exp_grp_exp_item_s_reserve_price_cur = 5        '화폐 
    Const C_exp_grp_exp_item_s_reserve_price_valid_from_dt = 6  '적용일 
    Const C_exp_grp_exp_item_s_reserve_price_reserve_price = 7  '적립단가 
    Const C_exp_grp_exp_item_s_reserve_price_ext1_qty = 8
    Const C_exp_grp_exp_item_s_reserve_price_ext2_qty = 9
    Const C_exp_grp_exp_item_s_reserve_price_ext1_amt = 10
    Const C_exp_grp_exp_item_s_reserve_price_ext2_amt = 11
    Const C_exp_grp_exp_item_s_reserve_price_ext1_cd = 12
    Const C_exp_grp_exp_item_s_reserve_price_ext2_cd = 13
    Const C_exp_grp_exp_item_b_item_spec = 14
    
     
    'View Name : exp b_biz_partner
    Const C_exp_b_biz_partner_bp_cd = 0
    Const C_exp_b_biz_partner_bp_nm = 1

    'View Name : exp b_currency
    Const C_exp_b_currency_currency = 0
    Const C_exp_b_currency_currency_desc = 1

    'View Name : exp b_item
    Const C_exp_b_item_item_cd = 0
    Const C_exp_b_item_item_nm = 1

    'View Name : exp b_unit_of_measure
    Const C_exp_b_unit_of_measure_unit = 0
    Const C_exp_b_unit_of_measure_unit_nm = 1
    
    'View Name : imp_next s_reserve_price
    Const C_imp_next_s_reserve_price_bp_cd = 0
    Const C_imp_next_s_reserve_price_item_cd = 1
    Const C_imp_next_s_reserve_price_cur = 2
    Const C_imp_next_s_reserve_price_unit = 3
    Const C_imp_next_s_reserve_price_valid_from_dt = 4
    

    Dim imp_s_reserve_price
    Dim exp_b_unit_of_measure
    Dim exp_b_currency
    Dim exp_grp
    Dim exp_b_item
    Dim exp_b_biz_partner
    Dim exp_next_s_reserve_price
    Dim imp_next_s_reserve_price
    Dim lgIntformatCount
    
    reDim imp_next_s_reserve_price(4)
    
    reDim imp_s_reserve_price(3)
    
	imp_s_reserve_price(C_imp_s_reserve_price_bp_cd) = Trim(Request("txtconBiz_partner"))
	imp_s_reserve_price(C_imp_s_reserve_price_item_cd) = Trim(Request("txtconItem_cd"))
	imp_s_reserve_price(C_imp_s_reserve_price_unit) = Trim(Request("txtconSales_unit"))
	imp_s_reserve_price(C_imp_s_reserve_price_cur) =Trim(Request("txtconCurrency"))

    
	If  StrPrevKey <> "" Then
		arrValue = Split(StrPrevKey, gColSep)
		           
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_bp_cd) = Trim(arrValue(C_imp_next_s_reserve_price_bp_cd))
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_item_cd) = Trim(arrValue(C_imp_next_s_reserve_price_item_cd))
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_cur) = Trim(arrValue(C_imp_next_s_reserve_price_cur))
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_unit) = Trim(arrValue(C_imp_next_s_reserve_price_unit))
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_valid_from_dt) = Trim(arrValue(C_imp_next_s_reserve_price_valid_from_dt))
	 
	Else 
	    imp_next_s_reserve_price(C_imp_next_s_reserve_price_bp_cd) = ""
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_item_cd) = ""
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_cur) = ""
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_unit) = ""
		imp_next_s_reserve_price(C_imp_next_s_reserve_price_valid_from_dt) = ""
		
    End If
   
    Set iPS1G120 = Server.CreateObject("PS1G120.cListReservePrcSvr")
   
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If 
  
    Call iPS1G120.SListReservePrcSvr(gStrGlobalCollection,C_SHEETMAXROWS_D, imp_next_s_reserve_price,imp_s_reserve_price, _
                                   exp_b_unit_of_measure,exp_b_currency,exp_grp,exp_b_item,exp_b_biz_partner, _
                                   exp_next_s_reserve_price)
    
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " Parent.frm1.txtconBiz_partner_nm.value   = """ & ConvSPChars(exp_b_biz_partner(C_exp_b_biz_partner_bp_nm)) & """" & vbCr
    Response.Write " Parent.frm1.txtconItem_nm.value   = """ & ConvSPChars(exp_b_item(C_exp_b_item_item_nm))                    & """" & vbCr        
    Response.Write " Parent.frm1.txtconSales_unit_nm.value   = """ & ConvSPChars(exp_b_unit_of_measure(C_exp_b_unit_of_measure_unit_nm))   & """" & vbCr  
    Response.Write "</Script> "																							& vbCr      
    
   	If CheckSYSTEMError(Err,True) = True Then
       Set iPS1G120 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   

     Set iPS1G120 = Nothing	
		   
	StrNext = exp_next_s_reserve_price(0)
	StrNext = StrNext & gColSep & exp_next_s_reserve_price(1)
	StrNext = StrNext & gColSep & exp_next_s_reserve_price(2)
	StrNext = StrNext & gColSep & exp_next_s_reserve_price(3)
    StrNext = StrNext & gColSep & exp_next_s_reserve_price(4)


	For LngRow = 0 To Ubound(exp_grp,1)
	
	
		If  LngRow < C_SHEETMAXROWS_D  Then
     	Else
	       StrNextKey = StrNext 		   		   
		   Exit For
        End If  
        
		lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_b_biz_partner_bp_cd))
		lgstrData = lgstrData & Chr(11) & ""
		lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_b_biz_partner_bp_nm))
        lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_b_item_item_cd))
        lgstrData = lgstrData & Chr(11) & ""
		lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_b_item_item_nm))
		lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_b_item_spec))
        lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_s_reserve_price_unit))
        lgstrData = lgstrData & Chr(11) & ""
        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(exp_grp(LngRow,C_exp_grp_exp_item_s_reserve_price_valid_from_dt))
		lgstrData = lgstrData & Chr(11) & ConvSPChars(exp_grp(LngRow,C_exp_grp_exp_item_s_reserve_price_cur))
		lgstrData = lgstrData & Chr(11) & ""
        lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(LngRow,C_exp_grp_exp_item_s_reserve_price_reserve_price), 0)
        lgstrData = lgstrData & Chr(11) & LngMaxRow + LngRow 
        lgstrData = lgstrData & Chr(11) & Chr(12)  
    Next
 
    Response.Write "<Script language=vbs> " & vbCr   
    
	Response.Write " Parent.frm1.txtHconBiz_partner.value  = """ & Request("txtconBiz_partner")     			 & """" & vbCr  
    Response.Write " Parent.frm1.txtHconItem_cd.value  = """ & Request("txtconItem_cd")					    	 & """" & vbCr
    Response.Write " Parent.frm1.txtHconSales_unit.value  = """ & Request("txtconSales_unit")					 & """" & vbCr
    Response.Write " Parent.frm1.txtHconValid_from_dt.value  = """ & Request("txtconValid_from_dt")	     		 & """" & vbCr
    Response.Write " Parent.frm1.txtHconCurrency.value  = """ & Request("txtconCurrency")			    		 & """" & vbCr
     
    Response.Write " Parent.ggoSpread.Source     = Parent.frm1.vspdData	"		 & vbCr
    Response.Write " Parent.frm1.vspdData.Redraw = False   "                     & vbCr      
    Response.Write " Parent.ggoSpread.SSShowDataByClip   """ & lgstrData & """ ,""F""" & vbCr

	If Ubound(exp_grp,1) < C_SHEETMAXROWS_D Then
		lgIntformatCount = Ubound(exp_grp,1) + 1 
	Else
		lgIntformatCount = Ubound(exp_grp,1)
	End If	

    Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & LngMaxRow + 1 & "," & LngMaxRow + lgIntformatCount & ",Parent.C_Cur,Parent.C_Item_Price,""C"" ,""I"",""X"",""X"")" & vbCr

'	Response.End 
    
    Response.Write " Parent.SetSpreadColor1 -1																     	  " & vbCr
    
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey										 & """" & vbCr  
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write " Parent.frm1.vspdData.Redraw = True   "                      & vbCr      
    Response.Write "</Script> "																							& vbCr      
  
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   
		                                                                    
	Dim iPS1G119
	Dim iErrorPosition
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status                                                            
     
	Set iPS1G119 = Server.CreateObject("PS1G119.cMaintReservePrcMulSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    Dim reqtxtSpread
    reqtxtSpread = Request("txtSpread")
    
    Call iPS1G119.MaintReservePrcMulSvr(gStrGlobalCollection,Trim(reqtxtSpread), iErrorPosition) 
          
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS1G119 = Nothing
       Exit Sub
	End If
	
    Set iPS1G119 = Nothing
    
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "           
              
End Sub    

%>
