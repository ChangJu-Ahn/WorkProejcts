<%@ LANGUAGE=VBSCript%>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s1931ma1_ko441
'*  4. Program Name         : 단가처리config등록(고객) 
'*  5. Program Desc         : 단가처리config등록(고객) 
'=============================================================================f==========================
%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%        

    On Error Resume Next    
	Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	Call LoadBasisGlobalInf()

    Call HideStatusWnd                                                               '☜: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query			
             Call SubBizQuery()
        Case CStr(UID_M0002)														'☜: Save,Update, Delete
             Call SubBizSaveMulti()
    End Select    
 
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
Sub SubBizQueryMulti()
	
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	    
    Dim StrNextKey  	
    Dim arrValue    

	Dim imp_s_bp_item_price(4)
	Dim imp_SBIP
	
    Dim imp_b_item 'As String
    Dim imp_b_biz_partner 'As String
    Dim imp_next_s_bp_item_price(4)
    Dim imp_next_SBIP
    Dim imp_next_b_item 'As String
    Dim imp_next_b_biz_partner 'As String           
        
    Dim exp_grp
	Dim exp_next_b_item
    Dim exp_next_b_biz_partner
    Dim exp_next_s_bp_item_price    
    Dim exp_b_biz_partner
    Dim exp_b_item
    Dim exp_deal_type_b_minor
    Dim exp_pay_meth_b_minor
    Dim exp_b_unit_of_measure
           
'    '''s_bp_item_price 
    Const C_deal_type = 0
    Const C_pay_meth = 1
    Const C_sales_unit = 2
    Const C_valid_from_dt = 3
    Const C_currency = 4

'	exp_grp에 담을 스프레드에 보일내용 
    Const C_exp_bp_cd = 0
    Const C_exp_bp_nm = 1
    Const C_exp_item_cd = 2
    Const C_exp_item_nm = 3
    Const C_exp_item_spce = 4
    Const C_exp_deal_type = 5
    Const C_exp_pay_meth = 6
    Const C_exp_sales_unit = 7
    Const C_exp_currency = 8
    Const C_exp_valid_from_dt = 9
    Const C_exp_item_price = 10
    Const C_exp_ext1_qty = 11
    Const C_exp_ext2_qty = 12
    Const C_exp_ext1_amt = 13
    Const C_exp_ext2_amt = 14
    Const C_exp_ext1_cd = 15
    Const C_exp_ext2_cd = 16
    Const C_exp_deal_type_nm = 17
    Const C_exp_pay_meth_nm = 18
    Const C_exp_Price_Flag = 19
    Const C_exp_remrk = 20
	
    Dim pS11128      

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status    

	
    imp_b_item = Trim(Request("txtconItem_cd"))
    imp_b_biz_partner = Trim(Request("txtconBp_cd"))
       
    iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	    
  	
  	
  	If iStrPrevKey <> "" then					
		arrValue = Split(iStrPrevKey, gColSep)
		imp_next_b_item = Trim(arrValue(0))	
		imp_next_b_biz_partner = Trim(arrValue(1))	
	else			
		imp_next_b_item = ""
		imp_next_b_biz_partner = ""
	End If 


    imp_SBIP = imp_s_bp_item_price    
    imp_next_SBIP =  imp_next_s_bp_item_price
	
	Set pS11128 = Server.CreateObject("PS1G104.CsLtBpItemProceSvr")
	
	if CheckSYSTEMError(Err,True) = True Then 
		Set pS11128 = Nothing
		Response.Write "<Script language=vbs> "			& vbCr       		
		Response.Write "Parent.frm1.txtconBp_cd.focus"	& vbCr       
		Response.Write "</Script> "						& vbCr          
		
		Exit Sub
	end if
  
	Call pS11128.S_LIST_BP_ITEM_PROCE_SVR(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), CStr(imp_b_item), CStr(imp_b_biz_partner), _
        imp_SBIP, CStr(imp_next_b_item), CStr(imp_next_b_biz_partner), imp_next_SBIP, _
        exp_grp)
    
    If CheckSYSTEMError(Err,True) = True Then 	
		Set pS11128 = Nothing
		Exit Sub
	end if
		
	iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count          
	For iLngRow = 0 To UBound(exp_grp,1)			
		
		If  iLngRow < C_SHEETMAXROWS_D  Then			
		Else	   		
		   StrNextKey = ConvSPChars(exp_grp(iLngRow, C_exp_item_cd))                    & gColSep '0		   		   
		   StrNextKey = StrNextKey & ConvSPChars(exp_grp(iLngRow, C_exp_bp_cd))         & gColSep '1
		   StrNextKey = StrNextKey & ConvSPChars(exp_grp(iLngRow, C_exp_deal_type))     & gColSep '2
		   StrNextKey = StrNextKey & ConvSPChars(exp_grp(iLngRow, C_exp_pay_meth))      & gColSep '3
		   
		   StrNextKey = StrNextKey & UNIConvDate(exp_grp(iLngRow, C_exp_valid_from_dt)) & gColSep '4
		   
		   StrNextKey = StrNextKey & ConvSPChars(exp_grp(iLngRow, C_exp_sales_unit))    & gColSep '5
		   StrNextKey = StrNextKey & ConvSPChars(exp_grp(iLngRow, C_exp_currency))                '6  				   
           Exit For
        End If        
		        
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp( iLngRow,C_exp_bp_cd)) 
		istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp( iLngRow,C_exp_bp_nm))
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_item_cd))
		istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_item_nm ))
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_item_spce ))
        
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_deal_type))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_deal_type_nm))				
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_pay_meth))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_pay_meth_nm))
		istrData = istrData & Chr(11) & UNIDateClientFormat(exp_grp(iLngRow, C_exp_valid_from_dt ))				

		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_sales_unit))
		istrData = istrData & Chr(11) & ""	
        istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_currency))
        istrData = istrData & Chr(11) & ""        

		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, C_exp_item_price), ggUnitCost.DecPoint, 0 )				
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_exp_Price_Flag))
		if ConvSPChars(exp_grp(iLngRow, C_exp_Price_Flag))="T" then
			istrData = istrData & Chr(11) & ConvSPChars("진단가")
		Else
			istrData = istrData & Chr(11) & ConvSPChars("가단가")		
		End if
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow,C_exp_remrk))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)               
    Next    
 
    
    Response.Write "<Script language=vbs> " & vbCr      
    
	Response.Write " Parent.frm1.txtconBp_nm.value   = """ & ConvSPChars(exp_b_biz_partner(0))     & """" & vbCr    
    
	Response.Write " Parent.frm1.vspdData.ReDraw = False													" & vbCr			        
    Response.Write " Parent.SetSpreadColor1 -1																" & vbCr    
    Response.Write " Parent.frm1.vspdData.ReDraw = True														" & vbCr
    
    Response.Write " Parent.frm1.txtHconBiz_partner.value  = """ & Trim(Request("txtconBiz_partner"))   & """" & vbCr
    Response.Write " Parent.frm1.txtHconItem_cd.value      = """ & Trim(Request("txtconItem_cd"))	    & """" & vbCr
    
    Response.Write " Parent.ggoSpread.Source               = Parent.frm1.vspdData									      " & vbCr
    Response.Write  "Parent.frm1.vspdData.Redraw = False  "      & vbCr      
 '   Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & istrData & """ ,""F""" & vbCr
 '   Response.Write  "Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & -1 & "," & -1  & ",Parent.C_Cur,Parent.C_Item_Price,""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey										 & """" & vbCr  
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write  "Parent.frm1.vspdData.Redraw = True  "       & vbCr      
    Response.Write "</Script> "																							& vbCr      
    
	Set pS11128 = Nothing	    
	
End Sub    

'============================================================================================================
Sub SubBizSaveMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
	Dim PS1G103
	Dim iErrorPosition	

	
	Set PS1G103 = Server.CreateObject("PS1G103.CsBpItemPrcMulSvr")	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If        

	Dim reqtxtSpread
	reqtxtSpread = Trim(Request("txtSpread"))
    Call PS1G103.S_MAINT_BP_ITEM_PRC_MUL_SVR(gStrGlobalCollection, cstr(reqtxtSpread),iErrorPosition )				
	
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set PS1G103 = Nothing
       Exit Sub
	End If	
	
    '-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	
    Set PS1G103 = Nothing    
        
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "               
    
End Sub
%>
