<%@ LANGUAGE=VBSCript%>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1312MB2
'*  4. Program Name         : 고객별품목할증등록 
'*  5. Program Desc         : 고객별품목할증등록 
'*  6. Comproxy List        : PS1G109.dll, PS1G110.dll
'*  7. Modified date(First) : 2000/04/09
'*  8. Modified date(Last)  : 2002/05/28
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Ryu KyungRae
'* 11. Comment              : 
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'**********************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Call HideStatusWnd                                                               '☜: Hide Processing message
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
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	
    Dim iS13128    
    '-----------------------------------------------
    ' Declare User Variable
    '-----------------------------------------------
    ' 고객 / 품목 / S_BP_ITEM_DC Table
    Dim i_B_Biz_Partner1
    Dim i_B_item1
    
    Dim i_SBp_item_dc1 
    ReDim i_SBp_item_dc1(2)    
    
    ' Next Page Variable
    Dim imp_next_s_bp_item_dc1
    ReDim imp_next_s_bp_item_dc1(3)
    Dim imp_next_b_item1
    Dim imp_next_b_biz_partner1    

    ' Reruen Call Variable
    Dim i_B_Biz_Partner
    Dim i_B_item    
    Dim i_SBp_item_dc    
    Dim imp_next_s_bp_item_dc
    Dim imp_next_b_item
    Dim imp_next_b_biz_partner

    ' Export Variables
    Dim exp_b_biz_partner
    Dim exp_b_item
    Dim exp_pay_meth_b_minor
    Dim exp_grp   
    
    Dim intGroupCount
    Dim StrNextKey  	
    Dim arrValue
    
    Const C_SHEETMAXROWS_D = 100
    
    ' exp_grp 저장 
    Const C_Exp_Bp_Cd               = 0
    Const C_Exp_Bp_Nm				= 1
    Const C_Exp_Item_Cd             = 2
    Const C_Exp_Item_Nm             = 3
    Const C_Exp_Item_Spec           = 4
    Const C_Exp_Pay_terms           = 5
    Const C_Exp_Pay_terms_nm        = 6
    Const C_Exp_Valid_from_dt       = 7
    Const C_Exp_Unit                = 8
    Const C_Exp_DC_BAS_Qty          = 9
    Const C_Exp_Dc_rate             = 10
    Const C_Exp_DC_Kind             = 11 
    Const C_Exp_DC_Kind_Nm          = 12
    Const C_Exp_Round_type          = 13
    Const C_Exp_Round_type_Nm       = 14
    Const C_Exp_ChgFlg				= 15

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    '-----------------------------------------------
    ' 고객 
    '-----------------------------------------------
    i_B_Biz_Partner1 = Trim(Request("txtconBiz_Partner"))
    i_B_Biz_Partner = i_B_Biz_Partner1
    '-----------------------------------------------
    ' 품목 
    '-----------------------------------------------
    i_B_item1 = Trim(Request("txtconItem_cd"))
    i_B_item = i_B_item1    
    '-----------------------------------------------
    ' S_BP_ITEM_DC TABLE 의 참조값 
    ' 0 : 결제방법 1 : 결제 단위 2 : 적용일 
    '-----------------------------------------------    
    i_SBp_item_dc1(0) = Trim(Request("txtconPay_terms"))
    i_SBp_item_dc1(1) = Trim(Request("txtconSales_unit"))
    i_SBp_item_dc1(2) = uniconvDate(Trim(Request("txtconValid_from_dt")))

    i_SBp_item_dc = i_SBp_item_dc1

	iStrPrevKey = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	
	
	If iStrPrevKey <> "" then
	    
		arrValue = Split(iStrPrevKey, gColSep)

        imp_next_s_bp_item_dc1(0) = arrValue(0)
        imp_next_s_bp_item_dc1(1) = arrValue(1)
        imp_next_s_bp_item_dc1(2) = arrValue(2)
        imp_next_s_bp_item_dc1(3) = arrValue(3)

        imp_next_b_item1 = arrValue(4)
        imp_next_b_biz_partner1 = arrValue(5)
        
        imp_next_s_bp_item_dc = imp_next_s_bp_item_dc1
        imp_next_b_item = imp_next_b_item1
        imp_next_b_biz_partner = imp_next_b_biz_partner1
        
	else			

        imp_next_s_bp_item_dc1(0) = ""
        imp_next_s_bp_item_dc1(1) = ""
        imp_next_s_bp_item_dc1(2) = ""
        imp_next_s_bp_item_dc1(3) = ""

        imp_next_b_item1 = ""
        imp_next_b_biz_partner1 = ""

        imp_next_s_bp_item_dc = imp_next_s_bp_item_dc1
        imp_next_b_item = imp_next_b_item1
        imp_next_b_biz_partner = imp_next_b_biz_partner1

	End If

    
	set iS13128 = CREATEOBJECT("PS1G110.cSListBpItemDcSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   
   	  		
	' Call the Dll   		
    Call iS13128.S_LIST_BP_ITEM_DC_SVR ( gStrGlobalCollection, _
                           cint(C_SHEETMAXROWS_D), _
                           imp_next_s_bp_item_dc, imp_next_b_item, imp_next_b_biz_partner, _
                           i_SBp_item_dc, i_B_item, i_B_Biz_Partner, _
                           exp_b_item, exp_b_biz_partner, exp_pay_meth_b_minor, _
                           exp_grp)
	
	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " Parent.frm1.txtconBiz_Partner_nm.value = """ & ConvSPChars(exp_b_biz_partner(1))		& """" & vbCr    
    Response.Write " Parent.frm1.txtconItem_nm.value		= """ & ConvSPChars(exp_b_item(1))				& """" & vbCr    
    Response.Write " Parent.frm1.txtconPay_terms_nm.value   = """ & ConvSPChars(exp_pay_meth_b_minor(1))	& """" & vbCr    
    Response.Write "</Script> "	& vbCr      
    
	If CheckSYSTEMError(Err,TRUE) = True Then
       Set iS13128 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   
    
    Set iS13128 = Nothing	
        
    iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count      
        
    For iLngRow = 0 To UBound(exp_grp,1)
    		
        If  iLngRow < C_SHEETMAXROWS_D  Then
        Else

            StrNextKey = ConvSPChars(exp_grp(iLngRow, C_Exp_Pay_terms))
            StrNextKey = StrNextKey & gColSep & ConvSPChars(exp_grp(iLngRow, C_Exp_Unit))
            StrNextKey = StrNextKey & gColSep & UNINumClientFormat(exp_grp(iLngRow, C_Exp_DC_BAS_Qty), ggAmtOfMoney.DecPoint, 0)
            StrNextKey = StrNextKey & gColSep & UNIConvDate(exp_grp(iLngRow, C_Exp_Valid_from_dt))
            StrNextKey = StrNextKey & gColSep & ConvSPChars(exp_grp(iLngRow, C_Exp_Item_Cd))
            StrNextKey = StrNextKey & gColSep & ConvSPChars(exp_grp(iLngRow, C_Exp_Bp_Cd))
            
            Exit For
        End If 
        
        If  (iLngMaxRow >= C_SHEETMAXROWS_D) AND iLngRow = 0 Then

        else        
        
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Bp_Cd))
            istrdata = istrdata & Chr(11) & ""
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Bp_Nm))
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Item_Cd))
            istrdata = istrdata & Chr(11) & ""
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Item_Nm))
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Item_Spec))
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Pay_terms))
            istrdata = istrdata & Chr(11) & ""
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Pay_terms_nm))
            istrdata = istrdata & Chr(11) & UNIDateClientFormat(exp_grp(iLngRow, C_Exp_Valid_from_dt))
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Unit))
            istrdata = istrdata & Chr(11) & ""
            istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, C_Exp_DC_BAS_Qty), ggAmtOfMoney.DecPoint, 0)
            istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, C_Exp_Dc_rate), ggAmtOfMoney.DecPoint, 0)
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_DC_Kind))
            istrdata = istrdata & Chr(11) & ""
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_DC_Kind_Nm))
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Round_type))
            istrdata = istrdata & Chr(11) & ""
            istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, C_Exp_Round_type_Nm))
            istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
            istrData = istrData & Chr(11) & Chr(12)       
        end if

    Next 
    
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " Parent.frm1.txtHconBiz_Partner.value   = """ & Trim(Request("txtconBiz_Partner")) & """" & vbCr    
    Response.Write " Parent.frm1.txtHconItem_cd.value = """ & Trim(Request("txtconItem_cd")) & """" & vbCr    
    Response.Write " Parent.frm1.txtHconPay_terms.value   = """ & Trim(Request("txtconPay_terms")) & """" & vbCr    
    Response.Write " Parent.frm1.txtHconValid_from_dt.value = """ & Trim(Request("txtconValid_from_dt")) & """" & vbCr    
    Response.Write " Parent.frm1.txtHconSales_unit.value   = """ & Trim(Request("txtconSales_unit")) & """" & vbCr    
    Response.Write " Parent.frm1.txtconBiz_Partner_nm.value   = """ & ConvSPChars(exp_b_biz_partner(1)) & """" & vbCr    
    Response.Write " Parent.frm1.txtconItem_nm.value = """ & ConvSPChars(exp_b_item(1)) & """" & vbCr    
    Response.Write " Parent.frm1.txtconPay_terms_nm.value   = """ & ConvSPChars(exp_pay_meth_b_minor(1)) & """" & vbCr    
    Response.Write " Parent.SetSpreadColor1 -1	                                                          " & vbCr 		
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData									      " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip        """ & istrData										     & """" & vbCr
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey										 & """" & vbCr  
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write "</Script> "	& vbCr      
        
End Sub    

'============================================================================================================
Sub SubBizSaveMulti()   

Dim S13121m	
Dim iErrorPosition
	
On Error Resume Next                                                                 '☜: Protect system from crashing
Err.Clear																			 '☜: Clear Error status                                                            
     
	Set S13121m = Server.CreateObject("PS1G109.cSBpItemDcMulSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    Dim reqtxtSpread
    reqtxtSpread = Request("txtSpread")
    
    Call S13121m.S_MAINT_BP_ITEM_DC_MUL_SVR  (gStrGlobalCollection, _
												Trim(reqtxtSpread), iErrorPosition)    
												      
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set S13121m = Nothing
       Exit Sub
	End If
	
    Set S13121m = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr   
                  
End Sub    

%>
