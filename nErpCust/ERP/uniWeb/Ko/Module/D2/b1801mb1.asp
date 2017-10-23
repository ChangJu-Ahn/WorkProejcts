<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : b1801mb1
'*  4. Program Name         : 경비항목설정 
'*  5. Program Desc         : 경비항목설정 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/08/02
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              : 2002/11/15 : UI성능 적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf()

Dim lgOpModeCRUD
On Error Resume Next                                                            
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
		
lgOpModeCRUD = Request("txtMode")                                                '☜: Read Operation Mode (CRUD)
	    
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query        
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update        
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                         '☜: Delete        
End Select

   
'============================================================================================================
Sub SubBizQueryMulti()
    
    Dim LngRow	
	Dim LngMaxRow
	
	Dim lgstrData
	
	
	Dim StrPrevKey
	
    Dim iPD2GS62    
    
    Dim intGroupCount
    
    Dim StrNext
    Dim StrNextKey  	
    Dim iarrValue
    Const C_SHEETMAXROWS_D  = 100
    Const C_Cost     = "단순경비"
    Const C_Material = "물대포함"
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	LngMaxRow       = CLng(Request("txtMaxRows"))                                  '☜: Fetechd Count
	StrPrevKey      =Trim(Request("lgStrPrevKey"))                                 '☜: Next Key


    'View Name : import_next b_trade_charge
    Const S090_I1_charge_cd = 0
    Const S090_I1_module_type = 1
    Const S090_I1_cost_flag = 2
    
    'View Name : import b_trade_charge
    Const S090_I2_charge_cd = 0
    Const S090_I2_module_type = 1
    
    'View Name : exp a_jnl_item
    Const S092_E1_jnl_cd = 0
    Const S092_E1_jnl_nm = 1    
  
    'Group Name : export_group
    Const S092_EG1_E1_a_acct_trans_type_trans_nm = 0
    Const S092_EG1_E2_a_jnl_item_jnl_nm = 1
    Const S092_EG1_E4_b_trade_charge_charge_cd = 2
    Const S092_EG1_E4_b_trade_charge_module_type = 3
    Const S092_EG1_E4_b_trade_charge_cost_flag = 4
    Const S092_EG1_E4_b_trade_charge_std_rate = 5
    Const S092_EG1_E4_b_trade_charge_acct_trans_type = 6
    Const S092_EG1_E4_b_trade_charge_distribution_flag = 7
    Const S092_EG1_E4_b_trade_charge_ext1_qty = 8
    Const S092_EG1_E4_b_trade_charge_ext2_qty = 9
    Const S092_EG1_E4_b_trade_charge_ext3_qty = 10
    Const S092_EG1_E4_b_trade_charge_ext1_amt = 11
    Const S092_EG1_E4_b_trade_charge_ext2_amt = 12
    Const S092_EG1_E4_b_trade_charge_ext3_amt = 13
    Const S092_EG1_E4_b_trade_charge_ext1_cd = 14
    Const S092_EG1_E4_b_trade_charge_ext2_cd = 15
    Const S092_EG1_E4_b_trade_charge_ext3_cd = 16
    Const S092_EG1_E4_b_trade_charge_vat_flg = 17
    
    'View Name : export_next b_trade_charge
    Const S092_E2_charge_cd = 0
    Const S092_E2_module_type = 1
    Const S092_E2_cost_flag = 2
        
    Dim I1_b_trade_charge
    Dim I2_b_trade_charge
        
    Dim E2_b_trade_charge 
    Dim EG1_export_group
    Dim E1_a_jnl_item
    
    Dim strtxtModuleType
    Dim strtxtChargeItem
    Dim strtxtRadio
    Redim I2_b_trade_charge(S090_I2_module_type)
             
	I2_b_trade_charge(S090_I2_charge_cd)   = FilterVar(Trim(Request("txtChargeItem")), "", "SNM")
	I2_b_trade_charge(S090_I2_module_type) = Trim(Request("txtModuleType"))  
	strtxtModuleType  = Trim(Request("txtModuleType"))
	strtxtChargeItem = Request("txtChargeItem")
	strtxtRadio = Request("txtRadio")
	
    
	
    Redim I1_b_trade_charge(S090_I1_cost_flag)
     
	If  StrPrevKey <> "" Then
		
		iarrValue = Split(StrPrevKey, gColSep)
		           
		I1_b_trade_charge(S090_I1_charge_cd)   = Trim(iarrValue(S090_I1_charge_cd))
		I1_b_trade_charge(S090_I1_module_type) = Trim(iarrValue(S090_I1_module_type))
		I1_b_trade_charge(S090_I1_cost_flag)   = Trim(iarrValue(S090_I1_cost_flag))
		 
	Else 
	    I1_b_trade_charge(S090_I1_charge_cd)   = ""
		I1_b_trade_charge(S090_I1_module_type) = ""
		I1_b_trade_charge(S090_I1_cost_flag)   = ""
		
    End If
    
    Set iPD2GS62 = Server.CreateObject("PD2GS62.cBLtTradeChargeSvr")
  
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If 

 
    Call iPD2GS62.B_LIST_TRADE_CHARGE_SVR (gStrGlobalCollection,C_SHEETMAXROWS_D,I1_b_trade_charge, _
                                       I2_b_trade_charge,E1_a_jnl_item,EG1_export_group ,E2_b_trade_charge) 
    
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " Parent.frm1.txtChargeItemNm.value   = """ & ConvSPChars(E1_a_jnl_item(S092_E1_jnl_nm))      & """" & vbCr
    Response.Write "</Script> "	
               
   	If CheckSYSTEMError(Err,True) = True Then
       Set iPD2GS62 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   

     Set iPD2GS62 = Nothing	
		   
	StrNext = E2_b_trade_charge(S092_E2_charge_cd)
	StrNext = StrNext & gColSep & E2_b_trade_charge(S092_E2_module_type)
	StrNext = StrNext & gColSep & E2_b_trade_charge(S092_E2_cost_flag)
		  
	For LngRow = 0 To Ubound(EG1_export_group,1)
		If  LngRow < C_SHEETMAXROWS_D  Then
     	Else
	       StrNextKey = StrNext 
	          Exit For
        End If  
 
   '단순경비여부 
	 	If EG1_export_group(LngRow,S092_EG1_E4_b_trade_charge_cost_flag) = "C" Then
	 		lgstrData = lgstrData & Chr(11) & C_Cost
	 	Else EG1_export_group(LngRow,S092_EG1_E4_b_trade_charge_cost_flag) = "M" 
	 		lgstrData = lgstrData & Chr(11) & C_Material
        End If	 	
   '경비항목코드 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,S092_EG1_E4_b_trade_charge_charge_cd))
   '경비항목버튼 
		lgstrData = lgstrData & Chr(11) & ""
   '경비항목명 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,S092_EG1_E2_a_jnl_item_jnl_nm))
   '회계거래 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,S092_EG1_E4_b_trade_charge_acct_trans_type))
   '회계거래버튼 
		lgstrData = lgstrData & Chr(11) & ""
   '회계거래유형명 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,S092_EG1_E1_a_acct_trans_type_trans_nm))
   'VAT 여부 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,S092_EG1_E4_b_trade_charge_vat_flg))
	    lgstrData = lgstrData & Chr(11) & LngMaxRow + LngRow 
        lgstrData = lgstrData & Chr(11) & Chr(12)  
       
    Next
    
    Response.Write "<Script language=vbs> " & vbCr   
    
    If strtxtModuleType = "S" Then
		Response.Write " Parent.frm1.rdoModuleFlag_S.checked = True                                                       " & vbCr	
	Else
		Response.Write " Parent.frm1.rdoModuleFlag_M.checked = True						            					      " & vbCr
	End If
		
	
	Response.Write " Parent.frm1.txtHChargeItem.value    = """ & ConvSPChars(strtxtChargeItem)         	 & """" & vbCr  
    Response.Write " Parent.frm1.txtHRadio.value         = """ & strtxtRadio				         	 & """" & vbCr
         
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData					                      " & vbCr
    Response.Write " Parent.ggoSpread.SSShowData        """ & lgstrData										     & """" & vbCr
    
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey										 & """" & vbCr  
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write "</Script> "																							& vbCr      
  
End Sub    


'============================================================================================================
Sub SubBizSaveMulti()   
		                                                                    
	Dim iPD2GS61
	Dim strtxtSpread
	Dim iErrorPosition
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status                                                            
    strtxtSpread = Trim(Request("txtSpread"))
    
	Set iPD2GS61 = Server.CreateObject("PD2GS61.cBTradeChargeSvr")
	

    If CheckSYSTEMError(Err,True) = True Then    
       Exit Sub
    End If
    
    Call iPD2GS61.B_MAINT_TRADE_CHARGE_SVR (gStrGlobalCollection, strtxtSpread, iErrorPosition) 
    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPD2GS61 = Nothing
       Exit Sub
	End If
	
    Set iPD2GS61 = Nothing
    
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "           
              
End Sub    

%>
