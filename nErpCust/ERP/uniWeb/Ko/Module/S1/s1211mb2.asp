<%@ Language=VBSCript%>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1211MB2
'*  4. Program Name         : 고객품목등록 
'*  5. Program Desc         : 고객품목등록 
'*  6. Comproxy List        : PS1G105.dll, PS1G106.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/21 : Grid성능 적용, Kang Jun Gu
'*                            2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%

Err.Clear                                                                        '☜: Clear Error status
On Error Resume Next                                                             '☜: Protect system from crashing

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")

Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
		Call SubBizQueryMulti()
	Case CStr(UID_M0002)
		Call SubBizSaveMulti()
End Select

'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    
    Dim ILngRow
	Dim ILngMaxRow							' 현재 그리드의 최대Row
	Dim lgStrPrevKey1
	Dim lgStrPrevKey2
	
    Dim iS1G106 
    Dim I1_next_b_item
    Dim I2_next_b_biz_partner
    Dim I3_b_item
    Dim I4_b_biz_partner
'    Dim E1_b_item
    Dim E2_b_biz_partner
'    Dim E3_s_bp_item
    Dim EG1_exp_grp
    
	Dim IntGroupCount															'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim StrNextKey1							' BP_CD   다음 값 
	Dim StrNextKey2							' ITEM_CD 다음 값 
	Dim strData
	
	Const C_SHEETMAXROWS_D  = 100
	
	' EG1_exp_grp 저장 
	Const C_b_biz_partner_bp_cd = 0			'고객 
    Const C_b_biz_partner_bp_nm = 1			'고객명 
    Const C_b_item_item_cd = 2				'품목 
    Const C_b_item_item_nm = 3				'품목명 
    Const C_b_item_item_spec = 4			'품목규격 
    Const C_s_bp_item_bp_item_cd = 5		'고객품목 
    Const C_s_bp_item_bp_item_nm = 6		'고객품목명 
    Const C_s_bp_item_bp_item_spec = 7		'고객품목규격 
    Const C_s_bp_item_bp_unit = 8			'고객품목단위 
    Const C_s_bp_item_bp_unit_nm = 9		'고객품목단위명 
    Const C_s_bp_item_ext1_qty = 10			
    Const C_s_bp_item_ext2_qty = 11
    Const C_s_bp_item_ext1_amt = 12
    Const C_s_bp_item_ext2_amt = 13
    Const C_s_bp_item_ext1_cd = 14
    Const C_s_bp_item_ext2_cd = 15
    Const C_s_bp_item_bp_cd = 16
    Const C_s_bp_item_item_cd = 17
    
	I4_b_biz_partner = Trim(Request("txtconBp_cd"))

	Redim I3_b_item(5)
	I3_b_item(0) = Trim(Request("txtconItem_cd"))
	I3_b_item(1) = Trim(Request("txtconItem_Nm"))
	I3_b_item(2) = Trim(Request("txtconItem_spec"))
	I3_b_item(3) = Trim(Request("txtConCustItem_cd"))
	I3_b_item(4) = Trim(Request("txtConCustItem_nm"))
	I3_b_item(5) = Trim(Request("txtConCustItem_spec"))

'	I3_b_item = Trim(Request("txtconItem_cd"))
	
	lgStrPrevKey1 = Trim(Request("lgStrPrevKey1"))                                       '☜: Next Key	
	lgStrPrevKey2 = Trim(Request("lgStrPrevKey2"))                                       '☜: Next Key	
	
	If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" then					
		I1_next_b_item = lgStrPrevKey1
		I2_next_b_biz_partner = lgStrPrevKey2
	Else
		I1_next_b_item = ""
		I2_next_b_biz_partner = ""
	End if
	
	Set iS1G106 = Server.CreateObject("PS1G106.cSListBpItemSvr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   
   
	Call iS1G106.S_LIST_BP_ITEM_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D,  _ 
									I1_next_b_item, I2_next_b_biz_partner, _
									I3_b_item, I4_b_biz_partner,  _ 
									E2_b_biz_partner, EG1_exp_grp)

    Response.Write "<Script language=vbs> "														   & vbCr
    Response.Write " With Parent.frm1 "															   & vbCr      
    Response.Write " .txtconBp_nm.value    = """ & ConvSPChars(E2_b_biz_partner(1))				   & """" & vbCr    
'    Response.Write " .txtconItem_nm.value  = """ & ConvSPChars(E1_b_item(1))		    		   & """" & vbCr    
'    Response.Write " .txtconItem_spec.value  = """ & ConvSPChars(E1_b_item(2))		    		   & """" & vbCr    

    Response.Write " End With"																	   & vbCr
    Response.Write "</Script> "																	   & vbCr      
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iS1G106 = Nothing
		Response.Write "<Script Language=vbscript>"			& vbCr
	    Response.Write "parent.frm1.txtconBp_cd.focus"		& vbCr    
	    Response.Write "</Script>"							& vbCr	
        Exit Sub
    End If   
            
	Set iS1G106 = Nothing	
        
    ILngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count      ]

	For ILngRow = 0 To UBound(EG1_exp_grp, 1)
		If  ILngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey1 = ConvSPChars(EG1_exp_grp(ILngRow, C_b_item_item_cd))
		   StrNextKey2 = ConvSPChars(EG1_exp_grp(ILngRow, C_b_biz_partner_bp_cd))  
		   Exit For
        End If  
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_b_biz_partner_bp_cd))			'고객 
			strData = strData & Chr(11)
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_b_biz_partner_bp_nm))			'고객명 
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_b_item_item_cd))				'품목 
			strData = strData & Chr(11)
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_b_item_item_nm))				'품목명 
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_b_item_item_spec))				'품목규격 
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_s_bp_item_bp_item_cd))			'고객품목 
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_s_bp_item_bp_item_nm))			'고객품목명 
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_s_bp_item_bp_unit))			'고객품목단위 
			strData = strData & Chr(11)
			strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, C_s_bp_item_bp_item_spec))		'고객품목규격 
				
			strData = strData & Chr(11) & ILngMaxRow + ILngRow
			strData = strData & Chr(11) & Chr(12)
    Next   
     
    Response.Write "<Script language=vbs> "																			& vbCr
    Response.Write " With Parent "																					& vbCr      
    Response.Write " .frm1.txtHconBp_cd.value			= """ & ConvSPChars(Trim(Request("txtconBp_cd")))			& """" & vbCr    
    Response.Write " .frm1.txtHconItem_cd.value			= """ & ConvSPChars(Trim(Request("txtconItem_cd")))			& """" & vbCr   
    Response.Write " .frm1.txtHconItem_nm.value			= """ & ConvSPChars(Trim(Request("txtconItem_nm")))			& """" & vbCr   
    Response.Write " .frm1.txtHconItem_spec.value		= """ & ConvSPChars(Trim(Request("txtconItem_spec")))		& """" & vbCr   
    Response.Write " .frm1.txtHconCustItem_cd.value		= """ & ConvSPChars(Trim(Request("txtConCustItem_cd")))		& """" & vbCr   
    Response.Write " .frm1.txtHconcustItem_nm.value		= """ & ConvSPChars(Trim(Request("txtConCustItem_nm")))		& """" & vbCr   
    Response.Write " .frm1.txtHConCustItem_Spec.value	= """ & ConvSPChars(Trim(Request("txtConCustItem_spec")))		& """" & vbCr   
    Response.Write " .frm1.vspdData.ReDraw = False																	" & vbCr
    Response.Write " .SetSpreadColor1 -1																			" & vbCr
    Response.Write " .frm1.vspdData.ReDraw = True 																	" & vbCr
    Response.Write " .ggoSpread.Source          = .frm1.vspdData"													& vbCr
    Response.Write " .ggoSpread.SSShowDataByClip        """ & strData												& """" & vbCr

    Response.Write " .lgStrPrevKey1             = """ & StrNextKey1													& """" & vbCr
    Response.Write " .lgStrPrevKey2             = """ & StrNextKey2													& """" & vbCr    

    Response.Write " .DbQueryOk "															 						& vbCr   
    Response.Write " End With"																						& vbCr
    Response.Write "</Script> "																						& vbCr      
End Sub   

'============================================================================================================
Sub SubBizSaveMulti()   
	
	Dim iS1G105
	Dim iErrorPosition
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear	
    
	Set iS1G105 = Server.CreateObject("PS1G105.cSBpItemMultiSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    Dim reqtxtSpread
    reqtxtSpread = Request("txtSpread")
    Call iS1G105.S_MAINT_BP_ITEM_MULTI_SVR(gStrGlobalCollection, Trim(reqtxtSpread), iErrorPosition)
    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
	   Set iS1G105 = Nothing
		%>
		<Script Language=vbscript>
			Call Parent.SubSetErrPos("<%=iErrorPosition%>")
		</Script>
		<%																
       Exit Sub
	End If
	
	Set iS1G105 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "           
      
End Sub    

%>
