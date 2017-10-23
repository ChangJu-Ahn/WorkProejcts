<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1311rb1
'*  4. Program Name         : Bom 정보 Ref
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/08/29
'*  8. Modified date(Last)  : 2003/06/12
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :	
'* 14. Meno					: Biz logic of BOM정보 
'**********************************************************************************************
Dim lgOpModeCRUD
 
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")                                                               '☜: Hide Processing message
	
lgOpModeCRUD  = Request("txtMode") 
									                                              '☜: Read Operation Mode (CRUD)
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         Call  SubBizQueryMulti()
End Select
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next
    Err.Clear			
	
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iPP1C001
	
	Dim iStrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	
	Dim TmpBuffer
	Dim iMax
	Dim iIntLoopCount
	Dim iTotalStr
	
	Dim I1_b_plant_plant_cd 
    Dim I2_b_item_item_cd
    Dim E1_p_bom_header 
    Dim E2_b_item 
    Dim E3_B_Plant
    Dim EG1_exp_group
    Dim E4_p_bom_for_explosion  'next seq
	
	Const M313_E1_bom_no = 0    
    Const M313_E1_description = 1
    Const M313_E1_major_flg = 2
    Const M313_E1_valid_from_dt = 3
    Const M313_E1_valid_to_dt = 4

    Const M313_E2_item_cd = 0    
    Const M313_E2_item_nm = 1
    Const M313_E2_formal_nm = 2
    Const M313_E2_spec = 3
    Const M313_E2_item_acct = 4
    Const M313_E2_item_class = 5
    Const M313_E2_hs_cd = 6
    Const M313_E2_hs_unit = 7
    Const M313_E2_unit_weight = 8
    Const M313_E2_unit_of_weight = 9
    Const M313_E2_basic_unit = 10
    Const M313_E2_phantom_flg = 11
    Const M313_E2_draw_no = 12
    Const M313_E2_blanket_pur_flg = 13
    Const M313_E2_base_item_cd = 14
    Const M313_E2_proportion_rate = 15
    Const M313_E2_valid_flg = 16

    Const M313_E3_plant_cd = 0    
    Const M313_E3_plant_nm = 1

    Const M313_EG1_E1_item_cd = 0    
    Const M313_EG1_E1_item_nm = 1
    Const M313_EG1_E2_plant_cd = 2   
    Const M313_EG1_E2_user_id = 3
    Const M313_EG1_E2_seq = 4
    Const M313_EG1_E2_prnt_node = 5
    Const M313_EG1_E2_own_node = 6
    Const M313_EG1_E2_material_flg = 7
    Const M313_EG1_E2_level_cd = 8
    Const M313_EG1_E2_prnt_item_cd = 9
    Const M313_EG1_E2_prnt_bom_no = 10
    Const M313_EG1_E2_child_item_seq = 11
    Const M313_EG1_E2_child_item_cd = 12
    Const M313_EG1_E2_child_bom_no = 13
    Const M313_EG1_E2_prnt_item_qty = 14
    Const M313_EG1_E2_prnt_item_unit = 15
    Const M313_EG1_E2_child_item_qty = 16
    Const M313_EG1_E2_child_item_unit = 17
    Const M313_EG1_E2_loss_rate = 18
    Const M313_EG1_E2_safety_lt = 19
    Const M313_EG1_E2_supply_type = 20
    Const M313_EG1_E2_bom_flg = 21
    Const M313_EG1_E2_remark = 22
    Const M313_EG1_E2_valid_from_dt = 23
    Const M313_EG1_E2_valid_to_dt = 24
	
	Dim L1_p_bom_detail
	Redim L1_p_bom_detail(1)
	
	lgStrPrevKey = Request("lgStrPrevKey")
	I2_b_item_item_cd 	= UCase(Trim(Request("txtItemCd")))
    I1_b_plant_plant_cd = UCase(Trim(Request("txtPlantCd")))
    '=======================================
	'200704 KSJ 추가(BOM적용유효일추가)
    L1_p_bom_detail(0) = Trim(Request("txtFrDt"))
    L1_p_bom_detail(1) = Trim(Request("txtToDt"))
    '=======================================
	
	Set iPP1C001 = CreateObject("PP1C001.cPListMajorBom")
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPP1C001 = Nothing												
		Exit Sub														
	End if
    Call iPP1C001.P_LIST_MAJOR_BOM(gStrGlobalCollection, I2_b_item_item_cd, _
                            I1_b_plant_plant_cd, L1_p_bom_detail, Cstr(E4_p_bom_for_explosion), _
                            EG1_exp_group, E3_B_Plant, E2_b_item, E1_p_bom_header)
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPP1C001 = Nothing												
		Exit Sub														
	End if
    Set iPP1C001 = Nothing
       
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtPlantCd.value = """ & UCase(Request("txtPlantCd"))   & """" & vbCr
	Response.Write "	.frm1.txtPlantNm.value = """ & ConvSPChars(E3_B_Plant(M313_E3_plant_nm))  & """" & vbCr
	Response.Write "	.frm1.txtItemCd.value = """ & ConvSPChars(E2_b_item(M313_E2_item_cd))  & """" & vbCr
	Response.Write "	.frm1.txtItemNm.value = """ & ConvSPChars(E2_b_item(M313_E2_item_nm))  & """" & vbCr
	Response.Write "	.frm1.txtBomno.value = """ & ConvSPChars(E1_p_bom_header(M313_E1_bom_no))  & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr
	
	If lgStrPrevKey = StrNextKey And UBound(EG1_exp_group, 1) < 0 Then
		Exit Sub														
	End If
	
	iLngMaxRow = Request("txtMaxRows")											                                
    GroupCount = UBound(EG1_exp_group,1)
    
	If EG1_exp_group(GroupCount, M313_EG1_E2_child_item_seq) = E4_p_bom_for_explosion Then
		StrNextKey = ""
	Else
		StrNextKey = E4_p_bom_for_explosion
	End If
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr	
	
	Response.Write " .frm1.txtPlantCd.value     = """ & ConvSPChars(E3_B_Plant(M313_E3_plant_cd))  & """" & vbCr
	Response.Write " .frm1.txtPlantNm.value     = """ & ConvSPChars(E3_B_Plant(M313_E3_plant_nm))  & """" & vbCr
	Response.Write " .frm1.txtItemCd.value		= """ & ConvSPChars(E2_b_item(M313_E2_item_cd)) & """" & vbCr
	Response.Write " .frm1.txtItemNm.value		= """ & ConvSPChars(E2_b_item(M313_E2_item_nm))    & """" & vbCr
	Response.Write " .frm1.txtBomno.value		= """ & ConvSPChars(E1_p_bom_header(M313_E1_bom_no))   & """" & vbCr
	
	iIntLoopCount = 0
	iMax = UBound(EG1_exp_group,1)
	ReDim TmpBuffer(iMax)
	
	
	For iLngRow = 0 To iMax
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = E4_p_bom_for_explosion  'next값...
           Exit For
        End If  
        
		istrData = ""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M313_EG1_E1_item_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M313_EG1_E1_item_nm))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M313_EG1_E2_prnt_item_qty), 4, 0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M313_EG1_E2_prnt_item_unit))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M313_EG1_E2_child_item_qty), 4, 0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M313_EG1_E2_child_item_unit))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M313_EG1_E2_supply_type))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M313_EG1_E2_loss_rate))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow, M313_EG1_E2_valid_from_dt))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow, M313_EG1_E2_valid_to_dt))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
           
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next  
	
	iTotalStr = Join(TmpBuffer, "")    
	
	Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr	    & """" & vbCr	
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 
	Response.Write " .frm1.hdnPlantCd.value	= """ & ConvSPChars(UCase(Request("txtPlantCd"))) & """" & vbCr
	Response.Write " .frm1.hdnItemCd.value  = """ & ConvSPChars(UCase(Request("txtitemCd")))  & """" & vbCr
	Response.Write " .frm1.hdnFrDt.value  = """ & ConvSPChars(UCase(Request("txtFrDt")))  & """" & vbCr
	Response.Write " .frm1.hdnToDt.value  = """ & ConvSPChars(UCase(Request("txtToDt")))  & """" & vbCr
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr
    
	
		
End Sub    
%>
