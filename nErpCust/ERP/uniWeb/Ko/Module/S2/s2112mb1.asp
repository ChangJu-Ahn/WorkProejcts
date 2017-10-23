<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2112MB1
'*  4. Program Name         : 공장별 판매계획조정 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS2G133.dll, PS2G134.dll, PS2G136.dll
'*  7. Modified date(First) : 2000/04/08
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Mr  Cho
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->

<%

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
Call HideStatusWnd                     

Dim PS2G134												'ComProxy Dll 사용 변수 
Dim PS2G133												'ComProxy Dll 사용 변수 
Dim PS2G136
Dim iErrorPosition

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngMaxRow							' 현재 그리드의 최대Row
Dim LngRow,iRow
Dim GroupCount          

Dim strCountryCD
Dim CommandSent
'type MissMach
Dim RequestTxtSpread

Dim I1_s_item_sales_plan_by_plant
Dim I2_s_item_sales_plan_by_plant
Dim I3_s_item_sales_plan_by_plant
Dim I4_s_item_sales_plan
Dim I5_s_item_sales_plan
Dim I6_b_item
Dim I7_b_item
Dim I8_b_plant

Dim E1_ief_supplied
Dim E2_b_plant
Dim E3_s_item_sales_plan_by_plant
Dim E4_b_item
Dim E5_b_item
Dim E6_s_cfm_item_sales_plan
Dim EG1_exp_plgrp
Dim EG2_exp_grp
Dim prGroupView


Redim I4_s_item_sales_plan(1)
Redim I5_s_item_sales_plan(1)

Const S235_I4_sp_year = 0
Const S235_I4_sp_month = 1
Const S235_I5_sp_year = 0 
Const S235_I5_sp_month = 1

Const S235_E1_total_real = 0 
Const S235_E2_plant_nm = 0   
Const S235_E3_sp_dt = 0    
Const S235_E4_item_cd = 0  
Const S235_E5_basic_unit = 0 
Const S235_E5_item_nm = 1
Const S235_E6_cfm_flag = 0 

Const S235_EG2_E2_item_cd = 0    ' : exp_item b_item
Const S235_EG2_E2_item_nm = 1
Const S235_EG2_E3_sp_dt = 2		  ' : exp_item s_item_sales_plan_by_plant
Const S235_EG2_E3_plan_qty = 3
Const S235_EG2_E3_req_sts = 4
Const S235_EG2_E3_plan_bunit_qty = 5
Const S235_EG2_E3_cfm_qty = 6
Const S235_EG2_E3_bunit_cfm_qty = 7
Const S235_EG2_E2_spec = 8

Const C_SHEETMAXROWS_D  = 100

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	Dim StrNextKey
	Dim arrNextKey
    
	I4_s_item_sales_plan(S235_I4_sp_year) = Trim(Request("txtPlanFromYear"))
	I4_s_item_sales_plan(S235_I4_sp_month) = Trim(Request("txtPlanFromMonth"))

	If Request("txtPlanFromDt") <> "" Then
	    I1_s_item_sales_plan_by_plant = UNIConvDate(Request("txtPlanFromDt"))
	End IF    

	
	I5_s_item_sales_plan(S235_I5_sp_year) = Trim(Request("txtPlanToYear"))
	I5_s_item_sales_plan(S235_I5_sp_month) = Trim(Request("txtPlanToMonth"))

	If Request("txtPlanToDt") <> "" Then
	    I2_s_item_sales_plan_by_plant = UNIConvDate(Request("txtPlanToDt"))
	End IF

	I7_b_item = Trim(Request("txtItem_code"))
	I8_b_plant = Trim(Request("txtPlant_code"))

	If Trim(Request("lgStrPrevKey")) <> "" then
		arrNextKey = Split(Trim(Request("lgStrPrevKey")), gColSep)
		I3_s_item_sales_plan_by_plant = Trim(arrNextKey(0))
		I6_b_item = Trim(arrNextKey(1))
	End if
		
	CommandSent = "QUERY"
			
	Set PS2G134 = Server.CreateObject("PS2G134.cSListItemSPByPlant")
    
    Call PS2G134.S_LIST_ITEM_SALES_PLAN_BY_PLANT_SVR(gStrGlobalCollection,CommandSent,C_SHEETMAXROWS_D, _
            I1_s_item_sales_plan_by_plant,I2_s_item_sales_plan_by_plant, _
            I3_s_item_sales_plan_by_plant,I4_s_item_sales_plan, _
            I5_s_item_sales_plan,I6_b_item,I7_b_item,I8_b_plant, _
            E1_ief_supplied, E2_b_plant, E3_s_item_sales_plan_by_plant, _
            E4_b_item ,E5_b_item,E6_s_cfm_item_sales_plan,EG1_exp_plgrp, _
            EG2_exp_grp,prGroupView)
    
    If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "125000" then    
	
		If CheckSYSTEMError(Err,True) = True Then
		   prGroupView = -1
		   Set PS2G134 = Nothing
%>
<Script Language=vbscript>
	parent.frm1.txtPlant_code.focus
	parent.frm1.txtPlant_code_nm.value = ""
	
	parent.frm1.btnConfirm.disabled = True
	parent.SetToolBar("11000000000011")
</Script>
<%		   
		   Response.End
		End If   
	
	Elseif cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "229916" then
		
		If CheckSYSTEMError(Err,True) = True Then
		   prGroupView = -1
		   Set PS2G134 = Nothing
%>
<Script Language=vbscript>
	parent.frm1.txtItem_code.focus
	parent.frm1.txtItem_code_nm.value = ""
	
	parent.frm1.btnConfirm.disabled = True
	parent.SetToolBar("11000000000011")
</Script>
<%	
		   Response.End
		End If   
	Else 
		
		If CheckSYSTEMError(Err,True) = True Then
		   prGroupView = -1
		   Set PS2G134 = Nothing
%>
<Script Language=vbscript>
	parent.frm1.txtItem_code.focus
	parent.frm1.txtItem_code_nm.value = "<%=ConvSPChars(E5_b_item(S235_E5_item_nm))%>"
	parent.frm1.txtPlant_code_nm.value	= "<%=ConvSPChars(E2_b_plant(S235_E2_plant_nm))%>"
	parent.frm1.btnConfirm.disabled = True
	parent.SetToolBar("11000000000011")
</Script>
<%
			Response.End
		End If   
	End if	
	
	LngMaxRow = Request("txtMaxRows")
    
    GroupCount = prGroupView
	
	If EG2_exp_grp(GroupCount,S235_EG2_E3_sp_dt) = E3_s_item_sales_plan_by_plant(S235_E3_sp_dt) And _
		EG2_exp_grp(GroupCount,S235_EG2_E2_item_cd) = E4_b_item(S235_E4_item_cd) Then
		StrNextKey = ""
	Else
		StrNextKey = E3_s_item_sales_plan_by_plant(S235_E3_sp_dt) & gColSep & E4_b_item(S235_E4_item_cd)
	End If

%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	strTemp = CInt("<%=Request("strTemp")%>")
	strData = ""

	With parent																	

		.frm1.txtSpread.value			= ""
		.frm1.txtSpread.value			= "<%=ConvSPChars(E6_s_cfm_item_sales_plan(S235_E6_cfm_flag))%>"
		.frm1.txtItem_code_nm.value	    = "<%=ConvSPChars(E5_b_item(S235_E5_item_nm))%>"
		.frm1.txtPlant_code_nm.value	= "<%=ConvSPChars(E2_b_plant(S235_E2_plant_nm))%>"

		.frm1.txtPlan_buom_qty.value	= "<%=UNINumClientFormat(E1_ief_supplied(S235_E1_total_real), ggQty.DecPoint, 0)%>"
		
		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                
	
<%      	
		For LngRow = 0 To GroupCount		
%>
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(EG2_exp_grp(LngRow, S235_EG2_E3_sp_dt))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(EG2_exp_grp(LngRow, S235_EG2_E2_item_cd))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(EG2_exp_grp(LngRow, S235_EG2_E2_item_nm))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(EG2_exp_grp(LngRow, S235_EG2_E2_spec))%>"			
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(LngRow, S235_EG2_E3_cfm_qty), ggQty.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(LngRow, S235_EG2_E3_bunit_cfm_qty), ggQty.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(LngRow, S235_EG2_E3_plan_qty), ggQty.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(LngRow, S235_EG2_E3_plan_bunit_qty), ggQty.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(EG2_exp_grp(LngRow, S235_EG2_E3_req_sts))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>                                
			strData = strData & Chr(11) & Chr(12)
<%
	    Next
%>
		'2002-08-24소스 마감전 추가 
				
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip strData

	    .frm1.vspdData.ReDraw = False
<%
		For iRow = 0 To GroupCount
%>		 
			.ggoSpread.SSSetProtected .C_SPDT, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_ITEM_CD, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			'2002-11-09 추가 
			.ggoSpread.SSSetProtected .C_ITEM_POP, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			
			.ggoSpread.SSSetProtected .C_ITEM_NM, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_Spec, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>			
			.ggoSpread.SSSetProtected .C_PlanCfmQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>			
			.ggoSpread.SSSetProtected .C_PlanBunitCfmQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>

			If "<%=EG2_exp_grp(iRow, S235_EG2_E3_plan_qty)%>" = 0 Then
				.ggoSpread.SSSetProtected .C_PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetRequired .C_PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
				strTemp = strTemp + 1
			End If

			.ggoSpread.SSSetProtected .C_PlanBasicQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_ReqSts, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>

<%
		Next
%>
		
	    .frm1.vspdData.ReDraw = True
	    
	    If strTemp = 0 OR "<%=GroupCount%>" = -1 Then
			.frm1.btnConfirm.disabled = True
			.SetToolBar("11000000000111")
		Else
			
			.frm1.btnConfirm.disabled = False
			.SetToolBar("11101111000111")
		End If
	    
		.lgStrPrevKey = "<%=StrNextKey%>"
		
		'2002-09-02일 
		.frm1.strHTemp.value = strTemp
		
		.frm1.strHFromYear.value = "<%=ConvSPChars(Request("txtPlanFromYear"))%>"
		.frm1.strHFromMonth.value = "<%=ConvSPChars(Request("txtPlanFromMonth"))%>"
		.frm1.txtHPlanFromDt.value = "<%=ConvSPChars(Request("txtPlanFromDt"))%>"
		.frm1.strHToYear.value = "<%=ConvSPChars(Request("txtPlanToYear"))%>"
		.frm1.strHToMonth.value = "<%=ConvSPChars(Request("txtPlanToMonth"))%>"
		.frm1.txtHPlanToDt.value = "<%=ConvSPChars(Request("txtPlanToDt"))%>"
		.frm1.HItemCd.value = "<%=ConvSPChars(Request("txtItem_code"))%>"
		.frm1.HPlantCd.value = "<%=ConvSPChars(Request("txtPlant_code"))%>"
		.DbQueryOk

	End With
</Script>	
<%    
     
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
	
	I4_s_item_sales_plan(S235_I4_sp_year) = Trim(Request("strHFromYear"))
	I4_s_item_sales_plan(S235_I4_sp_month) = Trim(Request("strHFromMonth"))
	I5_s_item_sales_plan(S235_I5_sp_year) = Trim(Request("strHToYear"))
	I5_s_item_sales_plan(S235_I5_sp_month) = Trim(Request("strHToMonth"))
    
    I8_b_plant = UCase(Trim(Request("txtPlant_code")))
	
	RequestTxtSpread = Request("txtSpread")
				
    Set PS2G133 = Server.CreateObject("PS2G133.cSDailySP")
    
    Call PS2G133.S_MAINT_DAILY_SALES_PLAN(gStrGlobalCollection, _
            RequestTxtSpread,I4_s_item_sales_plan,I5_s_item_sales_plan,I8_b_plant,iErrorPosition)
    	    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
     Set PS2G133 = Nothing
     Response.End

	End If	
 	
    Set PS2G133 = Nothing                '☜: Unload Comproxy
    

%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%					

Case CStr(2580)


'***계획일 *******
	I4_s_item_sales_plan(S235_I4_sp_year) = Trim(Request("txtPlanFromYear"))
	I4_s_item_sales_plan(S235_I4_sp_month) = Trim(Request("txtPlanFromMonth"))

	If Request("txtPlanFromDt") <> "" Then
	    I1_s_item_sales_plan_by_plant = UNIConvDate(Request("txtPlanFromDt"))
	End IF    
	
	I5_s_item_sales_plan(S235_I5_sp_year) = Trim(Request("txtPlanToYear"))
	I5_s_item_sales_plan(S235_I5_sp_month) = Trim(Request("txtPlanToMonth"))

	If Request("txtPlanToDt") <> "" Then
	    I2_s_item_sales_plan_by_plant = UNIConvDate(Request("txtPlanToDt"))
	End IF
	
	I7_b_item = UCase(Trim(Request("txtItem_code")))
	I8_b_plant = UCase(Trim(Request("txtPlant_code")))
    
    Set PS2G136 = Server.CreateObject("PS2G136.cSConfirmItemSPByPt")
     
	Call PS2G136.S_CONFIRM_ITEM_SALES_PLAN_BY_PT(gStrGlobalCollection, _
            I1_s_item_sales_plan_by_plant ,I2_s_item_sales_plan_by_plant, _
            I4_s_item_sales_plan,I5_s_item_sales_plan,I7_b_item,I8_b_plant)	
	
	If CheckSYSTEMError(Err,True) = True Then
       Set PS2G136 = Nothing
       Response.End
    End If   
    Set PS2G136 = Nothing	    
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		.btnConfirm_Ok
        .DbSaveOk
	End With
</Script>
<%					
End Select
%>
