<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2112MB2
'*  4. Program Name         : 확정품목별 판매계획 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS2G131.dll, PS2G132.dll, PS2G135.dll
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Mr  Cho
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : 3rd Coding
'*                            -2001/01/03 : 5th Coding
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
Call HideStatusWnd                     

Dim PS2G132			'ComProxy Dll 사용 변수(query)
Dim PS2G131			'Save
Dim PS2G135			'Split
Dim iErrorPosition
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey										' 다음 값 
Dim lgStrPrevKey									' 이전 값 
Dim LngMaxRow										' 현재 그리드의 최대Row
Dim LngRow, iRow
Dim GroupCount          

Dim strCountryCD
Dim arrNextKey
'type MissMach
Dim RequestTxtSpread

Dim I1_ief_supplied
Dim I2_b_item
Dim I3_b_item
Dim I4_s_cfm_item_sales_plan

Dim E1_b_item
Dim E2_b_item
Dim EG1_exp_grp
Dim LG1_loc_grp
Dim prGroupView

Redim I4_s_cfm_item_sales_plan(1)

Const lsSPLIT  = "SPLIT"							'strMode 값:공장별배분작업 

Const S226_I4_sp_year = 0
Const S226_I4_sp_month = 1

Const S226_E1_item_cd = 0

Const S226_E2_item_cd = 0
Const S226_E2_item_nm = 1

Const S226_EG1_E1_s_cfm_item_sales_plan_plan_unit = 0
Const S226_EG1_E2_s_wks_msp_plan_qty1 = 1
Const S226_EG1_E2_s_wks_msp_plan_qty2 = 2
Const S226_EG1_E2_s_wks_msp_plan_qty3 = 3
Const S226_EG1_E2_s_wks_msp_plan_qty4 = 4
Const S226_EG1_E2_s_wks_msp_plan_qty5 = 5
Const S226_EG1_E2_s_wks_msp_plan_qty6 = 6
Const S226_EG1_E2_s_wks_msp_plan_qty7 = 7
Const S226_EG1_E2_s_wks_msp_plan_qty8 = 8
Const S226_EG1_E2_s_wks_msp_plan_qty9 = 9
Const S226_EG1_E2_s_wks_msp_plan_qty10 = 10
Const S226_EG1_E2_s_wks_msp_plan_qty11 = 11
Const S226_EG1_E2_s_wks_msp_plan_qty12 = 12
Const S226_EG1_E2_s_wks_msp_split_flag1 = 13
Const S226_EG1_E2_s_wks_msp_split_flag2 = 14
Const S226_EG1_E2_s_wks_msp_split_flag3 = 15
Const S226_EG1_E2_s_wks_msp_split_flag4 = 16
Const S226_EG1_E2_s_wks_msp_split_flag5 = 17
Const S226_EG1_E2_s_wks_msp_split_flag6 = 18
Const S226_EG1_E2_s_wks_msp_split_flag7 = 19
Const S226_EG1_E2_s_wks_msp_split_flag8 = 20
Const S226_EG1_E2_s_wks_msp_split_flag9 = 21
Const S226_EG1_E2_s_wks_msp_split_flag10 = 22
Const S226_EG1_E2_s_wks_msp_split_flag11 = 23
Const S226_EG1_E2_s_wks_msp_split_flag12 = 24
Const S226_EG1_E3_b_item_item_cd = 25
Const S226_EG1_E3_b_item_item_nm = 26
Const S226_EG1_E3_b_item_spec = 27

Const S226_LG1_L1_b_item_item_cd = 0
Const S226_LG1_L1_b_item_item_nm = 1

Const C_SHEETMAXROWS_D  = 100

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    
	lgStrPrevKey = Request("lgStrPrevKey")
        
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I3_b_item = Trim(Request("txtConItemCd"))	
	I4_s_cfm_item_sales_plan(S226_I4_sp_year) = Trim(Request("txtConSpYear"))
	I1_ief_supplied = Trim(Request("txtCfmFlag"))
	
	If lgStrPrevKey <> "" then
		I2_b_item = lgStrPrevKey
	End if

	Set PS2G132 = Server.CreateObject("PS2G132.cSListCfmItemSP")
	
	Call PS2G132.S_LIST_CFM_ITEM_SALES_PLAN(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_ief_supplied, I2_b_item, _
            I3_b_item, I4_s_cfm_item_sales_plan, E1_b_item, E2_b_item, EG1_exp_grp,LG1_loc_grp,prGroupView)
	'-----------------------
	'조건부 조건명 
	'-----------------------
	If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "229916" then    
		
			If CheckSYSTEMError(Err,True) = True Then
			    prGroupView = -1
				Set PS2G132 = Nothing
	%>
	<Script Language=vbscript>
		parent.frm1.txtConItemCd.focus
		parent.frm1.txtConItemNm.value = ""
		
	</Script>
	<%		   
			   Response.End
			End If 	
	Else 
			
			If CheckSYSTEMError(Err,True) = True Then
			   prGroupView = -1
			   Set PS2G132 = Nothing
	%>

	<Script Language=vbscript>
		parent.frm1.txtConSpYear.focus
		parent.frm1.txtConItemNm.Value = "<%=ConvSPChars(E2_b_item(S226_E2_item_nm))%>"
	</Script>

	<%
				Response.End
			End If   
	End if	
	

	LngMaxRow = Request("txtMaxRows")
    GroupCount = prGroupView
	
	If LG1_loc_grp(GroupCount,S226_LG1_L1_b_item_item_cd) = E1_b_item(S226_E1_item_cd) Then
		StrNextKey = ""
	Else
		StrNextKey = E1_b_item(S226_E1_item_cd)
    End If

%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow       
    Dim strTemp
    Dim strData

	<%'----------------------------------
	' 확정품목별판매계획 Grid의 내용을 표시한다.
	'---------------------------------- %>
	strData = ""

	With parent																	'☜: 화면 처리 ASP 를 지칭함 

		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                

<%      
		For LngRow = 0 To GroupCount
%>

			<% ' 품목 %>
			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S226_EG1_E3_b_item_item_cd))%>"
			strData = strData & Chr(11) & ""															
			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S226_EG1_E3_b_item_item_nm))%>"						
			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S226_EG1_E3_b_item_spec))%>"
			<% ' 계획단위 %>
			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S226_EG1_E1_s_cfm_item_sales_plan_plan_unit))%>"
			strData = strData & Chr(11) & ""
			<% ' 년계획 수량 합계 %>
			strData = strData & Chr(11) & ""
			<% ' 1월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty1), ggQty.DecPoint, 0)%>"
			<% ' 2월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty2), ggQty.DecPoint, 0)%>"
			<% ' 3월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty3), ggQty.DecPoint, 0)%>"
			<% ' 4월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty4), ggQty.DecPoint, 0)%>"
			<% ' 5월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty5), ggQty.DecPoint, 0)%>"
			<% ' 6월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty6), ggQty.DecPoint, 0)%>"
			<% ' 7월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty7), ggQty.DecPoint, 0)%>"
			<% ' 8월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty8), ggQty.DecPoint, 0)%>"
			<% ' 9월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty9), ggQty.DecPoint, 0)%>"
			<% ' 10월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty10), ggQty.DecPoint, 0)%>"
			<% ' 11월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty11), ggQty.DecPoint, 0)%>"
			<% ' 12월 계획량 %>
			strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_plan_qty12), ggQty.DecPoint, 0)%>"

			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag1)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag2)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag3)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag4)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag5)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag6)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag7)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag8)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag9)%>"					
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag10)%>"				
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag11)%>"				
			strData = strData & Chr(11) & "<%=EG1_exp_grp(LngRow,S226_EG1_E2_s_wks_msp_split_flag12)%>"				

			strData = strData & Chr(11) & LngMaxRow + <%=LngRow+1%>                                
			strData = strData & Chr(11) & Chr(12)
			
<%
	    Next
%>
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip strData

	    .frm1.vspdData.ReDraw = False
<%
		For iRow = 0 To GroupCount
%>		 
			.ggoSpread.SSSetProtected .C_ItemCode, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_ItemName, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_Spec, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_PlanUnit, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			.ggoSpread.SSSetProtected .C_PlanUnitPopup, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>			
			.ggoSpread.SSSetProtected .C_YearQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag1)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_01PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_01PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag2)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_02PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_02PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag3)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_03PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_03PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag4)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_04PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_04PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>				
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag5)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_05PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_05PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>				
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag6)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_06PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_06PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag7)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_07PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_07PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag8)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_08PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_08PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag9)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_09PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_09PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag10)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_10PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_10PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag11)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_11PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_11PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			End If

			If UCase("<%=EG1_exp_grp(iRow,S226_EG1_E2_s_wks_msp_split_flag12)%>") = "N" Then
				.ggoSpread.SSSetRequired .C_12PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>
			Else
				.ggoSpread.SSSetProtected .C_12PlanQty, LngMaxRow + <%=iRow+1%>, LngMaxRow + <%=iRow+1%>				
			End If

<%
		Next
%>
	    
	    .frm1.vspdData.ReDraw = True
		
		.lgStrPrevKey = "<%=StrNextKey%>"
    
		.frm1.txtConItemNm.Value = "<%=ConvSPChars(E2_b_item(S226_E2_item_nm))%>"
		.frm1.HItemCd.value = "<%=ConvSPChars(Request("txtConItemCd"))%>"
		.frm1.HConSpYear.value = "<%=ConvSPChars(Request("txtConSpYear"))%>"

		.DbQueryOk
		
	    
	End With
</Script>	
<%    
      
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
			
	I4_s_cfm_item_sales_plan(S226_I4_sp_year) = Trim(Request("HConSpYear"))
    
    RequestTxtSpread = Request("txtSpread")
    
    Set PS2G131 = Server.CreateObject("PS2G131.cSCfmItemSP")
        
	Call PS2G131.S_CFM_ITEM_SALES_PLAN(gStrGlobalCollection,"SAVE",RequestTxtSpread,I4_s_cfm_item_sales_plan,,iErrorPosition)
	    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		
		Set PS2G131 = Nothing
		Response.End
     
	End If	
 	
    Set PS2G131 = Nothing                '☜: Unload Comproxy

%>
<Script Language=vbscript>
	With parent					'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%					

Case CStr(lsSPLIT)		'공장별배분작업 
	
	'Dim cntItem, arrItemTemp
	'arrItemTemp = Split(Request("txtItemArrary"), gColSep)

	'For cntItem = 1 To Request("txtItemCount")
	'	Response.Write "ItemBItemItemCd" & cntItem & "=>" & UCASE(TRIM(arrItemTemp(cntItem-1))) &"<BR>"
	'Next
		
    'Response.Write "lsPlanMonth=>"&TRIM(Request("lsPlanMonth")) &"<BR>"
    'Response.Write "lsPlanUnit=>"&TRIM(Request("lsPlanUnit")) &"<BR>"
    'Response.Write "HConSpYear=>"&TRIM(Request("HConSpYear")) &"<BR>"
	'Response.Write "HItemCd=>"&TRIM(Request("HItemCd")) &"<BR>"
        
    'Response.End
    
    I4_s_cfm_item_sales_plan(S226_I4_sp_year) = Trim(Request("HConSpYear"))
    I4_s_cfm_item_sales_plan(S226_I4_sp_month) = Trim(Request("lsPlanMonth"))
    
    I3_b_item = Trim(Request("HItemCd"))
    
    Set PS2G135 = Server.CreateObject("PS2G135.cSSplitItemSP")
        
	Call PS2G135.S_SPLIT_ITEM_SALES_PLAN(gStrGlobalCollection,I3_b_item,I4_s_cfm_item_sales_plan)
	    
    If CheckSYSTEMError(Err,True) = True Then
       Set PS2G135 = Nothing
       Response.End
       'Exit Sub
    End If   
 	
    Set PS2G135 = Nothing                '☜: Unload Comproxy
	
%>
<Script Language=vbscript>
	With parent																'☜: 화면 처리 ASP 를 지칭함 
		'.btnSplit_Ok
		.DbSaveOk
	End With
</Script>
<%		

End Select

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>

