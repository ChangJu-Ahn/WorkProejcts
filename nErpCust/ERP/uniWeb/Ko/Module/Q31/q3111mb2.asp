<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q3111MB2
'*  4. Program Name         : X-Bar & R 관리도 
'*  5. Program Desc         : 평균 및 범위에 대한 관리도 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- ChartFX용 상수를 사용하기 위한 Include 지정 -->
<!-- #include file="../../inc/CfxIE.inc" -->

<Script Language=vbscript>
	Dim strVar1
	Dim strVar2
	Dim strVar3
	Dim strVar4
	Dim strVar5
	Dim strVar6
	Dim strVar7
	

	Dim	TempstrPlantCd
	Dim TempstrItemCd
	Dim TempstrInspItemCd
	
	TempstrPlantCd		= "<%=Request("txtPlantCd")%>"
	TempstrItemCd		= "<%=Request("txtItemCd")%>"
	TempstrInspItemCd	= "<%=Request("txtInspItemCd")%>"	
	
	'공장명 불러오기 
	Call parent.CommonQueryRs("PLANT_CD,PLANT_NM","B_PLANT","PLANT_CD =  " & parent.FilterVar(TempstrPlantCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtPlantNm.Value = strVar2
	
	'품목명 불러오기 
	Call parent.CommonQueryRs("ITEM_CD,ITEM_NM","B_ITEM","ITEM_CD =  " & parent.FilterVar(TempstrItemCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtItemNm.Value = strVar2	
	
	'검사항목명 불러오기 
	Call parent.CommonQueryRs("INSP_ITEM_CD,INSP_ITEM_NM","Q_INSPECTION_ITEM","INSP_ITEM_CD =  " & parent.FilterVar(TempstrInspItemCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtInspItemNm.Value = strVar2
</Script>
<%													
On Error Resume Next

'Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")
	
Dim Conn
	
Dim strPlantCd
Dim strInspItemCd
Dim strInspClassCd
Dim strYrDt1
Dim strYrDt2
Dim strItemCd
	
Dim lgdblData() 
		
Dim lgXbar()
Dim lgXbarbar
Dim lgR()
Dim lgRbar
	
Dim lgMaxXbar
Dim lgMinXbar
Dim lgMaxR
Dim lgMinR
		
Dim lglngNumberOfData
Dim lgintSizeOfSubgroup
Dim lglngNumberOfSubgroup
	
'검사규격 
Dim lgstrInspSpec
Dim lgdblLSL
Dim lgdblUSL
Dim lgMsmtUnitCd
	
'계수 
Dim lgdblA2
Dim lgdblD3
Dim lgdblD4
	
'Xbar 관리한계 
'/* SCR 213: 관리한계 계산이 틀림 관련 - START */
Dim lgstrMthdOfCL
Dim lgintCntOfSubGroupForCL 
'/* SCR 213: 관리한계 계산이 틀림 관련 - END */
Dim lgstrLCL
Dim lgstrUCL
	
Dim lgdblCL
Dim lgdblLCL
Dim lgdblUCL
		
'범위 관리한계 
Dim lgdblR_CL
Dim lgdblR_LCL
Dim lgdblR_UCL
	
Dim lgblnRet
Dim i
		
Dim lgintDecimal
	
Dim strMark
	
'Request
lgblnRet = Request_QueryData
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If
	
'데이타 얻기 
lgblnRet = Get_Data
If lgblnRet = False Then 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If
	
'Subgroup수 구하기 
lgblnRet = CalForNumOfSubgroup
If lgblnRet = False Then 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If
    
'군별 평균 및 군들의 평균 구하기 
lgblnRet = CalForAvgOfSubgroup
If lgblnRet = False Then 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If
	
'군 개개의 범위 및 범위들의 평균 구하기 
lgblnRet = CalForRangeOfSubgroup
If lgblnRet = False Then 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If
    	
'Xbar의 관리한계 구하기 
lgblnRet = CalForControlLimit
If lgblnRet = False Then 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If
       
'R의 관리한계 구하기 
lgblnRet = CalForR_ControlLimit
If lgblnRet = False Then 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If
%>
<Script Language=vbscript>
Dim lgblnRet
Dim lgOKFlag
	
lgOKFlag = True
	
<%'----------------------------------------------%>
<%'기준 DATA DISPLAY %>
lgblnRet = Display_InspStand
If lgblnRet = False Then lgOKFlag = False
	
<%'-------------------- SPREAD --------------------------%>
<%'SPREAD에 DATA DISPLAY %>
lgblnRet = DisplayData_OnSpread
If lgblnRet = False Then lgOKFlag = False
	
    	
If lgOKFlag = True Then 
	Call Parent.DbQueryOk
End If
    	
<%'################################################################################################################
'############################################ CLIENT SIDE FUNCTION ##############################################
'################################################################################################################%>

<%'/*****************************************************
'/	기준 데이타 Display
'/*****************************************************%>
Function Display_InspStand()
	Display_InspStand = False
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1
		.txtInspSpec.Value = "<%=UniNumClientFormat(lgstrInspSpec, lgintDecimal, 0)%>"
		.txtLSL.Value = "<%=UniNumClientFormat(lgdblLSL, lgintDecimal ,0)%>"
		.txtUSL.Value = "<%=UniNumClientFormat(lgdblUSL, lgintDecimal ,0)%>"
		.txtMeasmtUnitCd.Value = "<%=lgMsmtUnitCd%>"
	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	Display_InspStand = True
	
End Function
<%'/*****************************************************
'/	Spread에 데이타 Display
'/*****************************************************%>
Function DisplayData_OnSpread()
	
	DisplayData_OnSpread = False
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1
		.vspdData.MaxCols = <%=lglngNumberOfSubgroup%>				<%'스프레드의 칼럼수 설정 %>

		Parent.ggoSpread.Source = .vspdData
<%
		For i = 0 To lglngNumberOfSubgroup - 1					'스프레드 헤더 보여주기 
			If i = 0 then
				strMark = "st"
			Elseif i = 1 then
				strMark = "nd"
			Elseif i = 2 then
				strMark = "rd"
			Else
				strMark = "th"
			End If
%>
			Parent.ggoSpread.SSSetEdit (<%=i%> + 1), "<%=CStr(i + 1) & strMark%>", 8, 1, -1, 15 
			.vspdData.Row = 1							<%'스프레드에 평균과 범위 넣어주기 %>
			.vspdData.Text = "<%=UNINumClientFormat(lgXBar(i), lgintDecimal, 0)%>"
			.vspdData.Row = 2
			.vspdData.Text = "<%=UNINumClientFormat(lgR(i), lgintDecimal, 0)%>"
<%
		Next
%>
	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	DisplayData_OnSpread = True
	
End Function
<%'/*****************************************************
'/	ChartFX1(Xbar Chart)의 환경 설정 
'/*****************************************************%>
Function Setting_ChartFX1()
	Dim sngTempMin
	Dim sngTempMax 
	Dim sngTempDiffStep
	
	Setting_ChartFX1 = False
	
	Err.Clear
	On Error Resume Next
	With Parent.frm1.ChartFX1
		<%'ToolBar 속성 %>
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>						<%'틀바를 새로운 창으로 보이기 %>
		.ToolBarObj.Left = 15								<%'틀바의 왼쪽 위치 %> 
		.ToolBarObj.Top = 10								<%'틀바의 상단 위치 %> 
		
		<%'Y축 표시값(개수)의 소수점 이하 자리수 설정 %>
		.Axis(<%=AXIS_Y%>).Decimals = <%=lgintDecimal%>
		
		'y축의 Min/Max값 및 Step 구하기 
		If <%= (lgMinXBar > lgdblLCL) %> Then
			sngTempMin = parent.UNICDbl("<%=UNINumClientFormat(lgdblLCL, lgintDecimal, 0)%>")
		Else
			sngTempMin = parent.UNICDbl("<%=UNINumClientFormat(lgMinXBar, lgintDecimal, 0)%>")
		End If
	
		If <%=(lgMaxXBar < lgdblUCL)%> Then
			sngTempMax = parent.UNICDbl("<%=UNINumClientFormat(lgdblUCL, lgintDecimal, 0)%>")
		Else
			'군별 평균의 최대값보다 1퍼센트 큰 값을 Y축의 최소값으로 설정 
			sngTempMax = parent.UNICDbl("<%=UNINumClientFormat(lgMaxXBar, lgintDecimal, 0)%>")
		End If
		sngTempDiffStep = (sngTempMax - sngTempMin) / 10
		.Axis(<%=AXIS_Y%>).Min = sngTempMin - sngTempDiffStep 
		.Axis(<%=AXIS_Y%>).Max = sngTempMax + sngTempDiffStep 
	End With
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Setting_ChartFX1 = True
	
End Function

<%'/*****************************************************
'/	ChartFX2(R Chart)의 환경 설정 
'/*****************************************************%>
Function Setting_ChartFX2()
	Dim sngTempMin
	Dim sngTempMax 
	Dim sngTempDiffStep
	
	Setting_ChartFX2 = False
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1.ChartFX2
		<%'ToolBar 속성 %>
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>						<%'틀바를 새로운 창으로 보이기 %>
		.ToolBarObj.Left = 15								<%'틀바의 왼쪽 위치 %> 
		.ToolBarObj.Top = 10								<%'틀바의 상단 위치 %> 
		
		<%'Y축 표시값(개수)의 소수점 이하 자리수 설정 %>
		.Axis(<%=AXIS_Y%>).Decimals = <%=lgintDecimal%>		
		
		'y축의 Min/Max값 및 Step 구하기 
		If <%= (lgMinR > lgdblR_LCL)%> Then
			sngTempMin = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_LCL, lgintDecimal, 0)%>")
		Else
			sngTempMin = parent.UNICDbl("<%=UNINumClientFormat(lgMinR, lgintDecimal, 0)%>")
		End If
    
		If <%= (lgMaxR < lgdblR_UCL)%> Then
			sngTempMax = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
		Else
			sngTempMax = parent.UNICDbl("<%=UNINumClientFormat(lgMaxR, lgintDecimal, 0)%>")
		End If
		
		sngTempDiffStep = (sngTempMax - sngTempMin) / 10
		.Axis(<%=AXIS_Y%>).Min = sngTempMin - sngTempDiffStep 
		If .Axis(<%=AXIS_Y%>).Min < 0 Then
			.Axis(<%=AXIS_Y%>).Min = 0
		End If
		.Axis(<%=AXIS_Y%>).Max = sngTempMax + sngTempDiffStep 
		.Axis(<%=AXIS_Y%>).STEP = (.Axis(<%=AXIS_Y%>).Max - .Axis(<%=AXIS_Y%>).Min) / 5		'Y축 (Max - Min) / 10으로 설정 
	End With
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Setting_ChartFX2 = True
End Function

<%'/*****************************************************
'/	ChartFX1(Xbar Chart) 그리기 
'/*****************************************************%>
Function Draw_XbarChart()
	
	Draw_XbarChart = False
	
	Err.Clear
	On Error Resume Next
	
'	With Parent.frm1.ChartFX1
		
'		.OpenDataEx COD_VALUES, 1, <%=lglngNumberOfSubgroup%>				'차트 FX와의 데이터 채널 열어주기 
'			'첫번째 계열(Xbar) 값 설정 
'			.Series(0).MarkerShape = <%=MK_CIRCLE%>
'			.Series(0).LineStyle = <%=CHART_SOLID%>
<%
		Dim YValue0, sInsSQL
		Dim blnRet
	    'DB 연결 
	    blnRet = DBConnect
						
		sInsSQL = "DELETE FROM Q_TMP_CHART_XBAR_ANALYSIS"
		Conn.Execute sInsSQL
%>
	
<%
			For i = 0 to lglngNumberOfSubgroup - 1
				If i = 0 then
					strMark = "st"
				Elseif i = 1 then
					strMark = "nd"
				Elseif i = 2 then
					strMark = "rd"
				Else
					strMark = "th"
				End If
%>
'				.Legend(<%=i%>) = "<%=CStr(i+1) & strMark %>"
'				.ValueEx(0, <%=i%>) =  parent.UNICDbl("<%=UNINumClientFormat(lgXBar(i), lgintDecimal, 0)%>")
<%
				YValue0 = UNINumClientFormat(lgXBar(i), lgintDecimal, 0)
				
				sInsSQL =			" INSERT INTO Q_TMP_CHART_XBAR_ANALYSIS (XVALUE, YVALUE, X_CL, X_UCL, X_LCL) "
				sInsSQL = sInsSQL & " VALUES ( " & FilterVar((i+1) & strMark,"","S") & "," 
				sInsSQL = sInsSQL & 		   FilterVar(YValue0, "", "S") & ", "
				sInsSQL = sInsSQL & 		   FilterVar(UNINumClientFormat(lgdblCL, lgintDecimal, 0), "", "S") & ", "
				sInsSQL = sInsSQL & 		   FilterVar(UNINumClientFormat(lgdblUCL, lgintDecimal, 0), "", "S") & ", "
				sInsSQL = sInsSQL & 		   FilterVar(UNINumClientFormat(lgdblLCL, lgintDecimal, 0), "", "S") & ") "

				Conn.Execute sInsSQL
				
%>
<%
			Next
%>
'		.CloseData COD_VALUES
		
'		'UCL, LCL, CL을 위한 Constant line(s)
'		.OpenDataEx <%=COD_CONSTANTS%>, 3, 0 					
'			.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblUCL, lgintDecimal, 0)%>")
'			.ConstantLine(0).Axis = <%=AXIS_Y%>
'			.ConstantLine(0).Label = "UCL = " & "<%=UNINumClientFormat(lgdblUCL, lgintDecimal, 0)%>"
'			.ConstantLine(0).LineColor = RGB(255, 0, 0)
'			.ConstantLine(1).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblLCL, lgintDecimal, 0)%>")
'			.ConstantLine(1).Axis = <%=AXIS_Y%>
'			.ConstantLine(1).Label = "LCL = " & "<%=UNINumClientFormat(lgdblLCL, lgintDecimal, 0)%>"
'			.ConstantLine(1).LineColor = RGB(255, 0, 0)
'			.ConstantLine(2).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblCL, lgintDecimal, 0)%>")
'			.ConstantLine(2).Axis = <%=AXIS_Y%>
'			.ConstantLine(2).Label = "CL = " & "<%=UNINumClientFormat(lgdblCL, lgintDecimal, 0)%>"
'			.ConstantLine(2).LineColor = RGB(0, 0, 0)
'		.CloseData <%=COD_CONSTANTS%>					'차트 FX와의 채널(Constant Line을 위한) 닫아주기 
		
'		.OpenDataEx <%=COD_STRIPES%>, 2, 0 					' Two Color stripes
'			.Stripe(0).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblUCL, lgintDecimal, 0)%>")
'			.Stripe(0).To = .Axis(<%=AXIS_Y%>).Max
'			.Stripe(0).Color = RGB(255, 100, 255)
'			.Stripe(0).Axis = <%=AXIS_Y%>
'			.Stripe(1).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblLCL, lgintDecimal, 0)%>")
'			.Stripe(1).To = .Axis(<%=AXIS_Y%>).Min
'			.Stripe(1).Color = RGB(255, 100, 255)
'			.Stripe(1).Axis = <%=AXIS_Y%>
'		.CloseData <%=COD_STRIPES%>					'차트 FX와의 채널(줄무늬를 위한) 닫아주기 
		
'		.Axis(<%=AXIS_X%>).Visible = True 
'		.Axis(<%=AXIS_Y%>).Visible = True 
'		.Series(0).Visible = True 
'	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	Draw_XbarChart = True
	
End Function		

<%'/*****************************************************
'/	ChartFX2(R Chart) 그리기 
'/*****************************************************%>
Function Draw_RChart()
	Draw_RChart = False
	
	Err.Clear
	On Error Resume Next
	
'	With Parent.frm1.ChartFX2
		
'		.OpenDataEx COD_VALUES, 1, <%=lglngNumberOfSubgroup%>				'차트 FX와의 데이터 채널 열어주기 
'			'첫번째 계열(Xbar) 값 설정 
'			.Series(0).MarkerShape = <%=MK_CIRCLE%>
'			.Series(0).LineStyle = <%=CHART_SOLID%>
<%
		
		sInsSQL = "DELETE FROM Q_TMP_CHART_R_ANALYSIS"
		Conn.Execute sInsSQL
%>

<%
			For i = 0 to lglngNumberOfSubgroup - 1
				If i = 0 then
					strMark = "st"
				Elseif i = 1 then
					strMark = "nd"
				Elseif i = 2 then
					strMark = "rd"
				Else
					strMark = "th"
				End If
%>
'				.Legend(<%=i%>) = "<%=CStr(i+1) & strMark %>"
'				.ValueEx(0, <%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgR(i), lgintDecimal, 0)%>")
<%
				YValue0 = UNINumClientFormat(lgR(i), lgintDecimal, 0)
				
				sInsSQL =			" INSERT INTO Q_TMP_CHART_R_ANALYSIS (XVALUE, YVALUE, R_CL, R_UCL, R_LCL) "
				sInsSQL = sInsSQL & " VALUES ( " & FilterVar((i+1) & strMark,"","S") & "," 
				sInsSQL = sInsSQL & 		   FilterVar(YValue0, "", "S") & ", "
				sInsSQL = sInsSQL & 		   FilterVar(UNINumClientFormat(lgdblR_CL, lgintDecimal, 0), "", "S") & ", "
				sInsSQL = sInsSQL & 		   FilterVar(UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0), "", "S") & ", "
				sInsSQL = sInsSQL & 		   FilterVar(UNINumClientFormat(lgdblR_LCL, lgintDecimal, 0), "", "S") & ") "

				Conn.Execute sInsSQL
						
			Next
%>
'		.CloseData <%=COD_VALUES%>
		
'		'UCL, LCL, CL을 위한 Constant line(s)
'		If lgdblR_LCL = 0 Then
'			.OpenDataEx <%=COD_CONSTANTS%>, 2, 0 					
'				.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
'				.ConstantLine(0).Axis = <%=AXIS_Y%>
'				.ConstantLine(0).Label = "UCL = " &  "<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>"
'				.ConstantLine(0).LineColor = RGB(255, 0, 0)
'				.ConstantLine(1).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_CL, lgintDecimal, 0)%>")
'				.ConstantLine(1).Axis = <%=AXIS_Y%>
'				.ConstantLine(1).Label = "CL = " &  "<%=UNINumClientFormat(lgdblR_CL, lgintDecimal, 0)%>"
'				.ConstantLine(1).LineColor = RGB(0, 0, 0)
'			.CloseData <%=COD_CONSTANTS%>					'차트 FX와의 채널(Constant Line을 위한) 닫아주기 
			
'			.OpenDataEx <%=COD_STRIPES%>, 1, 0 					' Two Color stripes
'				.Stripe(0).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
'				.Stripe(0).To = .Axis(<%=AXIS_Y%>).Max
'				.Stripe(0).Color = RGB(255, 100, 255)
'				.Stripe(0).Axis = <%=AXIS_Y%>
'			.CloseData <%=COD_STRIPES%>					'차트 FX와의 채널(줄무늬를 위한) 닫아주기 
'		Else
'			.OpenDataEx <%=COD_CONSTANTS%>, 3, 0 					
'				.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
'				.ConstantLine(0).Axis = <%=AXIS_Y%>
'				.ConstantLine(0).Label = "UCL = " &  "<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>"
'				.ConstantLine(0).LineColor = RGB(255, 0, 0)
'				.ConstantLine(1).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_LCL, lgintDecimal, 0)%>")
'				.ConstantLine(1).Axis = <%=AXIS_Y%>
'				.ConstantLine(1).Label = "LCL = " & "<%=UNINumClientFormat(lgdblR_LCL, lgintDecimal, 0)%>"
'				.ConstantLine(1).LineColor = RGB(255, 0, 0)
'				.ConstantLine(2).Value = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_CL, lgintDecimal, 0)%>")
'				.ConstantLine(2).Axis = <%=AXIS_Y%>
'				.ConstantLine(2).Label = "CL = " & "<%=UNINumClientFormat(lgdblR_CL, lgintDecimal, 0)%>"
'				.ConstantLine(2).LineColor = RGB(0, 0, 0)
'			.CloseData <%=COD_CONSTANTS%>					'차트 FX와의 채널(Constant Line을 위한) 닫아주기 
			
'			.OpenDataEx <%=COD_STRIPES%>, 2, 0 					' Two Color stripes
'				.Stripe(0).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
'				.Stripe(0).To = .Axis(<%=AXIS_Y%>).Max
'				.Stripe(0).Color = RGB(255, 100, 255)
'				.Stripe(0).Axis = <%=AXIS_Y%>
'				.Stripe(1).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_LCL, lgintDecimal, 0)%>")
'				.Stripe(1).To = .Axis(<%=AXIS_Y%>).Min
'				.Stripe(1).Color = RGB(255, 100, 255)
'				.Stripe(1).Axis = <%=AXIS_Y%>
'			.CloseData <%=COD_STRIPES%>					'차트 FX와의 채널(줄무늬를 위한) 닫아주기 
'		End If
		
'		.Axis(<%=AXIS_X%>).Visible = True 
'		.Axis(<%=AXIS_Y%>).Visible = True 
'		.Series(0).Visible = True 
'	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Draw_RChart = True
End Function	
    
</Script>   
<%
'################################################################################################################
'############################################ SERVER SIDE FUNCTION ##############################################
'################################################################################################################

'/*****************************************************
'/ 입력 데이타 얻기 
'/*****************************************************
Function Request_QueryData()
	Request_QueryData = False
	strPlantCd  = Request("txtPlantCd")
	strInspClassCd = Request("cboInspClassCd")
	strYrDt1= UNIConvDate(Request("txtYrDt1"))
	strYrDt2= UNIConvDate(Request("txtYrDt2"))
	strItemCd = Request("txtItemCd")
	strInspItemCd = Request("txtInspItemCd")
	If Trim(Request("txtPartSampleCnt")) = "" Then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건 값이 비었습니다 
		Exit Function
	End If
	lgintSizeOfSubgroup = CInt(Request("txtPartSampleCnt"))
	
	If strPlantCd="" or strInspClassCd = "" or strYrDt1="" or strYrDt2="" or strItemCd="" or strInspItemCd="" then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건 값이 비었습니다 
		Exit Function
	End IF
	
	If lgintSizeOfSubgroup < 2 or lgintSizeOfSubgroup > 20 Then								'군별 시료수는 2에서 25사이의 수를 이용하게 되어 있다.
		'아래는 지나치게 적거나 많은 군별 시료수를 입력했을 경우를 위해 임의로 준 메시지이다.
		Call DisplayMsgBox("229907", vbOKOnly, "", "", I_MKSCRIPT)	'적절한 군별 시료수를 입력하십시오 
		Exit Function
	End IF
	Request_QueryData = True
End Function

'/*****************************************************
'/ 조회 데이타 얻기 
'/*****************************************************
Function Get_Data()
    Dim i
    Dim blnRet
    
    Get_Data = False
    
    'DB 연결 
    blnRet = DBConnect
    If blnRet = False Then Exit Function
    
    '소수 자릿수 얻기 
    blnRet = Get_Decimal
    If blnRet = False Then Exit Function
    
	'Check Input Data
    blnRet = Check_InputData
    If blnRet = False Then Exit Function    
    
    '검사기준 정보 얻기 
    blnRet = Get_InspStandard
    If blnRet = False Then Exit Function
    
    '측정치 얻기 
    blnRet = Get_MeasuredValues
    If blnRet = False Then Exit Function
    
    '계수값 얻기 
    blnRet = Get_Parameters
    If blnRet = False Then Exit Function
    
    'DB 연결 끊기 
    blnRet = DBClose
    If blnRet = False Then Exit Function
    
    Get_Data = True
End Function

'/*****************************************************
'/ Database 연결 
'/*****************************************************
Function DBConnect()
	DBConnect = False
	
	'Object 생성 
	Set Conn = Server.CreateObject("ADODB.Connection")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	
		Exit Function					
	End If
	
	
	' ODBC Data source 열기 
	With Conn
		.ConnectionString  = gADODBConnString		
		.ConnectionTimeout = 180
		
		.Open
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
			Conn.Close
			Set Conn = Nothing
			Exit Function	
		End If
	End With
	
	DBConnect = True
End Function


'/*****************************************************
'/ Database 연결 끊기 
'/*****************************************************
Function DBClose()
	DBClose = False
	
	Err.Clear
	On Error Resume Next
	
	Conn.Close
	Set Conn = Nothing		
	
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function
	End If
	
	DBClose = True
End Function

'/*****************************************************
'/ 입력 데이타 체크 
'/*****************************************************
Function Check_InputData()
	Dim RS
	Dim strSql
	Check_InputData = False
	
	Err.Clear
	On Error Resume Next
	
	
            
	Set RS = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	'공장 체크 
	If strPlantCd <> "" Then
		strSql = "SELECT PLANT_CD " &_
				"FROM B_PLANT " &_
				"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S")
        
        RS.Open  strSql, Conn, 1			'adOpenKeyset
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			If CheckSYSTEMError(Err,True) = False Then
		       Call CheckSQLError(Conn,True)
		    End If
		    RS.Close
		    Set RS = Nothing											'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing
			Exit Function
		End If
		
		'레코드가 하나도 없다면 
		If RS.EOF or RS.BOF then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)	'공장이 존재하지 않습니다.
			RS.Close
			Set RS = Nothing												'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'☜: ComProxy Unload
			Exit Function
		End If
	RS.Close
	End if
	
	'품목 체크 
	If strItemCd <> "" Then
		strSql = "SELECT ITEM_CD " &_
				"FROM B_ITEM_BY_PLANT " &_
				"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S")

        RS.Open  strSql, Conn, 1			'adOpenKeyset
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			If CheckSYSTEMError(Err,True) = False Then
		       Call CheckSQLError(Conn,True)
		    End If
		    RS.Close
		    Set RS = Nothing											'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing
			Exit Function
		End If
		
		'레코드가 하나도 없다면 
		If RS.EOF or RS.BOF then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)	'공장이 존재하지 않습니다.
			RS.Close
			Set RS = Nothing												'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'☜: ComProxy Unload
			Exit Function
		End If
	RS.Close
	End if

	'검사항목 체크 
	If strInspItemCd <> "" Then
		strSql = "SELECT INSP_ITEM_CD " &_
				"FROM Q_INSPECTION_STANDARD_BY_ITEM " &_
				"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") &_
				" AND INSP_ITEM_CD = " & FilterVar(strInspItemCd, "''", "S")
        
        RS.Open  strSql, Conn, 1			'adOpenKeyset
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			If CheckSYSTEMError(Err,True) = False Then
		       Call CheckSQLError(Conn,True)
		    End If
		    RS.Close
		    Set RS = Nothing											'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing
			Exit Function
		End If
		
		'레코드가 하나도 없다면 
		If RS.EOF or RS.BOF then
			Call DisplayMsgBox("220201", vbOKOnly, "", "", I_MKSCRIPT)	'공장이 존재하지 않습니다.
			RS.Close
			Set RS = Nothing												'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'☜: ComProxy Unload
			Exit Function
		End If
	
	RS.Close
	End if
	
	Set RS = Nothing	
	
	Check_InputData = True
End Function


'/*****************************************************
'/ 검사기준 데이타 얻기 
'/*****************************************************
Function Get_InspStandard()
	Dim RS
	Dim strSql
	Get_InspStandard = False
	
	Err.Clear
	On Error Resume Next
	
	'/* SCR 213: 관리한계 계산이 틀림 관련 - START */
	strSql = "SELECT INSP_SPEC, LSL, USL, MTHD_OF_CL_CAL, CALCULATED_QTY, LCL, UCL, MEASMT_UNIT_CD, INSP_UNIT_INDCTN " &_
              "FROM Q_INSPECTION_STANDARD_BY_ITEM " &_
              "WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & "" &_
              "AND INSP_CLASS_CD = " & FilterVar(strInspClassCd, "''", "S") & "" &_
              "AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") & "" &_
              "AND INSP_ITEM_CD = " & FilterVar(strInspItemCd, "''", "S") & ""
    '/* SCR 213: 관리한계 계산이 틀림 관련 - END */          
	Set RS = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	RS.Open  strSql, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
           Call CheckSQLError(Conn,True)
        End If
		'Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Set RS = Nothing												'☜: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	'레코드가 하나도 없다면 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	   '조건에 맞는 검사결과가 없습니다 
		RS.Close
		Set RS = Nothing												'☜: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		'아래는 임의로 준 메시지 
		Exit Function
	End If
	
	If Trim(RS(8)) <> "3" Then
		Call DisplayMsgBox("229939", vbOKOnly, "", "", I_MKSCRIPT)	'검사단위품질표시가 특성치가 아닙니다.
		Exit Function
	End If
	
	If Trim(RS(0)) = "" Then
		Call DisplayMsgBox("220706", vbOKOnly, "", "", I_MKSCRIPT)	'검사규격이 입력되어 있지 않습니다 
		Exit Function
	End If
	
	lgstrInspSpec = RS(0)
	
	If Trim(RS(1)) = "" AND Trim(RS(2)) = "" Then
		Call DisplayMsgBox("229911", vbOKOnly, "", "", I_MKSCRIPT)	'상한/하한규격 중 적어도 하나는 존재해야 합니다 
		Exit Function
	End If

	lgdblLSL = UNICDbl(RS(1), 0)
	lgdblUSL = UNICDbl(RS(2), 0)
	
	'/* SCR 213: 관리한계 계산이 틀림 관련 - START */
	lgstrMthdOfCL = UCase(Trim(RS(3)))
	SELECT CASE lgstrMthdOfCL
		CASE "S"	'부분계산 
			lgintCntOfSubGroupForCL = RS(4)
			lgstrLCL = ""
			lgstrUCL = ""
		CASE "T"	'목표치 
			lgstrLCL = Trim(RS(5))
			lgstrUCL = Trim(RS(6))
		CASE ELSE	'전체계산: "C" Or ""
			lgstrLCL = ""
			lgstrUCL = ""
	END SELECT
		
	lgMsmtUnitCd = Trim(RS(7))
	'/* SCR 213: 관리한계 계산이 틀림 관련 - END */
	RS.Close
	Set RS = Nothing
	
	Get_InspStandard = True
End Function

'/*****************************************************
'/ 측정치 데이타 얻기 
'/*****************************************************
Function Get_MeasuredValues()
	Dim RS
	Dim strSql
	Get_MeasuredValues = False
	
	Err.Clear
	On Error Resume Next
	
	strSql = "SELECT A.MEAS_VALUE" &_
		" FROM (Q_Inspection_Measured_Values  A LEFT OUTER JOIN Q_Inspection_Details B" &_
	     	" ON A.Insp_Req_No = B.Insp_Req_No AND A.Insp_Result_No = B.Insp_Result_No" &_
	     	" AND A.INSP_ITEM_CD = B.INSP_ITEM_CD AND A.INSP_SERIES = B.INSP_SERIES)" &_
	     	" LEFT OUTER JOIN Q_Inspection_Result C " &_
	     	" ON A.Insp_Req_No = C.Insp_Req_No AND A.Insp_Result_No = C.Insp_Result_No" &_
		" WHERE C.Plant_Cd = " & FilterVar(strPlantCd, "''", "S") & "" &_
		" AND C.INSP_CLASS_CD = " & FilterVar(strInspClassCd, "''", "S") & "" &_
		" AND C.Item_Cd = " & FilterVar(strItemCd, "''", "S") & "" &_
	      	" AND B.Insp_Item_Cd = " & FilterVar(strInspItemCd, "''", "S") & "" &_
	      	" AND C.Insp_DT BETWEEN  " & FilterVar(strYrDt1, "''", "S") & " AND  " & FilterVar(strYrDt2, "''", "S") & "" &_
	      	" AND RTrim(LTrim(A.MEAS_VALUE)) <> ''" & _
	     " ORDER BY A.INSP_REQ_NO, A.INSP_RESULT_NO, A.INSP_SERIES, A.SAMPLE_NO"
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	RS.Open  strSql, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
           Call CheckSQLError(Conn,True)
        End If
		Set RS = Nothing												'☜: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	'레코드가 하나도 없다면 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'조건에 맞는 검사결과가 없습니다 
		RS.Close
		Set RS = Nothing												'☜: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If

	'레코드가 있다면 
	lglngNumberOfData = RS.RecordCount
	
	ReDim lgdblData(lglngNumberOfData - 1)
    
	For i = 0 To lglngNumberOfData - 1
		If Trim(RS(0)) = "" Then
			Call DisplayMsgBox("229910", vbOKOnly, "", "", I_MKSCRIPT)	'관리도를 그릴 수 없는 자료입니다.
			Exit Function
		Else
	    	lgdblData(i) = UNICDbl(RS(0), 0)
	    End If
	    RS.MoveNext
	Next
	
	RS.Close
	Set RS = Nothing
	
	Get_MeasuredValues = True
End Function

'/*****************************************************
'/ 계수표 데이타 얻기 
'/*****************************************************
Function Get_Parameters()
	Dim RS
	Dim strSql
	Get_Parameters = False
	
	strSql = "SELECT U_A2, U_D3, U_D4 FROM Q_PARAMETER " &_
		  "WHERE N = " & lgintSizeOfSubgroup
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	RS.Open strSql, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
           Call CheckSQLError(Conn,True)
        End If
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function		
	End If
	
	'레코드가 하나도 없다면 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("224501", vbOKOnly, "", "", I_MKSCRIPT)	'계수표에 자료가 없습니다 
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function		
	End If
	
	lgdblA2 = UNICDbl(RS(0), 0)
	lgdblD3 = UNICDbl(RS(1), 0)
	lgdblD4 = UNICDbl(RS(2), 0)
	
	RS.Close
	Set RS = Nothing
	
	Get_Parameters = True
End Function

'/*****************************************************
'/ 소수 자릿수 얻기 
'/*****************************************************
Function Get_Decimal()
	Get_Decimal = False
	
	lgintDecimal = 4
	
	Get_Decimal = True
End Function

'/*****************************************************
'/ Subgroup 수 구하기 
'/*****************************************************
Function CalForNumOfSubgroup()
    Dim intRest
    
    CalForNumOfSubgroup = False
    
    If lglngNumberOfData < lgintSizeOfSubgroup Then
    	Call DisplayMsgBox("229912", vbOKOnly, "", "", I_MKSCRIPT)	'조회된 자료의 수가 군별 시료수보다 적습니다.
    	Exit Function
    End If
    
    lglngNumberOfSubgroup = lglngNumberOfData \ lgintSizeOfSubgroup
    intRest = lglngNumberOfData Mod lgintSizeOfSubgroup
    If intRest <> 0 Then
        Call DisplayMsgBox("229913", vbOKOnly, "", "", I_MKSCRIPT)	'조회된 자료수가 군별 시료수의 배수가 아닙니다. 나머지 자료는 무시합니다 
	lglngNumberOfData = lglngNumberOfData - intRest
    End If
    
    CalForNumOfSubgroup = True
End Function

'/*****************************************************
'/ Subgroup 개개의 평균 및 Subgroup들의 평균 구하기 
'/*****************************************************
Function CalForAvgOfSubgroup()
	Dim i
	Dim j
	Dim SumOfSubgroup
	Dim SumOfSubgroups
	
	CalForAvgOfSubgroup = False
	
	On Error Resume Next
	
	ReDim lgXbar(lglngNumberOfSubgroup - 1)
	
	If lgstrMthdOfCL = "S" Then
		SumOfSubgroups = 0
		
		If lglngNumberOfSubgroup < lgintCntOfSubGroupForCL Then
			lgintCntOfSubGroupForCL = lglngNumberOfSubgroup
		End If
		
		For i = 0 To lgintCntOfSubGroupForCL - 1
			SumOfSubgroup = 0
			For j = 0 To lgintSizeOfSubgroup - 1
			    SumOfSubgroup = SumOfSubgroup + lgdblData(i * lgintSizeOfSubgroup + j)
			Next
			lgXbar(i) = SumOfSubgroup / lgintSizeOfSubgroup
			SumOfSubgroups = SumOfSubgroups + lgXbar(i)
			
			'Min/Max 구하기 
			If i = 0 Then
				lgMinXbar = lgXbar(0)
				lgMaxXbar = lgXbar(0)
			End If
			If lgMinXbar > lgXbar(i) Then
				lgMinXbar = lgXbar(i)
			End If
			
			If lgMaxXbar < lgXbar(i) Then
				lgMaxXbar = lgXbar(i)
			End If
		Next
    
		lgXbarbar = SumOfSubgroups / lgintCntOfSubGroupForCL
		
		If lglngNumberOfSubgroup > lgintCntOfSubGroupForCL Then
			For i = lgintCntOfSubGroupForCL To lglngNumberOfSubgroup - 1
				SumOfSubgroup = 0
				For j = 0 To lgintSizeOfSubgroup - 1
				    SumOfSubgroup = SumOfSubgroup + lgdblData(i * lgintSizeOfSubgroup + j)
				Next
				lgXbar(i) = SumOfSubgroup / lgintSizeOfSubgroup
				SumOfSubgroups = SumOfSubgroups + lgXbar(i)
				
				If lgMinXbar > lgXbar(i) Then
					lgMinXbar = lgXbar(i)
				End If
				
				If lgMaxXbar < lgXbar(i) Then
					lgMaxXbar = lgXbar(i)
				End If
			Next
		End If
	Else
		SumOfSubgroups = 0
		For i = 0 To lglngNumberOfSubgroup - 1
			SumOfSubgroup = 0
			For j = 0 To lgintSizeOfSubgroup - 1
			    SumOfSubgroup = SumOfSubgroup + lgdblData(i * lgintSizeOfSubgroup + j)
			Next
			lgXbar(i) = SumOfSubgroup / lgintSizeOfSubgroup
			SumOfSubgroups = SumOfSubgroups + lgXbar(i)
			
			'Min/Max 구하기 
			If i = 0 Then
				lgMinXbar = lgXbar(0)
				lgMaxXbar = lgXbar(0)
			End If
			If lgMinXbar > lgXbar(i) Then
				lgMinXbar = lgXbar(i)
			End If
			
			If lgMaxXbar < lgXbar(i) Then
				lgMaxXbar = lgXbar(i)
			End If
		Next
    
		lgXbarbar = SumOfSubgroups / lglngNumberOfSubgroup
	
	End If
	
    If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
    
    CalForAvgOfSubgroup = True
    
End Function

'/*****************************************************
'/ Subgroup 개개의 범위 및 Subgroup들의 범위 평균 구하기 
'/*****************************************************
Function CalForRangeOfSubgroup()
    Dim i
    Dim j
    
    Dim dblTemp
    Dim dblMin
    Dim dblMax
    
    Dim SumOfRange
    
    CalForRangeOfSubgroup = False
    
    On Error Resume Next
    
    ReDim lgR(lglngNumberOfSubgroup - 1)
    
    SumOfRange = 0
    dblMax = lgdblData(0)
    dblMin = lgdblData(0)
    
    If lgstrMthdOfCL = "S" Then
		
		If lglngNumberOfSubgroup < lgintCntOfSubGroupForCL Then
			lgintCntOfSubGroupForCL = lglngNumberOfSubgroup
		End If
		
		For i = 0 To lgintCntOfSubGroupForCL - 1
			For j = 0 To lgintSizeOfSubgroup - 1
		        dblTemp = Abs(lgdblData((i) * lgintSizeOfSubgroup + j))
		        				'Min/Max 구하기 
				If j = 0 Then
					dblMin = dblTemp
					dblMax = dblTemp
				Else
					If Abs(dblMax) < dblTemp Then
					    dblMax = lgdblData((i) * lgintSizeOfSubgroup + j)
					End If
		        
					If Abs(dblMin) > dblTemp Then
					    dblMin = lgdblData((i) * lgintSizeOfSubgroup + j)
					End If
				End If
		    Next
			    
		    lgR(i) = dblMax - dblMin
		    SumOfRange = SumOfRange + lgR(i)
	
		    'Min/Max 구하기 
		    If i = 0 Then
		    	lgMinR = lgR(0)
		    	lgMaxR = lgR(0)
		    End If
		    If lgMinR > lgR(i) Then
		    	lgMinR = lgR(i)
		    End If

			If lgMaxR < lgR(i) Then
		    	lgMaxR = lgR(i)
		    End If
		Next    
		
		lgRbar = SumOfRange / lgintCntOfSubGroupForCL
		
		If lglngNumberOfSubgroup > lgintCntOfSubGroupForCL Then
			For i = lgintCntOfSubGroupForCL To lglngNumberOfSubgroup - 1
			    For j = 0 To lgintSizeOfSubgroup - 1
				    dblTemp = Abs(lgdblData((i) * lgintSizeOfSubgroup + j))
				    				'Min/Max 구하기 
					If j = 0 Then
						dblMin = dblTemp
						dblMax = dblTemp
					Else
						If Abs(dblMax) < dblTemp Then
						    dblMax = lgdblData((i) * lgintSizeOfSubgroup + j)
						End If
				    
						If Abs(dblMin) > dblTemp Then
						    dblMin = lgdblData((i) * lgintSizeOfSubgroup + j)
						End If
					End If
				Next
			    
			    lgR(i) = dblMax - dblMin
			    SumOfRange = SumOfRange + lgR(i)
			    
	
			    If lgMinR > lgR(i) Then
			    	lgMinR = lgR(i)
			    End If

				If lgMaxR < lgR(i) Then
			    	lgMaxR = lgR(i)
			    End If
			Next    
		End If
    Else
		For i = 0 To lglngNumberOfSubgroup - 1
			For j = 0 To lgintSizeOfSubgroup - 1
		        dblTemp = Abs(lgdblData((i) * lgintSizeOfSubgroup + j))
		        				'Min/Max 구하기 
				If j = 0 Then
					dblMin = dblTemp
					dblMax = dblTemp
				Else
					If Abs(dblMax) < dblTemp Then
					    dblMax = lgdblData((i) * lgintSizeOfSubgroup + j)
					End If
		        
					If Abs(dblMin) > dblTemp Then
					    dblMin = lgdblData((i) * lgintSizeOfSubgroup + j)
					End If
				End If
		    Next
		    
		    lgR(i) = dblMax - dblMin
		    SumOfRange = SumOfRange + lgR(i)
		    'Min/Max 구하기 
		    If i = 0 Then
		    	lgMinR = lgR(0)
		    	lgMaxR = lgR(0)
		    End If
		    If lgMinR > lgR(i) Then
		    	lgMinR = lgR(i)
		    End If

			If lgMaxR < lgR(i) Then
		    	lgMaxR = lgR(i)
		    End If
		Next    
		
		lgRbar = SumOfRange / lglngNumberOfSubgroup
    End If
    
    If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
    CalForRangeOfSubgroup = True
    
End Function

'/*****************************************************
'/ Xbar Chart 관리한계 구하기 
'/*****************************************************
Function CalForControlLimit()
    CalForControlLimit = False
    
    On Error Resume Next
    
    SELECT CASE lgstrMthdOfCL
		CASE "S"
			lgdblCL = lgXbarbar
			lgdblLCL = lgXbarbar - (lgdblA2 * lgRbar)
			lgdblUCL = lgXbarbar + (lgdblA2 * lgRbar)
		CASE "T"
			lgdblCL = lgXbarbar
			If lgstrLCL <> ""  Then
				lgdblLCL = UNICDbl(lgstrLCL, 0)
			End If
			
			If lgstrUCL <> ""  Then
				lgdblUCL = UNICDbl(lgstrUCL, 0)
			End If
		CASE ELSE
			lgdblCL = lgXbarbar
			lgdblLCL = lgXbarbar - (lgdblA2 * lgRbar)
			lgdblUCL = lgXbarbar + (lgdblA2 * lgRbar)
    END SELECT 
    
    If Err.Number <> 0 Then
    	Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	
    	Exit Function
    End If
    
    CalForControlLimit = True
End Function

'/*****************************************************
'/ R Chart 관리한계 구하기 
'/*****************************************************
Function CalForR_ControlLimit()
    CalForR_ControlLimit = False
    
    On Error Resume Next
    
    lgdblR_CL = lgRbar
    lgdblR_LCL = lgdblD3 * lgRbar
    lgdblR_UCL = lgdblD4 * lgRbar
    
    If Err.Number <> 0 Then
    	Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	
    	Exit Function	
    End If
    
    CalForR_ControlLimit = True
       
End Function
%>
