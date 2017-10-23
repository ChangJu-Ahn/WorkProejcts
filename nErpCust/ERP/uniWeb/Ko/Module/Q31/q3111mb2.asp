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
'*  4. Program Name         : X-Bar & R ������ 
'*  5. Program Desc         : ��� �� ������ ���� ������ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- ChartFX�� ����� ����ϱ� ���� Include ���� -->
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
	
	'����� �ҷ����� 
	Call parent.CommonQueryRs("PLANT_CD,PLANT_NM","B_PLANT","PLANT_CD =  " & parent.FilterVar(TempstrPlantCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtPlantNm.Value = strVar2
	
	'ǰ��� �ҷ����� 
	Call parent.CommonQueryRs("ITEM_CD,ITEM_NM","B_ITEM","ITEM_CD =  " & parent.FilterVar(TempstrItemCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtItemNm.Value = strVar2	
	
	'�˻��׸�� �ҷ����� 
	Call parent.CommonQueryRs("INSP_ITEM_CD,INSP_ITEM_NM","Q_INSPECTION_ITEM","INSP_ITEM_CD =  " & parent.FilterVar(TempstrInspItemCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtInspItemNm.Value = strVar2
</Script>
<%													
On Error Resume Next

'Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
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
	
'�˻�԰� 
Dim lgstrInspSpec
Dim lgdblLSL
Dim lgdblUSL
Dim lgMsmtUnitCd
	
'��� 
Dim lgdblA2
Dim lgdblD3
Dim lgdblD4
	
'Xbar �����Ѱ� 
'/* SCR 213: �����Ѱ� ����� Ʋ�� ���� - START */
Dim lgstrMthdOfCL
Dim lgintCntOfSubGroupForCL 
'/* SCR 213: �����Ѱ� ����� Ʋ�� ���� - END */
Dim lgstrLCL
Dim lgstrUCL
	
Dim lgdblCL
Dim lgdblLCL
Dim lgdblUCL
		
'���� �����Ѱ� 
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
	
'����Ÿ ��� 
lgblnRet = Get_Data
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
	
'Subgroup�� ���ϱ� 
lgblnRet = CalForNumOfSubgroup
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
    
'���� ��� �� ������ ��� ���ϱ� 
lgblnRet = CalForAvgOfSubgroup
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
	
'�� ������ ���� �� �������� ��� ���ϱ� 
lgblnRet = CalForRangeOfSubgroup
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
    	
'Xbar�� �����Ѱ� ���ϱ� 
lgblnRet = CalForControlLimit
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
       
'R�� �����Ѱ� ���ϱ� 
lgblnRet = CalForR_ControlLimit
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
%>
<Script Language=vbscript>
Dim lgblnRet
Dim lgOKFlag
	
lgOKFlag = True
	
<%'----------------------------------------------%>
<%'���� DATA DISPLAY %>
lgblnRet = Display_InspStand
If lgblnRet = False Then lgOKFlag = False
	
<%'-------------------- SPREAD --------------------------%>
<%'SPREAD�� DATA DISPLAY %>
lgblnRet = DisplayData_OnSpread
If lgblnRet = False Then lgOKFlag = False
	
    	
If lgOKFlag = True Then 
	Call Parent.DbQueryOk
End If
    	
<%'################################################################################################################
'############################################ CLIENT SIDE FUNCTION ##############################################
'################################################################################################################%>

<%'/*****************************************************
'/	���� ����Ÿ Display
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
'/	Spread�� ����Ÿ Display
'/*****************************************************%>
Function DisplayData_OnSpread()
	
	DisplayData_OnSpread = False
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1
		.vspdData.MaxCols = <%=lglngNumberOfSubgroup%>				<%'���������� Į���� ���� %>

		Parent.ggoSpread.Source = .vspdData
<%
		For i = 0 To lglngNumberOfSubgroup - 1					'�������� ��� �����ֱ� 
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
			.vspdData.Row = 1							<%'�������忡 ��հ� ���� �־��ֱ� %>
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
'/	ChartFX1(Xbar Chart)�� ȯ�� ���� 
'/*****************************************************%>
Function Setting_ChartFX1()
	Dim sngTempMin
	Dim sngTempMax 
	Dim sngTempDiffStep
	
	Setting_ChartFX1 = False
	
	Err.Clear
	On Error Resume Next
	With Parent.frm1.ChartFX1
		<%'ToolBar �Ӽ� %>
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>						<%'Ʋ�ٸ� ���ο� â���� ���̱� %>
		.ToolBarObj.Left = 15								<%'Ʋ���� ���� ��ġ %> 
		.ToolBarObj.Top = 10								<%'Ʋ���� ��� ��ġ %> 
		
		<%'Y�� ǥ�ð�(����)�� �Ҽ��� ���� �ڸ��� ���� %>
		.Axis(<%=AXIS_Y%>).Decimals = <%=lgintDecimal%>
		
		'y���� Min/Max�� �� Step ���ϱ� 
		If <%= (lgMinXBar > lgdblLCL) %> Then
			sngTempMin = parent.UNICDbl("<%=UNINumClientFormat(lgdblLCL, lgintDecimal, 0)%>")
		Else
			sngTempMin = parent.UNICDbl("<%=UNINumClientFormat(lgMinXBar, lgintDecimal, 0)%>")
		End If
	
		If <%=(lgMaxXBar < lgdblUCL)%> Then
			sngTempMax = parent.UNICDbl("<%=UNINumClientFormat(lgdblUCL, lgintDecimal, 0)%>")
		Else
			'���� ����� �ִ밪���� 1�ۼ�Ʈ ū ���� Y���� �ּҰ����� ���� 
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
'/	ChartFX2(R Chart)�� ȯ�� ���� 
'/*****************************************************%>
Function Setting_ChartFX2()
	Dim sngTempMin
	Dim sngTempMax 
	Dim sngTempDiffStep
	
	Setting_ChartFX2 = False
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1.ChartFX2
		<%'ToolBar �Ӽ� %>
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>						<%'Ʋ�ٸ� ���ο� â���� ���̱� %>
		.ToolBarObj.Left = 15								<%'Ʋ���� ���� ��ġ %> 
		.ToolBarObj.Top = 10								<%'Ʋ���� ��� ��ġ %> 
		
		<%'Y�� ǥ�ð�(����)�� �Ҽ��� ���� �ڸ��� ���� %>
		.Axis(<%=AXIS_Y%>).Decimals = <%=lgintDecimal%>		
		
		'y���� Min/Max�� �� Step ���ϱ� 
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
		.Axis(<%=AXIS_Y%>).STEP = (.Axis(<%=AXIS_Y%>).Max - .Axis(<%=AXIS_Y%>).Min) / 5		'Y�� (Max - Min) / 10���� ���� 
	End With
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Setting_ChartFX2 = True
End Function

<%'/*****************************************************
'/	ChartFX1(Xbar Chart) �׸��� 
'/*****************************************************%>
Function Draw_XbarChart()
	
	Draw_XbarChart = False
	
	Err.Clear
	On Error Resume Next
	
'	With Parent.frm1.ChartFX1
		
'		.OpenDataEx COD_VALUES, 1, <%=lglngNumberOfSubgroup%>				'��Ʈ FX���� ������ ä�� �����ֱ� 
'			'ù��° �迭(Xbar) �� ���� 
'			.Series(0).MarkerShape = <%=MK_CIRCLE%>
'			.Series(0).LineStyle = <%=CHART_SOLID%>
<%
		Dim YValue0, sInsSQL
		Dim blnRet
	    'DB ���� 
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
		
'		'UCL, LCL, CL�� ���� Constant line(s)
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
'		.CloseData <%=COD_CONSTANTS%>					'��Ʈ FX���� ä��(Constant Line�� ����) �ݾ��ֱ� 
		
'		.OpenDataEx <%=COD_STRIPES%>, 2, 0 					' Two Color stripes
'			.Stripe(0).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblUCL, lgintDecimal, 0)%>")
'			.Stripe(0).To = .Axis(<%=AXIS_Y%>).Max
'			.Stripe(0).Color = RGB(255, 100, 255)
'			.Stripe(0).Axis = <%=AXIS_Y%>
'			.Stripe(1).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblLCL, lgintDecimal, 0)%>")
'			.Stripe(1).To = .Axis(<%=AXIS_Y%>).Min
'			.Stripe(1).Color = RGB(255, 100, 255)
'			.Stripe(1).Axis = <%=AXIS_Y%>
'		.CloseData <%=COD_STRIPES%>					'��Ʈ FX���� ä��(�ٹ��̸� ����) �ݾ��ֱ� 
		
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
'/	ChartFX2(R Chart) �׸��� 
'/*****************************************************%>
Function Draw_RChart()
	Draw_RChart = False
	
	Err.Clear
	On Error Resume Next
	
'	With Parent.frm1.ChartFX2
		
'		.OpenDataEx COD_VALUES, 1, <%=lglngNumberOfSubgroup%>				'��Ʈ FX���� ������ ä�� �����ֱ� 
'			'ù��° �迭(Xbar) �� ���� 
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
		
'		'UCL, LCL, CL�� ���� Constant line(s)
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
'			.CloseData <%=COD_CONSTANTS%>					'��Ʈ FX���� ä��(Constant Line�� ����) �ݾ��ֱ� 
			
'			.OpenDataEx <%=COD_STRIPES%>, 1, 0 					' Two Color stripes
'				.Stripe(0).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
'				.Stripe(0).To = .Axis(<%=AXIS_Y%>).Max
'				.Stripe(0).Color = RGB(255, 100, 255)
'				.Stripe(0).Axis = <%=AXIS_Y%>
'			.CloseData <%=COD_STRIPES%>					'��Ʈ FX���� ä��(�ٹ��̸� ����) �ݾ��ֱ� 
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
'			.CloseData <%=COD_CONSTANTS%>					'��Ʈ FX���� ä��(Constant Line�� ����) �ݾ��ֱ� 
			
'			.OpenDataEx <%=COD_STRIPES%>, 2, 0 					' Two Color stripes
'				.Stripe(0).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_UCL, lgintDecimal, 0)%>")
'				.Stripe(0).To = .Axis(<%=AXIS_Y%>).Max
'				.Stripe(0).Color = RGB(255, 100, 255)
'				.Stripe(0).Axis = <%=AXIS_Y%>
'				.Stripe(1).From = parent.UNICDbl("<%=UNINumClientFormat(lgdblR_LCL, lgintDecimal, 0)%>")
'				.Stripe(1).To = .Axis(<%=AXIS_Y%>).Min
'				.Stripe(1).Color = RGB(255, 100, 255)
'				.Stripe(1).Axis = <%=AXIS_Y%>
'			.CloseData <%=COD_STRIPES%>					'��Ʈ FX���� ä��(�ٹ��̸� ����) �ݾ��ֱ� 
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
'/ �Է� ����Ÿ ��� 
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
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���� ���� ������ϴ� 
		Exit Function
	End If
	lgintSizeOfSubgroup = CInt(Request("txtPartSampleCnt"))
	
	If strPlantCd="" or strInspClassCd = "" or strYrDt1="" or strYrDt2="" or strItemCd="" or strInspItemCd="" then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���� ���� ������ϴ� 
		Exit Function
	End IF
	
	If lgintSizeOfSubgroup < 2 or lgintSizeOfSubgroup > 20 Then								'���� �÷���� 2���� 25������ ���� �̿��ϰ� �Ǿ� �ִ�.
		'�Ʒ��� ����ġ�� ���ų� ���� ���� �÷���� �Է����� ��츦 ���� ���Ƿ� �� �޽����̴�.
		Call DisplayMsgBox("229907", vbOKOnly, "", "", I_MKSCRIPT)	'������ ���� �÷���� �Է��Ͻʽÿ� 
		Exit Function
	End IF
	Request_QueryData = True
End Function

'/*****************************************************
'/ ��ȸ ����Ÿ ��� 
'/*****************************************************
Function Get_Data()
    Dim i
    Dim blnRet
    
    Get_Data = False
    
    'DB ���� 
    blnRet = DBConnect
    If blnRet = False Then Exit Function
    
    '�Ҽ� �ڸ��� ��� 
    blnRet = Get_Decimal
    If blnRet = False Then Exit Function
    
	'Check Input Data
    blnRet = Check_InputData
    If blnRet = False Then Exit Function    
    
    '�˻���� ���� ��� 
    blnRet = Get_InspStandard
    If blnRet = False Then Exit Function
    
    '����ġ ��� 
    blnRet = Get_MeasuredValues
    If blnRet = False Then Exit Function
    
    '����� ��� 
    blnRet = Get_Parameters
    If blnRet = False Then Exit Function
    
    'DB ���� ���� 
    blnRet = DBClose
    If blnRet = False Then Exit Function
    
    Get_Data = True
End Function

'/*****************************************************
'/ Database ���� 
'/*****************************************************
Function DBConnect()
	DBConnect = False
	
	'Object ���� 
	Set Conn = Server.CreateObject("ADODB.Connection")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	
		Exit Function					
	End If
	
	
	' ODBC Data source ���� 
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
'/ Database ���� ���� 
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
'/ �Է� ����Ÿ üũ 
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
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	'���� üũ 
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
		    Set RS = Nothing											'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing
			Exit Function
		End If
		
		'���ڵ尡 �ϳ��� ���ٸ� 
		If RS.EOF or RS.BOF then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)	'������ �������� �ʽ��ϴ�.
			RS.Close
			Set RS = Nothing												'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'��: ComProxy Unload
			Exit Function
		End If
	RS.Close
	End if
	
	'ǰ�� üũ 
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
		    Set RS = Nothing											'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing
			Exit Function
		End If
		
		'���ڵ尡 �ϳ��� ���ٸ� 
		If RS.EOF or RS.BOF then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)	'������ �������� �ʽ��ϴ�.
			RS.Close
			Set RS = Nothing												'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'��: ComProxy Unload
			Exit Function
		End If
	RS.Close
	End if

	'�˻��׸� üũ 
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
		    Set RS = Nothing											'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing
			Exit Function
		End If
		
		'���ڵ尡 �ϳ��� ���ٸ� 
		If RS.EOF or RS.BOF then
			Call DisplayMsgBox("220201", vbOKOnly, "", "", I_MKSCRIPT)	'������ �������� �ʽ��ϴ�.
			RS.Close
			Set RS = Nothing												'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'��: ComProxy Unload
			Exit Function
		End If
	
	RS.Close
	End if
	
	Set RS = Nothing	
	
	Check_InputData = True
End Function


'/*****************************************************
'/ �˻���� ����Ÿ ��� 
'/*****************************************************
Function Get_InspStandard()
	Dim RS
	Dim strSql
	Get_InspStandard = False
	
	Err.Clear
	On Error Resume Next
	
	'/* SCR 213: �����Ѱ� ����� Ʋ�� ���� - START */
	strSql = "SELECT INSP_SPEC, LSL, USL, MTHD_OF_CL_CAL, CALCULATED_QTY, LCL, UCL, MEASMT_UNIT_CD, INSP_UNIT_INDCTN " &_
              "FROM Q_INSPECTION_STANDARD_BY_ITEM " &_
              "WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & "" &_
              "AND INSP_CLASS_CD = " & FilterVar(strInspClassCd, "''", "S") & "" &_
              "AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") & "" &_
              "AND INSP_ITEM_CD = " & FilterVar(strInspItemCd, "''", "S") & ""
    '/* SCR 213: �����Ѱ� ����� Ʋ�� ���� - END */          
	Set RS = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
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
		Set RS = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	'���ڵ尡 �ϳ��� ���ٸ� 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	   '���ǿ� �´� �˻����� �����ϴ� 
		RS.Close
		Set RS = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		'�Ʒ��� ���Ƿ� �� �޽��� 
		Exit Function
	End If
	
	If Trim(RS(8)) <> "3" Then
		Call DisplayMsgBox("229939", vbOKOnly, "", "", I_MKSCRIPT)	'�˻����ǰ��ǥ�ð� Ư��ġ�� �ƴմϴ�.
		Exit Function
	End If
	
	If Trim(RS(0)) = "" Then
		Call DisplayMsgBox("220706", vbOKOnly, "", "", I_MKSCRIPT)	'�˻�԰��� �ԷµǾ� ���� �ʽ��ϴ� 
		Exit Function
	End If
	
	lgstrInspSpec = RS(0)
	
	If Trim(RS(1)) = "" AND Trim(RS(2)) = "" Then
		Call DisplayMsgBox("229911", vbOKOnly, "", "", I_MKSCRIPT)	'����/���ѱ԰� �� ��� �ϳ��� �����ؾ� �մϴ� 
		Exit Function
	End If

	lgdblLSL = UNICDbl(RS(1), 0)
	lgdblUSL = UNICDbl(RS(2), 0)
	
	'/* SCR 213: �����Ѱ� ����� Ʋ�� ���� - START */
	lgstrMthdOfCL = UCase(Trim(RS(3)))
	SELECT CASE lgstrMthdOfCL
		CASE "S"	'�κа�� 
			lgintCntOfSubGroupForCL = RS(4)
			lgstrLCL = ""
			lgstrUCL = ""
		CASE "T"	'��ǥġ 
			lgstrLCL = Trim(RS(5))
			lgstrUCL = Trim(RS(6))
		CASE ELSE	'��ü���: "C" Or ""
			lgstrLCL = ""
			lgstrUCL = ""
	END SELECT
		
	lgMsmtUnitCd = Trim(RS(7))
	'/* SCR 213: �����Ѱ� ����� Ʋ�� ���� - END */
	RS.Close
	Set RS = Nothing
	
	Get_InspStandard = True
End Function

'/*****************************************************
'/ ����ġ ����Ÿ ��� 
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
		Set Conn = Nothing												'��: ComProxy Unload
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
		Set RS = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	'���ڵ尡 �ϳ��� ���ٸ� 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		RS.Close
		Set RS = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	'���ڵ尡 �ִٸ� 
	lglngNumberOfData = RS.RecordCount
	
	ReDim lgdblData(lglngNumberOfData - 1)
    
	For i = 0 To lglngNumberOfData - 1
		If Trim(RS(0)) = "" Then
			Call DisplayMsgBox("229910", vbOKOnly, "", "", I_MKSCRIPT)	'�������� �׸� �� ���� �ڷ��Դϴ�.
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
'/ ���ǥ ����Ÿ ��� 
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
		Set Conn = Nothing												'��: ComProxy Unload
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
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function		
	End If
	
	'���ڵ尡 �ϳ��� ���ٸ� 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("224501", vbOKOnly, "", "", I_MKSCRIPT)	'���ǥ�� �ڷᰡ �����ϴ� 
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
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
'/ �Ҽ� �ڸ��� ��� 
'/*****************************************************
Function Get_Decimal()
	Get_Decimal = False
	
	lgintDecimal = 4
	
	Get_Decimal = True
End Function

'/*****************************************************
'/ Subgroup �� ���ϱ� 
'/*****************************************************
Function CalForNumOfSubgroup()
    Dim intRest
    
    CalForNumOfSubgroup = False
    
    If lglngNumberOfData < lgintSizeOfSubgroup Then
    	Call DisplayMsgBox("229912", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ�� �ڷ��� ���� ���� �÷������ �����ϴ�.
    	Exit Function
    End If
    
    lglngNumberOfSubgroup = lglngNumberOfData \ lgintSizeOfSubgroup
    intRest = lglngNumberOfData Mod lgintSizeOfSubgroup
    If intRest <> 0 Then
        Call DisplayMsgBox("229913", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ�� �ڷ���� ���� �÷���� ����� �ƴմϴ�. ������ �ڷ�� �����մϴ� 
	lglngNumberOfData = lglngNumberOfData - intRest
    End If
    
    CalForNumOfSubgroup = True
End Function

'/*****************************************************
'/ Subgroup ������ ��� �� Subgroup���� ��� ���ϱ� 
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
			
			'Min/Max ���ϱ� 
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
			
			'Min/Max ���ϱ� 
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
'/ Subgroup ������ ���� �� Subgroup���� ���� ��� ���ϱ� 
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
		        				'Min/Max ���ϱ� 
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
	
		    'Min/Max ���ϱ� 
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
				    				'Min/Max ���ϱ� 
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
		        				'Min/Max ���ϱ� 
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
		    'Min/Max ���ϱ� 
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
'/ Xbar Chart �����Ѱ� ���ϱ� 
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
'/ R Chart �����Ѱ� ���ϱ� 
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
