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
'*  2. Function Name        : Cp,Cpk
'*  3. Program ID           : Q3111MB3
'*  4. Program Name         : �����ɷ��� 
'*  5. Program Desc         : 
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
Dim lgdblZi(632)
Dim lgdblFZi(632)
Dim lgPCL
Dim lgPLSL
Dim lgPUSL	
	
Dim lgdblFZMax
		
Dim lgCp				'Cp
Dim lgCpk			'Cpk
	
'�ִ�/�ּ� 
Dim lgdblRange
Dim lgdblMax
Dim lgdblMin
	
'�ִ�/�ּҰ��� 
Dim lgstrMaxTolerance
Dim lgstrMinTolerance
	
'���/ǥ������/+3�ñ׸�/-3�ñ׸� 
Dim lgdblAvg
Dim lgdblSigma
Dim lgdblP3Sigma
Dim lgdblM3Sigma
	
'����Ÿ��(�÷��)
Dim lglngNumberOfData
	
'�˻�԰� 
Dim lgdblInspSpec
Dim lgstrLSL
Dim lgstrUSL
Dim lgdblLSL
Dim lgdblUSL
Dim lgMsmtUnitCd
	
Dim lgSpecFlag		'U:���ѱ԰ݸ� , L:���ѱ԰ݸ� , B:���� �԰� 
	
'�Ҽ��� �ڸ��� 
Dim lgintDecimal
	
Dim lgblnRet
Dim i
	
Const PI = 3.14159265358979
		
'Request
lgblnRet = Request_QueryData
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If

'����Ÿ ��� 
lgblnRet = Get_Data
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If
	
'Cp/Cpk ���� 
lgblnRet = CalForCapabilityIndices
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If

'����ȭ 
lgblnRet = CalForStandadization
If lgblnRet = False Then 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Response.End
End If

%>
<Script Language=vbscript>
Dim lgblnRet
Dim lgOKFlag
		
lgOKFlag = True
	
'----------------------------------------------
'���� DATA DISPLAY %>
lgblnRet = Display_InspStand
If lgblnRet = False Then lgOKFlag = False
		
'-------------------- CHART --------------------------
'ChartFX �Ӽ� ���� 
'lgblnRet = Setting_ChartFX1
'If lgblnRet = False Then lgOKFlag = False
	
'Cp/Cpk �׸��� 
lgblnRet = Draw_Capablity
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
		.txtCp.Value = "<%=UNINumClientFormat(lgCp, lgintDecimal, 0)%>"
		If "<%=lgCpk%>" <> "" Then
			.txtCpk.Value = "<%=UNINumClientFormat(lgCpk, lgintDecimal, 0)%>"
		End If
		.txtInspSpec.Value = "<%=UniNumClientFormat(lgdblInspSpec, lgintDecimal, 0)%>"
		
		If "<%=lgstrLSL%>" = "" Then
			.txtLSL.Value = ""
		Else
			.txtLSL.Value = "<%=UniNumClientFormat(lgdblLSL, lgintDecimal ,0)%>"
		End If
		
		If "<%=lgstrUSL%>" = "" Then
			.txtUSL.Value = ""
		Else
			.txtUSL.Value = "<%=UniNumClientFormat(lgdblUSL, lgintDecimal ,0)%>"
		End If
		
		.txtSampleQty.Value = "<%=UniNumClientFormat(lglngNumberOfData, ggQty.DecPoint ,0)%>"	<%'�÷�� %>
		.txtMaxTol.Value = "<%=UniNumClientFormat(lgstrMaxTolerance, lgintDecimal ,0)%>"		<%'�ִ���� %>
		.txtMinTol.Value = "<%=UniNumClientFormat(lgstrMinTolerance, lgintDecimal ,0)%>"		<%'�ּҰ��� %>
		.txtMAX.Value = "<%=UniNumClientFormat(lgdblMax, lgintDecimal, 0)%>"			<%'�ִ밪 %>
		.txtMIN.Value = "<%=UniNumClientFormat(lgdblMin, lgintDecimal, 0)%>"			<%'�ִ밪 %>
		.txtAverage.Value = "<%=UniNumClientFormat(lgdblAvg, lgintDecimal, 0)%>"		<%'��� %>
		.txtRange.Value = "<%=UniNumClientFormat(lgdblRange, lgintDecimal, 0)%>"			<%'���� %>
		.txtStd.Value = "<%=UniNumClientFormat(lgdblSigma, lgintDecimal, 0)%>"			<%'ǥ������ %>
		.txtP3Sigma.Value = "<%=UniNumClientFormat(lgdblP3Sigma, lgintDecimal, 0)%>"	<%'+3�ñ׸� %>
		.txtM3Sigma.Value = "<%=UniNumClientFormat(lgdblM3Sigma, lgintDecimal, 0)%>"	<%'-3�ñ׸� %>
		.txtMeasmtUnitCd.Value = "<%=lgMsmtUnitCd%>"		<%'�������� %>
	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	Display_InspStand = True
	
End Function

<%'/*****************************************************
'/	ChartFX1(Cp/Cpk)�� ȯ�� ���� 
'/*****************************************************%>
Function Setting_ChartFX1()
	Dim sngTempDiffStep
	
	Setting_ChartFX1 = False
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1.ChartFX1
			
		<%'ToolBar �Ӽ� %>
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>				<%'Ʋ�ٸ� ���ο� â���� ���̱� %>
		.ToolBarObj.Left = 15								<%'Ʋ���� ���� ��ġ %> 
		.ToolBarObj.Top = 10								<%'Ʋ���� ��� ��ġ %> 
		
		<%'Y�� ǥ�ð�(����)�� �Ҽ��� ���� �ڸ��� ���� %>
		.Axis(<%=AXIS_Y%>).Decimals = <%=lgintDecimal%>
		
		.Volume = 100	
		.MarkerShape = <%=MK_NONE%>
		
		'Min/Max/Step ���ϱ� 
		.Axis(<%=AXIS_Y%>).Min = 0
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UniNumClientFormat(lgdblFZMax, lgintDecimal, 0)%>") * (11 / 10)							'�׷��� Y���� �ִ밪 ���� 
		.Axis(<%=AXIS_Y%>).STEP = .Axis(<%=AXIS_Y%>).Max / 10
		
	End With
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Setting_ChartFX1 = True
	
End Function

<%'/*****************************************************
'/	ChartFX1(Cp/Cpk) �׸��� 
'/*****************************************************%>
Function Draw_Capablity()
	
	Draw_Capablity = False
	
	Err.Clear
	On Error Resume Next
	
'''''	With Parent.frm1.ChartFX1
		
		Select Case "<%=lgSpecFlag%>"
			Case "B"
'''''				.OpenDataEx <%=COD_VALUES%>, 2, 633				'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
	Dim YValue0, YValue1, sInsSQL
	Dim blnRet
    'DB ���� 
    blnRet = DBConnect
					
					sInsSQL = "DELETE FROM Q_TMP_CHART_OPRN_CAPA"
					Conn.Execute sInsSQL
%>

<%

					For i = 0 to 632
						
%>

						
'''''						.Series(0).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
'''''						
'''''						If <%= (i >= 0) And (i < lgPLSL)%> Then
'''''							.Series(1).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
'''''						ElseIf <%= (i > lgPUSL) And (i <= 632)%> Then
'''''							.Series(1).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
'''''						Else
'''''							.Series(1).YValue(<%=i%>) = <%=CHART_HIDDEN%>
'''''						End If	


<%
						' ���ʱ԰��� �ִ� ��� ������ INSERT
						If lgSpecFlag = "B" then

							YValue0 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)

							If (i >= 0) And (i < lgPLSL) Then
								YValue1 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)
							ElseIf (i > lgPUSL) And (i <= 632) Then
								YValue1 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)
							Else
								YValue1 = 0
							End If	

							sInsSQL = "INSERT INTO Q_TMP_CHART_OPRN_CAPA  (XVALUE, YVALUE1, YVALUE2, LSL_VALUE, LSL_XPOS, Bar_VALUE, Bar_XPOS, USL_VALUE, USL_XPOS) "
							sInsSQL = sInsSQL & " VALUES (" & FilterVar(i, "", "S") & ", " & FilterVar(YValue0, "", "S") & ", " & FilterVar(YValue1, "", "S") & ", "
							sInsSQL = sInsSQL &  FilterVar(lgdblLSL, "", "S") & ", "  &  FilterVar(lgPLSL, "", "S") & ", " 
							sInsSQL = sInsSQL &  FilterVar(lgdblAvg, "", "S") & ", "  &  FilterVar(lgPCL, "", "S") & ", " 
							sInsSQL = sInsSQL &  FilterVar(lgdblUSL, "", "S") & ", " &  FilterVar(lgPUSL, "", "S") & ")"

							Conn.Execute sInsSQL
						End if
%>

<%
					Next
%>

'''''				.CloseData <%=COD_VALUES%>
				
				'USL, LSL, Avg�� ���� Constant line(s)
'''''				.OpenDataEx <%=COD_CONSTANTS%>, 3, 0 					
'''''					.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPLSL, lgintDecimal, 0)%>")
'''''					.ConstantLine(0).Axis = <%=AXIS_X%>
'''''					.ConstantLine(0).Label = "LSL = " & "<%=UNINumClientFormat(lgdblLSL, lgintDecimal, 0)%>"
'''''					.ConstantLine(0).LineColor = RGB(0, 255, 0)
'''''					.ConstantLine(0).Style = .ConstantLine(0).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''					.ConstantLine(1).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPUSL, lgintDecimal, 0)%>")
'''''					.ConstantLine(1).Axis = <%=AXIS_X%>
'''''					.ConstantLine(1).Label = "USL = " & "<%=UNINumClientFormat(lgdblUSL, lgintDecimal, 0)%>"
'''''					.ConstantLine(1).LineColor = RGB(0, 255, 0)
'''''					.ConstantLine(1).Style = .ConstantLine(1).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''					.ConstantLine(2).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPCL, lgintDecimal, 0)%>")
'''''					.ConstantLine(2).Axis = <%=AXIS_X%>
'''''					.ConstantLine(2).Label = "Xbar = " & "<%=UNINumClientFormat(lgdblAvg, lgintDecimal, 0)%>"
'''''					.ConstantLine(2).LineColor = RGB(0, 0, 0)
'''''					.ConstantLine(2).Style = .ConstantLine(2).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''				.CloseData <%=COD_CONSTANTS%>
			
			Case "U"
'''''				.OpenDataEx <%=COD_VALUES%>, 2, 633				'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
					For i = 0 to 632
%>
'''''						.Series(0).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
						
'''''						If <%= (i > lgPUSL) And (i <= 632) %> Then
'''''							.Series(1).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
'''''						End If

<%
						' ���ѱ԰ݸ� �ִ� ��� ������ INSERT
						If lgSpecFlag = "U" then

							YValue0 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)

							If  (i > lgPUSL) And (i <= 632) Then
								YValue1 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)
							End If	


							sInsSQL = "INSERT INTO Q_TMP_CHART_OPRN_CAPA  (XVALUE, YVALUE1, YVALUE2, LSL_VALUE, LSL_XPOS, Bar_VALUE, Bar_XPOS, USL_VALUE, USL_XPOS) "
							sInsSQL = sInsSQL & " VALUES (" & FilterVar(i, "", "S") & ", " & FilterVar(YValue0, "", "S") & ", " & FilterVar(YValue1, "", "S") & ", "
							sInsSQL = sInsSQL &  FilterVar("0", "", "S") & ", "  &  FilterVar("0", "", "S") & ", " 
							sInsSQL = sInsSQL &  FilterVar(lgdblAvg, "", "S") & ", "  &  FilterVar(lgPCL, "", "S") & ", " 
							sInsSQL = sInsSQL &  FilterVar(lgdblUSL, "", "S") & ", " &  FilterVar(lgPUSL, "", "S") & ")"

							Conn.Execute sInsSQL
						End If
%>


<%
					Next
%>
					
'''''				.CloseData <%=COD_VALUES%>
				'USL, Avg�� ���� Constant line(s)
'''''				.OpenDataEx <%=COD_CONSTANTS%>, 2, 0 					
'''''					.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPUSL, lgintDecimal, 0)%>")
'''''					.ConstantLine(0).Axis = <%=AXIS_X%>
'''''					.ConstantLine(0).Label = "USL = " & "<%=UNINumClientFormat(lgdblUSL, lgintDecimal, 0)%>"
'''''					.ConstantLine(0).LineColor = RGB(0, 255, 0)
'''''					.ConstantLine(0).Style = .ConstantLine(0).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''					.ConstantLine(1).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPCL, lgintDecimal, 0)%>")
'''''					.ConstantLine(1).Axis = <%=AXIS_X%>
'''''					.ConstantLine(1).Label = "Xbar = " & "<%=UNINumClientFormat(lgdblAvg, lgintDecimal, 0)%>"
'''''					.ConstantLine(1).LineColor = RGB(0, 0, 0)
'''''					.ConstantLine(1).Style = .ConstantLine(1).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''				.CloseData <%=COD_CONSTANTS%>
			
			Case "L"
				.OpenDataEx <%=COD_VALUES%>, 2, 633				'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
					For i = 0 to 632
%>
'''''						.Series(0).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
						
'''''						If <%= (i >= 0) And (i < lgPLSL)%> Then
'''''							.Series(1).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)%>")
'''''						End If

<%

						' ���ѱ԰ݸ� �ִ� ��� ������ INSERT
						If lgSpecFlag = "L" then

							YValue0 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)

							If  (i >= 0) And (i < lgPLSL) Then
								YValue1 = UNINumClientFormat(lgdblFZi(i), lgintDecimal, 0)
							End If	

							sInsSQL = "INSERT INTO Q_TMP_CHART_OPRN_CAPA  (XVALUE, YVALUE1, YVALUE2, LSL_VALUE, LSL_XPOS, Bar_VALUE, Bar_XPOS, USL_VALUE, USL_XPOS) "
							sInsSQL = sInsSQL & " VALUES (" & FilterVar(i, "", "S") & ", " & FilterVar(YValue0, "", "S") & ", " & FilterVar(YValue1, "", "S") & ", "
							sInsSQL = sInsSQL &  FilterVar(lgdblLSL, "", "S") & ", "  &  FilterVar(lgPLSL, "", "S") & ", " 
							sInsSQL = sInsSQL &  FilterVar(lgdblAvg, "", "S") & ", "  &  FilterVar(lgPCL, "", "S") & ", " 
							sInsSQL = sInsSQL &  FilterVar("0", "", "S") & ", " &  FilterVar("0", "", "S") & ")"

							Conn.Execute sInsSQL
						End If
%>

<%
					Next
%>
'''''				.CloseData <%=COD_VALUES%>
				'LSL, Avg�� ���� Constant line(s)
'''''				.OpenDataEx <%=COD_CONSTANTS%>, 2, 0 					
'''''					.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPLSL, lgintDecimal, 0)%>")
'''''					.ConstantLine(0).Axis = <%=AXIS_X%>
'''''					.ConstantLine(0).Label = "LSL = " & "<%=UNINumClientFormat(lgdblLSL, lgintDecimal, 0)%>"
'''''					.ConstantLine(0).LineColor = RGB(0, 255, 0)
'''''					.ConstantLine(0).Style = .ConstantLine(0).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''					.ConstantLine(1).Value = parent.UNICDbl("<%=UNINumClientFormat(lgPCL, lgintDecimal, 0)%>")
'''''					.ConstantLine(1).Axis = <%=AXIS_X%>
'''''					.ConstantLine(1).Label = "Xbar = " & "<%=UNINumClientFormat(lgdblAvg, lgintDecimal, 0)%>"
'''''					.ConstantLine(1).LineColor = RGB(0, 0, 0)
'''''					.ConstantLine(1).Style = .ConstantLine(1).Style Or &H4	<%'=CC_RIGHTALIGNED%>
'''''				.CloseData <%=COD_CONSTANTS%>
		End Select
		
'''''		.OpenDataEx <%=COD_COLORS%>, 2, 0 					
'''''			.Series(0).Color = RGB(100, 100, 255)
'''''			.Series(1).Color = RGB(255, 0, 0)
'''''		.CloseData <%=COD_COLORS%>
'''''		
'''''		.Series(0).Gallery = <%=LINES%>		'Cp�(ù��° �迭) �׷����� Ÿ�� 
'''''		
'''''		.Series(1).Gallery = <%=BAR%>		'�ι�° �迭 �׷����� Ÿ�� 
'''''		.Series(1).Border = True
'''''		.Series(1).BorderColor = RGB(255, 0, 0)
'''''		
'''''		.Series(0).Visible = True 
'''''		.Series(1).Visible = True 
'''''		
'''''	End With
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	Draw_Capablity = True
	
End Function		
</Script>   
<%

blnRet = DBClose


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
	
	If strPlantCd="" or strInspClassCd = "" or strYrDt1="" or strYrDt2="" or strItemCd="" or strInspItemCd="" then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���� ���� ������ϴ� 
		Exit Function
	End IF
	
	Request_QueryData = True
End Function

'/*****************************************************
'/ ��ȸ ����Ÿ ��� 
'/*****************************************************
Function Get_Data()
    Dim blnRet
    
    Get_Data = False
    
    'DB ���� 
    blnRet = DBConnect
    If blnRet = False Then Exit Function

    '�Ҽ� �ڸ��� ��� 
    blnRet = Get_Decimal
    If blnRet = False Then
    	Conn.Close
	Set Conn = Nothing	
    	Exit Function
   End If

	'Check Input Data
    blnRet = Check_InputData
    If blnRet = False Then Exit Function

    '�˻���� ���� ��� 
    blnRet = Get_InspStandard
    If blnRet = False Then
    	Exit Function
   End If

    '����ġ ��� 
    blnRet = Get_MeasuredValues
    If blnRet = False Then
    	Exit Function
   End If

    'DB ���� ���� 
    blnRet = DBClose
    If blnRet = False Then Exit Function
    	
    'ǥ������ 
    blnRet = Get_Sigma
    If blnRet = False Then
    	Exit Function
   End If

   '+3�ñ׸�/-3�ñ׸� 
    blnRet = Get_PM3Sigma
    If blnRet = False Then
    	Exit Function
   End If
   
   
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
'/ �Ҽ� �ڸ��� ��� 
'/*****************************************************
Function Get_Decimal()
	Get_Decimal = False
	
	lgintDecimal = 4
	
	Get_Decimal = True
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
	
	strSql = "SELECT INSP_SPEC, LSL, USL, MEASMT_UNIT_CD, INSP_UNIT_INDCTN " &_
              "FROM Q_INSPECTION_STANDARD_BY_ITEM " &_
              "WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & "" &_
              "AND INSP_CLASS_CD = " & FilterVar(strInspClassCd, "''", "S") & "" &_
              "AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") & "" &_
              "AND INSP_ITEM_CD = " & FilterVar(strInspItemCd, "''", "S") & ""              
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
        RS.Close
		Conn.Close
		Set RS = Nothing
		Set Conn = Nothing
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
	
	If Trim(RS(4)) <> "3" Then
		Call DisplayMsgBox("229939", vbOKOnly, "", "", I_MKSCRIPT)	'�˻����ǰ��ǥ�ð� Ư��ġ�� �ƴմϴ�.
		Exit Function
	End If
	
	If Trim(RS(0)) = "" Then
		Call DisplayMsgBox("220706", vbOKOnly, "", "", I_MKSCRIPT)	'�˻�԰��� �ԷµǾ� ���� �ʽ��ϴ� 
		Exit Function
	End If
	
	lgdblInspSpec = UNICDbl(RS(0), 0)
	
	If Trim(RS(1)) = "" AND Trim(RS(2)) = "" Then
		Call DisplayMsgBox("229911", vbOKOnly, "", "", I_MKSCRIPT)	'����/���ѱ԰� �� ��� �ϳ��� �����ؾ� �մϴ� 
		Exit Function
	End If
	
	lgstrLSL = Trim(RS(1))
	lgstrUSL = Trim(RS(2))
	lgdblLSL = UNICDbl(RS(1), 0)
	lgdblUSL = UNICDbl(RS(2), 0)
	lgMsmtUnitCd = Trim(RS(3))
		
	lgSpecFlag = "B"
	
	If lgstrLSL = "" Then
		lgstrMinTolerance = ""	
		lgSpecFlag = "U"
	Else
		lgstrMinTolerance = lgdblInspSpec - lgdblLSL
	End If

	If lgstrUSL = "" Then
		lgstrMaxTolerance = ""
		lgSpecFlag = "L"
	Else
		lgstrMaxTolerance = lgdblUSL - lgdblInspSpec
	End If

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
	Dim dblSum
	
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
	      	" AND RTrim(LTrim(A.MEAS_VALUE)) <> ''"
	
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
        RS.Close
		Conn.Close
		Set RS = Nothing
		Set Conn = Nothing
		Exit Function
	End If

	'���ڵ尡 �ϳ��� ���ٸ� 
	If RS.EOF or RS.BOF then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT) 	'���ǿ� �´� �˻����� �����ϴ� 
		RS.Close
		Set RS = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	'���ڵ尡 �ִٸ� 
	lglngNumberOfData = RS.RecordCount
	If lglngNumberOfData < 50 Then
		Call DisplayMsgBox("229917", vbOKOnly, "", "", I_MKSCRIPT) 	'�����ɷ��� �׸��� ���� ����Ÿ���� �����մϴ�'
		RS.Close
		Set RS = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing
		Exit Function
	End If
	
	ReDim lgdblData(lglngNumberOfData - 1)
    	dblSum = 0
	For i = 0 To lglngNumberOfData - 1
		If Trim(RS(0)) = "" Then
			Call DisplayMsgBox("229910", vbOKOnly, "", "", I_MKSCRIPT) 	'�������� �׸� �� ���� �ڷ��Դϴ�.
			Exit Function
		Else
	    		lgdblData(i) = Cdbl(RS(0))
	    		'Sum
	    		dblSum = dblSum + lgdblData(i)
	    		'Min/Max ��� 
	    		If i = 0 Then
	    			lgdblMin = lgdblData(0)
	    			lgdblMax = lgdblData(0)
	    		End If
	    		
	    		If lgdblMin > lgdblData(i) Then
	    			lgdblMin = lgdblData(i)
	    		End If
	    		
	    		If lgdblMax < lgdblData(i) Then
	    			lgdblMax = lgdblData(i)
	    		End If
	    	End If
	    	RS.MoveNext
	Next
	
	lgdblRange = lgdblMax - lgdblMin		'���� 
	lgdblAvg = dblSum / lglngNumberOfData		'��� 

	RS.Close
	Set RS = Nothing
	
	Get_MeasuredValues = True
End Function


'/*****************************************************
'/ ǥ������ ���ϱ� 
'/*****************************************************
Function Get_Sigma()
	Dim dblSum
	
	Get_Sigma = False
	
	Err.Clear
	On Error Resume Next
	
    dblSum = 0
	For i = 0 To lglngNumberOfData - 1
			dblSum = dblSum + (lgdblAvg - lgdblData(i)) ^ 2
	Next
		
	lgdblSigma = Sqr(dblSum / (lglngNumberOfData - 1))
	If lgdblSigma = 0 Then
		Call DisplayMsgBox("229914", vbOKOnly, "", "", I_MKSCRIPT)	'ǥ�������� 0 �Դϴ� 
		Exit Function
	End If
		
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Get_Sigma = True
End Function

'/*****************************************************
'/ +3Sigma / -3Sigma ���ϱ� 
'/*****************************************************
Function Get_PM3Sigma()
	Get_PM3Sigma= False
	
	Err.Clear
	On Error Resume Next
	
	lgdblP3Sigma = lgdblAvg + 3 * lgdblSigma
	lgdblM3Sigma = lgdblAvg - 3 * lgdblSigma
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Get_PM3Sigma = True
End Function
  
'/*****************************************************
'/ �����ɷ� ���� ���ϱ� 
'/*****************************************************
Function CalForCapabilityIndices()
	Dim k
	CalForCapabilityIndices = False
    	
    	Err.Clear
	On Error Resume Next
	
	'Cp / Cpk ���� 
	Select Case lgSpecFlag
		Case "B"
			k = Abs(((lgdblUSL + lgdblLSL) / 2) - lgdblAvg) / ((lgdblUSL - lgdblLSL) / 2)
			lgCp = (lgdblUSL - lgdblLSL) / (6 * lgdblSigma)
			
			If k < 1 Then
				lgCpk = (1 - k) * lgCp
			Else
				lgCpk = 0
			End If
		Case "U"
		    If lgdblAvg >= lgdblUSL Then
		    		lgCp = 0
		    Else
		    		lgCp = (lgdblUSL - lgdblAvg) / (3 * lgdblSigma)
		    End If
		    lgCpk = ""
		Case "L"
		    If lgdblAvg <= lgdblLSL Then
		    		lgCp = 0
		    Else
		    		lgCp = (lgdblAvg - lgdblLSL) / (3 * lgdblSigma)
		    End If
	End Select
	
    If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
    
   	CalForCapabilityIndices = True
    
End Function

'/*****************************************************
'/ ����ȭ 
'/*****************************************************
Function CalForStandadization()
	Dim M
	Dim Ps			'X���� ������ 
	Dim Pf			'X���� ���� 
	Dim Xi			'X�� Point
	
	Dim ZCL
	Dim ZUSL
	Dim ZLSL
	
	Dim BeforeZi
	Dim XInterval	
	
	CalForStandadization = False
    	
   	Err.Clear
	On Error Resume Next
	
	ZCL = 0						'����� ���Ժ�ȯ�� ���� Ȯ���е��Լ��� 0
	lgdblFZMax = 0
	lgPCL = 0
	lgPLSL = 0
	lgPUSL = 0
	'X���� ���۰� �� ��� 
    	Select Case lgSpecFlag
        	Case "B"      '���ʱ԰��� ��� 
            	M = (lgdblUSL + lgdblLSL) / 2			'�԰��� �߽�ġ 
                
            	Ps = M - 3.2 * lgdblSigma
            	If Ps > lgdblLSL Then Ps = lgdblLSL - 0.1 * (lgdblUSL - lgdblLSL)
                		
            	Pf = M + 3.2 * lgdblSigma
            	If Pf < lgdblUSL Then Pf = lgdblUSL + 0.1 * (lgdblUSL - lgdblLSL)
                
            	XInterval = (Pf - Ps) / 633		'633���� ��. �� ���� �Ÿ��� ���Ѵ�.
                		
	        	ZUSL = (lgdblUSL - lgdblAvg) / lgdblSigma								'�԰ݻ����� ���Ժ�ȯ 
				ZLSL = (lgdblLSL - lgdblAvg) / lgdblSigma								'�԰������� ���Ժ�ȯ 
			
				For i = 0 to 632										'�� ������ ���Ժ�ȯ�Ѵ�.
					Xi = Ps + i * XInterval
					lgdblZi(i) = (Xi - lgdblAvg) / lgdblSigma
					lgdblFZi(i) = (1 / Sqr(2 * PI)) * Exp(-0.5 * (lgdblZi(i) ^ 2))
						
					If i = 0 Then
						If lgdblZi(i) = ZLSL Then					'LSL�� �ش��ϴ� ���� ��ġ ��� 
							lgPLSL = i
						End If
	
						lgdblFZMax = lgdblFZi(i)
					Else
						If BeforeZi < ZLSL and lgdblZi(i) >= ZLSL Then					'LSL�� �ش��ϴ� ���� ��ġ ��� 
							lgPLSL = i
						ElseIf BeforeZi < ZCL and lgdblZi(i) >= ZCL Then				'��տ� �ش��ϴ� ���� ��ġ ��� 
							lgPCL = i
						ElseIf BeforeZi < ZUSL and lgdblZi(i) >= ZUSL Then					'USL�� �ش��ϴ� ���� ��ġ ��� 
							lgPUSL = i
						End If
	
						If lgdblFZMax <lgdblFZi(i) Then								'�׷��� Y���� Adm ������ ���� 
							lgdblFZMax = lgdblFZi(i)
						End If
					End If
					BeforeZi = lgdblZi(i)
				Next
			  		
        	Case "U"      '���ѱ԰ݸ� �ִ� ��� 
            	M = lgdblAvg
                
            	Ps = M - 3.2 * lgdblSigma
            	If Ps > lgdblMin Then Ps = lgdblMin - 0.1 * (lgdblUSL - lgdblMin)
                		
            	Pf = M + 3.2 * lgdblSigma
            	If Pf < lgdblUSL Then Pf = lgdblUSL + 0.1 * (lgdblUSL - lgdblMin)
                		
            	XInterval = (Pf - Ps) / 633		'633���� ��. �� ���� �Ÿ��� ���Ѵ�.
                		
            	ZUSL = (lgdblUSL - lgdblAvg) / lgdblSigma						'�԰ݻ����� ���Ժ�ȯ 

				For i = 0 to 632									'�� ������ ���Ժ�ȯ�Ѵ�.
					Xi = Ps + i * XInterval
					lgdblZi(i) = (Xi - lgdblAvg) / lgdblSigma
					lgdblFZi(i) = (1 / Sqr(2 * PI)) * Exp(-0.5 * (lgdblZi(i) ^ 2))
					If i = 0 Then
						If lgdblZi(i) = ZCL Then					'CL�� �ش��ϴ� ���� ��ġ ��� 
							lgPCL = i
						End If
	
						lgdblFZMax = lgdblFZi(i)
					Else
						If BeforeZi < ZCL and lgdblZi(i) >= ZCL Then					'��տ� �ش��ϴ� ���� ��ġ ��� 
							lgPCL = i
	
						ElseIf BeforeZi < ZUSL and lgdblZi(i) >= ZUSL Then				'USL�� �ش��ϴ� ���� ��ġ ��� 
							lgPUSL = i
						End If
	
						If lgdblFZMax < lgdblFZi(i) Then						'�׷��� Y���� Adm ������ ���� 
							lgdblFZMax = lgdblFZi(i)
						End If
					End If
					BeforeZi = lgdblZi(i)
				Next		
                
        		Case "L"      '���ѱ԰ݸ� �ִ� ��� 
                		M = lgdblAvg
                		
                		
                		Ps = M - 3.2 * lgdblSigma
                		If Ps > lgdblLSL Then Ps = lgdblLSL - 0.1 * (lgdblMax - lgdblLSL)
                		
                		Pf = M + 3.2 * lgdblSigma
                		If Pf < lgdblMax Then Pf = MaxValue + 0.1 * (lgdblMax - lgdblLSL)
				
				XInterval = (Pf - Ps) / 633		'633���� ��. �� ���� �Ÿ��� ���Ѵ�.
                		
                		ZLSL = (lgdblLSL - lgdblAvg) / lgdblSigma						'�԰������� ���Ժ�ȯ 
		                
				For i = 0 to 632									'�� ������ ���Ժ�ȯ�Ѵ�.
				Xi = Ps + i * XInterval
				lgdblZi(i) = (Xi - lgdblAvg) / lgdblSigma
				lgdblFZi(i) = (1 / Sqr(2 * PI)) * Exp(-0.5 * (lgdblZi(i) ^ 2))
				
				If i = 0 Then
					If lgdblZi(i) = ZLSL Then					'LSL�� �ش��ϴ� ���� ��ġ ��� 
						lgPLSL = i
					End If
	
					lgdblFZMax = lgdblFZi(i)
				Else
					If BeforeZi < ZLSL and lgdblZi(i) >= ZLSL Then					'LSL�� �ش��ϴ� ���� ��ġ ��� 
						
						lgPLSL = i
	
					ElseIf BeforeZi < ZCL and lgdblZi(i) >= ZCL Then					'��տ� �ش��ϴ� ���� ��ġ ��� 
						lgPCL = i
					End If
	
					If lgdblFZMax < lgdblFZi(i) Then								'�׷��� Y���� Adm ������ ���� 
						lgdblFZMax = lgdblFZi(i)
					End If
				End If
				BeforeZi = lgdblZi(i)
			Next
			
    	End Select
    				
    	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
    
    	CalForStandadization = True
    
End Function
%>
