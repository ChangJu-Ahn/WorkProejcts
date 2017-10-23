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
'*  1. Module Name          : Quality
'*  2. Function Name        :  
'*  3. Program ID              : q3111mb4.asp
'*  4. Program Name         : p ������ 
'*  5. Program Desc         : ��� �� ������ ���� ������ 
'*  6. Comproxy List         : 
'                             
'*  7. Modified date(First) : 2000/08/23
'*  8. Modified date(Last)  : 2001/01/03
'*  9. Modifier (First)     : Oh Youngjoon
'* 10. Modifier (Last)      : Yang Jaehee
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- ChartFX�� ����� ����ϱ� ���� Include ���� -->
<!-- #include file="../../inc/CfxIE.inc" -->
<!--
Parameter�ޱ� 
DB Connect
��Ʈ �ʱ�ȭ 
Data�������� 
�Ѱ˻��,�Ѻҷ��� ���ϱ� 
�ҷ��� ���ϱ� 
��պҷ��� ���ϱ� 
�ִ�ҷ��� ���ϱ� 
�ּҺҷ��� ���ϱ� 
UCL���ϱ� 
LCL���ϱ� 
ȭ�鿡 ���� Display
��Ʈ�׸��� 
-->

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
	
	TempstrPlantCd		= "<%=Request("txtPlantCd")%>"
	TempstrItemCd		= "<%=Request("txtItemCd")%>"	
	
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
</Script>
<%													
On Error Resume Next

'Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")
				

Dim i
Dim j

Dim Conn
Dim RS
Dim RS1
Dim RS2
Dim RS3

Dim LngRow
Dim strSql
Dim strSql1
Dim strSql2
Dim StrSql3

Dim intRecordCount1

Dim strPlantCd
Dim strInspItemCd
Dim strInspClassCd
Dim strYrDt1
Dim strYrDt2
Dim strItemCd
Dim ClsQty													'���� ���� 

' �����Ѱ� 
Dim lgdblSigma
Dim lgdblCL
Dim lgdblLCL()
Dim lgdblUCL()
Dim lgdblMaxUCL
Dim lgdblMinLCL
Dim lgtempLCL

Dim lgblnRet
Dim lgInspQty()
Dim lgDefectQty()
Dim lgDefectRatio()
Dim avgDefectRatio

On Error Resume Next

lgdblMinLCL = 0

'Request
lgblnRet = Request_QueryData
If lgblnRet = False Then 
	Call HideStatusWnd	
	Response.End
End If	
'DB Connect
lgblnRet = DBConnect
If lgblnRet = False Then 
	Call HideStatusWnd	
	Response.End
End If	
'Get Data
lgblnRet = Get_Data
If lgblnRet = False Then
	Call HideStatusWnd	
	Response.End
End If	

'Get Defect Ratio(Average,Min,Max)
lgblnRet = Get_DefectRatio
If lgblnRet = False Then 
	Call HideStatusWnd	
	Response.End
End If	
%>
<Script Language="VBScript">
	Dim MeasmtUnitCd											'���������ڵ� 
	Dim CL
	Dim UCL(<%=ClsQty%>)
	Dim LCL(<%=ClsQty%>)
	Dim InspQtyTotal											'�� �˻�� 
	Dim DefectQtyTotal											'�� �ҷ��� 
	Dim InspQty(<%=ClsQty%>)										'�� ���� �˻�� 
	Dim DefectQty(<%=ClsQty%>)										'�� ���� �ҷ��� 
	Dim p(<%=ClsQty%>)											'�� ���� �ҷ��� 
	Dim Maxp												'�ִ� �ҷ��� 
 	Dim Minp												'�ּ� �ҷ��� 
	Dim StDev(<%=ClsQty%>)												'ǥ������ 
	Dim MaxUCL
	Dim MinLCL
	Dim UCLSeries
	
	'Dim lgdblMaxUCL
	'Dim lgdblMinLCL
	
	Dim lgblnRet
	Dim lgOKFlag

	lgOKFlag = True
	
	<%'���� DATA DISPLAY %>
	lgblnRet = Display_InspStand
	If lgblnRet = False Then lgOKFlag = False
	
	<%'SPREAD�� DATA DISPLAY %>
	lgblnRet = DisplayData_OnSpread
	If lgblnRet = False Then lgOKFlag = False
	<%'ChartFX�� DATA DISPLAY %>
	'lgblnRet = Setting_chartFX1
	'If lgblnRet = False Then lgOKFlag = False
	
	lgblnRet = Draw_chartP
	If lgblnRet = False Then lgOKFlag = False
	
	If lgOKFlag = True Then 
    	Call Parent.DbQueryOk
    End If
    
<%'################################################################################################################
'############################################ CLIENT SIDE FUNCTION ##############################################
'################################################################################################################%>
Function Display_InspStand()
	Display_InspStand = False
	With Parent.frm1
		.txtInspQtyTotal.value = "<%=UNINumClientFormat(RS2(0), ggQty.DecPoint, 0)%>"
		.txtDefectQtyTotal.value = "<%=UNINumClientFormat(RS2(1), ggQty.DecPoint, 0)%>"
		.txtMinp.value = "<%=UNINumClientFormat(RS3(0), 2, 0)%>"
		.txtMaxp.value =  "<%=UNINumClientFormat(RS3(1), 2, 0)%>"
	End With
	
	Display_InspStand = True
	
End Function

Function DisplayData_OnSpread()
	DisplayData_OnSpread = False
	'Dim lgtempLCL
	
	Err.Clear
	On Error Resume Next
	
	With Parent.frm1
		.vspdData.MaxCols = <%=intRecordCount1 + 1%>				<%'���������� Į���� ���� %>

		Parent.ggoSpread.Source = .vspdData

<%
			For i = 0 To intRecordCount1 - 1					'�������� ��� �����ֱ� 
%>
				Parent.ggoSpread.SSSetEdit (<%=i%> + 1), (<%=i%> + 1), 8, 1, -1, 15 
<%
			Next
%>
			Parent.ggoSpread.SSSetEdit (<%=i%> + 1), "�� ��", 8, 1, -1, 15 
			
			<%i=1%>
			.vspdData.Row = <%=i%>
<%
			Redim lgdblUCL(intRecordCount1)
			Redim lgdblLCL(intRecordCount1)

			lgdblMaxUCL = 0
			lgdblMaxLCL = 0

			For j = 0 To intRecordCount1
%>
				.vspdData.Col = <%=j + 1%>					<%'�������忡 �˻�� �־��ֱ� %>
				.vspdData.Text =  "<%=UNINumClientFormat(lgInspQty(i - 1,j),   ggQty.DecPoint, 0)%>"
<% 

				lgdblSigma = SQR(avgDefectRatio * (1 - avgDefectRatio) / lgInspQty(i - 1,j))
				
				lgdblUCL(j) = avgDefectRatio + (3 * lgdblSigma)
				If lgdblMaxUCL <= lgdblUCL(j) Then
					lgdblMaxUCL = lgdblUCL(j)
				End if

				lgtempLCL = avgDefectRatio - (3 * lgdblSigma)

				If lgtempLCL >= 0 Then
					lgdblLCL(j) = lgtempLCL
					If lgdblMinLCL >= lgdblLCL(j) Then
						lgdblMinLCL = lgtempLCL
					End if
				Else
					lgdblLCL(j) = "-"
				End if
			Next
			
			i=2
%>
			.vspdData.Row = <%=i%>
<%
			For j = 0 To intRecordCount1
%>
				.vspdData.Col = <%=j + 1%>					<%'�������忡 �ҷ��� �־��ֱ� %>
				.vspdData.Text = "<%=UNINumClientFormat(lgInspQty(i - 1,j), ggQty.DecPoint, 0)%>"
<%
			Next
			
			i=3
%>
			.vspdData.Row = <%=i%>
<%
			Redim lgdefectRatio(intRecordCount1-1)
			
			For j = 0 To intRecordCount1
%>
				.vspdData.Col = <%=j + 1%>					<%'�������忡 �ҷ��� �־��ֱ� %>
				.vspdData.Text = "<%=UNINumClientFormat(lgInspQty(i - 1,j) * 100, 2, 0)%>"
<%				
				lgdefectRatio(j) = lgInspQty(i - 1,j)
				
				If lgdblMaxUCL <= lgdefectRatio(j) Then
					lgdblMaxUCL = lgdefectRatio(j)
				End if
				
				If lgdblMinUCL > lgdefectRatio(j) Then
					lgdblMinUCL = lgdefectRatio(j)
				End if
			Next

			i=4
%>
			.vspdData.Row = <%=i%>
<%
			'�հ� �κ��� ����Ÿ�� �� �ִ´�.
			For j = 0 To intRecordCount1 -1
%>
				.vspdData.Col = <%=j + 1%>					<%'�������忡 UCL �־��ֱ� %>
				.vspdData.Text = "<%=UNINumClientFormat(lgdblUCL(j) * 100, 2, 0)%>"
<%
			Next

			i=5
%>
			.vspdData.Row = <%=i%>
<%
			'�հ� �κ��� ����Ÿ�� �� �ִ´�.
			For j = 0 To intRecordCount1-1				'�������� ��� �����ֱ� 
%>
				.vspdData.Col = <%=j + 1%>					<%'�������忡 LCL �־��ֱ� %>
				If <%= (lgdblLCL(j) = "-")%> Then
					.vspdData.Text = "-"
				Else
					.vspdData.Text = "<%=UNINumClientFormat(lgdblLCL(j) * 100, 2, 0)%>"
				End If
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

Function Setting_chartFX1()
	Setting_chartFX1 = False
	With parent.frm1.chartFX1
		'.Gallery=1
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
	End With
	Setting_chartFX1 = True
End Function

Function Draw_chartP()
	Draw_chartP = False
<%	
	
	Err.Clear
	On Error Resume Next
	
	
	Dim TermRatio1
	Dim sInsSQL
	Dim YValue0, YValue1, YValue2 , XCL
	
	YValue2 = CDbl(0)

	lgblnRet = DBConnect
	'���̺� ���� 
	sInsSQL = " DELETE FROM Q_TMP_CHART_P "
	Conn.Execute sInsSQL

	If lgdblMinLCL < 0 Then
		lgdblMinLCL = 0
	End if

	TermRatio1 = UNINumClientFormat(lgdblMaxUCL - lgdblMinLCL, 4, 0)			'�ִ밪�� �ּҰ��� ���� 
	TermRatio1 = (TermRatio1 / 10) * 2				'�ִ밪�� �ּҰ��� ���̸� 10���		
	
	if TermRatio1=0 Then TermRatio1 = MaxDRatio1 / 10 * 2
	
	XCL = UNINumClientFormat(avgDefectRatio * 100, 2, 0)
	
	For j=0  to intRecordCount1 - 1
	
		YValue0 = UNINumClientFormat(lgInspQty(2,j) * 100, 2, 0)	'�ҷ��� 
		YValue1 = UNINumClientFormat(lgdblUCL(j) * 100, 2, 0)	'UCL 
		
		If CStr(lgdblLCL(j)) = "-" Then	
			YValue2 = 0         'LCL
			'Call ServerMesgBox("True : " & CStr(lVValue2) , vbInformation, I_MKSCRIPT)
		Else 
			
			YValue2 = UNINumClientFormat(lgdblLCL(j) * 100, 2, 0)		'LCL
			'Call ServerMesgBox("False : " & CStr(lgdblLCL(j) * 100) , vbInformation, I_MKSCRIPT)
		End if
			
			 
		sInsSQL = "INSERT INTO Q_TMP_CHART_P  (XVALUE, YVALUE1, YVALUE2, YVALUE3, X_CL) "
		sInsSQL = sInsSQL & " VALUES (" & FilterVar(j+1, "", "S") & ", " 
		sInsSQL	= sInsSQL & FilterVar(YValue0, "", "S") & ", " 
		sInsSQL	= sInsSQL & FilterVar(YValue1, "", "S") & ", " 
		sInsSQL	= sInsSQL & FilterVar(YValue2, "", "S") & ", "
		sInsSQL	= sInsSQL & FilterVar(XCL, "", "S") & ") "

		
		Conn.Execute sInsSQL
		
	Next
%>		
	'	With parent.frm1.chartFX1

		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		'.OpenDataEx <%=COD_VALUES%>,3, <%=intRecordCount1 %>
		'.Axis(<%=AXIS_Y%>).Max = (parent.UNICDbl("<%=UNINumClientFormat(lgdblMaxUCL, 4, 0)%>") + TermRatio1 ) * 100
		
		'If (parent.UNICDbl("<%=UNINumClientFormat(lgdblMinLCL, 4, 0)%>") - TermRatio1) < 0 Then
		'	.Axis(<%=AXIS_Y%>).Min = 0
		'Else
		'	.Axis(<%=AXIS_Y%>).Min = (parent.UNICDbl("<%=UNINumClientFormat(lgdblMinLCL, 4, 0)%>") - TermRatio1) * 100
		'End if
		'.Axis(<%=AXIS_Y%>).STEP = (.Axis(<%=AXIS_Y%>).Max - .Axis(<%=AXIS_Y%>).Min) / 10
		
		'.SerLeg(0) = "�ҷ���"
		'.SerLeg(1) = "UCL"
		'.SerLeg(2) = "LCL"
		
		'.Series(0).Gallery = 1
		'.Series(1).Gallery = 8
		'.Series(2).Gallery = 8
		
		'.Series(0).Visible = True
		'.Series(1).Visible = True
		'.Series(2).Visible = True
		
		
		'.AXIS(<%=AXIS_X%>).Label(<%=j%>) = <%= j + 1 %>
			'.Series(0).YValue(<%=j%>) = parent.UNICDbl("<%=UNINumClientFormat(lgInspQty(2,j) * 100, 2, 0)%>")	'�ҷ��� 
			'.Series(1).YValue(<%=j%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblUCL(j) * 100, 2, 0)%>")	'UCL
			'If <%= (lgdblLCL(j) = "-") %> Then	
			'	.Series(2).YValue(<%=j%>) = <%=CHART_HIDDEN%>          'LCL
			'Else 
			'	.Series(2).YValue(<%=j%>) = parent.UNICDbl("<%=UNINumClientFormat(lgdblLCL(j) * 100, 2, 0)%>")		'LCL
			'End if

		'.CloseData <%=COD_VALUES%>	' Close the VALUES channel
		
		'CL�� ���� Constant line(s)
		'.OpenDataEx <%=COD_CONSTANTS%>, 1, 0 					
		'	.ConstantLine(0).Value = parent.UNICDbl("<%=UNINumClientFormat(avgDefectRatio * 100, 2, 0)%>")
		'	.ConstantLine(0).Axis = <%=AXIS_Y%>
		'	.ConstantLine(0).Label = "CL = " &  "<%=UNINumClientFormat(avgDefectRatio * 100, 2, 0)%>"
		'	.ConstantLine(0).LineColor = RGB(255, 0, 0)			
		'.CloseData <%=COD_CONSTANTS%>					'��Ʈ FX���� ä��(Constant Line�� ����) �ݾ��ֱ� 
'	End With

<%	
	blnRet = DBClose
%>	

	Draw_chartP = True
End Function

</Script>

<%
'################################################################################################################
'############################################ SERVER SIDE FUNCTION ##############################################
'################################################################################################################

Function Request_QueryData()
	Request_QueryData = False
	
	strPlantCd  = Request("txtPlantCd")
	strInspClassCd = Request("cboInspClassCd")
	strYrDt1= UNIConvDate(Request("txtYrDt1"))
	strYrDt2= UNIConvDate(Request("txtYrDt2"))
	strItemCd = Request("txtItemCd")
	
	If strPlantCd="" or strInspClassCd = "" or strYrDt1="" or strYrDt2="" or strItemCd="" then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���� ���� ������ϴ� 
		Exit Function
	End IF
	
	Request_QueryData = True
End Function

Function DBConnect()
	DBConnect = False
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	If Err.Number <> 0 Then
		Set Conn = Nothing
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function
	End if

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





Function Get_Data()
	get_data = False
	strSql1 = "Select Insp_Qty, Defect_Qty, case Insp_Qty when 0 then Null else Defect_Qty/Insp_Qty end from Q_Inspection_Result Where Plant_Cd = " & FilterVar(strPlantCd, "''", "S") & " and Insp_Class_Cd = " & FilterVar(strInspClassCd, "''", "S") & " and " &_
	          "Item_Cd = " & FilterVar(strItemCd, "''", "S") & " and Insp_Dt Between '" & strYrDt1 & "' and '" & strYrDt2 & "'"
	
	strSql2 = "Select Sum(Insp_Qty), Sum(Defect_Qty), Sum(Defect_Qty)/Sum(Insp_Qty) from Q_Inspection_Result Where Plant_Cd = " & FilterVar(strPlantCd, "''", "S") & " and Insp_Class_Cd = " & FilterVar(strInspClassCd, "''", "S") & " and " &_
	          "Item_Cd = " & FilterVar(strItemCd, "''", "S") & " and Insp_Dt Between '" & strYrDt1 & "' and '" & strYrDt2 & "'"
    
	On Error Resume Next
	
'**********************�߰� �κн���*************************************
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
	
'**********************�߰� �κг�*************************************	
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set RS1 = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function														'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	RS1.Open  strSql1, Conn, 1			'adOpenKeyset
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
	If RS1.EOF or RS1.BOF then
		RS1.Close
		Set RS1 = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing
		'�Ʒ��� ���Ƿ� �� �޽��� 
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		Exit Function	
	End If
	
	'���ڵ尡 �ִٸ� 
	intRecordCount1 = RS1.RecordCount
	Redim lgInspQty(4, intRecordCount1)
	For i=0 to intRecordCount1 - 1
		lgInspQty(0,i) = CLng(RS1(0))
		lgInspQty(1,i) = CLng(RS1(1))
		lgInspQty(2,i) = CDbl(RS1(2))
		RS1.MoveNext
	Next
		
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set RS2 = Nothing
		RS1.Close
		Set RS1 = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	RS2.Open  strSql2, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		RS2.Close
		Set RS2 = Nothing
		RS1.Close
		Set RS1 = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing		
		If CheckSYSTEMError(Err,True) = False Then
           Call CheckSQLError(Conn,True)
        End If										'��: ComProxy Unload
		
		Exit Function													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	'���ڵ尡 �ϳ��� ���ٸ� 
	If RS2.EOF or RS2.BOF then
		RS2.Close
		Set RS2 = Nothing
		RS1.Close
		Set RS1 = Nothing												'��: ComProxy Unload
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Set Conn = Nothing												'��: ComProxy Unload
		'�Ʒ��� ���Ƿ� �� �޽��� 
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		Exit Function	
	End If
	lgInspQty(0,intRecordCount1) = CLng(RS2(0))
	lgInspQty(1,intRecordCount1) = CLng(RS2(1))
	lgInspQty(2,intRecordCount1) = CDbl(RS2(2))
	avgDefectRatio = CDbl(RS2(2))
	
	get_data = True
End Function

Function Get_DefectRatio()
	Get_DefectRatio = False
	
	StrSql3 = "Select Min(Defect_qty/Insp_qty)*100,Max(Defect_qty/Insp_qty)*100 from Q_Inspection_Result Where Plant_Cd = " & FilterVar(strPlantCd, "''", "S") & " and Insp_Class_Cd = " & FilterVar(strInspClassCd, "''", "S") & " and " &_
	          "Item_Cd = " & FilterVar(strItemCd, "''", "S") & " and Insp_Dt Between '" & strYrDt1 & "' and '" & strYrDt2 & "' and Insp_qty > 0"
	
	Set RS3 = Server.CreateObject("ADODB.RecordSet")
	
	If Err.Number <> 0 Then
		Set RS3 = Nothing
		Exit Function
	End if
	
	RS3.Open StrSql3,Conn,1
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
           Call CheckSQLError(Conn,True)
        End If
		Set RS3 = Nothing
		Exit Function													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	'���ڵ尡 �ϳ��� ���ٸ� 
	If RS3.EOF or RS3.BOF then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		RS3.Close
		Set RS3 = Nothing
		Exit Function
	End If
	
	Get_DefectRatio = True
End Function
%>
