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
'*  2. Function Name        : q11118ListMeaEquSvr
'*  3. Program ID           : q3412mb1.asp
'*  4. Program Name         : �����ҷ����� �ķ���ǥ 
'*  5. Program Desc         : �����ҷ����� 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2000/05/29
'*  8. Modified date(Last)  : 2000/05/29
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      :
'* 11. Comment              :
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
	

	Dim TempstrPlantCd
	Dim TempstrInspClassCd
	Dim TempstrItemCd
	Dim TempstrInspItemCd
	Dim TempstrDefectTypeCd
	
	
	TempstrPlantCd		= "<%=Request("txtPlantCd")%>"
	TempstrInspClassCd	= "P"
	TempstrItemCd		= "<%=Request("txtItemCd")%>"
	TempstrInspItemCd	= "<%=Request("txtInspItemCd")%>"	
	TempstrDefectTypeCd	= "<%=Request("txtDefectTypeCd")%>"
	
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

	'�ҷ������� �ҷ����� 
	Call parent.CommonQueryRs("DEFECT_TYPE_CD,DEFECT_TYPE_NM","Q_DEFECT_TYPE","PLANT_CD =  " & parent.FilterVar(TempstrPlantCd , "''", "S") & " AND INSP_CLASS_CD =  " & parent.FilterVar(TempstrInspClassCd , "''", "S") & " AND DEFECT_TYPE_CD =  " & parent.FilterVar(TempstrDefectTypeCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtDefectTypeNm.Value = strVar2
</Script>
<%													
On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")  

Dim Conn

Dim RS	'@@@���� 
Dim RS1
Dim RS2

Dim LngRow

Dim strSql	'@@@���� 
Dim strSql1
Dim strSql2

Dim aggrDTotal	'@@@���� 
Dim intRecordCount1	'@@@���� 
Dim intRecordCount2	'@@@���� 

Dim TempInspItemCd
Dim TempDefectTypeCd          
Dim strPlantCd
Dim strItemCd
Dim strInspItemCd
Dim strDefectTypeCd
Dim strYr1
Dim strYr2
Dim strMnth1
Dim strMnth2
Dim blnBeforeFlag
Dim blnNowFlag
Dim strSpdData1()
Dim strSpdData2()
Dim dblSpdData1()
Dim dblSpdData2()
Dim i
Dim j
Dim Total1
Dim Total2
Dim blnRet

' Receive datas from client
blnRet = DataReceive()
If blnRet = False Then
	Response.End
End If

' Connect the Database
blnRet = DBConnect()
If blnRet = False Then
	Response.End
End If

' Get datas from the database
blnRet = GetData()
If blnRet = False Then
	Response.End
End If

' Calculate datas for display
Call TransferData()

%>

<Script Language=vbscript>
' Diplay datas on spread sheet.
If <%=blnNowFlag%> Then
	Call DisplayOnSpreadForNow()
'	Call DrawChartForNow
End If

If <%=blnBeforeFlag%> Then
	Call DisplayOnSpreadForBefore()
'	Call DrawChartForBefore
End If

Call Parent.DBQueryOK

'=================================================================================
' Diplay datas on spread sheet for This Month.
'=================================================================================
Sub DisplayOnSpreadForNow()
	'------------------------------------------------------------------------------------------------------------------------
	'��� 
	'------------------------------------------------------------------------------------------------------------------------
	Parent.frm1.vspdData1.MaxCols = <%=intRecordCount1+1%>
		
	With Parent.ggoSpread
		.Source = Parent.frm1.vspdData1
<%
		'�������� ��� �����ֱ� 
		For i = 0 To intRecordCount1-1
%>									
			.SSSetEdit  <%=i + 2%>, "<%=ConvSPChars(strSpdData1(i))%>" , 8, 1, -1, 15 
<%
		Next
%>	
	End With		
	
	With Parent.frm1.vspdData1
<%
		For i = 0 to 3														'�������忡 �� �����ֱ� 
%>
			.Row = <%= i + 1%>
<%
			For j = 0 to intRecordCount1 - 1
%>
				.Col = <%=j+2%>
				Select Case <%=i%>
					Case 0, 2		'�ҷ���, �����ҷ��� 
						.Text = "<%=UNINumClientFormat(dblSpdData1(i,j), ggQty.DecPoint, 0)%>"
					Case Else		'������, ���������� 
						.Text = "<%=UNINumClientFormat(dblSpdData1(i,j), 2, 0)%>"
				End Select
<%
			Next
		Next
%>
		.Row = 1															'�� �����ֱ� 
		.Col = 1
		.Text = "<%=UNINumClientFormat(Total1, ggQty.DecPoint, 0)%>"		'�ҷ��� ��	
		.Row = 2
		.Text = "<%=UNINumClientFormat("100", 2, 0)%>"						'�ҷ��� �� 
	End With
End Sub

'=================================================================================
' Diplay datas on spread sheet for Before Month.
'=================================================================================
Sub DisplayOnSpreadForBefore()
	'------------------------------------------------------------------------------------------------------------------------
	'���� 
	'------------------------------------------------------------------------------------------------------------------------
	Parent.frm1.vspdData2.MaxCols = <%=intRecordCount2+1%>
		
	With Parent.ggoSpread
		.Source = Parent.frm1.vspdData2
<%
		'�������� ��� �����ֱ� 
		For i = 0 To intRecordCount2-1
%>									
			.SSSetEdit  <%=i + 2%>, "<%=ConvSPChars(strSpdData2(i))%>" , 8, 1, -1, 15 
<%
		Next
%>	
	End With		
	
	With Parent.frm1.vspdData2
<%
		For i = 0 to 3														'�������忡 �� �����ֱ� 
%>
			.Row = <%= i + 1%>
<%
			For j = 0 to intRecordCount2 - 1
%>
				.Col = <%=j+2%>
				Select Case <%=i%>
					Case 0, 2		'�ҷ���, �����ҷ��� 
						.Text = "<%=UNINumClientFormat(dblSpdData2(i,j), ggQty.DecPoint, 0)%>"
					Case Else		'������, ���������� 
						.Text = "<%=UNINumClientFormat(dblSpdData2(i,j), 2, 0)%>"
				End Select
<%
			Next
		Next
%>
		.Row = 1															'�� �����ֱ� 
		.Col = 1
		.Text = "<%=UNINumClientFormat(Total2, ggQty.DecPoint, 0)%>"		'�ҷ��� ��	
		.Row = 2
		.Text = "<%=UNINumClientFormat("100", 2, 0)%>"						'�ҷ��� �� 
	End With	
End Sub

'=================================================================================
' Draw the chart for This Month
'=================================================================================
Sub DrawChartForNow()
	With Parent.frm1.ChartFX1
		'---------------------------------------------------------------------
		'��ƮFX1 - ��� 
		'---------------------------------------------------------------------
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		.Axis(<%=AXIS_Y2%>).Visible = True
		.SerLeg(0) = "�ҷ���"
		.SerLeg(1) = "����������"
													
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(Total1, ggQty.DecPoint, 0)%>")
		.Axis(<%=AXIS_Y2%>).Max = 100
		.Axis(<%=AXIS_Y%>).Min = 0
		.Axis(<%=AXIS_Y2%>).Min = 0
		.Axis(<%=AXIS_Y%>).Decimals = 0											'Y�� �Ҽ��� �ڸ��� ���� 
		.Volume = 100												'Bar���� ���̸� 0���� ���� 
		.ToolBarObj.Visible = False													'���ٸ� ����Ʈ�� ���� �ֱ� 
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.OpenDataEx <%=COD_VALUES%>, 2, <%=intRecordCount1%>											'��ƮFX���� ä�� �����ֱ� 
<%
			For i = 0 to intRecordCount1 -1
%>
				.Axis(<%=AXIS_X%>).Label(<%=i%>) = "<%=ConvSPChars(strSpdData1(i))%>"
				.Series(0).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData1(0,i), ggQty.DecPoint, 0)%>")
				.Series(1).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData1(3,i), 2, 0)%>")
<%
			Next
%>
		.CloseData <%=COD_VALUES%>	' Close the VALUES channel
		
		.Series(0).Gallery = 2
		.Series(1).Gallery = 1
		
		.Series(0).Border = True
		
		.Series(0).Visible = True
		.Series(1).Visible = True
		
		.Series(1).YAxis = <%=AXIS_Y2%>		'To move the second series to a secondary Y axis
	End with
End Sub

'=================================================================================
' Draw the chart for Before Month
'=================================================================================
Sub DrawChartForBefore()
	'------------------------------------------------------------------------------------------------------------------------
	'����: ��ƮFX2 - ������ ��Ʈ 
	'------------------------------------------------------------------------------------------------------------------------
	With Parent.frm1.ChartFX2
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		.Axis(<%=AXIS_Y2%>).Visible = True
		.SerLeg(0) = "�ҷ���"
		.SerLeg(1) = "����������"
		
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(Total2, ggQty.DecPoint, 0)%>")
		.Axis(<%=AXIS_Y2%>).Max = 100
		.Axis(<%=AXIS_Y%>).Min = 0
		.Axis(<%=AXIS_Y2%>).Min = 0
		.Axis(<%=AXIS_Y%>).Decimals = 0												'Y�� �Ҽ��� �ڸ��� ���� 
		.Volume = 100																'Bar���� ���̸� 0���� ���� 
		.ToolBarObj.Visible = False													'���ٸ� ����Ʈ�� ���� �ֱ� 
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.OpenDataEx <%=COD_VALUES%>, 2, <%=intRecordCount2%>						'��ƮFX���� ä�� �����ֱ� 
<%
			' Code to set the data
			For i = 0 to intRecordCount2 -1
%>
				.Axis(<%=AXIS_X%>).Label(<%=i%>) = "<%=ConvSPChars(strSpdData2(i))%>"
				.Series(0).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData2(0,i), ggQty.DecPoint, 0)%>")
				.Series(1).YValue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData2(3,i), 2, 0)%>")
<%
			Next
%>
		.CloseData <%=COD_VALUES%>
		
		.Series(0).Gallery = 2
		.Series(1).Gallery = 1
		
		.Series(0).Border = True
			
		.Series(0).Visible = True
		.Series(1).Visible = True
		.Series(1).YAxis = <%=AXIS_Y2%>		'To move the second series to a secondary Y axis
	End with
End Sub

</Script>	

<%
Function DataReceive()
	DataReceive = False

	strPlantCd  = Request("txtPlantCd")
	strItemCd = Request("txtItemCd")

	strDefectTypeCd = Request("txtDefectTypeCd")
	TempDefectTypeCd = Request("txtDefectTypeCd")
	If strDefectTypeCd = "" Then
		strDefectTypeCd = "%"
	End if

	strInspItemCd = Request("txtInspItemCd")
	TempInspItemCd = Request("txtInspItemCd")	
	If strInspItemCd = "" Then
		strInspItemCd = "%"
	End if

	strYr1 = Request("txtYr")
	strMnth1 = Request("txtMnth")

	If strMnth1 = "01" Then
		strYr2 = CStr(CInt(strYr1) - 1)
		strMnth2 = "12"
	Else
		strYr2 = strYr1
		strMnth2 = Right("0" & CStr(CInt(strMnth1) - 1), 2)
	End If

	If strPlantCd="" or strItemCd="" or strYr1="" or strYr2="" or strMnth1="" or strMnth2="" Then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���� ���� ������ϴ� 
		Exit Function
	End If

	DataReceive = True
End Function

Function DBConnect()
	DBConnect = False

	On Error Resume Next		
	' Database ���� Object ���� 
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

Sub DBClose()
	On Error Resume Next	
	
	RS1.Close
	RS2.Close
	RS3.Close
	RS4.Close
	Conn.Close
	Set RS2 = Nothing
	Set RS1 = Nothing
	Set RS3 = Nothing
	Set RS4 = Nothing
	Set Conn = Nothing
	
End Sub

Function GetData()
	GetData = False
	On Error Resume Next	

	'��� 
	strSql1 = "SELECT Q_DEFECT_CAUSE.DEFECT_CAUSE_NM, Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) FROM Q_FINAL_DEFECT_CAUSE_TOTAL Left Outer Join " &_
	              "Q_DEFECT_CAUSE On Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_CAUSE_CD = Q_DEFECT_CAUSE.DEFECT_CAUSE_CD " &_
	              "Where Q_DEFECT_CAUSE.INSP_CLASS_CD=" & FilterVar("F", "''", "S") & "  and Q_FINAL_DEFECT_CAUSE_TOTAL.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.YR = '" & strYr1 & "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Mnth = '" & strMnth1 &_
	              "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Item_Cd = " & FilterVar(strItemCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.INSP_ITEM_CD LIKE " & FilterVar(strInspItemCd, "''", "S") & " and DEFECT_TYPE_CD LIKE " & FilterVar(strDefectTypeCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY > 0   " &_
	              "GROUP BY Q_DEFECT_CAUSE.DEFECT_CAUSE_NM Order By Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) Desc"
	'���� 
	strSql2 = "SELECT Q_DEFECT_CAUSE.DEFECT_CAUSE_NM, Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) FROM Q_FINAL_DEFECT_CAUSE_TOTAL Left Outer Join " &_
	              "Q_DEFECT_CAUSE On Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_CAUSE_CD = Q_DEFECT_CAUSE.DEFECT_CAUSE_CD " &_
	              "Where Q_DEFECT_CAUSE.Insp_Class_Cd=" & FilterVar("F", "''", "S") & "  and Q_FINAL_DEFECT_CAUSE_TOTAL.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.YR = '" & strYr2 & "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Mnth = '" & strMnth2 &_
	              "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Item_Cd = " & FilterVar(strItemCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.INSP_ITEM_CD Like " & FilterVar(strInspItemCd, "''", "S") & " and DEFECT_TYPE_CD LIKE " & FilterVar(strDefectTypeCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY > 0   " &_
	              "GROUP BY Q_DEFECT_CAUSE.DEFECT_CAUSE_NM Order By Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) Desc"

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
	
	
	'�ҷ����� 
	if	TempDefectTypeCd <> "" Then
	strSql = "SELECT DEFECT_TYPE_CD " &_
			"FROM Q_DEFECT_TYPE " &_
			"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND INSP_CLASS_CD = " & FilterVar("P", "''", "S") & " "  			
	'�����˻�(P),�����˻�(F),���ϰ˻�(S)


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
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)	'������ �������� �ʽ��ϴ�.
			RS.Close
			Set RS = Nothing												'��: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'��: ComProxy Unload
			Exit Function
		End If
	RS.Close
	End if	
	
	'�˻��׸� 
	If TempInspItemCd <> "" Then
		strSql = "SELECT INSP_ITEM_CD " &_
				"FROM Q_INSPECTION_STANDARD_BY_ITEM " &_
				"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") &_
				" AND INSP_ITEM_CD = " & FilterVar(TempInspItemCd, "''", "S")
        
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
	
'**********************�߰� �κг�*************************************	


	Set RS1 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	RS1.Open  strSql1, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
	       Call CheckSQLError(Conn,True)
	    End If
		Set RS1 = Nothing
		Conn.Close
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	Set RS2 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
		RS1.Close
		Conn.Close
		Set RS1 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	RS2.Open  strSql2, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
	       Call CheckSQLError(Conn,True)
	    End If
		RS1.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	If RS1.EOF or RS1.BOF Then
		blnNowFlag = False	
	Else
		blnNowFlag = True
	End If

	If RS2.EOF or RS2.BOF then
		blnBeforeFlag = False								'��: �����Ͻ� ���� ó���� ������ 
	Else
		blnBeforeFlag = True
	End If
	
	If blnBeforeFlag = False and blnNowFlag = False Then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing		
		Exit Function
	End If

	'���ڵ尡 �ִٸ� 
	intRecordCount1 = RS1.RecordCount
	intRecordCount2 = RS2.RecordCount
	
	GetData = True
End Function

Function TransferData()
	'------------------------------------------------------------------------------------------------------------------------
	'��� 
	'------------------------------------------------------------------------------------------------------------------------
	If blnNowFlag Then
		ReDim strSpdData1(intRecordCount1)
		ReDim dblSpdData1(3, intRecordCount1)
		
		Total1 = 0

		RS1.MoveFirst	
		For i = 0 to intRecordCount1 - 1 
			strSpdData1(i) = RS1(0)								'�ҷ����� 
			dblSpdData1(0, i) = CDbl(RS1(1))					'�̹��� �ҷ��� 
			Total1 = Total1 + dblSpdData1(0, i)					'�� �ҷ��� 
			RS1.MoveNext				
		Next
		
		aggrDTotal = 0
		For i = 0 to intRecordCount1 - 1
			If Total1 <> 0 Then
				dblSpdData1(1,i) = (dblSpdData1(0,i) / Total1) * 100		'������ 
				aggrDTotal = aggrDTotal + dblSpdData1(0,i) 					
				dblSpdData1(2,i) = aggrDTotal								'�����ҷ��� 
				dblSpdData1(3,i) = (aggrDTotal / Total1) * 100				'���� ������ 
			End If
		Next
	End If
	
	'------------------------------------------------------------------------------------------------------------------------
	'���� 
	'------------------------------------------------------------------------------------------------------------------------
	If blnBeforeFlag Then 
		ReDim strSpdData2(intRecordCount2)
		ReDim dblSpdData2(3, intRecordCount2)
		
		Total2 = 0
		
		RS2.MoveFirst	
		For i = 0 to intRecordCount2 - 1 
			strSpdData2(i) = RS2(0)								'�ŷ�ó�� 
			dblSpdData2(0, i) = CDbl(RS2(1))					'������ �ҷ��� 
			Total2 = Total2 + dblSpdData2(0, i)					'�� �ҷ��� 
			RS2.MoveNext
		Next

		aggrDTotal = 0
		For i = 0 to intRecordCount2 - 1
			If Total2 <> 0 Then
				dblSpdData2(1,i) = (dblSpdData2(0,i) / Total2) * 100		'������ 
				aggrDTotal = aggrDTotal + dblSpdData2(0,i)					'���� �ҷ��� 
				dblSpdData2(2,i) = aggrDTotal
				dblSpdData2(3,i) = (aggrDTotal / Total2) * 100					'���� ������ 
			End If
		Next
	End If
End Function
%>
