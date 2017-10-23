<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<%Option Explicit%>
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
'*  3. Program ID           : Q3411MB1
'*  4. Program Name         : ǰ������(�Ϻ�)
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
	

	Dim TempstrPlantCd
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

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim Conn

Dim RS	'@@@���� 
Dim RS1
Dim RS2
Dim RS3

Dim LngRow

Dim strSql	'@@@���� 
Dim strSql1
Dim strSql2
Dim strSql3

Dim intRecordCount1
Dim intRecordCount2

Dim lgDataFlag          
Dim strPlantCd
Dim strYr
Dim strMnth
Dim strItemCd
Dim TagetFlag

Dim strSpdData(7, 31)
Dim dblSpdData(7, 31)
	
Dim i, j
Dim p
Dim lgParameter												'�ҷ��� ��� 
Dim DayCount													'��¥�� �� 28, 29, 30, 31 
Dim Total
Dim TargetFlag
Dim TransTarget

Dim QMaxDRatio
Dim QMinDRatio
Dim LMaxDRatio
Dim LMinDRatio
Dim TermRatio

Dim blnRet

'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
Dim RS4
Dim strSql4
Dim lgTargetParameter
Dim lgMonthlyTargetValue
'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END *

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
%>
<SCRIPT Language=vbscript>
' Diplay datas on spread sheet.
	Call Parent.DBQueryFailed
</SCRIPT>
<%
	Response.End
End If

' Calculate datas for display
Call TransferData()

If lgDataFlag = "Q"  then
	Call CalculateDataForQ
Else
	Call CalculateDataForL
End If

Call DBClose		
		
%>	

<SCRIPT Language=vbscript>
' Diplay datas on spread sheet.
Call DisplayOnSpread()

' Draw the chart
If <%=(lgDataFlag = "Q")%>  then
'	Call DrawChartForQ
Else
'	Call DrawChartForL
End If

Call Parent.DBQueryOK

'=================================================================================
' Diplay datas on spread sheet.
'=================================================================================
Sub DisplayOnSpread()

	With Parent.frm1.vspdData
<%
		For i = 0 to 7
%>
			.Row = <%= i + 1%>					'�������忡 �� 
<%												'�ѷ��ֱ� 
			For j = 0 to DayCount
%>
				.Col = <%= j + 1%>

				Select Case <%=i%> 
					Case 5, 6, 7				'�˻� �ҷ���, ��Ʈ�ҷ���, ��ǥ 
						If "<%=strSpdData(i, j)%>" <> "" then
							.Text = "<%=UNINumClientFormat(dblSpdData(i, j), 2, 0)%>"
						End If
					Case Else					'���� 
						If "<%=strSpdData(i, j)%>" <> "" then
							.Text = "<%=UNINumClientFormat(dblSpdData(i, j), ggQty.DecPoint, 0)%>"
						End If
				End Select
<%
			Next
		Next
%>
	End With
	
End Sub

'=================================================================================
' Draw the chart for Inspection Defect Ratio.
'=================================================================================
Sub DrawChartForQ()
	With Parent.frm1.ChartFX1
		.Gallery = 1
		.Axis(<%=AXIS_Y%>).AutoScale = False
		.Chart3D = False	'2D
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		
		'Y���� Min/Max ���� 
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(QMinDRatio, 4, 0)%>") 
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(QMaxDRatio, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio, 4, 0)%>")
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.Title_(0) = "�Ϻ� �ҷ��� ���̵�"
		.SerLeg(0) = "�ҷ���"
		.SerLeg(1) = "��ǥġ"
		
		.OpenDataEx <%=COD_VALUES%>,2, <%=DayCount - 1%>								'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
		For i = 0 to DayCount - 1
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = "<%=Cstr(i+1)%>" & "��"
			'�ҷ��� 
			if  "<%=strSpdData(5, i)%>" = "" then
				.Series(0).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
			Else
				.Series(0).Yvalue(<%=i%>) =  parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(5, i), 4, 0)%>")
			End If
<%
		Next
%>
		.Series(0).Visible = True
		
		'��ǥġ�� �ִٸ� 
		If <%= (TargetFlag = True) %> Then					
<%
			For i = 0 to DayCount - 1
%>
				If  "<%=strSpdData(7, i)%>" = "" then
					.Series(1).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
				Else
					.Series(1).Yvalue(<%=i%>) =  parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(7, i), 4, 0)%>")
				End If
<%
			Next
%>
			.Series(1).Visible = True
		End If
		
		' Close the VALUES channel
		.CloseData <%=COD_VALUES%>											'��Ʈ FX���� ä�� �ݾ��ֱ� 
	End With
End Sub

'=================================================================================
' Draw the chart for lot rejection ratio.
'=================================================================================
Sub DrawChartForL()
	With Parent.frm1.ChartFX1
		.Gallery = 1
		.Axis(<%=AXIS_Y%>).AutoScale = False
		.Chart3D = False	'2D
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		
		'Y���� Min/Max ���� 
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(LMinDRatio, 4, 0)%>") 
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(LMaxDRatio, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio, 4, 0)%>")
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
			
		.Title_(0) = "�Ϻ� LOT���հݷ� ���̵�"
		.SerLeg(0) = "LOT���հݷ�"
		
		.OpenDataEx <%=COD_VALUES%>, 1, <%=DayCount - 1%>								'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
		For i = 0 to DayCount - 1
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = "<%=Cstr(i+1)%>" & "��"
			'LOT���հݷ� 
			if  "<%=strSpdData(6, i)%>" = "" then
				.Series(0).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
			Else
				.Series(0).Yvalue(<%=i%>) =  parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(6, i), 4, 0)%>")
			End If
<%
		Next
%>
		.Series(0).Visible = True

		' Close the VALUES channel
		.CloseData <%=COD_VALUES%>											'��Ʈ FX���� ä�� �ݾ��ֱ� 
	End With
End Sub

</SCRIPT>	
<%
Function DataReceive()
	DataReceive = False

	lgDataFlag = Request("txtDataFlag")
	strPlantCd  = Request("txtPlantCd")
	strYr=Request("txtYr")
	strMnth = Request("txtMnth")
	strItemCd = Request("txtItemCd")
	DayCount = CInt(Request("txtCTotal")) - 1
	
	If strPlantCd="" or strYr ="" or strMnth="" or strItemCd = "" Then
		'�Ʒ��� ���Ƿ� �� �޽����̴�.
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
	'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	RS1.Close
	RS2.Close
	Conn.Close
	Set RS2 = Nothing
	Set RS1 = Nothing
	Set Conn = Nothing
	'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END */
End Sub

Function GetData()
	GetData = False
	On Error Resume Next	
	
	' SQL ��� ���� 
	strSql1 = "SELECT Day(RELEASE_DT), Count(INSP_REQ_NO), Sum(LOT_SIZE), Sum(INSP_QTY), Sum(DEFECT_QTY) " & _
			   " From Q_INSPECTION_RESULT " & _
			  " WHERE PLANT_CD= " & FilterVar(strPlantCd, "''", "S") & _
			    " and ITEM_CD= " & FilterVar(strItemCd, "''", "S") & _
			    " and YEAR(RELEASE_DT) = " & strYr & " and MONTH(RELEASE_DT)= " & strMnth & _
			    " and INSP_CLASS_CD=" & FilterVar("F", "''", "S") & "  and status_flag=" & FilterVar("R", "''", "S") & "  " & _
			  " GROUP BY RELEASE_DT ORDER BY RELEASE_DT"
				     
	strSql2 = "SELECT Day(RELEASE_DT), Count(DECISION) " & _
			   " From Q_INSPECTION_RESULT " & _
			  " WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & _
			    " and ITEM_CD= " & FilterVar(strItemCd, "''", "S") & _
			    " and YEAR(RELEASE_DT)= " & strYr & " and MONTH(RELEASE_DT) = " & strMnth & _
				" and INSP_CLASS_CD=" & FilterVar("F", "''", "S") & "  and status_flag=" & FilterVar("R", "''", "S") & "  and DECISION=" & FilterVar("R", "''", "S") & "  " & _
			  " GROUP BY RELEASE_DT ORDER BY RELEASE_DT"

	'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	strSql3 = " SELECT a.DEFECT_RATIO_UNIT_CD,  b.PARAMETER " & _
				" From Q_DEFECT_RATIO_BY_INSP_CLASS a " & _
				" INNER JOIN Q_DEFECT_RATIO_UNIT b ON a.DEFECT_RATIO_UNIT_CD = b.DEFECT_RATIO_UNIT_CD " & _
			   " WHERE a.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & _
				 " AND a.INSP_CLASS_CD = " & FilterVar("F", "''", "S") 
	
	strSql4 = " SELECT C.TARGET_VALUE, B.PARAMETER " & _
			  " FROM Q_Yearly_Target A inner join Q_DEFECT_RATIO_UNIT B ON A.DEFECT_RATIO_UNIT_CD = B.DEFECT_RATIO_UNIT_CD " & _
				" INNER JOIN Q_MONTHLY_TARGET C ON A.PLANT_CD = C.PLANT_CD AND A.INSP_CLASS_CD = C.INSP_CLASS_CD AND A.YR = C.YR " & _
			   " Where A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & _
			   " and A.YR= " & FilterVar(strYr,"''","S") & _
			   " and A.INSP_CLASS_CD = " & FilterVar("F", "''", "S") & _
			     " AND C.MNTH = " & FilterVar(strMnth,"''","S")
	
	'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END */
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
		Conn.Close
		Set RS1 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If RS1.EOF or RS1.BOF Then											'���հ� ��Ʈ ���� ���� �� �� �����Ƿ� ���ǿ� ���� �ʴ´�.
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		RS1.Close
		Conn.Close
		Set RS1 = Nothing
		Set Conn = Nothing
		Exit Function
	End If
	'���ڵ尡 �ִٸ� 
	intRecordCount1 = RS1.RecordCount

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
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	'���ڵ尡 �ִٸ� 
	intRecordCount2 = RS2.RecordCount
	
	'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	
	Set RS3 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	RS3.Open  strSql3, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
	       Call CheckSQLError(Conn,True)
	    End If
		RS1.Close
		RS2.Close
		RS3.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set RS3 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	'�ҷ��������ڵ忡 ���� ����� ���ٸ� 
	If RS3.EOF or RS3.BOF then
		Call DisplayMsgBox("221205", vbOKOnly, "", "", I_MKSCRIPT)	'�ҷ��� ����� �����ϴ�.
		RS1.Close
		RS2.Close
		RS3.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set RS3 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If

	lgParameter = CDbl(RS3(1))
	
	RS3.Close
	Set RS3 = Nothing
	
	'������ǥ ���̺��� �ҷ��� ���� �� ��� �б� 
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	RS4.Open  strSql4, Conn, 1			'adOpenKeyset
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		If CheckSYSTEMError(Err,True) = False Then
	       Call CheckSQLError(Conn,True)
	    End If
		RS1.Close
		RS2.Close
		RS4.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set RS4 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	'����ǥ ���̺��� �ҷ��� ���� �� ����� ���ٸ� 
	If RS4.EOF or RS4.BOF then
		RS4.Close
		Set RS4 = Nothing
		
		TargetFlag = False
	Else
		TargetFlag = True
		lgMonthlyTargetValue = CDbl(RS4(0))
		lgTargetParameter = CDbl(RS4(1))
		
		RS4.Close
		Set RS4 = Nothing
	
		' ��ǥ���� ������ �ҷ��� ������ �ٸ��� ���� ȯ������ 
		If lgTargetParameter <> lgParameter Then
			TransTarget = lgMonthlyTargetValue / lgTargetParameter * lgParameter
		Else
			TransTarget = lgMonthlyTargetValue
		End If
	End If
	
	'/* [2005-10-24] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	

	GetData = True
End Function

Function TransferData()
	Dim intDay, intDay2
	Dim FirstFlag
	Dim Target
	
	For LngRow = 0 To intRecordCount1 -1
		
		intDay = CInt(RS1(0))
	
		strSpdData(0, intDay - 1) = RS1(1)	'Lot�� 
		
		If strSpdData(0, intDay - 1) = "0" Then
			strSpdData(0, intDay - 1) = ""	'Lot�� 
			strSpdData(1, intDay - 1) = ""	'���հ�Lot�� 
			strSpdData(2, intDay - 1) = ""	'�԰�� 
			strSpdData(3, intDay - 1) = ""	'�˻�� 
			strSpdData(4, intDay - 1) = ""	'�ҷ��� 
			
			dblSpdData(0, intDay - 1) = 0	'Lot�� 
			dblSpdData(1, intDay - 1) = 0	'���հ�Lot�� 
			dblSpdData(2, intDay - 1) = 0	'�԰�� 
			dblSpdData(3, intDay - 1) = 0	'�˻�� 
			dblSpdData(4, intDay - 1) = 0	'�ҷ��� 
		
		Else
			'���հ�Lot�� 
			strSpdData(1, intDay - 1) = "0"	
			dblSpdData(1, intDay - 1) = 0
			
			strSpdData(2, intDay - 1) = RS1(2)	'�԰�� 
			strSpdData(3, intDay - 1) = RS1(3)	'�˻�� 
			strSpdData(4, intDay - 1) = RS1(4)	'�ҷ��� 
			
			dblSpdData(0, intDay - 1) = CDbl(RS1(1))	'Lot�� 
			dblSpdData(2, intDay - 1) = CDbl(RS1(2))	'�԰�� 
			dblSpdData(3, intDay - 1) = CDbl(RS1(3))	'�˻�� 
			dblSpdData(4, intDay - 1) = CDbl(RS1(4))	'�ҷ��� 
		End If
		
		'�ҷ��� 
		If strSpdData(3, intDay -1) = "" Then
			strSpdData(5, intDay -1) = ""
		Else							
			'�˻���� 0�� �ƴ� ��츸..
			dblSpdData(5, intDay -1) = (dblSpdData(4, intDay -1) / dblSpdData(3, intDay -1)) * lgParameter
			strSpdData(5, intDay -1) = CStr(dblSpdData(5, intDay -1))
				
			If FirstFlag = False Then
				QMaxDRatio = dblSpdData(5, intDay -1)
				QMinDRatio = dblSpdData(5, intDay -1)
				FirstFlag = True
			End If
							
			If dblSpdData(5, intDay -1) > QMaxDRatio Then
				QMaxDRatio = dblSpdData(5, intDay -1)
			End If
					
			If dblSpdData(5, intDay -1) < QMinDRatio Then
				QMinDRatio = dblSpdData(5, intDay -1)
			End If
		End If
		
		'LOT���հݷ� 
		If strSpdData(0, intDay -1) = "" Then
			strSpdData(6, intDay -1) = ""
		Else							
			'LOT���� 0�� �ƴ� ��츸..
			dblSpdData(6, intDay -1) = 0
			strSpdData(6, intDay -1) = "0"
		End If
					
		RS1.MoveNext
	Next
	'���հ�Lot�� 
	For LngRow = 0 To intRecordCount2 -1
		
		intDay2 = CInt(RS2(0))
		
		strSpdData(1, intDay2 - 1) = RS2(1)	
		dblSpdData(1, intDay2 - 1) = CDbl(RS2(1))
		
		'LOT���հݷ� 
		If strSpdData(0, intDay2 -1) = "" Then
			strSpdData(6, intDay2 -1) = ""
		Else							
			'LOT���� 0�� �ƴ� ��츸..
			dblSpdData(6, intDay2 -1) = (dblSpdData(1, intDay2 -1) / dblSpdData(0, intDay2 -1)) * 100
			strSpdData(6, intDay2 -1) = CStr(dblSpdData(6, intDay2 -1))
			
			If FirstFlag = False Then
				LMaxDRatio = dblSpdData(6, intDay2 -1)
				LMinDRatio = dblSpdData(6, intDay2 -1)
				FirstFlag = True
			End If
												
			If dblSpdData(6, intDay2 -1) > LMaxDRatio Then
				LMaxDRatio = dblSpdData(6, intDay2 -1)
			End If
			If dblSpdData(6, intDay2 -1) < LMinDRatio Then
				LMinDRatio = dblSpdData(6, intDay2 -1)
			End If
		End If
		
		RS2.MoveNext
	Next
	
	For i = 0 to 4							'�հ� ���ϱ� 
		Total = 0
		For j = 0 to DayCount - 1
			Total = Total + dblSpdData(i,j)
		Next
		dblSpdData(i,DayCount) = Total
		strSpdData(i,DayCount) = CStr(Total)
	Next
	
	If dblSpdData(0,DayCount) = 0 Then
		' �հ迡 ���� �˻�ҷ��� ���ϱ� 
		dblSpdData(5,DayCount) = 0
		strSpdData(5,DayCount) = ""
		
		'�հ迡 ���� ��Ʈ���հݷ� ���ϱ�		
		dblSpdData(6,DayCount) = 0
		strSpdData(6,DayCount) = ""
	Else			
		' �հ迡 ���� �˻�ҷ��� ���ϱ� 
		If dblSpdData(3,DayCount) <> 0 and dblSpdData(4,DayCount) = 0 Then
			dblSpdData(5,DayCount) = 0
			strSpdData(5,DayCount) = "0"
		ElseIf dblSpdData(3,DayCount) <> 0 and dblSpdData(4,DayCount) <> 0 Then
			dblSpdData(5,DayCount) = (dblSpdData(4,DayCount) / dblSpdData(3,DayCount)) * lgParameter
			strSpdData(5,DayCount) = CStr(dblSpdData(5,DayCount))
		End If
			
		'�հ迡 ���� ��Ʈ���հݷ� ���ϱ�		
		If dblSpdData(0,DayCount) <> 0 and dblSpdData(1,DayCount) =  0 Then
			dblSpdData(6,DayCount) = 0
			strSpdData(6,DayCount) = "0"
		ElseIf dblSpdData(0,DayCount) <> 0 and dblSpdData(1,DayCount) <>  0 Then
			dblSpdData(6,DayCount) = (dblSpdData(1,DayCount) / dblSpdData(0,DayCount)) * 100
			strSpdData(6,DayCount) = CStr(dblSpdData(6,DayCount))
		End If
	End If
				
	'��ǥ���� �ִٸ� 
	
	If TargetFlag = True Then
		'����ǥ 
		strSpdData(7, DayCount) = TransTarget
		Target = TransTarget
		dblSpdData(7, DayCount) = Target				
		'�ϸ�ǥ 
		For i = 0 to DayCount - 1
			strSpdData(7,i) = TransTarget
			dblSpdData(7,i) = Target
		Next
		
		If QMaxDRatio < Target Then
			QMaxDRatio = Target					'��ǥġ�� Max���� Ŭ��� 
		End If
		If QMinDRatio > Target Then
			QMinDRatio = Target					'��ǥġ�� Min���� ���� ��� 
		End If
	End If
End Function

'�˻�ҷ����� ���� ����Ÿ ���ϱ� 
Function CalculateDataForQ()
	'ChartFX�� Min/Max/Step ���� 
	If QMaxDRatio = 0 Then
		QMaxDRatio = lgParameter / 10
	Else
		If QMaxDRatio + (QMaxDRatio / 10) > lgParameter Then
			QMaxDRatio = lgParameter
		Else
			QMaxDRatio = QMaxDRatio + (QMaxDRatio / 10)
		End If
		
		If QMinDRatio - (QMinDRatio / 10) < 0 Then
			QMinDRatio = 0
		Else
			QMinDRatio = QMinDRatio - (QMinDRatio / 10)
		End If
	End If	
			
	TermRatio = (QMaxDRatio - QMinDRatio) / 10
End Function

'��Ʈ�ҷ����� ���� ����Ÿ ���ϱ� 
Function CalculateDataForL()
	'ChartFX�� Min/Max/Step ���� 
	If LMaxDRatio = 0 Then
		LMaxDRatio = 10
	Else
		If LMaxDRatio + (LMaxDRatio / 10) > 100 Then
			LMaxDRatio = 100
		Else
			LMaxDRatio = LMaxDRatio + (LMaxDRatio / 10)
		End If
		
		If LMinDRatio - (LMinDRatio / 10) < 0 Then
			LMinDRatio = 0
		Else
			LMinDRatio = LMinDRatio - (LMinDRatio / 10)
		End If
	End If	
			
	TermRatio = (LMaxDRatio - LMinDRatio) / 10
End Function
%>
