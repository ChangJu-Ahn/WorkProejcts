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
'*  3. Program ID           : q3211mb2.asp
'*  4. Program Name         : ���԰˻�ǰ������(��/�б⺰)
'*  5. Program Desc         : ���԰˻�ǰ�����̸� ��ȸ�Ѵ�.
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2000/05/
'*  8. Modified date(Last)  : 2000/08/23
'*  9. Modifier (First)     : Oh Youngjoon
'* 10. Modifier (Last)      : Oh Youngjoon
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
	Dim TempstrItemCd
	Dim TempstrBpCd
	
	TempstrPlantCd		= "<%=Request("txtPlantCd")%>"
	TempstrItemCd		= "<%=Request("txtItemCd")%>"
	TempstrBpCd		= "<%=Request("txtBpCd")%>"
	
	'/* [2005-10-26] ������ ��ġ���� �ʽ��ϴ� ���� ���� ����: FilterVar --> parent.FilterVar - START */
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

	'����ó �ҷ����� 
	Call parent.CommonQueryRs("BP_CD,BP_NM","B_BIZ_PARTNER","BP_CD =  " & parent.FilterVar(TempstrBpCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtBpNm.Value = strVar2
	'/* [2005-10-26] ������ ��ġ���� �ʽ��ϴ� ���� ���� ����: FilterVar --> parent.FilterVar - END */
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

Dim Total	'@@@���� 

Dim intRecordCount1
Dim intRecordCount3

Dim TempBpCd
Dim lgDataFlag          
Dim strPlantCd
Dim strItemCd
Dim strYr
Dim strBpCd
Dim TargetFlag
Dim TransTargetM
Dim TransTargetY

Dim strSpdData(9, 12)
Dim dblSpdData(9, 12)

Dim i, j
Dim lgParameter
Dim QMaxDRatio1
Dim QMinDRatio1
Dim QMaxDRatio2
Dim QMinDRatio2
Dim LMaxDRatio1
Dim LMinDRatio1
Dim LMaxDRatio2
Dim LMinDRatio2
Dim TermRatio1
Dim TermRatio2


Dim blnRet

'/* [2005-10-26] ������ ��ġ���� �ʽ��ϴ� ���� ���� ����: FilterVar --> parent.FilterVar - START */
Dim RS4
Dim strSql4
Dim lgTargetParameter
Dim lgYearlyTargetValue
'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END */

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
    Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write " Call parent.DBQueryErr " & vbCrLf
	Response.Write "</Script>" & vbCrLf
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

<Script Language=vbscript>

' Diplay datas on spread sheet.
Call DisplayOnSpread()



Call Parent.DBQueryOK

'=================================================================================
' Diplay datas on spread sheet.
'=================================================================================
Sub DisplayOnSpread()

	With Parent.frm1.vspdData
<%
		For i = 0 to 7
%>
			.Row = <%=i + 1%>					'�������忡 �� 
<%												'�ѷ��ֱ� 
			For j = 0 to 12
%>
				.Col = <%=j + 1%>

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
		'��ƮFX1 - ���� ��Ʈ 
		.Gallery = 1		'������ �׷��� 
		.Axis(<%=AXIS_Y%>).AutoScale = False
		.Chart3D = False	'2D
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		
		'Y���� Min/Max ���� 
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(QMinDRatio1, 4, 0)%>") 
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(QMaxDRatio1, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio1, 4, 0)%>")
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.Title_(0) = "���� �ҷ��� ���̵�"
		.SerLeg(0) = "�ҷ���"
		.SerLeg(1) = "��ǥġ"
		
		.OpenDataEx <%=COD_VALUES%>, 2, 12					'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
		For i = 0 to 11
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = Cstr(<%=i+1%>) & "��"
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
			For i = 0 to 11
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
		.CloseData <%=COD_VALUES%>						'��Ʈ FX���� ä�� �ݾ��ֱ� 
	End With
	
	With Parent.frm1.ChartFX2
		'��ƮFX2 - ������ ��Ʈ 
		.Gallery = 2
		.Axis(<%=AXIS_Y%>).AutoScale = False
		.Chart3D = False	'2D
		.Series(0).visible = True
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(QMinDRatio2, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(QMaxDRatio2, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio2, 4, 0)%>")

		
		.Title_(0) = "�б⺰ �ҷ��� ���̵�"
		.SerLeg(0) = "�ҷ���"

		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.OpenDataEx <%=COD_VALUES%>,1,4
<%
		For i = 0 to 3
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = Cstr(<%=i+1%>) & "�б�"
			if  "<%=strSpdData(8, i)%>" = "" then
				.Series(0).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
			Else
				.Series(0).Yvalue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(8, i), 4, 0)%>")
			End If
<%
		Next
%>	
		' Close the VALUES channel
		.CloseData <%=COD_VALUES%>
	End with
End Sub

'=================================================================================
' Draw the chart for lot rejection ratio.
'=================================================================================
Sub DrawChartForL()
	With Parent.frm1.ChartFX1
		'��ƮFX1 - ���� ��Ʈ 
		.Gallery = 1		'������ �׷��� 
		.Axis(<%=AXIS_Y%>).AutoScale = False
		.Chart3D = False	'2D
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		
		'Y���� Min/Max ���� 
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(LMinDRatio1, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(LMaxDRatio1, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio1, 4, 0)%>")
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.Title_(0) = "���� LOT���հݷ� ���̵�"
		.SerLeg(0) = "LOT���հݷ�"
		
		.OpenDataEx <%=COD_VALUES%>, 1, 12					'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
		For i = 0 to 11
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = Cstr(<%=i+1%>) & "��"
			'LOT���հݷ� 
			if  "<%=strSpdData(6, i)%>" = "" then
				.Series(0).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
			Else
				.ValueEx(0,<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(6, i), 4, 0)%>")
			End If
<%
		Next
%>
		.Series(0).Visible = True
		
		' Close the VALUES channel
		.CloseData <%=COD_VALUES%>						'��Ʈ FX���� ä�� �ݾ��ֱ� 
	End With
	
	With Parent.frm1.ChartFX2
		'��ƮFX2 - ������ ��Ʈ 
		.Gallery = 2
		.Axis(<%=AXIS_Y%>).AutoScale = False
		.Chart3D = False	'2D
		.Series(0).visible = True
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(LMinDRatio2, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(LMaxDRatio2, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio2, 4, 0)%>")
		
		.Title_(0) = "�б⺰ LOT���հݷ� ���̵�"
		.SerLeg(0) = "LOT���հݷ�"

		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.OpenDataEx <%=COD_VALUES%>,1,4
<%
		For i = 0 to 3
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = Cstr(<%=i+1%>) & "�б�"
			if  "<%=strSpdData(9, i)%>" = "" then	
				.Series(0).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
			Else
				.Series(0).Yvalue(<%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(9, i), 4, 0)%>")
			End If
<%
		Next
%>	
		' Close the VALUES channel
		.CloseData <%=COD_VALUES%>
	End with
End Sub

</Script>

<%
Function DataReceive()
	DataReceive = False

	lgDataFlag = Request("txtDataFlag")
	strPlantCd = Request("txtPlantCd")
	strItemCd = Request("txtItemCd")
	strYr = Request("txtYr")
	strBpCd = Request("txtBpCd")
	TempBpCd = Request("txtBpCd")

	If strPlantCd="" or strItemCd="" or strYr="" Then
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
	On Error Resume Next	
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	RS1.Close
	RS2.Close
	Conn.Close
	Set RS2 = Nothing
	Set RS1 = Nothing
	Set Conn = Nothing
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END */
End Sub

Function GetData()
	GetData = False
	On Error Resume Next	
	
	If strBpCd <> "" Then
		strSql1 = " SELECT Mnth, RECEIVING_LOT_CNT, REJT_LOT_CNT, LOT_SIZE_QTY, INSP_QTY, DEFECT_QTY " & _
					" FROM Q_Receiving_Inspection_Total " & _
				   " Where PLANT_CD= " & FilterVar(strPlantCd, "''", "S") & _
					 " and ITEM_CD= " & FilterVar(strItemCd, "''", "S") & _
					 " and BP_CD =" & FilterVar(strBpCd, "''", "S") & _
					 " and YR = " & strYr & _
				   " ORDER BY MNTH"
	Else
		strSql1 = " SELECT Mnth, Sum(RECEIVING_LOT_CNT), Sum(REJT_LOT_CNT), Sum(LOT_SIZE_QTY), Sum(INSP_QTY), Sum(DEFECT_QTY) " & _
					" FROM Q_Receiving_Inspection_Total " & _
				   " Where PLANT_CD= " & FilterVar(strPlantCd, "''", "S") & _
					 " and ITEM_CD= " & FilterVar(strItemCd, "''", "S") & _
					 " and YR = " & strYr & _
				   " GROUP BY MNTH ORDER BY MNTH"
	End If
	
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	strSql2 = " SELECT a.DEFECT_RATIO_UNIT_CD, b.PARAMETER " & _
			  " From Q_DEFECT_RATIO_BY_INSP_CLASS a " & _
			  " INNER JOIN Q_DEFECT_RATIO_UNIT b ON a.DEFECT_RATIO_UNIT_CD = b.DEFECT_RATIO_UNIT_CD " & _
			  " WHERE a.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & _
			  " AND a.INSP_CLASS_CD = " & FilterVar("R", "''", "S") & " "
			  
	strSql3 = " SELECT A.TARGET_VALUE, B.PARAMETER " & _
			  " FROM Q_YEARLY_TARGET A inner join Q_DEFECT_RATIO_UNIT B " & _
			  " ON A.DEFECT_RATIO_UNIT_CD = B.DEFECT_RATIO_UNIT_CD " & _
			  " WHERE A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & _
			  " AND A.YR = " & FilterVar(strYr,"''","S") & _
			  " AND A.INSP_CLASS_CD = " & FilterVar("R","''","S")
			  
	strSql4 = " SELECT Cast(MNTH AS INT), TARGET_VALUE " & _
			  " FROM Q_MONTHLY_TARGET " & _
			  " WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & _
			  " AND YR = " & FilterVar(strYr,"''","S") & _
			  " AND INSP_CLASS_CD = " & FilterVar("R","''","S") & _
			  " ORDER BY MNTH "
	
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END */
	
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
				"FROM B_Plant " &_
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
	
	
	If TempBpCd <> "" Then		 	 
	strSql = "SELECT BP_CD " &_
			"FROM B_BIZ_PARTNER " &_
			"WHERE (BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("S", "''", "S") & " ) AND	BP_CD = " & FilterVar(TempBpCd, "''", "S")
	
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
			Call DisplayMsgBox("229927", vbOKOnly, "", "", I_MKSCRIPT)	'������ �������� �ʽ��ϴ�.
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
	If RS1.EOF or RS1.BOF Then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'���ǿ� �´� �˻����� �����ϴ� 
		RS1.Close
		Conn.Close
		Set RS1 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	'���ڵ尡 �ִٸ� 
	intRecordCount1 = RS1.RecordCount
	
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	
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

	'�ҷ��������ڵ忡 ���� ����� ���ٸ� 
	If RS2.EOF or RS2.BOF then
		Call DisplayMsgBox("221205", vbOKOnly, "", "", I_MKSCRIPT)	'�ҷ��� ����� �����ϴ�.
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	lgParameter = CDbl(RS2(1))
	
	RS2.Close
	Set RS2 = Nothing
	
	'����ǥ ���̺��� �ҷ��� ���� �� ��� �б� 
	Set RS3 = Server.CreateObject("ADODB.Recordset")
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
	
	RS3.Open  strSql3, Conn, 1
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
		Set RS3 = Nothing
		Set Conn = Nothing												'��: ComProxy Unload
		Exit Function
	End If
	
	'����ǥ ���̺��� �ҷ��� ���� �� ����� ���ٸ� 
	If RS3.EOF or RS3.BOF then
		RS3.Close
		Set RS3 = Nothing
		
		TargetFlag = False
	Else
		TargetFlag = True
		
		lgYearlyTargetValue = CDbl(RS3(0))
		lgTargetParameter = CDbl(RS3(1))
		RS3.Close
		Set RS3 = Nothing

		' ��ǥ���� ������ �ҷ��� ������ �ٸ��� ���� ȯ������ 
		If lgTargetParameter <> lgParameter Then
			TransTargetY = lgYearlyTargetValue / lgTargetParameter * lgParameter
		Else
			TransTargetY = lgYearlyTargetValue
		End If

		Set RS4 = Server.CreateObject("ADODB.Recordset")
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

		RS4.Open  strSql4, Conn, 1
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
			Set RS4 = Nothing
			Set Conn = Nothing												'��: ComProxy Unload
			Exit Function
		End If

		'�� ��ǥ���� ���ٸ� 
		If RS4.EOF or RS4.BOF then
			intRecordCount3 = 0
		Else
			intRecordCount3 = RS4.RecordCount
		End If
	End If
	
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */
	
	GetData = True
End Function

Function TransferData()
	Dim FirstFlag
	Dim intMonth
	Dim Target
	
	FirstFlag = False				
	For LngRow = 0 To intRecordCount1 -1
		intMonth = CInt(RS1(0))
		
		strSpdData(0, intMonth - 1) = RS1(1)	'Lot�� 
		
		If strSpdData(0, intMonth - 1) = "0" Then
			strSpdData(0, intMonth - 1) = ""	'Lot�� 
			strSpdData(1, intMonth - 1) = ""	'���հ�Lot�� 
			strSpdData(2, intMonth - 1) = ""	'�԰�� 
			strSpdData(3, intMonth - 1) = ""	'�˻�� 
			strSpdData(4, intMonth - 1) = ""	'�ҷ��� 
			
			dblSpdData(0, intMonth - 1) = 0	'Lot�� 
			dblSpdData(1, intMonth - 1) = 0	'���հ�Lot�� 
			dblSpdData(2, intMonth - 1) = 0	'�԰�� 
			dblSpdData(3, intMonth - 1) = 0	'�˻�� 
			dblSpdData(4, intMonth - 1) = 0	'�ҷ��� 

		Else
			strSpdData(1, intMonth - 1) = RS1(2)	'���հ�Lot�� 
			strSpdData(2, intMonth - 1) = RS1(3)	'�԰�� 
			strSpdData(3, intMonth - 1) = RS1(4)	'�˻�� 
			strSpdData(4, intMonth - 1) = RS1(5)	'�ҷ��� 
			
			dblSpdData(0, intMonth - 1) = CDbl(RS1(1))	'Lot�� 
			dblSpdData(1, intMonth - 1) = CDbl(RS1(2))	'���հ�Lot�� 
			dblSpdData(2, intMonth - 1) = CDbl(RS1(3))	'�԰�� 
			dblSpdData(3, intMonth - 1) = CDbl(RS1(4))	'�˻�� 
			dblSpdData(4, intMonth - 1) = CDbl(RS1(5))	'�ҷ��� 
		End If

		'�ҷ��� 
		If strSpdData(3, intMonth -1) = "" Then
			strSpdData(5, intMonth -1) = ""
		Else							
			'�˻���� 0�� �ƴ� ��츸..
			dblSpdData(5, intMonth -1) = (dblSpdData(4, intMonth -1) / dblSpdData(3, intMonth -1)) * lgParameter
			strSpdData(5, intMonth -1) = CStr(dblSpdData(5, intMonth -1))
				
			If FirstFlag = False Then
				QMaxDRatio1 = dblSpdData(5, intMonth -1)
				QMinDRatio1 = dblSpdData(5, intMonth -1)
				FirstFlag = True
			End If
							
			If dblSpdData(5, intMonth -1) > QMaxDRatio1 Then
				QMaxDRatio1 = dblSpdData(5, intMonth -1)
			End If
					
			If dblSpdData(5, intMonth -1) < QMinDRatio1 Then
				QMinDRatio1 = dblSpdData(5, intMonth -1)
			End If
		End If

		'LOT���հݷ� 
		If strSpdData(0, intMonth -1) = "" Then
			strSpdData(6, intMonth -1) = ""
		Else							
			'LOT���� 0�� �ƴ� ��츸..
			dblSpdData(6, intMonth -1) = (dblSpdData(1, intMonth -1) / dblSpdData(0, intMonth -1)) * 100
			strSpdData(6, intMonth -1) = CStr(dblSpdData(6, intMonth -1))
			
			If FirstFlag = False Then
				LMaxDRatio1 = dblSpdData(6, intMonth -1)
				LMinDRatio1 = dblSpdData(6, intMonth -1)
				FirstFlag = True
			End If
												
			If dblSpdData(6, intMonth -1) > LMaxDRatio1 Then
				LMaxDRatio1 = dblSpdData(6, intMonth -1)
			End If
			If dblSpdData(6, intMonth -1) < LMinDRatio1 Then
				LMinDRatio1 = dblSpdData(6, intMonth -1)
			End If
		End If
			
		RS1.MoveNext
	Next
			
	For i = 0 to 4							'�հ� ���ϱ� 
		Total = 0
		For j = 0 to 11
			Total = Total + dblSpdData(i,j)
		Next
		dblSpdData(i,12) = Total
		strSpdData(i,12) = CStr(Total)
	Next
	
	If dblSpdData(0,12) = 0 Then
		' �հ迡 ���� �˻�ҷ��� ���ϱ� 
		dblSpdData(5,12) = 0
		strSpdData(5,12) = ""
		
		'�հ迡 ���� ��Ʈ���հݷ� ���ϱ�		
		dblSpdData(6,12) = 0
		strSpdData(6,12) = ""
	Else
		' �հ迡 ���� �˻�ҷ��� ���ϱ� 
		If dblSpdData(3,12) <> 0 and dblSpdData(4,12) = 0 Then
			dblSpdData(5,12) = 0
			strSpdData(5,12) = ""
		ElseIf dblSpdData(3,12) <> 0 and dblSpdData(4,12) <> 0 Then
			dblSpdData(5,12) = (dblSpdData(4,12) / dblSpdData(3,12)) * lgParameter
			strSpdData(5,12) = CStr(dblSpdData(5,12))
		End If
			
		'�հ迡 ���� ��Ʈ���հݷ� ���ϱ�		
		If dblSpdData(0,12) <> 0 and dblSpdData(1,12) =  0 Then
			dblSpdData(6,12) = 0
			strSpdData(6,12) = 0
		ElseIf dblSpdData(0,12) <> 0 and dblSpdData(1,12) <>  0 Then
			dblSpdData(6,12) = (dblSpdData(1,12) / dblSpdData(0,12)) * 100
			strSpdData(6,12) = CStr(dblSpdData(6,12))
		End If
	End If
	
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - START */			
	'��ǥ���� �ִٸ� 
	If TargetFlag = True Then
		'����ǥ 
		strSpdData(7, 12) = TransTargetY
		dblSpdData(7, 12) = TransTargetY				
		
		'����ǥ 
		Target = 0
		For LngRow = 0 To intRecordCount3 -1
			If CDbl(RS2(4)) <> lgParameter Then
				TransTargetM = CDbl(RS4(1)) / lgTargetParameter * lgParameter
			Else
				TransTargetM = CDbl(RS4(1))
			End If
	
			strSpdData(7, RS4(0) - 1) = TransTargetM
			dblSpdData(7, RS4(0) - 1) = TransTargetM				
			If dblSpdData(7, RS4(0)-1) > Target Then
				Target = dblSpdData(7, RS4(0) - 1)
			End If
			RS4.MoveNext
		Next
		
		If QMaxDRatio1 < Target Then
			QMaxDRatio1 = Target					'��ǥġ�� Max���� Ŭ��� 
		End If
		If QMinDRatio1 > Target Then
			QMinDRatio1 = Target					'��ǥġ�� Min���� ���� ��� 
		End If
	End If
	'/* [2005-10-26] ������ǥġ ���� ��� Null�� ���� �߻����� ���� - END */
End Function

'�˻�ҷ����� ���� ����Ÿ ���ϱ� 
Function CalculateDataForQ()
	Dim FirstFlag
	
	'ChartFX1�� Min/Max/Step ���� 
	If QMaxDRatio1 = 0 Then
		QMaxDRatio1 = lgParameter / 10
	Else
		If QMaxDRatio1 + (QMaxDRatio1 / 10) > lgParameter Then
			QMaxDRatio1 = lgParameter
		Else
			QMaxDRatio1 = QMaxDRatio1 + (QMaxDRatio1 / 10)
		End If
	
		If QMinDRatio1 - (QMinDRatio1 / 10) < 0 Then
			QMinDRatio1 = 0
		Else
			QMinDRatio1 = QMinDRatio1 - (QMinDRatio1 / 10)
		End If
	End If	
			
	TermRatio1 = (QMaxDRatio1 - QMinDRatio1) / 10

	'�б⺰ �˻�ҷ��� ����ϱ� 
	'�˻��� Lot�� �ִ� ��쿡��..
	If (dblSpdData(0,0) + dblSpdData(0,1) + dblSpdData(0,2)) <> 0 Then
		If (dblSpdData(3,0) + dblSpdData(3,1) + dblSpdData(3,2)) <> 0 Then
			dblSpdData(8,0) = ((dblSpdData(4,0) + dblSpdData(4,1) + dblSpdData(4,2)) / (dblSpdData(3,0) + dblSpdData(3,1) + dblSpdData(3,2))) * lgParameter		'1�б� 
			strSpdData(8,0) = CStr(dblSpdData(8,0))
		Else
			dblSpdData(8,0) = 0
			strSpdData(8,0) = "0"
		End If
	Else
		strSpdData(8,0) = ""
	End If
	
	If (dblSpdData(0,3) + dblSpdData(0,4) + dblSpdData(0,5)) <> 0 Then
		If (dblSpdData(3,3) + dblSpdData(3,4) + dblSpdData(3,5))  <> 0 Then
			dblSpdData(8,1) = ((dblSpdData(4,3) + dblSpdData(4,4) + dblSpdData(4,5)) / (dblSpdData(3,3) + dblSpdData(3,4) + dblSpdData(3,5))) * lgParameter		'2�б� 
			strSpdData(8,1) = CStr(dblSpdData(8,1))
		Else
			dblSpdData(8,1) = 0
			strSpdData(8,1) = "0"
		End If
	Else
		strSpdData(8,1) = ""
	End If
	
	If (dblSpdData(0,6) + dblSpdData(0,7) + dblSpdData(0,8)) <> 0 Then
		If (dblSpdData(3,6) + dblSpdData(3,7) + dblSpdData(3,8)) <> 0 Then
			dblSpdData(8,2) = ((dblSpdData(4,6) + dblSpdData(4,7) + dblSpdData(4,8)) / (dblSpdData(3,6) + dblSpdData(3,7) + dblSpdData(3,8))) * lgParameter		'3�б� 
			strSpdData(8,2) = CStr(dblSpdData(8,2))
		Else
			dblSpdData(8,2) = 0
			strSpdData(8,2) = "0"
		End If
	Else
		strSpdData(8,2) = ""
	End If
	
	If (dblSpdData(0,6) + dblSpdData(0,7) + dblSpdData(0,8)) <> 0 Then					
		If (dblSpdData(3,9) + dblSpdData(3,10) + dblSpdData(3,11)) <> 0 Then
			dblSpdData(8,3) = ((dblSpdData(4,9) + dblSpdData(4,10) + dblSpdData(4,11)) / (dblSpdData(3,9) + dblSpdData(3,10) + dblSpdData(3,11))) * lgParameter	'4�б� 
			strSpdData(8,3) = CStr(dblSpdData(8,3))
		Else
			dblSpdData(8,3) = 0
			strSpdData(8,3) = "0"
		End If
	Else
		strSpdData(8,3) = ""
	End If				
	
	FirstFlag = False
	For i = 0 to 3		
		If strSpdData(8,i) <> "" Then
			If FirstFlag = False Then
				QMaxDRatio2 = dblSpdData(8, i)
				QMinDRatio2 = dblSpdData(8, i)
				FirstFlag = True
			End if
						
			If dblSpdData(8,i) > QMaxDRatio2 Then
				QMaxDRatio2 = dblSpdData(8,i)
			End If
			If dblSpdData(8,i) < QMinDRatio2 Then
				QMinDRatio2 = dblSpdData(8,i)
			End If
		End If
	Next

'***** chart2�� 1�� ���� �����ϰ� ������. 03/07/25 AJJ
	'ChartFX2�� Min/Max/Step ���� 
'	If QMaxDRatio2 = 0 Then
'		QMaxDRatio2 = lgParameter / 10
'	Else
'		If QMaxDRatio2 + (QMaxDRatio2 / 10) > lgParameter Then
'			QMaxDRatio2 = lgParameter
'		Else
'			QMaxDRatio2 = QMaxDRatio2 + (QMaxDRatio2 / 10)
'		End If
		
'		If QMinDRatio2 - (QMinDRatio2 / 10) < 0 Then
'			QMinDRatio2 = 0
'		Else
'			QMinDRatio2 = QMinDRatio2 - (QMinDRatio2 / 10)
'		End If
'	End If	
			
'	TermRatio2 = (QMaxDRatio2 - QMinDRatio2) / 10
'*****
	QMaxDRatio2 = QMaxDRatio1
	QMinDRatio2 = QMinDRatio1
	TermRatio2 = TermRatio1
	
End Function

'��Ʈ�ҷ����� ���� ����Ÿ ���ϱ� 
Function CalculateDataForL()
	Dim FirstFlag
	
	'ChartFX1�� Min/Max/Step ���� 
	If LMaxDRatio1 = 0 Then
		LMaxDRatio1 = 10
	Else
		If LMaxDRatio1 + (LMaxDRatio1 / 10) > 100 Then
			LMaxDRatio1 = 100
		Else
			LMaxDRatio1 = LMaxDRatio1 + (LMaxDRatio1 / 10)
		End If
		
		If LMinDRatio1 - (LMinDRatio1 / 10) < 0 Then
			LMinDRatio1 = 0
		Else
			LMinDRatio1 = LMinDRatio1 - (LMinDRatio1 / 10)
		End If
	End If	
			
	TermRatio1 = (LMaxDRatio1 - LMinDRatio1) / 10
	
	'�б⺰ Lot���հݷ� ����ϱ� 
	If (dblSpdData(0,0) + dblSpdData(0,1) + dblSpdData(0,2)) <> 0 Then
		dblSpdData(9,0) = ((dblSpdData(1,0) + dblSpdData(1,1) + dblSpdData(1,2)) / (dblSpdData(0,0) + dblSpdData(0,1) + dblSpdData(0,2))) * 100
		strSpdData(9,0) = CStr(dblSpdData(9,0))
	Else
		strSpdData(9,0) = ""
	End If
	
	If (dblSpdData(0,3) + dblSpdData(0,4) + dblSpdData(0,5)) <> 0 Then
		dblSpdData(9,1) = ((dblSpdData(1,3) + dblSpdData(1,4) + dblSpdData(1,5)) / (dblSpdData(0,3) + dblSpdData(0,4) + dblSpdData(0,5))) * 100
		strSpdData(9,1) = CStr(dblSpdData(9,1))
	Else
		strSpdData(9,1) = ""
	End If
	
	If (dblSpdData(0,6) + dblSpdData(0,7) + dblSpdData(0,8)) <> 0 Then
		dblSpdData(9,2) = ((dblSpdData(1,6) + dblSpdData(1,7) + dblSpdData(1,8)) / (dblSpdData(0,6) + dblSpdData(0,7) + dblSpdData(0,8))) * 100
		strSpdData(9,2) = CStr(dblSpdData(9,2))
	Else
		strSpdData(9,2) = ""
	End If
	
	If (dblSpdData(0,9) + dblSpdData(0,10) + dblSpdData(0,11)) <> 0 Then
		dblSpdData(9,3) = ((dblSpdData(1,9) + dblSpdData(1,10) + dblSpdData(1,11)) / (dblSpdData(0,9) + dblSpdData(0,10) + dblSpdData(0,11))) * 100
		strSpdData(9,3) = CStr(dblSpdData(9,3))
	Else
		strSpdData(9,3) = ""
	End If
				
	FirstFlag = False
	For i = 0 to 3		
		If strSpdData(9,i) <> "" Then
			If FirstFlag = False Then
				LMaxDRatio2 = dblSpdData(9, i)
				LMinDRatio2 = dblSpdData(9, i)
				FirstFlag = True
			End if
						
			If dblSpdData(9,i) > LMaxDRatio2 Then
				LMaxDRatio2 = dblSpdData(9,i)
			End If
			If dblSpdData(9,i) < LMinDRatio2 Then
				LMinDRatio2 = dblSpdData(9,i)
			End If
		End If
	Next

'***** chart2�� 1�� ���� �����ϰ� ������. 03/07/25 AJJ
	'ChartFX2�� Min/Max/Step ���� 
'	If LMaxDRatio2 = 0 Then
'		LMaxDRatio2 = 10
'	Else
'		If LMaxDRatio2 + (LMaxDRatio2 / 10) > 100 Then
'			LMaxDRatio2 = 100
'		Else
'			LMaxDRatio2 = LMaxDRatio2 + (LMaxDRatio2 / 10)
'		End If
		
'		If LMinDRatio2 - (LMinDRatio2 / 10) < 0 Then
'			LMinDRatio2 = 0
'		Else
'			LMinDRatio2 = LMinDRatio2 - (LMinDRatio2 / 10)
'		End If
'	End If	
		
'	TermRatio2 = (LMaxDRatio2 - LMinDRatio2) / 10
'*****
	LMaxDRatio2 = LMaxDRatio1
	LMinDRatio2 = LMinDRatio1
	TermRatio2 = TermRatio1
End Function


%>
