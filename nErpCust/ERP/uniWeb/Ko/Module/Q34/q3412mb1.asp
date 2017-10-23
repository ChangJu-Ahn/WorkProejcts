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
'*  4. Program Name         : 최종불량원인 파레토표 
'*  5. Program Desc         : 최종불량원인 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2000/05/29
'*  8. Modified date(Last)  : 2000/05/29
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      :
'* 11. Comment              :
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

	'불량유형명 불러오기 
	Call parent.CommonQueryRs("DEFECT_TYPE_CD,DEFECT_TYPE_NM","Q_DEFECT_TYPE","PLANT_CD =  " & parent.FilterVar(TempstrPlantCd , "''", "S") & " AND INSP_CLASS_CD =  " & parent.FilterVar(TempstrInspClassCd , "''", "S") & " AND DEFECT_TYPE_CD =  " & parent.FilterVar(TempstrDefectTypeCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtDefectTypeNm.Value = strVar2
</Script>
<%													
On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")  

Dim Conn

Dim RS	'@@@변경 
Dim RS1
Dim RS2

Dim LngRow

Dim strSql	'@@@변경 
Dim strSql1
Dim strSql2

Dim aggrDTotal	'@@@변경 
Dim intRecordCount1	'@@@변경 
Dim intRecordCount2	'@@@변경 

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
	'당월 
	'------------------------------------------------------------------------------------------------------------------------
	Parent.frm1.vspdData1.MaxCols = <%=intRecordCount1+1%>
		
	With Parent.ggoSpread
		.Source = Parent.frm1.vspdData1
<%
		'스프레드 헤더 보여주기 
		For i = 0 To intRecordCount1-1
%>									
			.SSSetEdit  <%=i + 2%>, "<%=ConvSPChars(strSpdData1(i))%>" , 8, 1, -1, 15 
<%
		Next
%>	
	End With		
	
	With Parent.frm1.vspdData1
<%
		For i = 0 to 3														'스프레드에 값 보여주기 
%>
			.Row = <%= i + 1%>
<%
			For j = 0 to intRecordCount1 - 1
%>
				.Col = <%=j+2%>
				Select Case <%=i%>
					Case 0, 2		'불량수, 누적불량수 
						.Text = "<%=UNINumClientFormat(dblSpdData1(i,j), ggQty.DecPoint, 0)%>"
					Case Else		'점유율, 누적점유율 
						.Text = "<%=UNINumClientFormat(dblSpdData1(i,j), 2, 0)%>"
				End Select
<%
			Next
		Next
%>
		.Row = 1															'계 보여주기 
		.Col = 1
		.Text = "<%=UNINumClientFormat(Total1, ggQty.DecPoint, 0)%>"		'불량수 계	
		.Row = 2
		.Text = "<%=UNINumClientFormat("100", 2, 0)%>"						'불량률 계 
	End With
End Sub

'=================================================================================
' Diplay datas on spread sheet for Before Month.
'=================================================================================
Sub DisplayOnSpreadForBefore()
	'------------------------------------------------------------------------------------------------------------------------
	'전월 
	'------------------------------------------------------------------------------------------------------------------------
	Parent.frm1.vspdData2.MaxCols = <%=intRecordCount2+1%>
		
	With Parent.ggoSpread
		.Source = Parent.frm1.vspdData2
<%
		'스프레드 헤더 보여주기 
		For i = 0 To intRecordCount2-1
%>									
			.SSSetEdit  <%=i + 2%>, "<%=ConvSPChars(strSpdData2(i))%>" , 8, 1, -1, 15 
<%
		Next
%>	
	End With		
	
	With Parent.frm1.vspdData2
<%
		For i = 0 to 3														'스프레드에 값 보여주기 
%>
			.Row = <%= i + 1%>
<%
			For j = 0 to intRecordCount2 - 1
%>
				.Col = <%=j+2%>
				Select Case <%=i%>
					Case 0, 2		'불량수, 누적불량수 
						.Text = "<%=UNINumClientFormat(dblSpdData2(i,j), ggQty.DecPoint, 0)%>"
					Case Else		'점유율, 누적점유율 
						.Text = "<%=UNINumClientFormat(dblSpdData2(i,j), 2, 0)%>"
				End Select
<%
			Next
		Next
%>
		.Row = 1															'계 보여주기 
		.Col = 1
		.Text = "<%=UNINumClientFormat(Total2, ggQty.DecPoint, 0)%>"		'불량수 계	
		.Row = 2
		.Text = "<%=UNINumClientFormat("100", 2, 0)%>"						'불량률 계 
	End With	
End Sub

'=================================================================================
' Draw the chart for This Month
'=================================================================================
Sub DrawChartForNow()
	With Parent.frm1.ChartFX1
		'---------------------------------------------------------------------
		'차트FX1 - 당월 
		'---------------------------------------------------------------------
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		.Axis(<%=AXIS_Y2%>).Visible = True
		.SerLeg(0) = "불량수"
		.SerLeg(1) = "누적점유율"
													
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(Total1, ggQty.DecPoint, 0)%>")
		.Axis(<%=AXIS_Y2%>).Max = 100
		.Axis(<%=AXIS_Y%>).Min = 0
		.Axis(<%=AXIS_Y2%>).Min = 0
		.Axis(<%=AXIS_Y%>).Decimals = 0											'Y축 소숫점 자리수 지정 
		.Volume = 100												'Bar들의 사이를 0으로 지정 
		.ToolBarObj.Visible = False													'툴바를 디폴트로 보여 주기 
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.OpenDataEx <%=COD_VALUES%>, 2, <%=intRecordCount1%>											'차트FX와의 채널 열어주기 
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
	'전월: 차트FX2 - 오른쪽 차트 
	'------------------------------------------------------------------------------------------------------------------------
	With Parent.frm1.ChartFX2
		.Axis(<%=AXIS_X%>).Visible = True
		.Axis(<%=AXIS_Y%>).Visible = True
		.Axis(<%=AXIS_Y2%>).Visible = True
		.SerLeg(0) = "불량수"
		.SerLeg(1) = "누적점유율"
		
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(Total2, ggQty.DecPoint, 0)%>")
		.Axis(<%=AXIS_Y2%>).Max = 100
		.Axis(<%=AXIS_Y%>).Min = 0
		.Axis(<%=AXIS_Y2%>).Min = 0
		.Axis(<%=AXIS_Y%>).Decimals = 0												'Y축 소숫점 자리수 지정 
		.Volume = 100																'Bar들의 사이를 0으로 지정 
		.ToolBarObj.Visible = False													'툴바를 디폴트로 보여 주기 
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.OpenDataEx <%=COD_VALUES%>, 2, <%=intRecordCount2%>						'차트FX와의 채널 열어주기 
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
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건 값이 비었습니다 
		Exit Function
	End If

	DataReceive = True
End Function

Function DBConnect()
	DBConnect = False

	On Error Resume Next		
	' Database 연결 Object 생성 
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

	'당월 
	strSql1 = "SELECT Q_DEFECT_CAUSE.DEFECT_CAUSE_NM, Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) FROM Q_FINAL_DEFECT_CAUSE_TOTAL Left Outer Join " &_
	              "Q_DEFECT_CAUSE On Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_CAUSE_CD = Q_DEFECT_CAUSE.DEFECT_CAUSE_CD " &_
	              "Where Q_DEFECT_CAUSE.INSP_CLASS_CD=" & FilterVar("F", "''", "S") & "  and Q_FINAL_DEFECT_CAUSE_TOTAL.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.YR = '" & strYr1 & "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Mnth = '" & strMnth1 &_
	              "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Item_Cd = " & FilterVar(strItemCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.INSP_ITEM_CD LIKE " & FilterVar(strInspItemCd, "''", "S") & " and DEFECT_TYPE_CD LIKE " & FilterVar(strDefectTypeCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY > 0   " &_
	              "GROUP BY Q_DEFECT_CAUSE.DEFECT_CAUSE_NM Order By Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) Desc"
	'전월 
	strSql2 = "SELECT Q_DEFECT_CAUSE.DEFECT_CAUSE_NM, Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) FROM Q_FINAL_DEFECT_CAUSE_TOTAL Left Outer Join " &_
	              "Q_DEFECT_CAUSE On Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_CAUSE_CD = Q_DEFECT_CAUSE.DEFECT_CAUSE_CD " &_
	              "Where Q_DEFECT_CAUSE.Insp_Class_Cd=" & FilterVar("F", "''", "S") & "  and Q_FINAL_DEFECT_CAUSE_TOTAL.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.YR = '" & strYr2 & "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Mnth = '" & strMnth2 &_
	              "' and Q_FINAL_DEFECT_CAUSE_TOTAL.Item_Cd = " & FilterVar(strItemCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.INSP_ITEM_CD Like " & FilterVar(strInspItemCd, "''", "S") & " and DEFECT_TYPE_CD LIKE " & FilterVar(strDefectTypeCd, "''", "S") &_
	              " and Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY > 0   " &_
	              "GROUP BY Q_DEFECT_CAUSE.DEFECT_CAUSE_NM Order By Sum(Q_FINAL_DEFECT_CAUSE_TOTAL.DEFECT_QTY) Desc"

'**********************추가 부분시작*************************************
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
	
	
	'불량유형 
	if	TempDefectTypeCd <> "" Then
	strSql = "SELECT DEFECT_TYPE_CD " &_
			"FROM Q_DEFECT_TYPE " &_
			"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND INSP_CLASS_CD = " & FilterVar("P", "''", "S") & " "  			
	'공정검사(P),최종검사(F),출하검사(S)


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
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)	'공장이 존재하지 않습니다.
			RS.Close
			Set RS = Nothing												'☜: ComProxy Unload
			Conn.Close
			Set Conn = Nothing												'☜: ComProxy Unload
			Exit Function
		End If
	RS.Close
	End if	
	
	'검사항목 
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
	
'**********************추가 부분끝*************************************	


	Set RS1 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
		Conn.Close
		Set Conn = Nothing												'☜: ComProxy Unload
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
		Set Conn = Nothing												'☜: ComProxy Unload
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
		Set Conn = Nothing												'☜: ComProxy Unload
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
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If

	If RS1.EOF or RS1.BOF Then
		blnNowFlag = False	
	Else
		blnNowFlag = True
	End If

	If RS2.EOF or RS2.BOF then
		blnBeforeFlag = False								'☜: 비지니스 로직 처리를 종료함 
	Else
		blnBeforeFlag = True
	End If
	
	If blnBeforeFlag = False and blnNowFlag = False Then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'조건에 맞는 검사결과가 없습니다 
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing		
		Exit Function
	End If

	'레코드가 있다면 
	intRecordCount1 = RS1.RecordCount
	intRecordCount2 = RS2.RecordCount
	
	GetData = True
End Function

Function TransferData()
	'------------------------------------------------------------------------------------------------------------------------
	'당월 
	'------------------------------------------------------------------------------------------------------------------------
	If blnNowFlag Then
		ReDim strSpdData1(intRecordCount1)
		ReDim dblSpdData1(3, intRecordCount1)
		
		Total1 = 0

		RS1.MoveFirst	
		For i = 0 to intRecordCount1 - 1 
			strSpdData1(i) = RS1(0)								'불량원인 
			dblSpdData1(0, i) = CDbl(RS1(1))					'이번달 불량수 
			Total1 = Total1 + dblSpdData1(0, i)					'총 불량수 
			RS1.MoveNext				
		Next
		
		aggrDTotal = 0
		For i = 0 to intRecordCount1 - 1
			If Total1 <> 0 Then
				dblSpdData1(1,i) = (dblSpdData1(0,i) / Total1) * 100		'점유율 
				aggrDTotal = aggrDTotal + dblSpdData1(0,i) 					
				dblSpdData1(2,i) = aggrDTotal								'누적불량수 
				dblSpdData1(3,i) = (aggrDTotal / Total1) * 100				'누적 점유율 
			End If
		Next
	End If
	
	'------------------------------------------------------------------------------------------------------------------------
	'전월 
	'------------------------------------------------------------------------------------------------------------------------
	If blnBeforeFlag Then 
		ReDim strSpdData2(intRecordCount2)
		ReDim dblSpdData2(3, intRecordCount2)
		
		Total2 = 0
		
		RS2.MoveFirst	
		For i = 0 to intRecordCount2 - 1 
			strSpdData2(i) = RS2(0)								'거래처명 
			dblSpdData2(0, i) = CDbl(RS2(1))					'지난달 불량수 
			Total2 = Total2 + dblSpdData2(0, i)					'총 불량수 
			RS2.MoveNext
		Next

		aggrDTotal = 0
		For i = 0 to intRecordCount2 - 1
			If Total2 <> 0 Then
				dblSpdData2(1,i) = (dblSpdData2(0,i) / Total2) * 100		'점유율 
				aggrDTotal = aggrDTotal + dblSpdData2(0,i)					'누적 불량수 
				dblSpdData2(2,i) = aggrDTotal
				dblSpdData2(3,i) = (aggrDTotal / Total2) * 100					'누적 점유율 
			End If
		Next
	End If
End Function
%>
