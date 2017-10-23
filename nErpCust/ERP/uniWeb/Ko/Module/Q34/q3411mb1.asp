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
'*  4. Program Name         : 품질추이(일별)
'*  5. Program Desc         : 
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
	

	Dim TempstrPlantCd
	Dim TempstrItemCd	
	
	TempstrPlantCd		= "<%=Request("txtPlantCd")%>"
	TempstrItemCd		= "<%=Request("txtItemCd")%>"
	
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
Dim RS3

Dim LngRow

Dim strSql	'@@@변경 
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
Dim lgParameter												'불량률 모수 
Dim DayCount													'날짜의 수 28, 29, 30, 31 
Dim Total
Dim TargetFlag
Dim TransTarget

Dim QMaxDRatio
Dim QMinDRatio
Dim LMaxDRatio
Dim LMinDRatio
Dim TermRatio

Dim blnRet

'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - START */
Dim RS4
Dim strSql4
Dim lgTargetParameter
Dim lgMonthlyTargetValue
'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - END *

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
			.Row = <%= i + 1%>					'스프레드에 값 
<%												'뿌려주기 
			For j = 0 to DayCount
%>
				.Col = <%= j + 1%>

				Select Case <%=i%> 
					Case 5, 6, 7				'검사 불량률, 로트불량률, 목표 
						If "<%=strSpdData(i, j)%>" <> "" then
							.Text = "<%=UNINumClientFormat(dblSpdData(i, j), 2, 0)%>"
						End If
					Case Else					'수량 
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
		
		'Y축의 Min/Max 설정 
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(QMinDRatio, 4, 0)%>") 
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(QMaxDRatio, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio, 4, 0)%>")
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
		.Title_(0) = "일별 불량률 추이도"
		.SerLeg(0) = "불량률"
		.SerLeg(1) = "목표치"
		
		.OpenDataEx <%=COD_VALUES%>,2, <%=DayCount - 1%>								'차트 FX와의 데이터 채널 열어주기 
<%
		For i = 0 to DayCount - 1
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = "<%=Cstr(i+1)%>" & "일"
			'불량률 
			if  "<%=strSpdData(5, i)%>" = "" then
				.Series(0).Yvalue(<%=i%>) = <%=CHART_HIDDEN%>
			Else
				.Series(0).Yvalue(<%=i%>) =  parent.UNICDbl("<%=UNINumClientFormat(dblSpdData(5, i), 4, 0)%>")
			End If
<%
		Next
%>
		.Series(0).Visible = True
		
		'목표치가 있다면 
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
		.CloseData <%=COD_VALUES%>											'차트 FX와의 채널 닫아주기 
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
		
		'Y축의 Min/Max 설정 
		.Axis(<%=AXIS_Y%>).Min = parent.UNICDbl("<%=UNINumClientFormat(LMinDRatio, 4, 0)%>") 
		.Axis(<%=AXIS_Y%>).Max = parent.UNICDbl("<%=UNINumClientFormat(LMaxDRatio, 4, 0)%>")
		.Axis(<%=AXIS_Y%>).Step = parent.UNICDbl("<%=UNINumClientFormat(TermRatio, 4, 0)%>")
		
		' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points
			
		.Title_(0) = "일별 LOT불합격률 추이도"
		.SerLeg(0) = "LOT불합격률"
		
		.OpenDataEx <%=COD_VALUES%>, 1, <%=DayCount - 1%>								'차트 FX와의 데이터 채널 열어주기 
<%
		For i = 0 to DayCount - 1
%>
			.Axis(<%=AXIS_X%>).Label(<%=i%>) = "<%=Cstr(i+1)%>" & "일"
			'LOT불합격률 
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
		.CloseData <%=COD_VALUES%>											'차트 FX와의 채널 닫아주기 
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
		'아래는 임의로 준 메시지이다.
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
	'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - START */
	RS1.Close
	RS2.Close
	Conn.Close
	Set RS2 = Nothing
	Set RS1 = Nothing
	Set Conn = Nothing
	'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - END */
End Sub

Function GetData()
	GetData = False
	On Error Resume Next	
	
	' SQL 명령 수행 
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

	'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - START */
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
	
	'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - END */
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
		Conn.Close
		Set RS1 = Nothing
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If RS1.EOF or RS1.BOF Then											'불합격 로트 수는 없을 수 도 있으므로 조건에 넣지 않는다.
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'조건에 맞는 검사결과가 없습니다 
		RS1.Close
		Conn.Close
		Set RS1 = Nothing
		Set Conn = Nothing
		Exit Function
	End If
	'레코드가 있다면 
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
		RS2.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	'레코드가 있다면 
	intRecordCount2 = RS2.RecordCount
	
	'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - START */
	
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
		Set Conn = Nothing												'☜: ComProxy Unload
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
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If

	'불량률단위코드에 대한 모수가 없다면 
	If RS3.EOF or RS3.BOF then
		Call DisplayMsgBox("221205", vbOKOnly, "", "", I_MKSCRIPT)	'불량률 모수가 없습니다.
		RS1.Close
		RS2.Close
		RS3.Close
		Conn.Close
		Set RS1 = Nothing
		Set RS2 = Nothing
		Set RS3 = Nothing
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If

	lgParameter = CDbl(RS3(1))
	
	RS3.Close
	Set RS3 = Nothing
	
	'연월목표 테이블에서 불량률 단위 및 모수 읽기 
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
		Set Conn = Nothing												'☜: ComProxy Unload
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
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
	'연목표 테이블에서 불량률 단위 및 모수가 없다면 
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
	
		' 목표값의 단위가 불량률 단위와 다를때 단위 환산해줌 
		If lgTargetParameter <> lgParameter Then
			TransTarget = lgMonthlyTargetValue / lgTargetParameter * lgParameter
		Else
			TransTarget = lgMonthlyTargetValue
		End If
	End If
	
	'/* [2005-10-24] 연월목표치 없을 경우 Null값 오류 발생관련 수정 - START */
	

	GetData = True
End Function

Function TransferData()
	Dim intDay, intDay2
	Dim FirstFlag
	Dim Target
	
	For LngRow = 0 To intRecordCount1 -1
		
		intDay = CInt(RS1(0))
	
		strSpdData(0, intDay - 1) = RS1(1)	'Lot수 
		
		If strSpdData(0, intDay - 1) = "0" Then
			strSpdData(0, intDay - 1) = ""	'Lot수 
			strSpdData(1, intDay - 1) = ""	'불합격Lot수 
			strSpdData(2, intDay - 1) = ""	'입고수 
			strSpdData(3, intDay - 1) = ""	'검사수 
			strSpdData(4, intDay - 1) = ""	'불량수 
			
			dblSpdData(0, intDay - 1) = 0	'Lot수 
			dblSpdData(1, intDay - 1) = 0	'불합격Lot수 
			dblSpdData(2, intDay - 1) = 0	'입고수 
			dblSpdData(3, intDay - 1) = 0	'검사수 
			dblSpdData(4, intDay - 1) = 0	'불량수 
		
		Else
			'불합격Lot수 
			strSpdData(1, intDay - 1) = "0"	
			dblSpdData(1, intDay - 1) = 0
			
			strSpdData(2, intDay - 1) = RS1(2)	'입고수 
			strSpdData(3, intDay - 1) = RS1(3)	'검사수 
			strSpdData(4, intDay - 1) = RS1(4)	'불량수 
			
			dblSpdData(0, intDay - 1) = CDbl(RS1(1))	'Lot수 
			dblSpdData(2, intDay - 1) = CDbl(RS1(2))	'입고수 
			dblSpdData(3, intDay - 1) = CDbl(RS1(3))	'검사수 
			dblSpdData(4, intDay - 1) = CDbl(RS1(4))	'불량수 
		End If
		
		'불량률 
		If strSpdData(3, intDay -1) = "" Then
			strSpdData(5, intDay -1) = ""
		Else							
			'검사수가 0이 아닌 경우만..
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
		
		'LOT불합격률 
		If strSpdData(0, intDay -1) = "" Then
			strSpdData(6, intDay -1) = ""
		Else							
			'LOT수가 0이 아닌 경우만..
			dblSpdData(6, intDay -1) = 0
			strSpdData(6, intDay -1) = "0"
		End If
					
		RS1.MoveNext
	Next
	'불합격Lot수 
	For LngRow = 0 To intRecordCount2 -1
		
		intDay2 = CInt(RS2(0))
		
		strSpdData(1, intDay2 - 1) = RS2(1)	
		dblSpdData(1, intDay2 - 1) = CDbl(RS2(1))
		
		'LOT불합격률 
		If strSpdData(0, intDay2 -1) = "" Then
			strSpdData(6, intDay2 -1) = ""
		Else							
			'LOT수가 0이 아닌 경우만..
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
	
	For i = 0 to 4							'합계 구하기 
		Total = 0
		For j = 0 to DayCount - 1
			Total = Total + dblSpdData(i,j)
		Next
		dblSpdData(i,DayCount) = Total
		strSpdData(i,DayCount) = CStr(Total)
	Next
	
	If dblSpdData(0,DayCount) = 0 Then
		' 합계에 대한 검사불량률 구하기 
		dblSpdData(5,DayCount) = 0
		strSpdData(5,DayCount) = ""
		
		'합계에 대한 로트불합격률 구하기		
		dblSpdData(6,DayCount) = 0
		strSpdData(6,DayCount) = ""
	Else			
		' 합계에 대한 검사불량률 구하기 
		If dblSpdData(3,DayCount) <> 0 and dblSpdData(4,DayCount) = 0 Then
			dblSpdData(5,DayCount) = 0
			strSpdData(5,DayCount) = "0"
		ElseIf dblSpdData(3,DayCount) <> 0 and dblSpdData(4,DayCount) <> 0 Then
			dblSpdData(5,DayCount) = (dblSpdData(4,DayCount) / dblSpdData(3,DayCount)) * lgParameter
			strSpdData(5,DayCount) = CStr(dblSpdData(5,DayCount))
		End If
			
		'합계에 대한 로트불합격률 구하기		
		If dblSpdData(0,DayCount) <> 0 and dblSpdData(1,DayCount) =  0 Then
			dblSpdData(6,DayCount) = 0
			strSpdData(6,DayCount) = "0"
		ElseIf dblSpdData(0,DayCount) <> 0 and dblSpdData(1,DayCount) <>  0 Then
			dblSpdData(6,DayCount) = (dblSpdData(1,DayCount) / dblSpdData(0,DayCount)) * 100
			strSpdData(6,DayCount) = CStr(dblSpdData(6,DayCount))
		End If
	End If
				
	'목표값이 있다면 
	
	If TargetFlag = True Then
		'월목표 
		strSpdData(7, DayCount) = TransTarget
		Target = TransTarget
		dblSpdData(7, DayCount) = Target				
		'일목표 
		For i = 0 to DayCount - 1
			strSpdData(7,i) = TransTarget
			dblSpdData(7,i) = Target
		Next
		
		If QMaxDRatio < Target Then
			QMaxDRatio = Target					'목표치가 Max보다 클경우 
		End If
		If QMinDRatio > Target Then
			QMinDRatio = Target					'목표치가 Min보다 작을 경우 
		End If
	End If
End Function

'검사불량률에 대한 데이타 구하기 
Function CalculateDataForQ()
	'ChartFX의 Min/Max/Step 산출 
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

'로트불량률에 대한 데이타 구하기 
Function CalculateDataForL()
	'ChartFX의 Min/Max/Step 산출 
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
