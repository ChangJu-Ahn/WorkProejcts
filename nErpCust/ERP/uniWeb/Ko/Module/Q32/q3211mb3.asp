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
'*  3. Program ID           : q3211mb3.asp
'*  4. Program Name         : Worst 10
'*  5. Program Desc         : Worst 10
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2000/04/24
'*  8. Modified date(Last)  : 2000/04/24
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
	Dim TempstrItemCd
	
	TempstrPlantCd		= "<%=Request("txtPlantCd")%>"
	TempstrItemCd		= "<%=Request("txtItemCd")%>"
	
	'/* [2005-10-27] 형식이 일치하지 않습니다 오류 관련 수정: FilterVar --> parent.FilterVar - START */
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
	'/* [2005-10-27] 형식이 일치하지 않습니다 오류 관련 수정: FilterVar --> parent.FilterVar - END */	
</Script>
<%													
On Error Resume Next

'Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
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

Dim Total	'@@@변경 
Dim blnRet	'@@@변경 

Dim intRecordCount1
Dim intRecordCount2	'@@@변경 
Dim intRecordCount3

Dim TempItemCd
Dim lgDataFlag          
Dim strPlantCd
Dim strItemCd
Dim strYr1
Dim strMnth1
Dim strYr2
Dim strMnth2


Dim strSpdData(7, 9)
Dim dblSpdData(6, 9)
Dim i, j
Dim lgParameter	
Dim QMaxDRatio
Dim QMinDRatio
Dim LMaxDRatio
Dim LMinDRatio
Dim TermRatio

' Receive datas from client
blnRet = DataReceive()
If blnRet = False Then
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If

' Connect the Database
blnRet = DBConnect()
If blnRet = False Then
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Response.End
End If

' Get datas from the database

blnRet = GetData()



If blnRet = False Then
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

    Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write " Call parent.DBQueryErr " & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
End If

' Calculate datas for display
Call TransferData()

If lgDataFlag = "Q"  then
	Call CalculateDataForQ
	Call DrawChartForQ
Else
	Call CalculateDataForL
	Call DrawChartForL
End If

Call DBClose		

%>

<Script Language=vbscript>

' Diplay datas on spread sheet.
Call DisplayOnSpread()

If <%=(lgDataFlag = "Q")%>  then
	Call DrawChartForQ
Else
	Call DrawChartForL
End If


Call Parent.DBQueryOK

'=================================================================================
' Diplay datas on spread sheet.
'=================================================================================
Sub DisplayOnSpread()

	With Parent.frm1.vspdData
		'거래처 
		.Row = 1					
<%											
		For j = 0 to 9
%>
			.Col = <%=j + 1%>
			.Text = "<%=strSpdData(7, j)%>"
<%
		Next

		For i = 0 to 6
%>
			.Row = <%=i + 2%>		
<%											
			For j = 0 to 9
%>
				.Col = <%=j + 1%>

				Select Case <%=i%> 
					Case 5, 6				'검사 불량률, 로트불량률 
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


<%

If lgDataFlag = "Q"  then

	Dim YValue
	dim  strSQLQ
    blnRet = DBConnect
	strSQLQ = "DELETE FROM Q_TMP_CHART_WORST"
	Conn.Execute strSQLQ
%>

<%	
		For i = 0 to 9
				'LOT불합격률 
				if  strSpdData(5, i)  = "" then
					 YValue = 0
				Else
					  YValue =  dblSpdData(5, i)
				End If
				
				
				strSQLQ = "INSERT INTO Q_TMP_CHART_WORST  (XVALUE, YVALUE) "
				strSQLQ = strSQLQ & " VALUES (" & FilterVar(i+1 , "", "S") & ", " & YValue  & ") "
			
		
				Conn.Execute strSQLQ			

		Next
End iF
%>
	
	
	
End Sub

'=================================================================================
' Draw the chart for lot rejection ratio.
'=================================================================================
Sub DrawChartForL()
<%
If lgDataFlag <> "Q"  then

	Dim YValue0, YValue1, sInsSQL

    blnRet = DBConnect
					
		sInsSQL = "DELETE FROM Q_TMP_CHART_WORST"
		Conn.Execute sInsSQL
%>

<%	
		For i = 0 to 9
			
			'LOT불합격률 
			if  strSpdData(6, i)  = "" then
				YValue0 = 0
			Else
				  YValue0 =  dblSpdData(6, i)
			End If
				sInsSQL = "INSERT INTO Q_TMP_CHART_WORST  (XVALUE, YVALUE) "
				sInsSQL = sInsSQL & " VALUES (" & FilterVar(i+1 , "", "S") & ", " & YValue0 & ") "

				Conn.Execute sInsSQL			

		Next
End If
%>

End Sub




</Script>   
<%

blnRet = DBClose


Function DataReceive()
	DataReceive = False

	lgDataFlag = Request("txtDataFlag")

	strPlantCd  = Request("txtPlantCd")
	strItemCd = Request("txtItemCd")
	TempItemCd = Request("txtItemCd")
	If strItemCd = "" Then
		strItemCd = "%"
	End if
	strYr1=Request("txtYr1")
	strMnth1 = Request("txtMnth1")
	strYr2=Request("txtYr2")
	strMnth2 = Request("txtMnth2")

	If strPlantCd = "" or strYr1 = "" or strMnth1 = "" or  strYr2 = "" or strMnth2 = "" Then
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
	Conn.Close
	Set RS2 = Nothing
	Set RS1 = Nothing
	Set RS3 = Nothing
	Set Conn = Nothing
	
End Sub

Function GetData()
	GetData = False
	On Error Resume Next	
	
	strSql1 = "SELECT Sum(RECEIVING_LOT_CNT), Sum(INSP_QTY) FROM Q_RECEIVING_INSPECTION_TOTAL WHERE YR + MNTH BETWEEN '" & strYr1 & "' + '" & strMnth1 & "' and '" & strYr2 & "' + '" & strMnth2 & "' and PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " and Item_cd LIKE " & FilterVar(strItemCd, "''", "S") & " GROUP BY BP_CD ORDER BY BP_CD"

	If lgDataFlag = "L" Then
		strSql2 = "SELECT TOP 10 B_BIZ_PARTNER.BP_NM, Sum(Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT), Sum(Q_RECEIVING_INSPECTION_TOTAL.REJT_LOT_CNT), Sum(Q_RECEIVING_INSPECTION_TOTAL.LOT_SIZE_QTY), " &_
		              "Sum(Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY), Sum(Q_RECEIVING_INSPECTION_TOTAL.DEFECT_QTY) " &_
		              "FROM Q_RECEIVING_INSPECTION_TOTAL Left Outer Join B_BIZ_PARTNER " &_
		              "On Q_RECEIVING_INSPECTION_TOTAL.BP_CD = B_BIZ_PARTNER.BP_CD WHERE Q_RECEIVING_INSPECTION_TOTAL.YR + Q_RECEIVING_INSPECTION_TOTAL.MNTH " &_
		              "BETWEEN '" & strYr1 & "' + '" & strMnth1 & "' and '" & strYr2 & "' + '" & strMnth2 & "' and Q_RECEIVING_INSPECTION_TOTAL.Plant_cd= " & FilterVar(strPlantCd, "''", "S") & " and " &_
		              "Q_RECEIVING_INSPECTION_TOTAL.Item_cd LIKE " & FilterVar(strItemCd, "''", "S") & "  AND Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT > 0 AND Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY > 0 " &_
		              "GROUP BY B_BIZ_PARTNER.BP_NM ORDER BY Sum(Q_RECEIVING_INSPECTION_TOTAL.REJT_LOT_CNT)/Sum(Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT) DESC" &_
		              ", Sum(Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT) ASC" &_
		              ", Sum(Q_RECEIVING_INSPECTION_TOTAL.DEFECT_QTY)/Sum(Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY) DESC" &_
		              ", Sum(Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY) ASC" &_
		              ", Sum(Q_RECEIVING_INSPECTION_TOTAL.REJT_LOT_CNT) ASC"
					  
	Else
		strSql2 = "SELECT TOP 10 B_BIZ_PARTNER.BP_NM, Sum(Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT), Sum(Q_RECEIVING_INSPECTION_TOTAL.REJT_LOT_CNT), Sum(Q_RECEIVING_INSPECTION_TOTAL.LOT_SIZE_QTY), " &_
		              "Sum(Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY), Sum(Q_RECEIVING_INSPECTION_TOTAL.DEFECT_QTY) " &_
		              "FROM Q_RECEIVING_INSPECTION_TOTAL Left Outer Join B_BIZ_PARTNER " &_
		              "On Q_RECEIVING_INSPECTION_TOTAL.BP_CD = B_BIZ_PARTNER.BP_CD WHERE Q_RECEIVING_INSPECTION_TOTAL.YR + Q_RECEIVING_INSPECTION_TOTAL.MNTH " &_
		              "BETWEEN '" & strYr1 & "' + '" & strMnth1 & "' and '" & strYr2 & "' + '" & strMnth2 & "' and Q_RECEIVING_INSPECTION_TOTAL.Plant_cd= " & FilterVar(strPlantCd, "''", "S") & " and " &_
		              "Q_RECEIVING_INSPECTION_TOTAL.Item_cd LIKE " & FilterVar(strItemCd, "''", "S") & " AND Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT > 0 AND Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY > 0 " &_
		              "GROUP BY B_BIZ_PARTNER.BP_NM ORDER BY Sum(Q_RECEIVING_INSPECTION_TOTAL.DEFECT_QTY)/Sum(Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY) DESC" &_
					  ", Sum(Q_RECEIVING_INSPECTION_TOTAL.INSP_QTY) ASC" &_
					  ", Sum(Q_RECEIVING_INSPECTION_TOTAL.REJT_LOT_CNT)/Sum(Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT) DESC" &_
					  ", Sum(Q_RECEIVING_INSPECTION_TOTAL.RECEIVING_LOT_CNT) ASC" &_
					  ", Sum(Q_RECEIVING_INSPECTION_TOTAL.LOT_SIZE_QTY) ASC"
	End If

	strSql3 = "SELECT PARAMETER From Q_DEFECT_RATIO_UNIT WHERE DEFECT_RATIO_UNIT_CD = (SELECT DEFECT_RATIO_UNIT_CD From Q_DEFECT_RATIO_BY_INSP_CLASS WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " and INSP_CLASS_CD = " & FilterVar("R", "''", "S") & " )"

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
	If TempItemCd <> "" Then
		strSql = "SELECT ITEM_CD " &_
				"FROM B_ITEM_BY_PLANT " &_
				"WHERE PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(TempItemCd, "''", "S")

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
		Exit Function														'☜: 비지니스 로직 처리를 종료함 
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
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If

	'레코드가 하나도 없다면 
	If RS1.EOF or RS1.BOF Then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'조건에 맞는 검사결과가 없습니다 
		RS1.Close
		Conn.Close
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function													'☜: 비지니스 로직 처리를 종료함 
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
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function														'☜: 비지니스 로직 처리를 종료함 
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
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set RS2 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function														'☜: 비지니스 로직 처리를 종료함 
	End If

	'레코드가 하나도 없다면 
	If RS2.EOF or RS2.BOF Then
		Call DisplayMsgBox("229902", vbOKOnly, "", "", I_MKSCRIPT)	'조건에 맞는 검사결과가 없습니다.
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set RS2 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function													'☜: 비지니스 로직 처리를 종료함 
	Else	'레코드가 있다면 
		intRecordCount2 = RS2.RecordCount
	End If

	Set RS3 = Server.CreateObject("ADODB.Recordset")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
		RS1.Close
		RS2.Close
		Conn.Close
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set RS2 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function													'☜: 비지니스 로직 처리를 종료함 
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
		Conn.Close
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set RS2 = Nothing												'☜: ComProxy Unload
		Set RS3 = Nothing												'☜: ComProxy Unload
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
		Set RS1 = Nothing												'☜: ComProxy Unload
		Set RS2 = Nothing												'☜: ComProxy Unload
		Set RS3 = Nothing												'☜: ComProxy Unload
		Set Conn = Nothing												'☜: ComProxy Unload
		Exit Function
	End If

	lgParameter = CSng(RS3(0))	

	GetData = True
End Function

Function TransferData()
	Dim FirstFlag
	
	FirstFlag = False				
	For LngRow = 0 To intRecordCount2 -1
			
		strSpdData(7, LngRow) = RS2(0)						'거래처명 
		
		If RS2(1) = "0" Then
			strSpdData(0, LngRow) = ""	'Lot수 
			strSpdData(1, LngRow) = ""	'불합격Lot수 
			strSpdData(2, LngRow) = ""	'입고수 
			strSpdData(3, LngRow) = ""	'검사수 
			strSpdData(4, LngRow) = ""	'불량수 
			
			dblSpdData(0, LngRow) = 0	'Lot수 
			dblSpdData(1, LngRow) = 0	'불합격Lot수 
			dblSpdData(2, LngRow) = 0	'입고수 
			dblSpdData(3, LngRow) = 0	'검사수 
			dblSpdData(4, LngRow) = 0	'불량수 

		Else
			strSpdData(0, LngRow) = RS2(1)	'Lot수 
			strSpdData(1, LngRow) = RS2(2)	'불합격Lot수 
			strSpdData(2, LngRow) = RS2(3)	'입고수 
			strSpdData(3, LngRow) = RS2(4)	'검사수 
			strSpdData(4, LngRow) = RS2(5)	'불량수 
			
			dblSpdData(0, LngRow) = CDbl(RS2(1))	'Lot수 
			dblSpdData(1, LngRow) = CDbl(RS2(2))	'불합격Lot수 
			dblSpdData(2, LngRow) = CDbl(RS2(3))	'입고수 
			dblSpdData(3, LngRow) = CDbl(RS2(4))	'검사수 
			dblSpdData(4, LngRow) = CDbl(RS2(5))	'불량수 
		End If
			
		'불량률 
		If strSpdData(3, LngRow) = "" Then
			strSpdData(5, LngRow) = ""
		Else							
			'검사수가 0이 아닌 경우만..
			dblSpdData(5, LngRow) = (dblSpdData(4, LngRow) / dblSpdData(3, LngRow)) * lgParameter
			
	
			strSpdData(5, LngRow) = CStr(dblSpdData(5, LngRow))
				
			If FirstFlag = False Then
				QMaxDRatio = dblSpdData(5, LngRow)
				QMinDRatio = dblSpdData(5, LngRow)
				FirstFlag = True
			End If
							
			If dblSpdData(5, LngRow) > QMaxDRatio Then
				QMaxDRatio = dblSpdData(5, LngRow)
			End If
					
			If dblSpdData(5, LngRow) < QMinDRatio Then
				QMinDRatio = dblSpdData(5, LngRow)
			End If
		End If


		'LOT불합격률 
		If strSpdData(0, LngRow) = "" Then
			strSpdData(6, LngRow) = ""
		Else							
			'LOT수가 0이 아닌 경우만..
			dblSpdData(6, LngRow) = (dblSpdData(1, LngRow) / dblSpdData(0, LngRow)) * 100
			strSpdData(6, LngRow) = CStr(dblSpdData(6, LngRow))
			
			If FirstFlag = False Then
				LMaxDRatio = dblSpdData(6, LngRow)
				LMinDRatio = dblSpdData(6, LngRow)
				FirstFlag = True
			End If
												
			If dblSpdData(6, LngRow) > LMaxDRatio Then
				LMaxDRatio = dblSpdData(6, LngRow)
			End If
			If dblSpdData(6, LngRow) < LMinDRatio Then
				LMinDRatio = dblSpdData(6, LngRow)
			End If
		End If
			
		RS2.MoveNext
	Next	

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
