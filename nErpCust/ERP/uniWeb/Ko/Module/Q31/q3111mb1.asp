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
'*  3. Program ID           : Q3111MB1
'*  4. Program Name         : 히스토그램 
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
	
Dim lgintNumOfPeriod		'구간수 
Dim lgintMaxNumOfFrequency
Dim lgintMinNumOfFrequency
Dim lgdblFromPeriod()
Dim lgdblToPeriod()
Dim lgdblCenerValue()
Dim lgNumberOfFrequency()
	
'최대/최소 
Dim lgdblRange
Dim lgdblMax
Dim lgdblMin
	
'최대/최소공차 
Dim lgstrMaxTolerance
Dim lgstrMinTolerance
	
'평균/표준편차/+3시그마/-3시그마 
Dim lgdblAvg
Dim lgdblSigma
Dim lgdblP3Sigma
Dim lgdblM3Sigma
	
'데이타수(시료수)
Dim lglngNumberOfData
	
'검사규격 
Dim lgdblInspSpec
Dim lgstrLSL
Dim lgstrUSL
Dim lgdblLSL
Dim lgdblUSL
Dim lgstrLCL
Dim lgstrUCL
Dim lgMsmtUnitCd
	
'소수점 자리수 
Dim lgintDecimal
	
Dim lgblnRet
Dim i
	
'Request
lgblnRet = Request_QueryData
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If
	
'데이타 얻기 
lgblnRet = Get_Data
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If
	
'구간수 / 구간의 중앙값 / 구간 시작값 / 구간 끝 값 구하기 
lgblnRet = CalForPeriod
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If

'구간별 도수 구하기 
lgblnRet = CalForNumberOfFrequency
If lgblnRet = False Then 
	Call HideStatusWnd
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
	
<%'-------------------- CHART --------------------------%>
<%'ChartFX 속성 설정 %>
'lgblnRet = Setting_ChartFX1
'If lgblnRet = False Then lgOKFlag = False

<%'Histogram 그리기 %>
'lgblnRet = Draw_Histogram
'If lgblnRet = False Then lgOKFlag = False

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
		.txtInspSpec.Value = "<%=UniNumClientFormat(lgdblInspSpec, lgintDecimal, 0)%>"
		.txtLSL.Value = "<%=UniNumClientFormat(lgdblLSL, lgintDecimal ,0)%>"
		.txtUSL.Value = "<%=UniNumClientFormat(lgdblUSL, lgintDecimal ,0)%>"
		.txtSampleQty.Value = "<%=UniNumClientFormat(lglngNumberOfData, ggQty.DecPoint ,0)%>"	<%'시료수 %>
		.txtMaxTol.Value = "<%=UniNumClientFormat(lgstrMaxTolerance, lgintDecimal ,0)%>"		<%'최대공차 %>
		.txtMinTol.Value = "<%=UniNumClientFormat(lgstrMinTolerance, lgintDecimal ,0)%>"		<%'최소공차 %>
		.txtMAX.Value = "<%=UniNumClientFormat(lgdblMax, lgintDecimal, 0)%>"			<%'최대값 %>
		.txtMIN.Value = "<%=UniNumClientFormat(lgdblMin, lgintDecimal, 0)%>"			<%'최대값 %>
		.txtAverage.Value = "<%=UniNumClientFormat(lgdblAvg, lgintDecimal, 0)%>"		<%'평균 %>
		.txtRange.Value = "<%=UniNumClientFormat(lgdblRange, lgintDecimal, 0)%>"			<%'범위 %>
		.txtStd.Value = "<%=UniNumClientFormat(lgdblSigma, lgintDecimal, 0)%>"			<%'표준편차 %>
		.txtP3Sigma.Value = "<%=UniNumClientFormat(lgdblP3Sigma, lgintDecimal, 0)%>"	<%'+3시그마 %>
		.txtM3Sigma.Value = "<%=UniNumClientFormat(lgdblM3Sigma, lgintDecimal, 0)%>"	<%'-3시그마 %>
		.txtMeasmtUnitCd.Value = "<%=lgMsmtUnitCd%>"		<%'측정단위 %>
	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	Display_InspStand = True
	
End Function

<%'/*****************************************************
'/	ChartFX1(Histogram)의 환경 설정 
'/*****************************************************%>
Function Setting_ChartFX1()
	Dim sngTempDiffStep
	Dim intR
	Dim intRest
	Setting_ChartFX1 = False
	
	Err.Clear
	'On Error Resume Next
	
	With Parent.frm1.ChartFX1
		'ToolBar 속성 
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>				'틀바를 새로운 창으로 보이기 
		.ToolBarObj.Left = 15								'틀바의 왼쪽 위치 
		.ToolBarObj.Top = 10								'틀바의 상단 위치 
		
		.Grid = <%=CHART_HORZGRID%>							'X축 수평 Grid
		
		.Axis(<%=AXIS_Y%>).GridColor = RGB(100, 100, 100)
		.Axis(<%=AXIS_Y%>).Decimals = 0		
		
		.Volume = 100	
		
		'Min/Max/Step 구하기 
		intR = <%=lgintMaxNumOfFrequency%>
		intRest = intR \ 10
		intR = intR + 10 - intRest
		sngTempDiffStep = CInt(intR / 10)
		
		.Axis(<%=AXIS_Y%>).Min = 0
		.Axis(<%=AXIS_Y%>).Max = <%=lgintMaxNumOfFrequency%> + sngTempDiffStep 
		.Axis(<%=AXIS_Y%>).STEP = sngTempDiffStep
		
	End With
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Setting_ChartFX1 = True
	
End Function

<%'/*****************************************************
'/	ChartFX1(Histogram) 그리기 
'/*****************************************************%>
Function Draw_Histogram()
	
	Draw_Histogram = False
	
	Err.Clear
	On Error Resume Next
	
'	With Parent.frm1.ChartFX1
'		.OpenDataEx <%=COD_VALUES%>, 1, <%=lgintNumOfPeriod%>				'차트 FX와의 데이터 채널 열어주기 
<%
	Dim YValue0, XValue0, sInsSQL
	Dim blnRet
	'DB 연결 
	blnRet = DBConnect
						
	sInsSQL = "DELETE FROM Q_TMP_CHART_HISTOGRAM"
	Conn.Execute sInsSQL	
%>
<%
			For i = 0 to lgintNumOfPeriod - 1
%>
'				.Legend(<%=i%>) = "<%=UNINumClientFormat(lgdblCenerValue(i), lgintDecimal, 0)%>"
'				'@@@
'				'.ValueEx(0, <%=i%>) = UNICDbl("<%=UNINumClientFormat(lgNumberOfFrequency(i), lgintDecimal, 0)%>")
'				.ValueEx(0, <%=i%>) = parent.UNICDbl("<%=UNINumClientFormat(lgNumberOfFrequency(i), lgintDecimal, 0)%>")
<%
				XValue0 = UNINumClientFormat(lgdblCenerValue(i), lgintDecimal, 0)
				YValue0 = UNINumClientFormat(lgNumberOfFrequency(i), lgintDecimal, 0)
				
				sInsSQL =			" INSERT INTO Q_TMP_CHART_HISTOGRAM (XVALUE, YVALUE ) "
				sInsSQL = sInsSQL &	"	   VALUES ( " & FilterVar(XValue0,"","S") & ", "
				sInsSQL = sInsSQL &		   FilterVar(YValue0,"","S") & ") "				
				
				Conn.Execute sInsSQL
%>				

<%
			Next
%>
'		.CloseData <%=COD_VALUES%>
		
'		.Axis(<%=AXIS_X%>).Visible = True 
'		.Axis(<%=AXIS_Y%>).Visible = True 
'		.Series(0).Visible = True 
'	End With
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	Draw_Histogram = True
	
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
	
	If strPlantCd="" or strInspClassCd = "" or strYrDt1="" or strYrDt2="" or strItemCd="" or strInspItemCd="" then
		Call DisplayMsgBox("229903", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건 값이 비었습니다 
		Exit Function
	End IF
	
	Request_QueryData = True
End Function

'/*****************************************************
'/ 조회 데이타 얻기 
'/*****************************************************
Function Get_Data()
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

    'DB 연결 끊기 
    blnRet = DBClose
    If blnRet = False Then Exit Function
    
    '표준편차/+3시그마/-3시그마 
    blnRet = Get_Sigma
    If blnRet = False Then Exit Function

   '+3시그마/-3시그마 
    blnRet = Get_PM3Sigma
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
	'On Error Resume Next
	
	Conn.Close
	Set Conn = Nothing		
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function
	End If
	
	DBClose = True
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
	'On Error Resume Next
	
	strSql = "SELECT INSP_SPEC, LSL, USL, LCL, UCL, MEASMT_UNIT_CD, INSP_UNIT_INDCTN " &_
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
        RS.Close
        Set RS = Nothing											'☜: ComProxy Unload
		Conn.Close
		Set Conn = Nothing
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
	
	If Trim(RS(6)) <> "3" Then
		Call DisplayMsgBox("229939", vbOKOnly, "", "", I_MKSCRIPT)	'검사단위품질표시가 특성치가 아닙니다.
		Exit Function
	End If
	
	If Trim(RS(0)) = "" Then
		Call DisplayMsgBox("220706", vbOKOnly, "", "", I_MKSCRIPT)	'검사규격이 입력되어 있지 않습니다 
		Exit Function
	End If
	
	lgdblInspSpec = UNICDbl(RS(0), 0)
	
	If Trim(RS(1)) = "" AND Trim(RS(2)) = "" Then
		Call DisplayMsgBox("229911", vbOKOnly, "", "", I_MKSCRIPT)	'상한/하한규격 중 적어도 하나는 존재해야 합니다 
		Exit Function
	End If
	
	lgstrLSL = Trim(RS(1))
	lgstrUSL = Trim(RS(2))
	lgdblLSL = UNICDbl(lgstrLSL, 0)
	lgdblUSL = UNICDbl(lgstrUSL, 0)
	lgstrLCL = Trim(RS(3))
	lgstrUCL = Trim(RS(4))
	lgMsmtUnitCd = Trim(RS(5))
	
	If lgstrLSL = "" Then
		lgstrMinTolerance = ""	
	Else
		lgstrMinTolerance = lgdblInspSpec - lgdblLSL
	End If

	If lgstrUSL = "" Then
		lgstrMaxTolerance = ""
	Else
		lgstrMaxTolerance = lgdblUSL - lgdblInspSpec
	End If

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
	Dim dblSum
	
	Get_MeasuredValues = False
	
	Err.Clear
	'On Error Resume Next
	
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
        RS.Close
		Conn.Close
		Set RS = Nothing
		Set Conn = Nothing
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
	If lglngNumberOfData < 50 Then
		Call DisplayMsgBox("229904", vbOKOnly, "", "", I_MKSCRIPT)	'히스토그램을 그리기 위한 데이타수가 부족합니다 
		RS.Close
		Set RS = Nothing												'☜: ComProxy Unload
		Conn.Close
		Set Conn = Nothing
		Exit Function
	End If
	
	ReDim lgdblData(lglngNumberOfData - 1)
    	dblSum = 0
	For i = 0 To lglngNumberOfData - 1
		If Trim(RS(0)) = "" Then
			Call DisplayMsgBox("229910", vbOKOnly, "", "", I_MKSCRIPT)	'관리도를 그릴 수 없는 자료입니다.
			Exit Function
		Else
	    		lgdblData(i) = UNICDbl(RS(0), 0)
	    		'Sum
	    		dblSum = dblSum + lgdblData(i)
	    		'Min/Max 계산 
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
	
	lgdblRange = lgdblMax - lgdblMin		'범위 
	lgdblAvg = dblSum / lglngNumberOfData		'평균 

	RS.Close
	Set RS = Nothing
	
	Get_MeasuredValues = True
End Function


'/*****************************************************
'/ 표준편차 구하기 
'/*****************************************************
Function Get_Sigma()
	Dim dblSum
	
	Get_Sigma = False
	
	Err.Clear
	'On Error Resume Next
	
    	dblSum = 0
	For i = 0 To lglngNumberOfData - 1
			dblSum = dblSum + (lgdblAvg - lgdblData(i)) ^ 2
	Next
		
	lgdblSigma = Sqr(dblSum / (lglngNumberOfData - 1))
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Get_Sigma = True
End Function

'/*****************************************************
'/ +3Sigma / -3Sigma 구하기 
'/*****************************************************
Function Get_PM3Sigma()
	Get_PM3Sigma= False
	
	Err.Clear
	'On Error Resume Next
	
	lgdblP3Sigma = lgdblAvg + 3 * lgdblSigma
	lgdblM3Sigma = lgdblAvg - 3 * lgdblSigma
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	Get_PM3Sigma = True
End Function
  
'/*****************************************************
'/ 구간수 / 구간의 중앙값 / 구간 시작값 / 구간 끝 값 구하기 
'/*****************************************************
Function CalForPeriod()
	Dim dblHistoAdjust
	Dim dblMinOfPeriod
	Dim dblMaxOfPeriod
	Dim dblPeriod
	Dim i
	CalForPeriod = False
    	
    	Err.Clear
	'On Error Resume Next
	
    'Sturges 방법으로 히스토그램을 그리기 위한 구간의수 계산 
    lgintNumOfPeriod = CInt((3322 / 1000) * (Log(lglngNumberOfData)/ Log(10)) + 1)						
    	
	'히스토 그램의 구간을 구하기 위해 필요한 값 계산.
	'최대값과 최소값의 차이를 구간 수로 나누고 그 값에 0.15를 곱한다.
	'이 값을 최소값에서 빼고, 최대값에 더한 값의 차이가 히스토그램의 구간이다.
	dblHistoAdjust = (lgdblRange / lgintNumOfPeriod) * (15 / 100)
	
	'구간 최대값 / 구간 최소값 / Histogram의 범위 
	dblMinOfPeriod = lgdblMin - dblHistoAdjust
	dblMaxOfPeriod = lgdblMax + dblHistoAdjust
	
	dblPeriod = dblMaxOfPeriod - dblMinOfPeriod
	
	'구간의 중앙값/시작값/끝값 
	Redim lgdblCenerValue(lgintNumOfPeriod - 1)
	Redim lgdblFromPeriod(lgintNumOfPeriod - 1)
	Redim lgdblToPeriod(lgintNumOfPeriod - 1)
	
	For i = 0 To lgintNumOfPeriod - 1
		lgdblCenerValue(i) = dblMinOfPeriod + (dblPeriod / lgintNumOfPeriod) * (i + 1) - (1 / 2) * (dblPeriod / lgintNumOfPeriod)
		lgdblFromPeriod(i) = lgdblCenerValue(i) - (1 / 2) * (dblPeriod / lgintNumOfPeriod)
		lgdblToPeriod(i) = lgdblCenerValue(i) + (1 / 2) * (dblPeriod / lgintNumOfPeriod)
	Next 
	
    	If Err.Number <> 0 Then
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
    
    	CalForPeriod = True
End Function

'/*****************************************************
'/ 구간별 도수 구하기 
'/*****************************************************
Function CalForNumberOfFrequency()
	Dim i
	Dim lngL
	Dim lngCount
	Dim intMax
	Dim intMin
	
	CalForNumberOfFrequency = False
	Err.Clear
	On Error Resume Next
	
	ReDim lgNumberOfFrequency(lgintNumOfPeriod - 1)
	
	For i = 0 To lgintNumOfPeriod - 1
		lngCount = 0
		For lngL = 0 To lglngNumberOfData - 1
			If lgdblData(lngL) >= lgdblFromPeriod(i) And lgdblData(lngL) < lgdblToPeriod(i) Then
				lngCount = lngCount + 1
			End If
		Next 
		lgNumberOfFrequency(i) = lngCount
		
		If i = 0 Then
			lgintMinNumOfFrequency = lgNumberOfFrequency(0)
			lgintMaxNumOfFrequency = lgNumberOfFrequency(0)
		End If
		
		If lgintMinNumOfFrequency > lgNumberOfFrequency(i) Then
			lgintMinNumOfFrequency = lgNumberOfFrequency(i)
		End If
		
		If lgintMaxNumOfFrequency < lgNumberOfFrequency(i) Then
			lgintMaxNumOfFrequency = lgNumberOfFrequency(i)
		End If
	Next 
	
	If Err.Number <> 0 Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Exit Function	
	End If
	
	CalForNumberOfFrequency = True
End Function
%>
