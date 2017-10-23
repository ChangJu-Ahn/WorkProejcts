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
'*  4. Program Name         : ������׷� 
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
	
Dim lgintNumOfPeriod		'������ 
Dim lgintMaxNumOfFrequency
Dim lgintMinNumOfFrequency
Dim lgdblFromPeriod()
Dim lgdblToPeriod()
Dim lgdblCenerValue()
Dim lgNumberOfFrequency()
	
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
Dim lgstrLCL
Dim lgstrUCL
Dim lgMsmtUnitCd
	
'�Ҽ��� �ڸ��� 
Dim lgintDecimal
	
Dim lgblnRet
Dim i
	
'Request
lgblnRet = Request_QueryData
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If
	
'����Ÿ ��� 
lgblnRet = Get_Data
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If
	
'������ / ������ �߾Ӱ� / ���� ���۰� / ���� �� �� ���ϱ� 
lgblnRet = CalForPeriod
If lgblnRet = False Then 
	Call HideStatusWnd
	Response.End
End If

'������ ���� ���ϱ� 
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
<%'���� DATA DISPLAY %>
lgblnRet = Display_InspStand
If lgblnRet = False Then lgOKFlag = False
	
<%'-------------------- CHART --------------------------%>
<%'ChartFX �Ӽ� ���� %>
'lgblnRet = Setting_ChartFX1
'If lgblnRet = False Then lgOKFlag = False

<%'Histogram �׸��� %>
'lgblnRet = Draw_Histogram
'If lgblnRet = False Then lgOKFlag = False

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
		.txtInspSpec.Value = "<%=UniNumClientFormat(lgdblInspSpec, lgintDecimal, 0)%>"
		.txtLSL.Value = "<%=UniNumClientFormat(lgdblLSL, lgintDecimal ,0)%>"
		.txtUSL.Value = "<%=UniNumClientFormat(lgdblUSL, lgintDecimal ,0)%>"
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
'/	ChartFX1(Histogram)�� ȯ�� ���� 
'/*****************************************************%>
Function Setting_ChartFX1()
	Dim sngTempDiffStep
	Dim intR
	Dim intRest
	Setting_ChartFX1 = False
	
	Err.Clear
	'On Error Resume Next
	
	With Parent.frm1.ChartFX1
		'ToolBar �Ӽ� 
		.ToolBarObj.Docked = <%=TGFP_FLOAT%>				'Ʋ�ٸ� ���ο� â���� ���̱� 
		.ToolBarObj.Left = 15								'Ʋ���� ���� ��ġ 
		.ToolBarObj.Top = 10								'Ʋ���� ��� ��ġ 
		
		.Grid = <%=CHART_HORZGRID%>							'X�� ���� Grid
		
		.Axis(<%=AXIS_Y%>).GridColor = RGB(100, 100, 100)
		.Axis(<%=AXIS_Y%>).Decimals = 0		
		
		.Volume = 100	
		
		'Min/Max/Step ���ϱ� 
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
'/	ChartFX1(Histogram) �׸��� 
'/*****************************************************%>
Function Draw_Histogram()
	
	Draw_Histogram = False
	
	Err.Clear
	On Error Resume Next
	
'	With Parent.frm1.ChartFX1
'		.OpenDataEx <%=COD_VALUES%>, 1, <%=lgintNumOfPeriod%>				'��Ʈ FX���� ������ ä�� �����ֱ� 
<%
	Dim YValue0, XValue0, sInsSQL
	Dim blnRet
	'DB ���� 
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

    'DB ���� ���� 
    blnRet = DBClose
    If blnRet = False Then Exit Function
    
    'ǥ������/+3�ñ׸�/-3�ñ׸� 
    blnRet = Get_Sigma
    If blnRet = False Then Exit Function

   '+3�ñ׸�/-3�ñ׸� 
    blnRet = Get_PM3Sigma
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
        Set RS = Nothing											'��: ComProxy Unload
		Conn.Close
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
	
	If Trim(RS(6)) <> "3" Then
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
'/ ����ġ ����Ÿ ��� 
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
	
	'���ڵ尡 �ִٸ� 
	lglngNumberOfData = RS.RecordCount
	If lglngNumberOfData < 50 Then
		Call DisplayMsgBox("229904", vbOKOnly, "", "", I_MKSCRIPT)	'������׷��� �׸��� ���� ����Ÿ���� �����մϴ� 
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
			Call DisplayMsgBox("229910", vbOKOnly, "", "", I_MKSCRIPT)	'�������� �׸� �� ���� �ڷ��Դϴ�.
			Exit Function
		Else
	    		lgdblData(i) = UNICDbl(RS(0), 0)
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
'/ +3Sigma / -3Sigma ���ϱ� 
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
'/ ������ / ������ �߾Ӱ� / ���� ���۰� / ���� �� �� ���ϱ� 
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
	
    'Sturges ������� ������׷��� �׸��� ���� �����Ǽ� ��� 
    lgintNumOfPeriod = CInt((3322 / 1000) * (Log(lglngNumberOfData)/ Log(10)) + 1)						
    	
	'������ �׷��� ������ ���ϱ� ���� �ʿ��� �� ���.
	'�ִ밪�� �ּҰ��� ���̸� ���� ���� ������ �� ���� 0.15�� ���Ѵ�.
	'�� ���� �ּҰ����� ����, �ִ밪�� ���� ���� ���̰� ������׷��� �����̴�.
	dblHistoAdjust = (lgdblRange / lgintNumOfPeriod) * (15 / 100)
	
	'���� �ִ밪 / ���� �ּҰ� / Histogram�� ���� 
	dblMinOfPeriod = lgdblMin - dblHistoAdjust
	dblMaxOfPeriod = lgdblMax + dblHistoAdjust
	
	dblPeriod = dblMaxOfPeriod - dblMinOfPeriod
	
	'������ �߾Ӱ�/���۰�/���� 
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
'/ ������ ���� ���ϱ� 
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
