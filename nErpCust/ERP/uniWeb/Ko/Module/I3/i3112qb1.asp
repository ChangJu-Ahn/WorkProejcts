<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- ChartFX용 상수를 사용하기 위한 Include 지정 -->
<!-- #include file="../../inc/CfxIE.inc" -->
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Changing Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3112QB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : PI3G140
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/05/25
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : KiHong Han
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "I", "NOCOOKIE","QB")
								
On Error Resume Next
'Call HideStatusWnd

'export fields
Const E1_yr = 0
Const E1_mnth = 1
Const E1_pernicious_stock_qty = 2
Const E1_pernicious_stock_amt = 3
Const E1_longterm_stock_qty = 4
Const E1_longterm_stock_amt = 5

Const CHRT_HIDDN = 0

Dim PI3G140		

Dim iLngCnt
Dim iLngRow
Dim iLngMaxRows

Dim iStrData
Dim TmpBuffer
Dim iTotalStr

Dim iStrPlantCd
Dim iStrFrYr
Dim iStrFrMnth
Dim iStrToYr
Dim iStrToMnth
Dim iStrQueryTargetClass
Dim iStrQueryTargetCd

Dim iVarPlantNm
Dim iVarQueryTargetNm
Dim iVarLongTermStockCalPeriod
Dim iVarPerniciousStockCalPeriod
Dim iArrExport

Dim Conn

iStrPlantCd = Request("txtPlantCd")
iStrFrYr = Request("txtFrYr")
iStrFrMnth = Request("txtFrMnth")
iStrToYr = Request("txtToYr")
iStrToMnth = Request("txtToMnth")
iStrQueryTargetClass = Request("txtQueryTargetClass")
iStrQueryTargetCd = Request("txtQueryTargetCd")

Set PI3G140 = Server.CreateObject("PI3G140.cILiLongtermInvChg")
If CheckSystemError(Err,True) Then					
	Call HideStatusWnd
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Call PI3G140.I_LI_LONGTERM_INV_CHANGING_SVR(gstrGlobalCollection, _
											iStrPlantCd, _
											iStrFrYr, _
											iStrFrMnth, _
											iStrToYr, _
											iStrToMnth, _
											iStrQueryTargetClass, _
											iStrQueryTargetCd, _
											iVarPlantNm, _
											iVarQueryTargetNm, _
											iVarLongTermStockCalPeriod, _
											iVarPerniciousStockCalPeriod, _
											iArrExport)			

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Set PI3G140 = Nothing
	Call HideStatusWnd
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PI3G140 = Nothing

iLngCnt = UBound(iArrExport, 1)

ReDim TmpBuffer1(iLngCnt)
ReDim TmpBuffer2(iLngCnt)
ReDim TmpBuffer3(iLngCnt)
ReDim TmpBuffer4(iLngCnt)
ReDim TmpBuffer5(iLngCnt)

'칼럼 헤더	
For iLngRow = 0 To iLngCnt
	TmpBuffer1(iLngRow) = iArrExport(iLngRow, E1_yr) & "-" & iArrExport(iLngRow, E1_mnth) & Chr(11)
Next

'악성재고수량 
For iLngRow = 0 To iLngCnt
	If Trim(iArrExport(iLngRow, E1_pernicious_stock_qty)) = "" Or Isnull(iArrExport(iLngRow, E1_pernicious_stock_qty)) Then
		TmpBuffer2(iLngRow) = "" & Chr(11)
	Else
		TmpBuffer2(iLngRow) = UniConvNumberDBToCompany(iArrExport(iLngRow, E1_pernicious_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & Chr(11)
	End If
	
Next

'악성재고금액 
For iLngRow = 0 To iLngCnt
	If Trim(iArrExport(iLngRow, E1_pernicious_stock_amt)) = "" Or Isnull(iArrExport(iLngRow, E1_pernicious_stock_amt)) Then
		TmpBuffer3(iLngRow) = "" & Chr(11)
	Else
		TmpBuffer3(iLngRow) = UniConvNumberDBToCompany(iArrExport(iLngRow, E1_pernicious_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & Chr(11)
	End If
Next

'장기보관재고수량	
For iLngRow = 0 To iLngCnt
	If Trim(iArrExport(iLngRow, E1_longterm_stock_qty)) = "" Or Isnull(iArrExport(iLngRow, E1_longterm_stock_qty)) Then
		TmpBuffer4(iLngRow) = "" & Chr(11)
	Else
		TmpBuffer4(iLngRow) = UniConvNumberDBToCompany(iArrExport(iLngRow, E1_longterm_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & Chr(11)
	End If
Next

'장기보관재고금액 
For iLngRow = 0 To iLngCnt
	If Trim(iArrExport(iLngRow, E1_longterm_stock_amt)) = "" Or Isnull(iArrExport(iLngRow, E1_longterm_stock_amt)) Then
		TmpBuffer5(iLngRow) = "" & Chr(11)
	Else
		TmpBuffer5(iLngRow) = UniConvNumberDBToCompany(iArrExport(iLngRow, E1_longterm_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & Chr(11)
	End If
Next
	
iTotalStr = Join(TmpBuffer1, "") & Chr(12) _
		& Join(TmpBuffer2, "") & Chr(12) _
		& Join(TmpBuffer3, "") & Chr(12) _
		& Join(TmpBuffer4, "") & Chr(12) _
		& Join(TmpBuffer5, "") & Chr(12)
		
	
Response.Write "<Script language=vbs> " & vbCr         

Response.Write " Dim strData  " & vbCr
Response.Write " Dim i  " & vbCr	
Response.Write " strData = Replace( """ & iTotalStr & """, Chr(11), Chr(9)) " & vbCr
Response.Write " strData = Replace(strData, Chr(12), Chr(13)) " & vbCr

Response.Write " With Parent.frm1.vspdData " & vbCr
Response.Write "	.Redraw = False " & vbCr

Response.Write "	.MaxCols = " & iLngCnt + 1 & vbCr
Response.Write "	.Col = 1 " & vbCr
Response.Write "	.Col2 = " & iLngCnt + 1 & vbCr
Response.Write "	.Row = 0 " & vbCr
Response.Write "	.Row2 = .MaxRows " & vbCr
Response.Write "	.Clip = strData " & vbCr

Response.Write "	.Redraw = True " & vbCr

Response.Write " End With " & vbCr
	
''Chart Display
'Response.Write " With Parent.frm1.ChartFX1 " & vbCr
'Response.Write "	'.Gallery = 1 " & vbCr
'Response.Write "	'.Chart3D = False " & vbCr			'2D
'Response.Write "	.Title_(2) =  " & """" & "장기재고추이도(수량)" & """" & vbCr
'Response.Write "	.Axis(" & AXIS_Y & ").Decimals = " & ggQty.DecPoint & vbCr
'Response.Write "	.Axis(" & AXIS_Y & ").Format = " & AF_NUMBER & vbCr

'Response.Write " 	.Axis(" & AXIS_X & ").Visible = True " & vbCr
'Response.Write " 	.Axis(" & AXIS_Y & ").Visible = True " & vbCr

'' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points " 
'Response.Write " 	.SerLeg(0) = " & """" & "악성재고수량" & """" & vbCr
'Response.Write " 	.SerLeg(1) = " & """" & "장기보관재고수량" & """" & vbCr

'Response.Write " 	.OpenDataEx " & COD_VALUES & ", 2, " & iLngCnt & vbCr					'차트 FX와의 데이터 채널 열어주기 

'For iLngRow = 0 to iLngCnt
'	'X축 라벨 
'	Response.Write "		.Axis(" & AXIS_X & ").Label(" & iLngRow & ") = """ & iArrExport(iLngRow, E1_yr) & "-" & iArrExport(iLngRow, E1_mnth) & """" & vbCr
'	'악성재고수량 
'	If Trim(iArrExport(iLngRow, E1_pernicious_stock_qty)) = "" Or Isnull(iArrExport(iLngRow, E1_pernicious_stock_qty)) Then
'		Response.Write "			.Series(0).Yvalue(" & iLngRow & ") = " & CHRT_HIDDN & vbCr
'	Else
'		Response.Write "			.Series(0).Yvalue(" & iLngRow & ") =  parent.UNICDbl(""" & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_pernicious_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """) " & vbCr
'	End If
'Next
'Response.Write "		.Series(0).Visible = True " & vbCr

'For iLngRow = 0 to iLngCnt
'	If Trim(iArrExport(iLngRow, E1_longterm_stock_qty)) = "" Or Isnull(iArrExport(iLngRow, E1_longterm_stock_qty)) Then
'		Response.Write "		.Series(1).Yvalue(" & iLngRow & ") = " & CHRT_HIDDN & vbCr
'	Else
'		Response.Write "		.Series(1).Yvalue(" & iLngRow & ") = parent.UNICDbl(""" & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_longterm_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """) " & vbCr
'	End If
'Next

'Response.Write "		.Series(1).Visible = True " & vbCr

'' Close the VALUES channel
'Response.Write "	.CloseData " & COD_VALUES & vbCr	

'Response.Write " 	.RecalcScale " & vbCr
'Response.Write " End With " & vbCr
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

Dim YValue0, YValue1, XValue0, sInsSQL
Dim blnRet

blnRet = DBConnect

sInsSQL = "DELETE FROM I_TEMP_CHART_LONGTERM"
Conn.Execute sInsSQL

For iLngRow = 0 to iLngCnt
	
	'X축 라벨 
	XValue0 = iArrExport(iLngRow, E1_yr) & "-" & iArrExport(iLngRow, E1_mnth)
	'악성재고수량 
	If Trim(iArrExport(iLngRow, E1_pernicious_stock_qty)) = "" Or Isnull(iArrExport(iLngRow, E1_pernicious_stock_qty)) Then
		YValue0 = CHRT_HIDDN
	Else
		YValue0 = CDBL(UniConvNumberDBToCompany(iArrExport(iLngRow, E1_pernicious_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0))
	End If

	If Trim(iArrExport(iLngRow, E1_longterm_stock_qty)) = "" Or Isnull(iArrExport(iLngRow, E1_longterm_stock_qty)) Then
		YValue1 = CHRT_HIDDN
	Else
		YValue1 = CDBL(UniConvNumberDBToCompany(iArrExport(iLngRow, E1_longterm_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0))
	End If
	
	
	sInsSQL =			" INSERT INTO I_TEMP_CHART_LONGTERM (ROWNUM, XVALUE, YVALUE1, YVALUE2 ) "
	sInsSQL = sInsSQL & " VALUES (" & iLngRow & ", " & FilterVar(XValue0, "''", "S") & ", "
	sInsSQL = sInsSQL & FilterVar(YValue0, "", "SNM") & ", " & FilterVar(YValue1, "", "SNM") & ") "
	
	Conn.Execute sInsSQL
Next


Response.Write " With Parent " & vbCr
Response.Write "	.frm1.txtPlantNm.Value = """ & ConvSPChars(iVarPlantNm) & """" & vbCr
Response.Write "	.frm1.txtQueryTargetNm.Value = """ & ConvSPChars(iVarQueryTargetNm) & """" & vbCr
Response.Write "	.frm1.txtLongtermStockCalPeriod.Value = """ & iVarLongTermStockCalPeriod & """" & vbCr
Response.Write "	.frm1.txtPerniciousStockCalPeriod.Value = """ & iVarPerniciousStockCalPeriod & """" & vbCr
Response.Write "	.DbQueryOK " & vbCr
Response.Write "	.frm1.vspdData.focus " & vbCr
Response.Write " End With " & vbCr	   

Response.Write "</Script> " & vbCr 
	
'Call ServerMesgBox("CHART LAST: " & Err.number & Err.Description, vbInformation, I_MKSCRIPT)
	
Response.End

%>
