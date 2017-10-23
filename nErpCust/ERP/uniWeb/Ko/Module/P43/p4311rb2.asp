<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4311rb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         : List Onhand Stock Detail
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter 선언 
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strQryMode
Dim strSlCd
Dim strNextSlCd		
Dim strLotNo
Dim strLotSubNo

Dim i

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	If Request("txtSlCd") <> "" Then
		Redim UNISqlId(0)
		Redim UNIValue(0, 0)

		UNISqlId(0) = "180000sad"	
	
		UNIValue(0, 0) = FilterVar(UCase(Request("txtSlCd")), "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

		If (rs0.EOF And rs0.BOF) Then
			rs0.Close
			Set rs0 = Nothing
			%>
			<Script Language=vbscript>
				Parent.txtSLNm.value = ""
			</Script>	
			<%
			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			Response.End													'☜: 비지니스 로직 처리를 종료함 
		Else
			%>
			<Script Language=vbscript>
				Parent.txtSLNm.value = "<%=ConvSPChars(rs0("SL_NM"))%>"
			</Script>	
			<%
			rs0.Close
			Set rs0 = Nothing
		End If
	Else
		%>
		<Script Language=vbscript>
			Parent.txtSLNm.value = ""
		</Script>	
		<%
	End IF
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "p4311rb2"

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtChildItemCd")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	
	strSlCd = FilterVar(UCase(Request("txtMajorSlCd")), "''", "S")
	
	If Request("lgStrPrevKey3") <> "" Then
		strNextSlCd = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
		strLotNo = FilterVar(UCase(Request("lgStrPrevKey4")), "''", "S")
		strLotSubNo = FilterVar(UCase(Trim(Request("lgStrPrevKey5"))),"" & FilterVar("0", "''", "S") & " ","S")
	Else
		strNextCd = "''"
		strLotNo = "''"
		strLotSubNo = "''"
	End If

	If strSlCd <> "''" Then	
		If strLotNo <> "''" Then
			UNIValue(0, 4) = " a.sl_cd = " & strSlCd & " and (a.lot_no >= " & strLotNo & " or (a.lot_no = " & strLotNo & " and a.lot_sub_no >= " & strLotSubNo & " ))"
		Else
			UNIValue(0, 4) = " a.sl_cd = " & strSlCd
		End If
	Else
		If strLotNo <> "''" Then
			UNIValue(0, 4) = " (a.sl_cd > " & strNextSlCd & " or (a.sl_cd >= " & strNextSlCd & " and a.lot_no > " & strLotNo & " ) or (a.sl_cd >= " & strNextSlCd & " and a.lot_no >= " & strLotNo & " and a.lot_sub_no >= " & strLotSubNo & " )) "
		Else
			UNIValue(0, 4) = "|"
		End If
	End If
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .vspdData2.MaxRows									'Save previous Maxrow
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If
		
		For i=0 to rs0.RecordCount-1 
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BLOCK_INDICATOR"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_ON_HAND_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PREV_GOOD_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_INSP_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_TRNS_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "0"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey3 = "<%=ConvSPChars(rs0("SL_CD"))%>"
		.lgStrPrevKey4 = "<%=ConvSPChars(rs0("LOT_NO"))%>"
		.lgStrPrevKey5 = "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbDtlQueryOk(LngMaxRow)									'☆: 조회 성공후 실행로직 


End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
