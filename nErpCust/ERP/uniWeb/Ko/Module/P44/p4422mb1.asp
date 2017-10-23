<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4422mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/10/23
'*  7. Modified date(Last)  : 2003/05/26
'*  8. Modifier (First)     : Park, Bumsoo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")


On Error Resume Next								'☜: 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter 선언 
Dim	rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim strPlantCd, strItemCd, strWcCd, strProdtOrderNo
Dim strOrderStatus, strTrackingNo, strItemGroupCd

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strFlag

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey1	
Dim i

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

	lgStrPrevKey1 = UCase(Trim(Request("lgStrPrevKey1")))	
	
	Redim UNISqlId(4)
	Redim UNIValue(4, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000sam"    
	UNISqlId(4) = "180000sas"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtItemGroupNm.value = ""
		parent.frm1.txtWCNm.value = ""
	</Script>	
	<%

	' Plant 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		strFlag = "ERROR_PLANT"	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If
	
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs4.EOF AND rs4.BOF Then
			strFlag = "ERROR_GROUP"
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs4("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	End If
	
	' Tracking No존재유무 Check
    IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			strFlag = "ERROR_TRACK"
		End If
	End IF
	
	' 작업장명 Display
	IF Request("txtWcCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			strFlag = "ERROR_WCCD"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtWCNm.value = "<%=ConvSPChars(rs2("WC_NM"))%>"
			</Script>	
			<%
		End If
	End IF		
	
	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs1.EOF And rs1.BOF) Then
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs1("ITEM_NM"))%>"
			</Script>	
			<%
		End If
	End IF
	
	rs0.Close	:	Set rs0 = Nothing
	rs1.Close	:	Set rs1 = Nothing
	rs2.Close	:	Set rs2 = Nothing
	rs3.Close	:	Set rs3 = Nothing
	rs4.Close	:	Set rs4 = Nothing
	
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantNm.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtItemNm.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_WCCD" Then		
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtWCNm.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End	
		End If
	End IF

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)
	
	UNISqlId(0) = "p4422mb1"
	
	IF Trim(Request("txtPlantCd")) = "" Then
	   strPlantCd = "|"
	ELSE
	   strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF

	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("cboOrderStatus")) = "" Then
	   strOrderStatus = "|"
	ELSE
	   strOrderStatus = " " & FilterVar(UCase(Request("cboOrderStatus")), "''", "S") & ""
	END IF
		
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	END IF	
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "n.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = Trim(strPlantCd)
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strProdtOrderNo
	UNIValue(0, 5) = strOrderStatus
	UNIValue(0, 6) = strTrackingNo
	UNIValue(0, 7) = strItemGroupCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
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
   	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount-1%>)
		
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("wipinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevprodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevgoodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("goodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("badqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("wipinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("base_unit"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevprodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevgoodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("goodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("badqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
		
	Next
%>
	
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData1
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.lgStrPrevKey1 = "<%=Trim(rs0("ITEM_CD"))%>"		

	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
	.frm1.hOrderStatus.value	= "<%=ConvSPChars(Request("cboOrderStatus"))%>"
	.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
