<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4311rb1.asp
'*  4. Program Name			: List Component Requirement (Reservation) (Query)
'*  5. Program Desc			: Used By Goods Issue For Production Order
'*  6. Comproxy List		: 
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2003/05/22
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Chen, JaeHyun
'* 11. Comment				: 
'********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim lgStrPrevKey
Dim lgStrPrevKey2

Dim i
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode         = Request("txtMode")												'☜ : 현재 상태를 받음 
lgStrPrevKey    = Request("lgStrPrevKey")
lgStrPrevKey2   = Request("lgStrPrevKey2")

On Error Resume Next
Err.Clear																	'☜: Protect system from crashing

	Redim UNISqlId(1)
	Redim UNIValue(1, 4)

	UNISqlId(0) = "p4111mb1"
	UNISqlId(1) = "p4311rb1"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	UNIValue(0, 3) = "|"
	UNIValue(0, 4) = "|"
	UNIValue(1, 0) = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")

	If Request("lgStrPrevKey") <> "" Then
		UNIValue(1, 1) = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
		UNIValue(1, 2) = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	Else
		UNIValue(1, 1) = "''"
		UNIValue(1, 2) = "''"
	End If
	If Request("lgStrPrevKey2") <> "" Then
		UNIValue(1, 3) = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
	Else
		UNIValue(1, 3) = "0"
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs0)
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.txtItemCd.value = "<%=ConvSPChars(rs1("ITEM_CD"))%>"
			parent.txtItemNm.value = "<%=ConvSPChars(rs1("ITEM_NM"))%>"
		</Script>	
		<%
		Set rs1 = Nothing
	End If
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("189500", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent
	LngMaxRow = .vspdData1.MaxRows
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
		
		For i = 0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"								'☆: Item Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"								'☆: Item Name
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"									'☆: Item Spec
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Req_Qty"),ggQty.DecPoint,0)%>"		'☆: Required Qty
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"								'☆: Base Unit
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Req_Dt"))%>"							'☆: Required Date
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"							'☆: Tracking No.
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Issued_Qty"),ggQty.DecPoint,0)%>"		'☆: Issued Qty
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Consumed_Qty"),ggQty.DecPoint,0)%>"	'☆: Issued Qty
				strData = strData & Chr(11) & ""																'☆: 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"									'☆: Storage Location
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Nm"))%>"									'☆: Storage Location
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Opr_No"))%>"									'☆: Operation No.
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Cd"))%>"									'☆: Work Center Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Seq"))%>"									'☆: Sequence
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Req_No"))%>"									'☆: Required No.
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Resv_Status"))%>"							'☆: Status
				strData = strData & Chr(11) & ""																'☆: Status Description
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd"))%>"								'☆: Issue Method
				strData = strData & Chr(11) & ""																'☆: Issue Method Description
		        strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs0("Opr_No"))%>"
		.lgStrPrevKey2 = "<%=Trim(rs0("Seq"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	

	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
