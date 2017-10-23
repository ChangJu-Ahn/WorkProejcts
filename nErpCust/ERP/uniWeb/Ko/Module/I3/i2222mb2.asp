<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4218mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2003/05/22
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
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
Call LoadInfTB19029B("Q", "P","NOCOOKIE","MB")

On Error Resume Next								'��: 

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3			'DBAgent Parameter ���� 
Dim strQryMode


Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgStrPrevKey2	
Dim i

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strPlantCd, strItemCd, strTrackingNo, strSLCd, strReqStartDt, strReqEndDt

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	strSLCd = FilterVar(UCase(Request("txtSLCd")), "''", "S")
	strReqStartDt = FilterVar(UniConvDate(Request("txtReqStartDt")), "''", "S") 
	strReqEndDt =  FilterVar(UniConvDate(Request("txtReqEndDt")), "''", "S")
	
	Redim UNISqlId(3)
	Redim UNIValue(3, 11)
	
	UNISqlId(0) = "I2222MB2_A"
	
	UNIValue(0, 0) = strPlantCd
	UNIValue(0, 1) = strReqStartDt
	UNIValue(0, 2) = strReqEndDt
	UNIValue(0, 3) = strItemCd
	UNIValue(0, 4) = strTrackingNo
	UNIValue(0, 5) = strSLCd
	UNIValue(0, 6) = strPlantCd
	UNIValue(0, 7) = strReqStartDt
	UNIValue(0, 8) = strReqEndDt
	UNIValue(0, 9) = strItemCd
	UNIValue(0, 10) = strTrackingNo
	UNIValue(0, 11) = strSLCd
	
	UNISqlId(1) = "I2222MB2_B"
	
	UNIValue(1, 0) = strPlantCd
	UNIValue(1, 1) = strReqStartDt
	UNIValue(1, 2) = strReqEndDt
	UNIValue(1, 3) = strItemCd
	UNIValue(1, 4) = strTrackingNo
	UNIValue(1, 5) = strSLCd
	UNIValue(1, 6) = strPlantCd
	UNIValue(1, 7) = strReqStartDt
	UNIValue(1, 8) = strReqEndDt
	UNIValue(1, 9) = strItemCd
	UNIValue(1, 10) = strTrackingNo
	UNIValue(1, 11) = strSLCd
	
	UNISqlId(2) = "I2222MB2_C"
	
	UNIValue(2, 0) = strPlantCd
	UNIValue(2, 1) = strReqStartDt
	UNIValue(2, 2) = strReqEndDt
	UNIValue(2, 3) = strItemCd
	UNIValue(2, 4) = strTrackingNo
	UNIValue(2, 5) = strSLCd
	
	UNISqlId(3) = "I2222MB2_D"
	
	UNIValue(3, 0) = strPlantCd
	UNIValue(3, 1) = strItemCd
	UNIValue(3, 2) = strSLCd
	UNIValue(3, 3) = strSLCd
	UNIValue(3, 4) = strTrackingNo

	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strDataA, strDataB, strDataC, strDataD
Dim A_TmpBuffer, B_TmpBuffer, C_TmpBuffer, D_TmpBuffer
Dim iTotalStrA, iTotalStrB, iTotalStrC, iTotalStrD
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	
<%  If Not (rs0.EOF And rs0.BOF ) Then
%>

		ReDim A_TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
		For i=0 to rs0.RecordCount-1 
%>
			strDataA = ""
			strDataA = strDataA & Chr(11) & "<%=ConvSPChars(rs0("pur_mfg"))%>"
			strDataA = strDataA & Chr(11) & "<%=ConvSPChars(rs0("order_no"))%>"
			strDataA = strDataA & Chr(11) & "<%=ConvSPChars(rs0("order_status"))%>"
			strDataA = strDataA & Chr(11) & "<%=UNIDateClientFormat(rs0("end_plan_dt"))%>"
			strDataA = strDataA & Chr(11) & "<%=UniNumClientFormat(rs0("order_qty"),ggQty.DecPoint,0)%>"
			strDataA = strDataA & Chr(11) & "<%=ConvSPChars(rs0("base_unit"))%>"
			strDataA = strDataA & Chr(11) & "<%=UniNumClientFormat(rs0("result_qty"),ggQty.DecPoint,0)%>"
			strDataA = strDataA & Chr(11) & "<%=UniNumClientFormat(rs0("pre_reciept_qty"),ggQty.DecPoint,0)%>"
			strDataA = strDataA & Chr(11) & "<%=UniNumClientFormat(rs0("reciept_qty"),ggQty.DecPoint,0)%>"
			strDataA = strDataA & Chr(11) & "<%=ConvSPChars(rs0("manager"))%>"
			strdataA = strDataA & Chr(11) & "<%=ConvSPChars(rs0("pur_mfg_flag"))%>"
			strDataA = strDataA & Chr(11) & LngMaxRow + <%=i%>
			strDataA = strDataA & Chr(11) & Chr(12)
			
			A_TmpBuffer(<%=i%>) = strDataA
<%		
			rs0.MoveNext
		Next
%>
		iTotalStrA = Join(A_TmpBuffer, "")	
<%
	End If	
			
	rs0.Close
	Set rs0 = Nothing
%>


	LngMaxRow = .frm1.vspdData3.MaxRows	
	
<%  If Not (rs1.EOF And rs1.BOF ) Then
%>
	
		ReDim B_TmpBuffer(<%=rs1.RecordCount - 1%>)	
<%  
		For i=0 to rs1.RecordCount-1 
%>
			strDataB = ""
			
			If "<%=ConvSPChars(rs1("module"))%>" = "P" Then
				strDataB = strDataB & Chr(11) & "����"
			Else
				strDataB = strDataB & Chr(11) & "����"
			End If		
			strDataB = strDataB & Chr(11) & "<%=ConvSPChars(rs1("order_no"))%>"
			strDataB = strDataB & Chr(11) & "<%=ConvSPChars(rs1("opr_no"))%>"
			strDataB = strDataB & Chr(11) & "<%=ConvSPChars(rs1("seq"))%>"
			strDataB = strDataB & Chr(11) & "<%=UNIDateClientFormat(rs1("req_dt"))%>"
			strDataB = strDataB & Chr(11) & "<%=ConvSPChars(rs1("to_locate"))%>"
			strDataB = strDataB & Chr(11) & "<%=ConvSPChars(rs1("to_loc_nm"))%>"
			strDataB = strDataB & Chr(11) & "<%=UniNumClientFormat(rs1("req_qty"),ggQty.DecPoint,0)%>"
			strDataB = strDataB & Chr(11) & "<%=ConvSPChars(rs1("base_unit"))%>"
			strDataB = strDataB & Chr(11) & "<%=UniNumClientFormat(rs1("issued_qty"),ggQty.DecPoint,0)%>"
			strDataB = strDataB & Chr(11) & "<%=UniNumClientFormat(rs1("consumed_qty"),ggQty.DecPoint,0)%>"
			strDataB = strDataB & Chr(11) & "<%=UniNumClientFormat(rs1("remain_qty"),ggQty.DecPoint,0)%>"
			strDataB = strDataB & Chr(11) & LngMaxRow + <%=i%>
			strDataB = strDataB & Chr(11) & Chr(12)
			
			B_TmpBuffer(<%=i%>) = strDataB
<%		
			rs1.MoveNext
		Next
%>
		iTotalStrB = Join(B_TmpBuffer, "")	
<%
	End If	
			
	rs1.Close
	Set rs1 = Nothing
%>

	LngMaxRow = .frm1.vspdData4.MaxRows	

<%  If Not (rs2.EOF And rs2.BOF ) Then
%>

		ReDim C_TmpBuffer(<%=rs2.RecordCount - 1%>)	
<%  
		For i=0 to rs2.RecordCount-1 
%>
			strDataC = ""
			strDataC = strDataC & Chr(11) & "<%=ConvSPChars(rs2("insp_class_cd"))%>"
			strDataC = strDataC & Chr(11) & "<%=ConvSPChars(rs2("insp_req_no"))%>"
			strDataC = strDataC & Chr(11) & "<%=ConvSPChars(rs2("insp_status"))%>"
			strDataC = strDataC & Chr(11) & "<%=UNIDateClientFormat(rs2("insp_req_dt"))%>"
			strDataC = strDataC & Chr(11) & "<%=UniNumClientFormat(rs2("lot_size"),ggQty.DecPoint,0)%>"
			strDataC = strDataC & Chr(11) & "<%=ConvSPChars(rs2("basic_unit"))%>"
			strDataC = strDataC & Chr(11) & "<%=UniNumClientFormat(rs2("goods_Qty"),ggQty.DecPoint,0)%>"
			strDataC = strDataC & Chr(11) & "<%=UniNumClientFormat(rs2("defectives_qty"),ggQty.DecPoint,0)%>"
			strDataC = strDataC & Chr(11) & LngMaxRow + <%=i%>
			strDataC = strDataC & Chr(11) & Chr(12)
			
			C_TmpBuffer(<%=i%>) = strDataC
<%		
			rs2.MoveNext
		Next
%>
		iTotalStrC = Join(C_TmpBuffer, "")	
<%
	End If	
			
	rs2.Close
	Set rs2 = Nothing
%>
	
	LngMaxRow = .frm1.vspdData5.MaxRows	

<%  If Not (rs3.EOF And rs3.BOF ) Then
%>

		ReDim D_TmpBuffer(<%=rs3.RecordCount - 1%>)	
<%  
		For i=0 to rs3.RecordCount-1 
%>
			strDataD = ""
			
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("sl_cd"))%>"
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("sl_nm"))%>"
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("tracking_no"))%>"
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("lot_no"))%>"
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("lot_sub_no"))%>"
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("block_indicator"))%>"
			strDataD = strDataD & Chr(11) & "<%=ConvSPChars(rs3("basic_unit"))%>"
			strDataD = strDataD & Chr(11) & "<%=UniNumClientFormat(rs3("good_on_hand_qty"),ggQty.DecPoint,0)%>"
			strDataD = strDataD & Chr(11) & "<%=UniNumClientFormat(rs3("stk_on_insp_qty"),ggQty.DecPoint,0)%>"
			strDataD = strDataD & Chr(11) & "<%=UniNumClientFormat(rs3("stk_on_trns_qty"),ggQty.DecPoint,0)%>"
			strDataD = strDataD & Chr(11) & "<%=UniNumClientFormat(rs3("prev_good_qty"),ggQty.DecPoint,0)%>"
			strDataD = strDataD & Chr(11) & "<%=UniNumClientFormat(rs3("prev_stk_on_insp_qty"),ggQty.DecPoint,0)%>"
			strDataD = strDataD & Chr(11) & "<%=UniNumClientFormat(rs3("prev_stk_in_trns_qty"),ggQty.DecPoint,0)%>"
			strDataD = strDataD & Chr(11) & LngMaxRow + <%=i%>
			strDataD = strDataD & Chr(11) & Chr(12)
			
			D_TmpBuffer(<%=i%>) = strDataD
<%		
			rs3.MoveNext
		Next
%>
		iTotalStrD = Join(D_TmpBuffer, "")	
<%
	End If	
			
	rs3.Close
	Set rs3 = Nothing
%>

	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowDataByClip iTotalStrA
	
	.ggoSpread.Source = .frm1.vspdData3
	.ggoSpread.SSShowDataByClip iTotalStrB
	
	.ggoSpread.Source = .frm1.vspdData4
	.ggoSpread.SSShowDataByClip iTotalStrC
	
	.ggoSpread.Source = .frm1.vspdData5
	.ggoSpread.SSShowDataByClip iTotalStrD
	
	.Dbquery2Ok()

End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
