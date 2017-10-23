<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2212mb2.asp
'*  4. Program Name         : MPS�����ٰ���ȸ 
'*  5. Program Desc         : MPS�����ٰ���ȸ 
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================-->

<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim strPlantCd, strItemCd, strTrackingNo
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
Dim i

Call HideStatusWnd

	Redim UNISqlId(0)
	Redim UNIValue(0, 5)
	
	UNISqlId(0) = "184200sab"
	
	strPlantCd		= FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	strItemCd		= FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	strTrackingNo	= FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")

	UNIValue(0, 0) = strPlantCd

	UNIValue(0, 1) = strItemCd
	UNIValue(0, 2) = strTrackingNo
	UNIValue(0, 3) = strPlantCd

	UNIValue(0, 4) = strItemCd
	UNIValue(0, 5) = strTrackingNo
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
					
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim arrVal
ReDim arrVal(0)

Dim f_resrv_qty
Dim f_req_qty
Dim f_sch_rcpt_qty
Dim f_on_hand_qty
Dim f_plan_qty
Dim f_rslt_qty
Dim f_mps_qty
    	
With parent
	LngMaxRow = .frm1.vspdData2.MaxRows
		
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("req_dt"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("sp_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("mps_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		
		f_resrv_qty = "<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_req_qty =  "<%=UniConvNumberDBToCompany(rs0("req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_sch_rcpt_qty = "<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_on_hand_qty = "<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_plan_qty = "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_inv_qty = "<%=UniConvNumberDBToCompany(rs0("inv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_mps_qty = "<%=UniConvNumberDBToCompany(rs0("mps_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"

		f_rslt_qty = parent.parent.UniCdbl(f_rslt_qty) + parent.parent.UNICdbl(f_plan_qty) + parent.parent.UNICdbl(f_on_hand_qty) _
					+ parent.parent.UNICdbl(f_mps_qty) + parent.parent.UNICdbl(f_sch_rcpt_qty) _
					- parent.parent.UNICdbl(f_resrv_qty) - parent.parent.UNICdbl(f_req_qty)

		f_rslt_qty = parent.UNIFormatNumber(f_rslt_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)
		
		<%		IF i=0 THEN %>
		   strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		<%		ELSE %>			
					strData = strData & Chr(11) & f_rslt_qty
		<%		END IF %>	

		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		ReDim Preserve arrVal(<%=i%>)
		arrVal(<%=i%>) = strData

<%		rs0.MoveNext
	Next
%>  
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData Join(arrVal,"")
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbDtlQueryOk
End With	
</Script>	

<%
Set ADF = Nothing
%>