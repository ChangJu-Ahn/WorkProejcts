<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2341rb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         : List MRP Info (Pegging)
'*  6. Modified date(First) : 2003-11-04
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Chen, Jae Hyun
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

On Error Resume Next

Dim ADF	
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0


Dim lgStrPrevKey	' 이전 값 
Dim i
Dim strTempDate		' 날짜 비교용 
Dim lgStrColorFlag

Call HideStatusWnd

On Error Resume Next

Dim strItemCd

	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
	UNISqlId(0) = "p2341rb2"
	
	strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	
	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = strItemCd
	UNIValue(0, 2) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 3) = strItemCd
		
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

Dim f_mps_qty
Dim sum_f_mps_qty
Dim f_resrv_qty
Dim sum_f_resrv_qty
Dim f_req_qty
Dim sum_f_req_qty
Dim f_sch_rcpt_qty
Dim sum_f_sch_rcpt_qty
Dim f_ttl_req_qty
Dim sum_f_ttl_req_qty
Dim f_on_hand_qty
Dim f_plan_qty
Dim sum_f_plan_qty
Dim f_rslt_qty

sum_f_mps_qty = 0
sum_f_resrv_qty = 0
sum_f_req_qty = 0
sum_f_sch_rcpt_qty = 0
sum_f_ttl_req_qty = 0
sum_f_plan_qty =0
    	
With parent
	LngMaxRow = .vspdData.MaxRows
	lgStrColorFlag = ""
	
<%	For i=0 to rs0.RecordCount-1 %>		
		
		strData = ""
<%		If	Trim(strTempDate) = Trim(UNIDateClientFormat(rs0("req_dt"))) Then %>
			strData = strData & Chr(11) & ""
<%		Else	%>
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("req_dt"))%>"
<%		End If	%>
		
<%		If Trim(rs0("prnt_item_cd")) <> "zzzzzzzzzzzzzzzzzz***" Then %>
			<%If Trim(rs0("prnt_item_cd")) <> "*" Then%>
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prnt_item_cd"))%>"
			<%Else%>
				strData = strData & Chr(11) & ""
			<%End If%>	
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("mps_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ttl_req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			<%IF i=0 THEN %>
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			<%ELSE %>
				strData = strData & Chr(11) & f_rslt_qty
			<%END IF %>
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
<%		Else %>
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""	'mps_qty
			strData = strData & Chr(11) & ""	'resrv_qty
			strData = strData & Chr(11) & ""	'req_qty
			strData = strData & Chr(11) & ""	'ttl_req_qty
			strData = strData & Chr(11) & ""	'sch_rcpt_qty
			strData = strData & Chr(11) & ""	'on_hand_qty
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			' to change spread color
			lgStrColorFlag = lgStrColorFlag & CStr(<%=i+1%>) & .PopupParent.gColSep & "1" & .PopupParent.gRowSep
			
<%		End If %>			

		f_resrv_qty = "<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_req_qty = "<%=UniConvNumberDBToCompany(rs0("req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_sch_rcpt_qty =  "<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_on_hand_qty =  "<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_plan_qty =  "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_inv_qty =  "<%=UniConvNumberDBToCompany(rs0("inv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		
		f_mps_qty = "<%=UniConvNumberDBToCompany(rs0("mps_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		f_ttl_req_qty = "<%=UniConvNumberDBToCompany(rs0("ttl_req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		
		f_rslt_qty = parent.parent.UniCdbl(f_rslt_qty) + parent.parent.UniCdbl(f_plan_qty) + parent.parent.UniCdbl(f_on_hand_qty) + parent.parent.UniCdbl(f_sch_rcpt_qty) - parent.parent.UniCdbl(f_resrv_qty) - parent.parent.UniCdbl(f_req_qty) - parent.parent.UniCdbl(f_mps_qty)
		f_rslt_qty = parent.parent.UNIFormatNumber(f_rslt_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)
		
		sum_f_mps_qty = sum_f_mps_qty + parent.parent.UniCdbl(f_mps_qty)
        sum_f_resrv_qty = sum_f_resrv_qty + parent.parent.UniCdbl(f_resrv_qty)
        sum_f_req_qty = sum_f_req_qty + parent.parent.UniCdbl(f_req_qty)
        sum_f_sch_rcpt_qty = sum_f_sch_rcpt_qty + parent.parent.UniCdbl(f_sch_rcpt_qty)
        sum_f_ttl_req_qty = sum_f_ttl_req_qty + parent.parent.UniCdbl(f_ttl_req_qty)
        sum_f_plan_qty = sum_f_plan_qty + parent.parent.UniCdbl(f_plan_qty)
		
		ReDim Preserve arrVal(<%=i%>)
		arrVal(<%=i%>) = strData
		
<%		strTempDate = UNIDateClientFormat(rs0("req_dt")) %>
		
<%		
		rs0.MoveNext
	Next
%>

        strData = ""
        strData = strData & Chr(11) & ""
        strData = strData & Chr(11) & "합계"
        strData = strData & Chr(11) & ""
        strData = strData & Chr(11) & ""
        strData = strData & Chr(11) & parent.parent.UNIFormatNumber(sum_f_mps_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)			'sum_f_mps_qty
        strData = strData & Chr(11) & parent.parent.UNIFormatNumber(sum_f_resrv_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)			'sum_f_resrv_qty
        strData = strData & Chr(11) & parent.parent.UNIFormatNumber(sum_f_req_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)			'sum_f_req_qty
        strData = strData & Chr(11) & parent.parent.UNIFormatNumber(sum_f_ttl_req_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)		'sum_f_ttl_req_qty
        strData = strData & Chr(11) & parent.parent.UNIFormatNumber(sum_f_sch_rcpt_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)		'sum_f_sch_rcpt_qty
        strData = strData & Chr(11) & f_rslt_qty
        strData = strData & Chr(11) & parent.parent.UNIFormatNumber(sum_f_plan_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)			'sum_f_plan_qty     
        
        strData = strData & Chr(11) & LngMaxRow + <%=i%>
        strData = strData & Chr(11) & Chr(12)
        
        ' to change spread color
		lgStrColorFlag = lgStrColorFlag & CStr(<%=i+1%>) & .PopupParent.gColSep & "2" & .PopupParent.gRowSep
        
		ReDim Preserve arrVal(<%=i%>)
		arrVal(<%=i%>) = strData        
        
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowData Join(arrVal,"")
		.lgStrColorFlag = lgStrColorFlag
		
	.DbDtlQueryOk
End With	
</Script>	
<%
rs0.Close
Set rs0 = Nothing
Set ADF = Nothing
%>
