<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711mb7.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/12/01
'*  7. Modified date(Last)  : 2001/12/01
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Park, Bum Soo 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
Call HideStatusWnd

On Error Resume Next

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter ���� 
Dim	rs0
Dim strQryMode
Dim strConsumedDtFrom, strConsumedDtTo, strItemCd, strWcCd, strProdtOrderNo, strResourceCd, strResourceGroupCd
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Err.Clear																'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0,10)

	UNISqlId(0) = "p4711mb7"

	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtConsumedDtFrom")) = "" Then
	   strConsumedDtFrom = "|"
	ELSE
	   strConsumedDtFrom = " " & FilterVar(UNIConvDate(Request("txtConsumedDtFrom")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtConsumedDtTo")) = "" Then
	   strConsumedDtTo = "|"
	ELSE
	   strConsumedDtTo = " " & FilterVar(UNIConvDate(Request("txtConsumedDtTo")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtResourceCd")) = "" Then
	   strResourceCd = "|"
	ELSE
	   strResourceCd = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtResourceGroupCd")) = "" Then
	   strResourceGroupCd = "|"
	ELSE
	   strResourceGroupCd = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	UNIValue(0, 4) = FilterVar(UCase(Request("txtOprNo")), "''", "S")
	UNIValue(0, 5) = strResourceCd
	UNIValue(0, 6) = strConsumedDtFrom
	UNIValue(0, 7) = strConsumedDtTo
	UNIValue(0, 8) = strResourceGroupCd
	UNIValue(0, 9) = strItemCd	
	UNIValue(0,10) = strWcCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
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
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData4.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_type"))%>"		
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("consumed_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("consumed_time"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_group_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_group_nm"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("valid_from_dt"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("valid_to_dt"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	iTotalStr = Join(TmpBuffer, "")	
	.ggoSpread.Source = .frm1.vspdData4
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.lgStrPrevKey8 = "<%=Trim(rs0("prodt_order_no"))%>"
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : �ð� �������� ���� 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</script>
