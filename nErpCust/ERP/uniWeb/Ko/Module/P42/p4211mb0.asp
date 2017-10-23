<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4211mb0.asp
'*  4. Program Name			: Lookup Item By Plant
'*  5. Program Desc			: Adjust Requirement (Query)
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2000/09/28
'*  8. Modified date(Last)	: 2002/06/29
'*  9. Modifier (First)		: Park, Bum Soo
'* 10. Modifier (Last)		: Park, Bum Soo
'* 11. Comment				:
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

On Error Resume Next

    Err.Clear															'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p4211mb0"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtItemCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		Call Parent.LookUpItemByPlantFail("<%=ConvSPChars(Request("txtItemCd"))%>", "<%=ConvSPChars(Request("txtRow"))%>")
		</Script>
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
    
%>
<Script Language=vbscript>

	With parent.frm1.vspdData2

		.Row = "<%=Request("txtRow")%>"

		If "<%=rs0("plant_valid_Flg")%>" = "N" or "<%=rs0("item_valid_Flg")%>" = "N" Then 'VALID_FLG
			
			Call parent.DisplayMsgBox("122619", "x", "x", "x") 
			
			Call Parent.LookUpItemByPlantFail("<%=Request("txtItemCd")%>", "<%=Request("txtRow")%>")
		Else
			If "<%=rs0("Phantom_Flg")%>" = "Y" Then 'PHANTOM_FLG
				
				Call parent.DisplayMsgBox("189214", "x", "x", "x")
				
			    Call Parent.LookUpItemByPlantFail("<%=ConvSPChars(Request("txtItemCd"))%>", "<%=ConvSPChars(Request("txtRow"))%>")
			Else
				.Col = parent.C_CompntNm
				.text = "<%=ConvSPChars(rs0("Item_Nm"))%>"
				.Col = parent.C_Spec
				.text = "<%=ConvSPChars(rs0("Spec"))%>"
				.Col = parent.C_Unit
				.text = "<%=ConvSPChars(rs0("Basic_Unit"))%>"

				If "<%=rs0("Tracking_Flg")%>" = "N" Then 'TRACKING_FLG
					.Col = parent.C_TrackingNo
					.Text = "*"
				Else
					.Col = parent.C_TrackingNo		
					.Value = parent.frm1.txtTrackingNo.Value
				End If

				.Col = parent.C_MajorSLCd
				.text = "<%=ConvSPChars(rs0("Issued_Sl_Cd"))%>"
				.Col = parent.C_MajorSLNm
				.text = "<%=ConvSPChars(rs0("Sl_Nm"))%>"
				.Col = parent.C_IssueMeth
				.text = "<%=ConvSPChars(rs0("Issue_Mthd"))%>"
				.Col = parent.C_IssueMethDesc
				.text = "<%=ConvSPChars(rs0("Issue_Desc"))%>"
			End If
		End If
	End With
	
    Parent.LookUpItemByPlantSuccess("<%=Request("txtRow")%>")
	
</Script>
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
