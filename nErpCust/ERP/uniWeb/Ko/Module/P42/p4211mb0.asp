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
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter 선언 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

On Error Resume Next

    Err.Clear															'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
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
		Response.End													'☜: 비지니스 로직 처리를 종료함 
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
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
