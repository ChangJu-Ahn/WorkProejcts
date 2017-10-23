<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4111mb1.asp 
'*  4. Program Name         : Called By P4111MA1 (Order Management - Single)
'*  5. Program Desc         : Lookup Production Order Header
'*  6. Modified date(First) : 2002/05/07
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : Park, BumSoo
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1	'DBAgent Parameter 선언 
Dim strProdtOrderNo, strProdtOrderNo_Next, strProdtOrderNo_Previous

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd


    Err.Clear															'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000sab"
	UNISqlId(1) = "p4111mb0"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtItemCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

    If rs1("Procur_Type") = "P" Then
    
		Call DisplayMsgBox("189209", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		%>
		<Script Language=vbscript>
		Parent.LookUpItemByPlantFail
		</Script>
		<%
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End																'☜: Process End
    End If

%>
<Script Language=vbscript>
	With parent.frm1

		If "<%=rs1("Item_Valid_Flg")%>" = "N" or "<%=rs1("Plant_Valid_Flg")%>" = "N" Then 'VALID_FLG
			.txtItemCd.Value = ""
			
			Call parent.DisplayMsgBox("122623", "x", "x", "x")
			
		Else
			If "<%=rs1("Tracking_Flg")%>" = "N" Then 'TRACKING_FLG
				.txtTrackingNo.ReadOnly = True
				.txtTrackingNo.classname = "protected"
				.txtTrackingNo.tabindex = "-1"
			Else
				.txtTrackingNo.ReadOnly = False
				.txtTrackingNo.classname = "required"
				.txtTrackingNo.tabindex = "1"
			End If	

			If "<%=rs1("Phantom_Flg")%>" = "Y" Then 'PHANTOM_FLG
				
				Call parent.DisplayMsgBox("189214", "x", "x", "x")
				
			Else
				.txtItemNm.Value		= "<%=ConvSPChars(rs1("Item_Nm"))%>"			'☆: Item Name
				.txtUnit.value			= "<%=ConvSPChars(rs1("Order_Unit_Mfg"))%>"		'☆: Item Name
				.txtProdLT.value		= "<%=rs1("Order_Lt_Mfg")%>"					'☆: Item Name
				.txtMaxLotQty.value		= "<%=rs1("Max_Mrp_Qty")%>"						'☆: Item Name
				.txtMinLotQty.value		= "<%=rs1("Min_Mrp_Qty")%>"						'☆: Item Name
				.txtRoundingQty.value	= "<%=rs1("Round_Qty")%>"						'☆: Item Name
				.txtSLCd.value			= "<%=ConvSPChars(rs1("Major_Sl_Cd"))%>"		'☆: Item Name
				.txtSLNm.value			= ""
				.txtBaseUnit.value		= "<%=ConvSPChars(rs1("Basic_Unit"))%>"			'☆: Basic Unit
				.txtSpecification.value	= "<%=ConvSPChars(rs1("Spec"))%>"
			End If
		End If

		Parent.LookUpItemByPlantSuccess

	End With
</Script>
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
