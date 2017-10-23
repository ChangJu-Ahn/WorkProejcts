<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4221mb1
'*  4. Program Name         : 
'*  5. Program Desc         : List Resource
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2		'DBAgent Parameter 선언 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim strFlag
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim i

'@Var_Declare

Call HideStatusWnd
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(2)
	Redim UNIValue(2, 1)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sae"
	UNISqlId(2) = "189700saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(2, 1) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    ' 자원그룹 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_RGGRP"
		%>
		<Script Language=vbscript>
			parent.frm1.txtResourceGroupNm.value = ""
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtResourceGroupNm.value = "<%=ConvSPChars(rs1("DESCRIPTION"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If
    
    ' Plant 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		strFlag = "ERROR_PLANT"
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs0.Close
		Set rs0 = Nothing
	End If

	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>			
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
		If strFlag = "ERROR_RGGRP" Then
			Call DisplayMsgBox("181700", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtResourceGroupCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
	End IF
      
	If rs2.EOF And rs2.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set rs2 = Nothing					
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If
%>
<Script Language=vbscript>
Dim LngLastRow
Dim LngMaxRow
Dim LngRow
Dim strTemp
Dim strData
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs2.RecordCount-1 
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs2("resource_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs2("description"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
<%		
			rs2.MoveNext
	Next
%>
	
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip strData
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hResourceGroupCd.value = "<%=ConvSPChars(Request("txtResourceGroupCd"))%>"
<%			
		rs2.Close
		Set rs2 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing															'☜: ActiveX Data Factory Object Nothing
%>
