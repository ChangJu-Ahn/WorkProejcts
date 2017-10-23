<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%Call LoadBasisGlobalInf%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b01mb5.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/11/15
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'☜: 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter 선언 
Dim strQryMode

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey1	
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strHsCd

lgStrPrevKey1 = UCase(Trim(Request("lgStrPrevKey1")))	
	
'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
	
Redim UNISqlId(0)
Redim UNIValue(0, 0)
	
UNISqlId(0) = "122600sab"
IF Request("txtHsCd") = "" Then
   strHsCd = "|"
ELSE
   strHsCd = FilterVar(Request("txtHsCd") , "''", "S")
END IF
	
UNIValue(0, 0) = strHsCd
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If rs0.EOF And rs0.BOF Then
'	Call DisplayMsgBox("126700", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	Set rs0 = Nothing					
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.LookUpHsNotOk" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtHsUnit.value = """ & ConvSPChars(rs0(3)) & """" & vbCrLf
Response.Write "</Script>" & vbCrLf
rs0.Close
Set rs0 = Nothing
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
Response.End
%>
