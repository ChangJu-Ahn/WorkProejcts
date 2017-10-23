<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3101bb9
'*  4. Program Name         : 실제원가 계산 
'*  5. Program Desc         : 실제원가 계산 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Cho Ig sung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf() 

'@Var_Declare
'--- Karrman_ADO
Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter 선언 
Dim strQryMode												'현재의 Query 상태를 위한 변수선언 

'Const DISCONNUPD  = "1"										'Disconnect + Update Mode
'Const DISCONNREAD = "2"										'Disconnect + ReadOnly Mode

'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------

Call HideStatusWnd 

																		'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow			' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          
Dim strPlantCd
Dim strInspClassCd
Dim strItemCd
   
lgStrPrevKey 	= Request("lgStrPrevKey")
LngMaxRow 	= Request("txtMaxRows")

'--- Karrman_ADO

strQryMode = Request("lgIntFlgMode")						'☜ : 현재 Query 상태를 받음 

Redim UNISqlId(0)
Redim UNIValue(0,2)

UNISqlId(0) = "C3101BA101"

UNIValue(0,0) = FilterVar(Request("txtYyyymm"),"''","S") 
'UNIValue(0,1) = ""
'UNIValue(0,2) = ""

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call ServerMesgBox("" , vbInformation, I_MKSCRIPT)
		
		rs0.Close
		Set rs0 = Nothing
					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	end if

	
'------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
%>

<Script Language=vbscript>
Dim strData
	
With Parent
	
<%
    For i=0 to rs0.RecordCount-1
%>
		strData = strData & Chr(11) & "0"

		if "<%=rs0("PROGRESS_YN")%>" = "" then 
  			strData = strData & Chr(11) & "N"	
		else
			strData = strData & Chr(11) & "<%=rs0("PROGRESS_YN")%>"
		end if

		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MINOR_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MINOR_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REFERENCE"))%>"
		strData = strData & Chr(11) & "<%=LngMaxRow + LngRow%>"
	       	strData = strData & Chr(11) & Chr(12)
<%
		rs0.MoveNext
	Next
%>    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData strData
		
	.lgStrPrevKey = ""
	<% ' Request값을 hidden input으로 넘겨줌 %>
	

<%
	rs0.Close
	Set rs0 = Nothing
%>	
	.DbQueryOk
		
	End with
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
