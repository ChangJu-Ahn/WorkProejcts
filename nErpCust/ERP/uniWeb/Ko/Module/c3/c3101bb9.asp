<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c3101bb9
'*  4. Program Name         : �������� ��� 
'*  5. Program Desc         : �������� ��� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Cho Ig sung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call LoadBasisGlobalInf() 

'@Var_Declare
'--- Karrman_ADO
Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 
Dim strQryMode												'������ Query ���¸� ���� �������� 

'Const DISCONNUPD  = "1"										'Disconnect + Update Mode
'Const DISCONNREAD = "2"										'Disconnect + ReadOnly Mode

'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------

Call HideStatusWnd 

																		'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim strMode													'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow			' ���� �׸����� �ִ�Row
Dim LngRow
Dim intGroupCount          
Dim strPlantCd
Dim strInspClassCd
Dim strItemCd
   
lgStrPrevKey 	= Request("lgStrPrevKey")
LngMaxRow 	= Request("txtMaxRows")

'--- Karrman_ADO

strQryMode = Request("lgIntFlgMode")						'�� : ���� Query ���¸� ���� 

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
					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
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
	<% ' Request���� hidden input���� �Ѱ��� %>
	

<%
	rs0.Close
	Set rs0 = Nothing
%>	
	.DbQueryOk
		
	End with
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
