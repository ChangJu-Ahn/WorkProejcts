<!--
'********************************************************************************************************
'*  1. Module Name          : Inventory											*
'*  2. Function Name        : DocumentNo Popup Business Part									*
'*  3. Program ID              : i1111bp1.asp												*
'*  4. Program Name         :															*
'*  5. Program Desc         : ���ҹ�ȣ�˾�													*
'*  7. Modified date(First) : 2000/04/18												*
'*  8. Modified date(Last)  : 2000/04/18										*
'*  9. Modifier (First)     : 										*
'* 10. Modifier (Last)      : 										*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              :																			*
'*                            2000/04/18 : Coding Start													*
'********************************************************************************************************-->
<%
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%								

err.clear
On Error Resume Next					
Call HideStatusWnd 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim LngRow
Dim GroupCount 
Dim i11118

lgStrPrevKey = Request("lgStrPrevKey")
	
Set i11118 = Server.CreateObject("i11118li.i11118ListGoodsMvmtSvr")
If Err.Number <> 0 Then
    Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '��:
	Set i11118 = Nothing																	'��: ComProxy UnLoad
	Response.End																				'��: Process End
End If

'set condition
i11118.ImportIGoodsMovementHeaderItemDocumentNo 	= Request("txtDocumentNo")
i11118.ImportIGoodsMovementHeaderDocumentYear 	 	= Request("hdnYear")	
i11118.ImportIGoodsMovementHeaderMovType 	 		= Request("txtMovType")
i11118.ImportIGoodsMovementHeaderTrnsTypeAsString   = Request("txtTrnsType")
i11118.ImportIGoodsMovementHeaderPlantCd            = Request("txtPlantCd")
i11118.DateFromProdWorkSetTempTimestamp             = UniConvDate(Request("txtFromDt"))			
i11118.DatetoProdWorkSetTempTimestamp               = UniConvDate(Request("txtToDt"))		


if Request("lgStrPrevKey") <> "" then
    i11118.ImportIGoodsMovementHeaderItemDocumentNo= Request("lgStrPrevKey")
end if


i11118.ServerLocation = ggServerIP
i11118.ComCfg = gConnectionString
i11118.Execute


'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
	Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '��:
	Set i11118 = Nothing																	'��: ComProxy UnLoad
	Response.End																				'��: Process End
End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If i11118.OperationStatusMessage <> MSG_OK_STR Then
		Call DisplayMsgBox(i11118.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		Set i11118 = Nothing																	'��: ComProxy UnLoad
		Response.End																				'��: Process End
	End If

	GroupCount = i11118.GroupExportCount 
    
    
    
	if i11118.ExportIGoodsMovementHeaderItemDocumentNo(GroupCount) = i11118.ExportNextIGoodsMovementHeaderItemDocumentNo Then
		StrNextKey = ""
	else
		StrNextKey = i11118.ExportNextIGoodsMovementHeaderItemDocumentNo
	end if
	
	if i11118.ExportIGoodsMovementHeaderPlantCd(GroupCount) = i11118.ExportNextIGoodsMovementHeaderPlantCd Then
		StrNextKey = ""
	else
		StrNextKey = i11118.ExportNextIGoodsMovementHeaderPlantCd
	end if

LngMaxRow = Request("txtMaxRows")

%>		    
<Script Language="vbscript">   
    Dim StrData
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim strTemp
    
	
With parent


	LngMaxRow = .frm1.vspdData.MaxRows	
	'Save previous Maxrow  
<%
	For LngRow = 1 To GroupCount
	
	
		'128���� �̻��� Data ������ ���� ��û�� �ʵ� �� ������ �ʵ� ��� 
		'7��° �ʵ��� ���� ��� 
%>
		strData = strData & Chr(11) & "<%=ConvSPChars(i11118.ExportIGoodsMovementHeaderItemDocumentNo(LngRow))%>"		'1
		strData = strData & Chr(11) & "<%=i11118.ExportIGoodsMovementHeaderDocumentYear(LngRow)%>"			'2
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(i11118.ExportIGoodsMovementHeaderDocumentDt(LngRow))%>"	'3
		strData = strData & Chr(11) & "<%=ConvSPChars(i11118.ExportIGoodsMovementHeaderMovType(LngRow))%>"	'4
		strData = strData & Chr(11) & "<%=ConvSPChars(i11118.ExportIGoodsMovementHeaderPlantCd(LngRow))%>"	'6
		strData = strData & Chr(11) & "<%=ConvSPChars(i11118.ExportIGoodsMovementHeaderDocumentText(LngRow))%>"	'8
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>   
		strData = strData & Chr(11) & Chr(12)
		
<%
   	Next
%>
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData strData
	.frm1.vspdData.focus
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	if .frm1.vspdData.MaxRows < .C_SHEETMAXROWS and .lgStrPrevKey <> "" Then
	   .DbQuery
	else
	   .hlgDocumentNo = "<%=ConvSPChars(Request("txtDocumentNo"))%>"
	   .hlgYear = "<%=Request("txtYear")%>"
	   .hlgFromDt = "<%=Request("txtFromDt")%>"	   
	   .hlgMovType = "<%=Request("txtToDt")%>"
	   .hlgTrnsType = "<%=ConvSPChars(Request("txtTrnsType"))%>"
	   
	   .DbQueryOk
	end if
		
End With

</Script>
<%
    Set i11118 = nothing
%>









