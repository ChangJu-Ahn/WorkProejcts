<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : �Աݹ��� 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/10/18
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Dim pAP0109											'��ȸ�� ComProxy Dll ��� ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount
Dim strItemSeq          

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strItemSeq = Request("txtItemSeq")

On Error Resume Next

Select Case strMode

Case CStr(UID_M0001)								'��: ���� ��ȸ/Prev/Next ��û�� ���� 

	lgStrPrevKey = Request("lgStrPrevKey")
	
    Set pAP0109 = Server.CreateObject("AP0109.ALookupPaymDcDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAP0109 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView
    pAP0109.ImprotAAllcPaymPaymNo = Request("txtAllcNo")
    pAP0109.ImportAPaymDcSeq = strItemSeq
    pAP0109.CommandSent = "lookup"
    
    'Call SvrMsgBox("Condition ->" & Request("txtAllcNo") & " : " & strItemSeq , vbInformation, I_MKSCRIPT)
    
    pAP0109.ServerLocation = ggServerIP
    
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAP0109 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
	pAP0109.ComCfg = gConnectionString
    pAP0109.Execute

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAP0109 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAP0109.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAP0109.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAP0109.ExportErrEabSqlCodeSqlcode, _
						    pAP0109.ExportErrEabSqlCodeSeverity, _
						    pAP0109.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAP0109.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAP0109 = Nothing
		Response.End
	End If
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
    GroupCount = pAP0109.ExportGroupCount

	' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
	' ����/���� �� ���, ���ƿ� �°� ó���� 
'	If pAP0109.ExportPIndReqIndReqmtNo(GroupCount) = pAP0109.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAP0109.ExportNextPMPSRequirementIndReqmtNo
'    End If
%>

<Script Language=vbscript>
    Dim lngMaxRows       
    Dim strData
	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		
		lngMaxRows = .frm1.vspdData3.MaxRows
		.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)		
<%      
	For LngRow = 1 To GroupCount
%>
	    strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExortAPaymDcDtlDtlSeq(LngRow))%>"        
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemCtrlCd(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemCtrlNm(LngRow))%>"
        If "<%=ConvSPChars(pAP0109.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" Then
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAP0109.ExortAPaymDcDtlCtrlVal(LngRow))%>"    '4  
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExortAPaymDcDtlCtrlVal(LngRow))%>"        
		ENd IF	       
        
        strData = strData & Chr(11) & ""    
        if "<%=ConvSPChars(pAP0109.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" then		
        	strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"  								'6
        ELSE	
			strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExortEabACtrlValRtnCtrlValC(LngRow))%>"  
		end if    
        
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExortAPaymDcSeq(LngRow))%>"                    
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemTblId(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemDataColmId(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemDataColmNm(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemColmDataType(LngRow))%>"
        strData = strData & Chr(11) & "<%=pAP0109.ExportItemACtrlItemDataLen(LngRow)%>"        
        strData = strData & Chr(11) & "<%=pAP0109.ExportAAssignAcctHqFg(LngRow)%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAP0109.ExportItemACtrlItemMajorCd(LngRow))%>"
        strData = strData & Chr(11) & "<%=LngRow%>"
        strData = strData & Chr(11) & Chr(12)        
        '
        .frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExortAPaymDcSeq(LngRow))%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExortAPaymDcDtlDtlSeq(LngRow))%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        If "<%=ConvSPChars(pAP0109.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" Then
			.frm1.vspdData3.Text =  "<%=UNIDateClientFormat(pAP0109.ExortAPaymDcDtlCtrlVal(LngRow))%>"    '4  
		Else
			.frm1.vspdData3.Text =  "<%=ConvSPChars(pAP0109.ExortAPaymDcDtlCtrlVal(LngRow))%>"        
		ENd IF	       
        
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        if "<%=ConvSPChars(pAP0109.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" then		
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"  								'6
        ELSE	
			.frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExortEabACtrlValRtnCtrlValC(LngRow))%>"  
		end if            
        
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExortAPaymDcSeq(LngRow))%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pAP0109.ExportItemACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=pAP0109.ExportAAssignAcctHqFg(LngRow)%>"
		.frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAP0109.ExportItemACtrlItemMajorCd(LngRow))%>"
<%      
    Next
%>    
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData strData
		
'		.lgStrPrevKey = "<%=StrNextKey%>"

'		If .frm1.vspdData2.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> "" Then	<% ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ %>
'			.DbQuery
'		Else
'			.frm1.hPlantCd.value = "<%=Request("txtPlantCd")%>"
'			.frm1.hReqStatus.value = "<%=Request("cboReqStatus")%>"			
'			.frm1.hFromReqrdDt.value = "<%=Request("txtFromReqrdDt")%>"
'			.frm1.hToReqrdDt.value = "<%=Request("txtToReqrdDt")%>"
'			.frm1.hItemCd.value = "<%=Request("txtItemCd")%>"
			
			.DbQueryOk2
'		End If
		
	End With
</Script>	
<% 
   
    Set pAP0109 = Nothing
End Select
%>
</Script>
