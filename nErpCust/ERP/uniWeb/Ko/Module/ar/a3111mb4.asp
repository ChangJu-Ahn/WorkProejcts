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

Dim pAR0139											'��ȸ�� ComProxy Dll ��� ���� 

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
	
    Set pAR0139 = Server.CreateObject("Ar0139.ALookupRcptDcDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAR0139 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView
    pAR0139.ImportAAllcRcptAllcNo = Trim(Request("txtAllcNo"))
    pAR0139.ImportARcptDcSeq = Trim(strItemSeq)
    pAR0139.CommandSent = "lookup"
    
   'Call SvrMsgBox("Condition ->" & Request("txtAllcNo") & " : " & Request("strItemSeq") , vbInformation, I_MKSCRIPT)
    
    pAR0139.ServerLocation = ggServerIP

	'-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAR0139 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
	
    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
	pAR0139.ComCfg = gConnectionString
    pAR0139.Execute

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAR0139 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAR0139.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAR0139.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAR0139.ExportErrEabSqlCodeSqlcode, _
						    pAR0139.ExportErrEabSqlCodeSeverity, _
						    pAR0139.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAR0139.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAR0139 = Nothing
		Response.End
	End If  
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
    GroupCount = pAR0139.ExportGroupCount

	' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
	' ����/���� �� ���, ���ƿ� �°� ó���� 
'	If pAR0139.ExportPIndReqIndReqmtNo(GroupCount) = pAR0139.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAR0139.ExportNextPMPSRequirementIndReqmtNo
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
	    strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportARcptDcDtlDtlSeq(LngRow))%>"        
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlCd(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlNm(LngRow))%>"
        if  "<%=pAR0139.ExportItemACtrlItemColmDataType(LngRow)%>" = "D" then
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAR0139.ExportARcptDcDtlCtrlVal(LngRow))%>"    '4          		    
		Else
 			strData = strData & Chr(11) & "<%=pAR0139.ExportARcptDcDtlCtrlVal(LngRow)%>"        
		End if                                  
        
        strData = strData & Chr(11) & ""   
        if "<%=pAR0139.ExportItemACtrlItemTblId(LngRow)%>" = "" and "<%=pAR0139.ExportItemACtrlItemColmDataType(LngRow)%>" = "D" then
		    strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"                              '6
		Else
 			strData = strData & Chr(11) & "<%=pAR0139.ExportEabACtrlValRtnCtrlValC(LngRow)%>"  
		End if                                  
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportARcptDcSeq(LngRow))%>"                    
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemTblId(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmId(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmNm(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemColmDataType(LngRow))%>"
        strData = strData & Chr(11) & "<%=pAR0139.ExportItemACtrlItemDataLen(LngRow)%>"        
        strData = strData & Chr(11) & "<%=pAR0139.ExportItemAAssignAcctHqFg(LngRow)%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemMajorCd(LngRow))%>"
        strData = strData & Chr(11) & "<%=LngRow%>"
        strData = strData & Chr(11) & Chr(12)        
        '
        .frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcSeq(LngRow))%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcDtlDtlSeq(LngRow))%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        if  "<%=pAR0139.ExportItemACtrlItemColmDataType(LngRow)%>" = "D" then
			.frm1.vspdData3.Text = "<%=UNIDateClientFormat(pAR0139.ExportARcptDcDtlCtrlVal(LngRow))%>"    '4          		    
		Else
 			.frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcDtlCtrlVal(LngRow))%>"
		End if                                  
        
        
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        if "<%=pAR0139.ExportItemACtrlItemTblId(LngRow)%>" = "" and "<%=ConvSPChars(pAR0139.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" then
		    .frm1.vspdData3.Text =  "(Format : YYYY-MM-DD)"                              '6
		Else
 			.frm1.vspdData3.Text = "<%=pAR0139.ExportEabACtrlValRtnCtrlValC(LngRow)%>"  
		End if          
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcSeq(LngRow))%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pAR0139.ExportItemACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=pAR0139.ExportItemAAssignAcctHqFg(LngRow)%>"
        .frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemMajorCd(LngRow))%>"

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
   
    Set pAR0139 = Nothing
End Select
%>
</Script>
