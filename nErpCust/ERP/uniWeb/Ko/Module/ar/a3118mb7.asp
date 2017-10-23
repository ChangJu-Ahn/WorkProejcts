<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : f6102mb7
'*  4. Program Name         :
'*  5. Program Desc         : ������ �����׸� ����Ÿ ��ȸ 
'*  6. Comproxy ����Ʈ     : Ar0099
'*  6. Modified date(First) : 2000/10/7
'*  7. Modified date(Last)  : �ۺ��� 
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

Dim pAr0099											'��ȸ�� ComProxy Dll ��� ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          
Dim strItemSeq
Dim AcctNm

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

On Error Resume Next

'Call SvrMsgBox("Condition ->" & Request("txtAcctCd") & " : " & Request("txtSttlNo") , vbInformation, I_MKSCRIPT)

Select Case strMode

Case CStr(UID_M0001)								'��: ���� ��ȸ/Prev/Next ��û�� ���� 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq = Request("txtSttlNo")
	
    Set pAr0099 = Server.CreateObject("Ar0099.ALookupArAdjustDtlSvr.1")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAr0099 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView      
    pAr0099.ImportAArAdjustAdjustNo = Trim(Request("txtAdjustNo"))
    pAr0099.ServerLocation = ggServerIP

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
    'pAr0099.ComCfg = "TCP Letitbe 2050"
    pAr0099.ComCfg = gConnectionString
    pAr0099.Execute
    
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAr0099 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	'-----------------------------------------
	'Com action result check area(DB,internal)
	'-----------------------------------------
	If Not (pAr0099.OperationStatusMessage = MSG_OK_STR) Then
	    Select Case pAr0099.OperationStatusMessage
            Case MSG_DEADLOCK_STR
                Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
            Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0099.ExportErrEabSqlCodeSqlcode, _
						    pAr0099.ExportErrEabSqlCodeSeverity, _
						    pAr0099.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
            Case Else
                Call DisplayMsgBox(pAr0099.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
            End Select

		Set pAr0099 = Nothing
		Response.End 
	End If  
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
   	GroupCount = pAr0099.ExportGroupCount

	' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
	' ����/���� �� ���, ���ƿ� �°� ó���� 
'	If pAr0099.ExportPIndReqIndReqmtNo(GroupCount) = pAr0099.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAr0099.ExportNextPMPSRequirementIndReqmtNo
'    End If
%>

<Script Language=vbscript>
    Dim lngMaxRows       
    Dim strData
    Dim lRows
    Dim tmpDrCrFg	
	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		
	lngMaxRows = .frm1.vspdData3.MaxRows
	.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)
<%      
	For LngRow = 1 To GroupCount
%>
<%'@ExportView - ����� %>

	    strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportAArAdjustDtlDtlSeq(LngRow))%>"     '1   
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemCtrlCd(LngRow))%>"          '2 
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemCtrlNm(LngRow))%>"          '3
        If "<%=ConvSPChars(pAr0099.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAr0099.ExportAArAdjustDtlCtrlVal(LngRow))%>"    '4  
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportAArAdjustDtlCtrlVal(LngRow))%>"    '4  
		ENd IF	
        strData = strData & Chr(11) & ""        												'5
        
		If "<%=ConvSPChars(pAr0099.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"  								'6									'6
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportEabACtrlValRtnCtrlValC(LngRow))%>" '6
		End If          							'6
        strData = strData & Chr(11) & "<%=strItemSeq%>"											'7	
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemTblId(LngRow))%>" 			'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemDataColmId(LngRow))%>"		'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemDataColmNm(LngRow))%>"		'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemColmDataType(LngRow))%>"    '11
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemDataLen(LngRow))%>"        	'12
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportAAssignAcctHqFg(LngRow))%>"		'13
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0099.ExportACtrlItemMajorCd(LngRow))%>"		'13
        strData = strData & Chr(11) & <%=LngRow%>												'14			
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportAApAdjustDtlDtlSeq(LngRow))%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        If "<%=ConvSPChars(pAr0099.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
			.frm1.vspdData3.Text = "<%=UNIDateClientFormat(pAr0099.ExportAApAdjustDtlCtrlVal(LngRow))%>"   
        ELSE
			.frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportAApAdjustDtlCtrlVal(LngRow))%>"   
        END IF
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        if "<%=ConvSPChars(pAr0099.ExportACtrlItemColmDataType(LngRow))%>" = "D" then		
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"  								'6
        ELSE	
			.frm1.vspdData3.Text =  "<%=ConvSPChars(pAr0099.ExportEabACtrlValRtnCtrlValC(LngRow))%>"
		end if
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemDataLen(LngRow))%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportAAssignAcctHqFg(LngRow))%>"
		.frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0099.ExportACtrlItemMajorCd(LngRow))%>"	
<%      
    Next
%>    
         
    .frm1.vspdData2.MaxRows = 0
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData
			
	.DbQueryOk2
		
	End With
</Script>	
<% 
   
    Set pAr0099 = Nothing
End Select
%>
</Script>
