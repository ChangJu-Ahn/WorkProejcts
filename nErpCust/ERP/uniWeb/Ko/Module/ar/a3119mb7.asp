<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : f6102mb7
'*  4. Program Name         :
'*  5. Program Desc         : ������ �����׸� ����Ÿ ��ȸ 
'*  6. Comproxy ����Ʈ     : Ar0119
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


Dim pAr0119											'��ȸ�� ComProxy Dll ��� ���� 

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

'Call SvrMsgBox("Condition ->" & Request("txtAdjustNo") & " : " & Request("txtSttlNo") , vbInformation, I_MKSCRIPT)

Select Case strMode

Case CStr(UID_M0001)								'��: ���� ��ȸ/Prev/Next ��û�� ���� 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq = Request("txtSttlNo")
	
    Set pAr0119 = Server.CreateObject("Ar0119.ALookupRcptAdjustDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------    
    If Err.Number <> 0 Then
		Set pAr0119 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView  
    
    pAr0119.ImportARcptAdjustAdjustNo = Trim(Request("txtAdjustNo"))
    pAr0119.ServerLocation = ggServerIP
	
    If Err.Number <> 0 Then
		Set pAr0119 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
	
    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
    'pAr0119.ComCfg = "TCP Letitbe 2050"
    pAr0119.ComCfg = gConnectionString
    pAr0119.Execute
    
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------    
    If Err.Number <> 0 Then
		Set pAr0119 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAr0119.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAr0119.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0119.ExportErrEabSqlCodeSqlcode, _
						    pAr0119.ExportErrEabSqlCodeSeverity, _
						    pAr0119.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAr0119.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAr0119 = Nothing
		Response.End
	End If      
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
   	GroupCount = pAr0119.ExportGroupCount
'Call ServerMesgBox(pAr0119.ExportGroupCount, vbCritical, I_MKSCRIPT)
	' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
	' ����/���� �� ���, ���ƿ� �°� ó���� 
'	If pAr0119.ExportPIndReqIndReqmtNo(GroupCount) = pAr0119.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAr0119.ExportNextPMPSRequirementIndReqmtNo
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

	    strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlDtlSeq(LngRow))%>"     '1   
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlCd(LngRow))%>"          '2 
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlNm(LngRow))%>"          '3
        If "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"    '4  
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"    '4  
		ENd IF	
        strData = strData & Chr(11) & ""        												'5
        
		If "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"  								'6
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportEabACtrlValRtnCtrlValC(LngRow))%>" '6
		End If          							'6
        strData = strData & Chr(11) & "<%=strItemSeq%>"											'7	
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemTblId(LngRow))%>" 			'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmId(LngRow))%>"		'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmNm(LngRow))%>"		'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>"    '11
        strData = strData & Chr(11) & "<%=pAr0119.ExportACtrlItemDataLen(LngRow)%>"        	'12
        strData = strData & Chr(11) & "<%=pAr0119.ExportAAssignAcctHqFg(LngRow)%>"		'13
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemMajorCd(LngRow))%>"			'13
        strData = strData & Chr(11) & <%=LngRow%>												'14			
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlDtlSeq(LngRow))%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        If "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
			.frm1.vspdData3.Text = "<%=UNIDateClientFormat(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"   
        ELSE
			.frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"   
        END IF
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        if "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" then		
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"  	
        ELSE	
			.frm1.vspdData3.Text =  "<%=ConvSPChars(pAr0119.ExportItemEabACtrlValRtnCtrlValC(LngRow))%>"
		end if
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pAr0119.ExportACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=pAr0119.ExportAAssignAcctHqFg(LngRow)%>"
		.frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemMajorCd(LngRow))%>"
<%      
    Next
%>    
         
    .frm1.vspdData2.MaxRows = GroupCount
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData
			
	.DbQueryOk2
		
	End With
</Script>	
<% 
   
    Set pAr0119 = Nothing
End Select
%>
</Script>
