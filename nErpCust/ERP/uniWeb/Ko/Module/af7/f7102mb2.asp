<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         : �������� �����׸� Setting(ȸ����� ab/AB0019mb1.asp�� ����)
'*  6. Modified date(First) : 2000/09/18
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Dim pP21011											'�Է�/������ ComProxy Dll ��� ���� 
Dim pAb0019											'��ȸ�� ComProxy Dll ��� ���� 

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

'Select Case strMode

'Case CStr(UID_M0001)								'��: ���� ��ȸ/Prev/Next ��û�� ���� 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq = Request("txtItemSeq")
	
    Set pAb0019 = Server.CreateObject("Ab0019.ALookupAcctSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAb0019 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView
    pAb0019.ImportAAcctAcctCd = Trim(Request("txtAcctCd"))   
    pAb0019.CommandSent = "lookupac"
    
    'Call SvrMsgBox("Condition ->" &Request("txtAcctCd")    & " : " & Request("txtItemSeq") , vbInformation, I_MKSCRIPT)
    
    pAb0019.ServerLocation = ggServerIP
    
	'-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAb0019 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
    pAb0019.ComCfg = gConnectionString
    pAb0019.Execute
    
    AcctNm = pAb0019.ExportAAcctAcctNm

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAb0019 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAb0019.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAb0019.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAb0019.ExportErrEabSqlCodeSqlcode, _
						    pAb0019.ExportErrEabSqlCodeSeverity, _
						    pAb0019.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAb0019.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAb0019 = Nothing
		Response.End
	End If
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
    	GroupCount = pAb0019.ExportGroupCount

	' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
	' ����/���� �� ���, ���ƿ� �°� ó���� 
'	If pAb0019.ExportPIndReqIndReqmtNo(GroupCount) = pAb0019.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAb0019.ExportNextPMPSRequirementIndReqmtNo
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
		strData = strData & Chr(11) & "<%=strItemSeq%>"
        strData = strData & Chr(11) & "<%=pAb0019.ExportAAcctCtrlAssnCtrlItemSeq(LngRow)%>"     '1   
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlCd(LngRow))%>"              '2 
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlNm(LngRow))%>"              '3
        strData = strData & Chr(11) & ""                                                        '4  
        strData = strData & Chr(11) & ""        						'5
        
		if "<%=ConvSPChars(pAb0019.ExportACtrlItemTblId(LngRow))%>" = "" and "<%=ConvSPChars(pAb0019.ExportACtrlItemColmDataType(LngRow))%>" = "D" then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"                              '6
		Else
 			strData = strData & Chr(11) & ""  						'6
		End if
        
        strData = strData & Chr(11) & "<%=strItemSeq%>"						'7	
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemTblId(LngRow))%>" 		'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemDataColmId(LngRow))%>"		'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemDataColmNm(LngRow))%>"		'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemColmDataType(LngRow))%>"        '11
        strData = strData & Chr(11) & "<%=pAb0019.ExportACtrlItemDataLen(LngRow)%>"        	'12

		if "<%=ConvSPChars(pAb0019.ExportAAcctCtrlAssnDrFg(LngRow))%>" = "Y" and "<%=ConvSPChars(pAb0019.ExportAAcctCtrlAssnCrFg(LngRow))%>" = "Y" then
			tmpDrCrFg = "DC" 
        elseif "<%=ConvSPChars(pAb0019.ExportAAcctCtrlAssnDrFg(LngRow))%>" = "Y" then
           tmpDrCrFg = "DR"
        elseif "<%=ConvSPChars(pAb0019.ExportAAcctCtrlAssnCrFg(LngRow))%>" = "Y" then 
           tmpDrCrFg = "CR"
        else
           tmpDrCrFg = ""
        end if
 
        strData = strData & Chr(11) & tmpDrCrFg		'13
        strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemMajorCd(LngRow))%>"						'14			
        strData = strData & Chr(11) & <%=LngRow%>
        strData = strData & Chr(11) & Chr(12)
        
        

        .ggoSpread.Source = .frm1.vspdData2
        .frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 0
        .frm1.vspdData3.Text = .ggoSpread.InsertFlag
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=pAb0019.ExportAAcctCtrlAssnCtrlItemSeq(LngRow)%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        .frm1.vspdData3.Text = ""
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text = ""
        .frm1.vspdData3.Col = 7
		if "<%=ConvSPChars(pAb0019.ExportACtrlItemTblId(LngRow))%>" = "" and "<%=ConvSPChars(pAb0019.ExportACtrlItemColmDataType(LngRow))%>" = "D" then
		        .frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"                              '6
		Else
 			.frm1.vspdData3.Text = ""						'6
		End if
		    
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pAb0019.ExportACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = tmpDrCrFg
        .frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text= "<%=ConvSPChars(pAb0019.ExportACtrlItemMajorCd(LngRow))%>"						'14			
<%      
    Next
%>    
    .frm1.vspdData2.MaxRows = 0
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData
		
	For lRows = 1 To .frm1.vspdData2.MaxRows
	    .frm1.vspdData2.Row = lRows
	    .frm1.vspdData2.Col = 0
	    .frm1.vspdData2.Text = .ggoSpread.InsertFlag
	Next
			
	'.frm1.vspdData.Row = .frm1.vspdData.ActiveRow
	'.frm1.vspdData.Col = 4  '�����ڵ�� 
	'.frm1.vspdData.Text = "<%=ConvSPChars(AcctNm)%>"

	.DbQueryOk3
		
	End With
	
<% 
   
    Set pAb0019 = Nothing
'End Select
%>
</Script>
