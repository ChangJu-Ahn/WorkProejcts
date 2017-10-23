<!-- #Include file="../inc/CommResponse.inc" -->
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<TITLE>Find</TITLE>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">
Const ActionActiveCell = 0

Dim lgobjDoc		' ������ ȭ�� ��ü ������ 
Dim lgFormType	' ������ ȭ�� Ÿ��(
Dim lgblnTab	' ������ ȭ���� Tab ���� 
Dim objSheet	' Multi ȭ���� ��� Sheet ��ü ������ 
Dim blnSheet
Dim vaSpread1	' Sheet ��ü�� �ӽ� ������ 

Dim lgiSrcIdx		' TEXT Index ��ġ 
Dim lgIntCol, lgLngRow ' Sheet Index ��ġ 
Dim lgiTabIdx		' ���� Tab��ġ 
Dim lgLogoName

lgLogoName = "<%=Request.Cookies("unierp")("gLogoName")%>"
'==============================================================================
' ȭ�� �ε尡 ������ �߻��ϴ� �̺�Ʈ 
'==============================================================================
Function window_onload()
	Dim arrParent
	Dim arr1, arr2, arr3

	arrParent = window.dialogArguments

	arr1 = arrParent(0)
	arr2 = arrParent(1)
	arr3 =arrParent(2)

	Set lgobjDoc = arr1(0) ' ������ ȭ�� ������Ŵ 
	lgFormType = arr2(0)   ' ȭ�� ���� 
	lgblnTab = arr3(0)     ' Tab ���� 

    txtChar.focus

	Select Case lgFormType
		Case C_SINGLE
			lgiSrcIdx = 0
			cboSheet.disabled = True
			chkSheet.disabled = True
		Case C_MULTI, C_SINGLEMULTI
			lgiSrcIdx = 0
			lgIntCol = 0 : lgLngRow = 0
			cboSheet.disabled = False
			chkSheet.disabled = False
			
			Call SetSheetInfo()	' ������ ȭ�鿡�� Sheet�� Columns�� ���� 
	End Select

	Call SetComboBox() ' by Shin hyoung jae, 2001/4/9
End Function

'==============================================================================
' Find ��ư�� Ŭ���ϸ� ó���ϴ� �̺�Ʈ 
'==============================================================================
Function btnFind_onclick()
	' by Shin hyoung jae, 2001/4/2
	If Trim(txtChar.Value)  = "" Then
		MsgBox "ã�� ���ڿ��� �Է��ϼ���", vbInformation,lgLogoName
		txtChar.focus()
		Exit Function
	End If
	
	window.setTimeout "vbscript:FindText()", 50
	
End Function

'==============================================================================
' Cancel ��ư�� Ŭ���ϸ� ó���ϴ� �̺�Ʈ 
'==============================================================================
Function btnCancel_onclick()
	self.close
End Function

'==============================================================================
' Tab ȭ���� �����ϴ� HTML Tag������ Find
'==============================================================================
Function FindText()
	Dim objDoc, blnFind
	Dim bIsSearch 

	blnFind = False
	bIsSearch = False
	
	Set objDoc = lgobjDoc.document.all
	
	If lgiSrcIdx = 0 And rdoUp.checked = True Then lgiSrcIdx = objDoc.length-1
	
	Do Until lgiSrcIdx > objDoc.length-1 Or lgiSrcIdx < 0
	
		Select Case objDoc(lgiSrcIdx).tagName
			Case "DIV"
				If rdoDn.checked = True Then
					lgiTabIdx = lgiTabIdx + 1					' Tab ���� 
				Else
					lgiTabIdx = lgiTabIdx - 1					' Tab ���� 
				End If
			Case "INPUT"
				If UCase(objDoc(lgiSrcIdx).Type) = "TEXT" Then	' Text�ڽ��� ��� 

					' by Shin hyoung jae, 2001/4/3
					If UCase(objDoc(lgiSrcIdx).Style.textTransform) = "UPPERCASE" Then
						bIsSearch = CheckCase(UCase(objDoc(lgiSrcIdx).Value))
					ElseIf UCase(objDoc(lgiSrcIdx).Style.textTransform) = "LOWERCASE" Then
						bIsSearch = CheckCase(LCase(objDoc(lgiSrcIdx).Value))
					Else
						bIsSearch = CheckCase(objDoc(lgiSrcIdx).Value)
					End If

					If bIsSearch Then	' ��/�ҹ��� ���� 
						Call CheckTabs(lgiTabIdx)		' Tab �̵� 
						Call objDoc(lgiSrcIdx).select()			' ���� Text�� Select
						blnFind = True
					End If
				End If
			Case "OBJECT"
				If UCase(objDoc(lgiSrcIdx).title) = "SPREAD" Then	' Sheet �� ��� 
					
					Set vaSpread1 = objDoc(lgiSrcIdx)				' Sheet ��ü ���� 
					
					If chkSheet.checked Then
						If cboSheet.value = "" Then
							MsgBox "COLUMN�� ������ �ֽʽÿ�!", vbInformation,lgLogoName
							Exit Function
						End If
					End If

					If lgLngRow <> 0 Or lgIntCol <> 0 Then
						If chkSheet.checked Then	
							If rdoUp.checked = True Then
								lgLngRow = vaSpread1.ActiveRow-1
							Else
								lgLngRow = vaSpread1.ActiveRow+1
							End If
						Else
							lgLngRow = vaSpread1.ActiveRow
						End If
						If rdoUp.checked = True Then
							lgIntCol = vaSpread1.ActiveCol-1
						Else
							lgIntCol = vaSpread1.ActiveCol+1
						End If
					ElseIf rdoUp.checked = True Then
						lgLngRow = vaSpread1.MaxRows: lgIntCol = vaSpread1.MaxCols-1
						vaSpread1.Row = lgLngRow
						vaSpread1.Col = lgIntCol
						vaSpread1.Action = ActionActiveCell	' Focus
					Else
						lgLngRow = 1: lgIntCol = 1
						vaSpread1.Row = lgLngRow
						vaSpread1.Col = lgIntCol
						vaSpread1.Action = ActionActiveCell	' Focus
					End If
					
					Do Until lgLngRow > vaSpread1.MaxRows Or lgLngRow < 0
						
						vaSpread1.Row = lgLngRow
						
						If chkSheet.checked Then	lgIntCol = CInt(cboSheet.value)
						vaSpread1.Col = lgIntCol
						
						If CheckCase(vaSpread1.Text) Then	' ��/�ҹ��� ���� 
							Call CheckTabs(lgiTabIdx)		' Tab �̵� 
							vaSpread1.Action = ActionActiveCell	' Focus
							blnFind = True
						End If
						
						' Columns ��/�� 
						If rdoUp.checked = True Then
							If chkSheet.checked Then
								lgLngRow = lgLngRow - 1
							Else
								lgIntCol = lgIntCol - 1
								If lgIntCol < 1 Then
									lgLngRow = lgLngRow - 1
									lgIntCol = vaSpread1.MaxCols-1
								End If
							End If
							If lgLngRow < 1 Then Exit Do
						Else
							If chkSheet.checked Then
								lgLngRow = lgLngRow + 1
							Else
								lgIntCol = lgIntCol + 1
								If lgIntCol > vaSpread1.MaxCols-1 Then  ' by Shin hyoung jae, 2001/ org >=
									lgLngRow = lgLngRow + 1
									lgIntCol = 1
								End If
							End If
							If lgLngRow > vaSpread1.MaxRows Then Exit Do
						End If
						
						If blnFind Then Exit Function
					Loop
					
				End If
		End Select
			
		If rdoDn.checked = True Then
			lgiSrcIdx = lgiSrcIdx + 1
		Else
			lgiSrcIdx = lgiSrcIdx - 1
		End If
		
		If blnFind Then Exit Do
	Loop

	If blnFind = False Then	' �������� �̾����� ���� ���� ã�⵵ �����ϸ� �޼��� ġ�� 
		MsgBox "ã�� ���ڿ��� �������� �ʽ��ϴ�.", vbInformation,lgLogoName
		Call lgobjDoc.document.selection.empty
		lgiSrcIdx = 0: lgLngRow = 0 : lgIntCol = 0: lgiTabIdx = 0
	End If
	
End Function

'==============================================================================
' ��/�ҹ��� ������ �� üũ�ϴ� Function
'==============================================================================
Function CheckCase(Byval strVal)
	CheckCase = False

	If chkType.checked = True Then	' ��/�ҹ��� ������ 
		If Instr(1, strVal , txtChar.value, 0) > 0 Then
			CheckCase = True
		End If
	Else
		If Instr(1, UCase(strVal) , UCase(txtChar.value)) > 0 Then
			CheckCase = True
		End If
	End If

End Function

'==============================================================================
' Tab�� �����ϴ� ȭ���� ��� �ش� Tab���� �̵� 
'==============================================================================
Function CheckTabs(Byval iTab)
	If Not lgblnTab Then Exit Function
	Select Case iTab
		Case 0
		Case 1
			Call lgobjDoc.ClickTab1
		Case 2
			Call lgobjDoc.ClickTab2
		Case 3
			Call lgobjDoc.ClickTab3
		Case 4
			Call lgobjDoc.ClickTab4
	End Select
End Function

'======================================================================================================
'	Function Name : SetCombo(pCombo, byval Code, byval Name)
'	Description : �޺��ڽ��� �����͸� Add�ϴ� �Լ� 
'	Parameters  :
'		pCombo	-	Combo Object Name(SELECT Tag Name)		
'		Code		-	Code
'		Name		-	Text Value
'======================================================================================================
Sub SetCombo(pCombo,  strValue,  strText)
	Dim objEl
			
	Set objEl = Document.CreateElement("OPTION")	
	objEl.Text = strText
	objEl.Value = strValue
				
	pcombo.Add(objEl)
	Set objEl = Nothing

End Sub

' by Shin hyoung jae, 2001/4/9
Sub SetComboBox()
	If  chkSheet.checked = False Then
		cboSheet.disabled = True
	Else
		cboSheet.disabled = False
	End If
End Sub

Function Document_onKeyDown()
	Dim KeyCode 
	KeyCode = window.event.keyCode
	Select Case KeyCode
		Case  13
			Call btnFind_onclick
		Case  27 
			self.close
	End Select
End  Function

</SCRIPT>
<Script language=jscript>
var objSheet


function SetSheetInfo()
{
	var i, j, iSheetCnt
	
	objSheet = lgObjDoc.document.all.tags("OBJECT");
	iSheetCnt = 0;
	
	for (i=0; i < objSheet.length; i++) {
		if (objSheet(i).title.toUpperCase() == "SPREAD") {
			++iSheetCnt;
			objSheet(i).Row = 0;
			for (j=1; j < objSheet(i).MaxCols; j++) {
				objSheet(i).Col = j;
				if (objSheet(i).Text != "")
				{
				  if (objSheet(i).ColHidden != true ) 
				  {
					SetCombo(cboSheet, j, objSheet(i).Text);
                  }
				}
			}
		}
	}
}

</Script>
</HEAD>

<BODY TABINDEX="-1" SCROLL=no>   
<TABLE WIDTH="98%" HEIGHT="90%" CELLSPACING="0" CELLPADDING="0" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
	<TR>
		<TD WIDTH="100%">
			<FIELDSET>
				<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=1>
					<TR>
						<TD CLASS="TD5"> ã�� ���ڿ� </TD>
						<TD CLASS="TD6"> <INPUT TYPE=TEXT NAME="txtChar" SIZE=20 MAXLENGTH=20></TD>
					</TR>
				</TABLE>
			</FIELDSET>
			<FIELDSET>
				<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=1>
					<TR>
						<TD CLASS="TD5"> ã�� ���� </TD>
						<TD CLASS="TD6">
							<INPUT TYPE=RADIO ID="rdoUp" NAME="rdoDirect" CLASS="RADIO"><LABEL FOR="rdoUp">����</LABEL>&nbsp;
							<INPUT TYPE=RADIO ID="rdoDn" NAME="rdoDirect" CLASS="RADIO" CHECKED><LABEL FOR="rdoDn">�Ʒ���</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5"> &nbsp;</TD>
						<TD CLASS="TD6">
							<INPUT TYPE=CHECKBOX CLASS="RADIO" ID="chkType"> <LABEL FOR="chkType">��/�ҹ��� ����</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5"> &nbsp;</TD>
						<TD CLASS="TD6">
							<INPUT TYPE=CHECKBOX CLASS="RADIO" ID="chkSheet" NAME="chkSheet" onClick="vbscript:SetComboBox()"> <LABEL FOR="chkSheet">SHEET COLUMN��</LABEL> 
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5"> &nbsp;</TD>
						<TD CLASS="TD6">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME="cboSheet" STYLE="width: 200"></SELECT>
						</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% ALIGN="CENTER">
			<TABLE CELLSPACING=10>
			 <TR>
				<TD><IMG SRC="../image/btnNext_off.gif" BORDER="0" Style="CURSOR: hand" ALT="���� ã��" NAME="btnFind" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/btnNext_on.gif',1)"></TD>
				<TD><IMG SRC="../image/btnCancel_off.gif" BORDER="0" Style="CURSOR: hand" ALT="�� ��" NAME="btnCancel" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/btnCancel_on.gif',1)"></TD>
			 </TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>

