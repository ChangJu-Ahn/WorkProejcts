<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : B/L���� 
'*  3. Program ID           : S5113MA8
'*  4. Program Name         : B/L��Ȳ��ȸ 
'*  5. Program Desc         : B/L��Ȳ��ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/08/02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              '��: indicates that All variables must be declared in advance
'========================================================================================================
Const BIZ_PGM_ID 		= "s5113MB8.asp"					                    '��: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID	= "s5211ma1"												'��: JUMP�� �����Ͻ� ���� ASP�� 

'========================================================================================================
Const C_MaxKey          = 6                                           '��: key count of SpreadSheet

Const C_PopApplicant	= 1
Const C_PopSalesGrp		= 2
Const C_PopForwarder	= 3
Const C_PopSoNo			= 4
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
Dim IsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
														    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
														    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim lgBlnOpenedFlag
Dim	lgBlnApplicantChg
Dim lgBlnSalesGrpChg
Dim	lgBlnForwarderChg

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'========================================================================================================
Function InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE					'Indicates that current mode is Create mode
    gblnWinEvent = False
    lgBlnFlgChgValue = False								'Indicates that no value changed
    lgSortKey        = 1   

    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""										'initializes Previous Key
	    
	lgBlnApplicantChg	= False		' ������ ���濩�� 
	lgBlnSalesGrpChg	= False		' �����׷� ���濩�� 
	lgBlnForwarderChg	= False		' ����ä�� ���濩�� 
End Function
'=======================================================================================================
Sub SetDefaultVal()
	With frm1
		.txtFromDt.Text = UNIDateClientFormat(UniConvDateAToB(UniConvDateToYYYYMM(EndDate, Parent.gDateFormat, "-") & "-01", Parent.gServerDateFormat ,Parent.gAPDateFormat))
		.txtToDt.Text = EndDate	
		.rdoPostfiFlagAll.checked = True
		.txtPostfiFlag.value = frm1.rdoPostfiFlagAll.value   
		If Parent.gSalesGrp <> "" Then
			.txtSalesGrp.value = Parent.gSalesGrp
			Call txtSalesGrp_Onchange()
		End If

		.txtFromDt.Focus
	End With
	lgBlnFlgChgValue = False
End Sub
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S5113MA8","S","A","V20030301", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	Call SetSpreadLock 
	
End Sub
'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	
'========================================================================================
Function CookiePage()

	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	
	If frm1.vspdData.ActiveRow > 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		WriteCookie CookieSplit , frm1.vspdData.Text
	Else
		WriteCookie CookieSplit , ""
	End If

End Function
'===========================================================================
Function OpenBLHdr()
	Dim iCalledAspName
	Dim iArrRet
	Dim iArrParam(1)

	On Error Resume Next

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtFromDt.focus
		Exit Function
	End IF

	IsOpenPop = True

	frm1.vspdData.row = frm1.vspddata.activerow
	frm1.vspdData.Col = GetKeyPos("A",1)
	iArrParam(0) = frm1.vspdData.Text
	frm1.vspdData.Col = GetKeyPos("A",2)			' B/L DOC No.
	iArrParam(1) = frm1.vspdData.Text
   
	iCalledAspName = AskPRAspName("s5113ra9")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5113ra9", "x")
		IsOpenPop = False
		exit Function
	end if

	iArrRet = window.showModalDialog(iCalledAspName,Array(window.parent, iArrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	frm1.vspdData.Focus
	IsOpenPop = False

End Function
'===========================================================================
Function OpenBLDtl()
	Dim iCalledAspName
	Dim iArrRet
	Dim iArrParam(5)
	
	On Error Resume Next

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtFromDt.focus
		Exit Function
	End IF

	IsOpenPop = True

	frm1.vspdData.row = frm1.vspddata.activerow
	frm1.vspdData.Col = GetKeyPos("A",1)
	iArrParam(0) = frm1.vspdData.Text
	frm1.vspdData.Col = GetKeyPos("A",2)			' B/L DOC No.
	iArrParam(1) = frm1.vspdData.Text
	frm1.vspdData.Col = GetKeyPos("A",3)			' Applicant
	iArrParam(2) = frm1.vspdData.Text
	frm1.vspdData.Col = GetKeyPos("A",4)			' Applicant Name
	iArrParam(3) = frm1.vspdData.Text
	frm1.vspdData.Col = GetKeyPos("A",5)			' Currency
	iArrParam(4) = frm1.vspdData.Text
	frm1.vspdData.Col = GetKeyPos("A",6)			' B/L Amount
	iArrParam(5) = frm1.vspdData.Text
   
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s5112ra8")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5112ra8", "x")
		IsOpenPop = False
		exit Function
	end if

	iArrRet = window.showModalDialog(iCalledAspName,Array(window.parent, iArrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	frm1.vspdData.focus
	IsOpenPop = False

End Function
'---------------------------------------------------------------------------------------------------------
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopApplicant												
		iArrParam(1) = "B_BIZ_PARTNER PARTNER"			<%' TABLE ��Ī %>
		iArrParam(2) = Trim(frm1.txtApplicant.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "PARTNER.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"	<%' Where Condition%>
		iArrParam(5) = "������"						<%' TextBox ��Ī %>
			
		iArrField(0) = "PARTNER.BP_CD"					<%' Field��(0)%>
		iArrField(1) = "PARTNER.BP_NM"					<%' Field��(1)%>
		    
		iArrHeader(0) = "������"						<%' Header��(0)%>
		iArrHeader(1) = "�����ڸ�"					<%' Header��(1)%>

	Case C_PopSalesGrp												
		iArrParam(1) = "B_SALES_GRP"						<%' TABLE ��Ī %>
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)		<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
		iArrParam(5) = "�����׷�"					<%' TextBox ��Ī %>
		
		iArrField(0) = "SALES_GRP"						<%' Field��(0)%>
		iArrField(1) = "SALES_GRP_NM"					<%' Field��(1)%>
    
	    iArrHeader(0) = "�����׷�"					<%' Header��(0)%>
	    iArrHeader(1) = "�����׷��"				<%' Header��(1)%>

	Case C_PopForwarder												
		iArrParam(1) = "dbo.b_biz_partner BP"
		iArrParam(2) = Trim(frm1.txtForwarder.value)
		iArrParam(3) = ""
		iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ", " & FilterVar("S", "''", "S") & " ) AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = frm1.txtForwarder.alt
			
		iArrField(0) = "ED15" & Parent.gColSep & "BP.bp_cd"	<%' Field��(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "BP.bp_nm"	<%' Field��(1)%>

		iArrHeader(0) = "����ȸ��"
		iArrHeader(1) = "����ȸ���"
		
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	Select Case pvIntWhere
	Case C_PopApplicant	
		frm1.txtApplicant.focus 
	Case C_PopSalesGrp		
		frm1.txtSalesGrp.focus  
	Case C_PopForwarder		
		frm1.txtForwarder.focus  
	End Select	

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
	End If	
	
End Function
'========================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub
'========================================================================================================
Function OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
'---------------------------------------------------------------------------------------------------------
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopApplicant
		frm1.txtApplicant.value = pvArrRet(0) 
		frm1.txtApplicantNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	Case C_PopForwarder
		frm1.txtForwarder.value = pvArrRet(0) 
		frm1.txtForwarderNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029											  '��: Load table , B_numeric_format
  
    'Html���� tag ���ڰ� 1�� 2�� �����ϴ� �κ� ����Format
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
  
  	Call InitVariables											  '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	lgBlnOpenedFlag = True
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
End Sub
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
'=========================================================================================
Function txtApplicant_Onchange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtApplicant.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopApplicant) Then
				.txtApplicant.value = ""
				.txtApplicantNm.value = ""
				.txtApplicant.focus
			ELSE
				.txtfromDt.focus
			End If
			txtApplicant_Onchange = False
		Else
			.txtApplicantNm.value = ""
		End If
	End With
End Function
'==========================================================================================
Function txtForwarder_Onchange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtForwarder.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("AB", "''", "S") & "", C_PopForwarder) Then
				.txtForwarder.value = ""
				.txtForwarderNm.value = ""
				.txtForwarder.focus
			ELSE
				.txtSalesGrp.focus
			End If
			txtForwarder_Onchange = False
		Else
			.txtForwarderNm.value = ""
		End If
	End With
End Function
'==========================================================================================
Function txtSalesGrp_Onchange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtSalesGrp.value = ""
				.txtSalesGrpNm.value = ""
				.txtSalesGrp.focus
			Else
				.rdoPostfiFlagAll.focus
			End If
			txtSalesGrp_Onchange = False
		Else
			.txtSalesGrpNm.value = ""
		End If
	End With
End Function
'==========================================================================================
Function txtApplicant_OnKeyDown()
	lgBlnApplicantChg = True
	lgBlnFlgChgValue = True
End Function
'==========================================================================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgChgValue = True
End Function
'==========================================================================================
Function txtForwarder_OnKeyDown()
	lgBlnForwarderChg = True
	lgBlnFlgChgValue = True
End Function
'====================================================================================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnApplicantChg Then
		iStrCode = Trim(frm1.txtApplicant.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopApplicant) Then
				Call DisplayMsgBox("970000", "X", frm1.txtApplicant.alt, "X")
				frm1.txtApplicantNm.value = ""
				frm1.txtApplicant.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtApplicantNm.value = ""
		End If
		lgBlnApplicantChg	= False
	End If

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
				frm1.txtSalesGrpNm.value = ""
				frm1.txtSalesGrp.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnForwarderChg Then
		iStrCode = Trim(frm1.txtForwarder.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("AB", "''", "S") & "", C_PopForwarder) Then
				Call DisplayMsgBox("970000", "X", frm1.txtForwarder.alt, "X")
				frm1.txtForwarderNm.value = ""
				frm1.txtForwarder.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtForwarderNm.value = ""
		End If
		lgBlnForwarderChg = False
	End If

End Function
'====================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' ���� Popup Display
		'ATH 09/18: �ٸ�ȭ����� �ϰ����� ����. ǥ�ص� �ƴѰ� ���� 
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function
'********************************************************************************************************
Sub rdoPostfiFlagAll_OnClick()

	frm1.txtPostfiFlag.value = frm1.rdoPostfiFlagAll.value 

End Sub

Sub rdoPostfiFlagNo_OnClick()

	frm1.txtPostfiFlag.value = frm1.rdoPostfiFlagNo.value 

End Sub

Sub rdoPostfiFlagYes_OnClick()

	frm1.txtPostfiFlag.value = frm1.rdoPostfiFlagYes.value 

End Sub
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("00000000001")
   
	gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
        
    If Row = 0 Then
		frm1.vspdData.ReDraw = False
		frm1.vspdData.OperationMode = 0

        If lgSortKey = 1 Then
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
		frm1.vspdData.ReDraw = True
	Else
		frm1.vspdData.ReDraw = False		
		frm1.vspdData.OperationMode = 3
		frm1.vspdData.ReDraw = True
    End If
  
End Sub
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub  
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
		If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If

End Sub
'========================================================================================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("M")   
		frm1.txtFromDt.Focus				
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If
End Sub
'=======================================================================================================
Sub txtFromDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call FncQuery()
End Sub

Sub txtToDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call FncQuery()
End Sub

Sub vspdData_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call FncQuery()
End Sub
'*********************************************************************************************************
Function FncQuery() 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

	End With
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    
	' ��ȸ���� ��ȿ�� check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
	
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = C_ApplicantNm
   
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		'�� Frm1������ frm1���� 
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
    End If   

    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    FncExit = True
End Function
'========================================================================================================
Function DbQuery() 

	Err.Clear														'��: Protect system from crashing
	DbQuery = False													'��: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & Parent.UID_M0001					<%'��: �����Ͻ� ó�� ASP�� ���� %>
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			' Scroll�� 
			strVal = strVal & "&txtForwarder=" & Trim(.txtHForwarder.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtApplicant=" & Trim(.txtHApplicant.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtPostfiFlag=" & Trim(.txtHPostfiFlag.value)
		Else
			' ó�� ��ȸ�� 
			strVal = strVal & "&txtForwarder=" & Trim(.txtForwarder.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtApplicant=" & Trim(.txtApplicant.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			strVal = strVal & "&txtPostfiFlag=" & Trim(.txtPostfiFlag.value)
		End If

        strVal = strVal & "&lgPageNo="		 & lgPageNo					'��: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With    
    
	Call RunMyBizASP(MyBizASP, strVal)									<%'��: �����Ͻ� ASP �� ���� %>
    DbQuery = True    

End Function
'=========================================================================================================
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			Call vspdData_Click(1, 1)
		End If
		lgIntFlgMode = Parent.OPMD_UMODE
	Else
		Call SetFocusToDocument("M")
		frm1.txtFromDt.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>B/L��Ȳ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenBLHdr">B/L ������</A> | <A href="vbscript:OpenBLDtl">B/L ��������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>B/L������</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s5113ma8_fpDateTime1_txtFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s5113ma8_fpDateTime2_txtToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtApplicant" ALT="������" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopApplicant">&nbsp;<INPUT NAME="txtApplicantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5>����ȸ��</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtForwarder" ALT="����ȸ��" SIZE=10 MAXLENGTH=20 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnForwarder" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopForwarder">&nbsp;<INPUT TYPE=TEXT NAME="txtForwarderNm" SIZE=20 TAG="14"></TD>
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" ALT="�����׷�" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPostfiFlag" id="rdoPostfiFlagAll" value=" " tag = "11" checked>
											<label for="rdoPostfiFlagAll">��ü</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoPostfiFlag" id="rdoPostfiFlagYes" value="Y" tag = "11">
											<label for="rdoPostfiFlagYes">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostfiFlag" id="rdoPostfiFlagNo" value="N" tag = "11">
											<label for="rdoPostfiFlagNo">��Ȯ��</label></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/s5113ma8_vaSpread1_vspdData.js'></script>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH="*" Align=right><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage">B/L���</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtPostfiFlag" tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHForwarder" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHApplicant" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPostfiFlag" tag="24" TABINDEX="-1">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
