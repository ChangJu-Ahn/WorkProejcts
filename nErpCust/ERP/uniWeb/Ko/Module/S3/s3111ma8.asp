<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : s3111ma8.asp	
'*  4. Program Name         : ������Ȳ��ȸ 
'*  5. Program Desc         : ������Ȳ��ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
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

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              '��: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                           
Dim lgMark                                                
Dim IscookieSplit 
Dim lsSoNo													'���ֹ�ȣ 
Dim lsCur

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3111mb8.asp"
Const BIZ_PGM_JUMP_ID	= "s3111ma1"
Const C_MaxKey          = 20                 
'--------------- ������ coding part(��������,End)-------------------------------------------------------------

'=============================================================================================================
Sub InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False                               'Indicates that no value changed    
    lgSortKey        = 1
    lgIntFlgMode = parent.OPMD_CMODE					   'Indicates that current mode is Create mode
End Sub

'=============================================================================================================
Sub SetDefaultVal()
	frm1.txtSOFrDt.text = StartDate
	frm1.txtSOToDt.text = EndDate
	frm1.txtRadio.value = frm1.rdoQueryFlg1.value 
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "MA") %>
End Sub


'=============================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3111QA1","S","A","V20030318", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock  
End Sub

'=============================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'=============================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If	
	Call OpenOrderByPopup("A")
End Sub

'=============================================================================================================
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub

'=============================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
		arrParam(5) = "�ֹ�ó"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "�ֹ�ó"							
		arrHeader(1) = "�ֹ�ó��"						

	Case 1
		arrParam(1) = "S_SO_TYPE_CONFIG"					
		arrParam(2) = Trim(frm1.txtSOType_cd.Value)			
		arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "					
		arrParam(5) = "��������"						
	
		arrField(0) = "SO_TYPE"						     	
		arrField(1) = "SO_TYPE_NM"							
    
		arrHeader(0) = "��������"						
		arrHeader(1) = "�������¸�"						

	Case 2

		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtdeal_type_cd.value)		
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & ""				
		arrParam(5) = "�Ǹ�����"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "�Ǹ�����"					
	    arrHeader(1) = "�Ǹ�������"					
    
	Case 3
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGroup_cd.Value)		
		arrParam(4) = ""									
		arrParam(5) = "�����׷�"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						
    
		arrHeader(0) = "�����׷�"						
		arrHeader(1) = "�����׷��"						

	Case 4
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtPayterms_cd.value)		
		
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""				
		arrParam(5) = "�������"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "�������"					
	    arrHeader(1) = "���������"					
		
	End Select

	arrParam(0) = arrParam(5)								
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	With frm1
		Select Case iWhere
		    Case 0
		    	.txtconBp_cd.focus  
		    Case 1
		    	.txtSOType_cd.focus   
		    Case 2
		    	.txtdeal_type_cd.focus
		    Case 3
		    	.txtSalesGroup_cd.focus
		    Case 4
		    	.txtPayTerms_cd.focus
		    End Select
	End With

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'=============================================================================================================
Function OpenSORef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If lsSoNo = "" Then
		Call DisplayMsgBox("203151", "X", "X", "X")		
		Exit Function
	End IF

	lgIsOpenPop = True

	arrParam(0) = lsSoNo								
	arrParam(1) = lsCur				
	  
	iCalledAspName = AskPRAspName("s3112ra7")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra7", "x")
		lgIsOpenPop = False
		exit Function
	end if

	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	lsSoNo = ""
	lsCur = ""

End Function

'=============================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_nm.value = arrRet(1) 
		Case 1
			.txtSOType_cd.value = arrRet(0) 
			.txtSOType_nm.value = arrRet(1)
		Case 2
			.txtdeal_type_cd.value = arrRet(0)
			.txtdeal_type_nm.value = arrRet(1)  
		Case 3
			.txtSalesGroup_cd.value = arrRet(0) 
			.txtSalesGroup_nm.value = arrRet(1)   
		Case 4
			.txtPayTerms_cd.value = arrRet(0) 
			.txtPayTerms_nm.value = arrRet(1)   
		End Select
	End With
End Function

'=============================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						

	If Kubun = 1 Then								

		If frm1.vspdData.ActiveRow > 0 Then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",7)
			WriteCookie CookieSplit , frm1.vspdData.Text
		Else
			WriteCookie CookieSplit , ""
		End If
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
		
		Dim iniSep

	'--------------- ������ coding part(�������,Start)---------------------------------------------------%>
		frm1.txtconBp_cd.value =  arrVal(0)
		frm1.txtconBp_nm.value =  arrVal(1)
		frm1.txtSOType_cd.value =  arrVal(2)
		frm1.txtSOType_nm.value = arrVal(3) 
		frm1.txtSalesGroup_cd.value =  arrVal(4)
		frm1.txtSalesGroup_nm.value = arrVal(5) 
		frm1.txtPayTerms_cd.value =  arrVal(6)
		frm1.txtPayTerms_nm.value = arrVal(7) 
		frm1.txtdeal_type_cd.value =  arrVal(8)
		frm1.txtdeal_type_nm.value = arrVal(9)
	'--------------- ������ coding part(�������,End)---------------------------------------------------%>

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function


'=============================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'��: ��ư ���� ���� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------   
	Call CookiePage(0)	
	frm1.txtSOType_cd.focus 		
	'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'=============================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=============================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
    If Row <= 0 Then
       
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If       
    
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	IscookieSplit = ""			
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",7)										'���ֹ�ȣ Col��ȣ�� �����Ѵ�.
	lsSoNo=frm1.vspdData.Text	
	
	frm1.vspdData.Col = GetKeyPos("A",12)										'���ֹ�ȣ Col��ȣ�� �����Ѵ�.
	lsCur=frm1.vspdData.Text		
	'--------------- ������ coding part(�������,End)------------------------------------------------------

    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    
End Sub

'=============================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	        	
    	If lgPageNo <> "" Then				    		
			Call DisableToolBar(parent.TBC_QUERY)
			Call DBQuery
    	End If
    End If    
End Sub

'=============================================================================================================
Sub rdoQueryFlg1_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg1.value
End Sub

Sub rdoQueryFlg2_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg2.value
End Sub

Sub rdoQueryFlg3_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg3.value
End Sub

'=============================================================================================================
Sub txtSOFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSOFrDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtSOFrDt.Focus
	End If
End Sub

Sub txtSOToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSOToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtSOToDt.Focus
	End If
End Sub

'=============================================================================================================
Sub txtSOFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtSOToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=============================================================================================================
Function FncQuery() 
    Dim IntRetCD
    FncQuery = False                                                        '��: Processing is NG    
    Err.Clear                                                               '��: Protect system from crashing
	
	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
	With frm1
		If .rdoQueryFlg1.checked = True Then
			.txtRadio.value = .rdoQueryFlg1.value
		ElseIf .rdoQueryFlg2.checked = True Then
			.txtRadio.value = .rdoQueryFlg2.value
		ElseIf .rdoQueryFlg3.checked = True Then
			.txtRadio.value = .rdoQueryFlg3.value
		End If		
	End With

    Call DbQuery															'��: Query db data

    FncQuery = True
    	
End Function

'=============================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=============================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=============================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                    
End Function


'=============================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'=============================================================================================================
Function DbQuery() 
	
	Dim strVal

    DbQuery = False
  
	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function
    
    Err.Clear                                                             
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
	If lgIntFlgMode = parent.OPMD_UMODE Then
	'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		strVal = strVal & "&txtconBp_cd=" & Trim(.HconBp_cd.value)
		strVal = strVal & "&txtSOType_cd=" & Trim(.HSOType_cd.value)
		strVal = strVal & "&txtSalesGroup_cd=" & Trim(.HSalesGroup_cd.value)
		strVal = strVal & "&txtPayTerms_cd=" & Trim(.HPayTerms_cd.value)
		strVal = strVal & "&txtdeal_type_cd=" & Trim(.Hdeal_type_cd.value)
		strVal = strVal & "&txtSOFrDt=" & Trim(.HSOFrDt.value)
		strVal = strVal & "&txtSOToDt=" & Trim(.HSOToDt.value)
		strVal = strVal & "&txtRadio=" & Trim(.HRadio.value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
		strVal = strVal & "&txtSOType_cd=" & Trim(.txtSOType_cd.value)
		strVal = strVal & "&txtSalesGroup_cd=" & Trim(.txtSalesGroup_cd.value)
		strVal = strVal & "&txtPayTerms_cd=" & Trim(.txtPayTerms_cd.value)
		strVal = strVal & "&txtdeal_type_cd=" & Trim(.txtdeal_type_cd.value)
		strVal = strVal & "&txtSOFrDt=" & Trim(.txtSOFrDt.text)
		strVal = strVal & "&txtSOToDt=" & Trim(.txtSOToDt.text)
		strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)		
	End If	
	'--------------- ������ coding part(�������,End)------------------------------------------------
        strVal = strVal & "&lgPageNo=" & lgPageNo	        
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        
        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'=============================================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	lgIntFlgMode = parent.OPMD_UMODE
	Call SetToolbar("11000000000111")  
    frm1.vspdData.Focus 
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������Ȳ��ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenSORef">���ֳ�����Ȳ</A></TD>
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
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSOType_cd" ALT="��������" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT NAME="txtSOType_nm" TYPE="Text" MAXLENGTH=20 SIZE=25 tag=14></TD>
									<TD CLASS="TD5" NOWRAP>�����׷�</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSalesGroup_cd" ALT="�����׷�" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT NAME="txtSalesGroup_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag=14></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>�ֹ�ó</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="�ֹ�ó" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH=20 SIZE=25 tag=14></TD>
									<TD CLASS="TD5" NOWRAP>�������</TD>
									<TD CLASS="TD6"><INPUT NAME="txtPayTerms_cd" ALT="�������" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 4">&nbsp;<INPUT NAME="txtPayTerms_nm" TYPE="Text" MAXLENGTH=20 SIZE=25 tag=14></TD>
								</TR>		
								<TR>	
									<TD CLASS=TD5>�Ǹ�����</TD>
									<TD CLASS=TD6><INPUT  NAME="txtdeal_type_cd" ALT="�Ǹ�����" TYPE="TEXT" MAXLENGTH=4 SIZE=10 TAG="11XXXU"  ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSORef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 2">&nbsp;<INPUT  NAME="txtdeal_type_nm" TYPE="TEXT" MAXLENGTH=20 SIZE=25 TAG=14></TD>									
								    <TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>	
								<TR>										
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s3111ma8_fpDateTime1_txtSOFrDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s3111ma8_fpDateTime2_txtSOToDt.js'></script>
									</TD> 
									<TD CLASS=TD5>Ȯ������</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="A" CHECKED ID="rdoQueryFlg1"><LABEL FOR="rdoQueryFlg1">��ü</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="Y" ID="rdoQueryFlg2"><LABEL FOR="rdoQueryFlg2">Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="N" ID="rdoQueryFlg3"><LABEL FOR="rdoQueryFlg3">��Ȯ��</LABEL>
									</TD>
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
										<script language =javascript src='./js/s3111ma8_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">���ֵ��</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%>
		                    FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14">

<INPUT TYPE=HIDDEN NAME="HconBp_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSOType_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGroup_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HPayTerms_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="Hdeal_type_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSOFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HSOToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadio" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
