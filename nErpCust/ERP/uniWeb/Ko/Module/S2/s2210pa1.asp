<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2210PA1
'*  4. Program Name         : ǰ�� Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/12/27
'*  8. Modified date(Last)  : 2002/12/27
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "S2210PB1.ASP"                              '��: Biz Logic ASP Name

Const C_MaxKey          = 6                                            '��: key count of SpreadSheet

Const C_PopItemGroupCd	= 1
Const C_PopPlantCd		= 2

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
														    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
														    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim lgArrParent
Dim lgStrInitQuery
Dim lgDtFromDt
Dim lgDtToDt
Dim	lgBlnItemGroupCdChg
Dim lgBlnPlantCdChg
Dim	lgBlnFlgConChgValue

lgArrParent = window.dialogArguments
Set PopupParent = lgArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================================================================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
End Function

'=======================================================================================================
Sub SetDefaultVal()
	Dim iArrParam
	Dim iArrReturn
	Dim iIntMaxIndex
	Dim iStrPlantCd
	
	iArrParam = lgArrParent(1)
	' 0 : ǰ���ڵ�, 1 : ��ȿ��(����), 2 : ��ȿ��(����), 3 : ����, 4 : ǰ�����, 5 : ǰ��׷� 
	With frm1
		.txtConItemCd.value = Trim(iArrParam(0))			' ǰ���ڵ� 
		
		iIntMaxIndex = UBound(iArrParam)
		If iIntMaxIndex >= 1 Then
			lgDtFromDt = Trim(iArrParam(1))					' ǰ����ȿ�� 
			lgDtToDt = Trim(iArrParam(2))					' ǰ����ȿ�� 
		End If
		If iIntMaxIndex >= 3 Then
			.txtConPlantCd.value = Trim(iArrParam(3))		' ���� 
			' ������ ���ڷ� ���� ��쿡�� ������ ������ �� ����.
			If Trim(iArrParam(3)) <> "" Then
				Call ggoOper.SetReqAttr(.txtConPlantCd ,"Q")
				.btnConPlantCd.disabled = True
			End If
		End If
		
		If iIntMaxIndex >= 4 Then
			.cboConItemAcct.value = Trim(iArrParam(4))		' ǰ����� 
			.txtConItemGroupCd.value = Trim(iArrParam(5))	' ǰ��׷� 
		End If
		
		If PopupParent.gPlant <> "" And Trim(.txtConPlantCd.value) = "" Then
			.txtConPlantCd.value = PopupParent.gPlant
		End If

		iStrPlantCd = .txtConPlantCd.value
		If iStrPlantCd <> "" Then
			iStrPlantCd = " " & FilterVar(iStrPlantCd, "''", "S") & ""
			Call GetCodeName(iStrPlantCd, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlantCd)
		End If
			
		lgStrInitQuery = Trim(iArrParam(0))
		
		.vspdData.OperationMode = 3
	End With

	lgBlnItemGroupCdChg = False
	lgBlnPlantCdChg = False
	lgBlnFlgConChgValue = False

	Redim iArrReturn(0)
	Self.Returnvalue = iArrReturn
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "PA") %>	
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S2210PA1","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")		
	Call SetSpreadLock 	
	    
End Sub

'========================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.SpreadLockWithOddEvenRowColor()
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    .vspdData.ReDraw = True

    End With
End Sub	

'========================================================================================================
Function OKClick()
	Dim iArrReturn
	With frm1
		If .vspdData.ActiveRow > 0 Then	
			Redim iArrReturn(4)
			.vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = GetKeyPos("A",1)		' ��ȹ�Ⱓ 
			iArrReturn(0) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",2)		' ��ȹ�Ⱓ���� 
			iArrReturn(1) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",3)		' ������ 
			iArrReturn(2) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",4)		' ������ 
			iArrReturn(3) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",5)		' �� 
			iArrReturn(4) = .vspdData.Text
			
			Self.Returnvalue = iArrReturn
		End If
	End With
	Err.Clear
	
	Self.Close()
End Function

'========================================================================================================
	Function CancelClick()
		Self.Close()
	End Function

'==========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		'ǰ����� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM", " B_MINOR ", " MAJOR_CD=" & FilterVar("P1001", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConItemAcct, lgF0,lgF1, PopUpParent.gColSep)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'======================================================================================================== 
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere

	Case C_PopItemGroupCd
		iArrParam(1) = "dbo.B_ITEM_GROUP "					<%' TABLE ��Ī %>
		iArrParam(2) = Trim(frm1.txtConItemGroupCd.value)	<%' Code Condition%>
		iArrParam(3) = ""									<%' Name Cindition%>
		iArrParam(4) = "LEAF_FLG = " & FilterVar("Y", "''", "S") & "  AND DEL_FLG = " & FilterVar("N", "''", "S") & " "	<%' Where Condition%>
		iArrParam(5) = frm1.txtConItemGroupCd.alt '"ǰ��׷�"		<%' TextBox ��Ī %>
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "ITEM_GROUP_CD"	<%' Field��(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "ITEM_GROUP_NM"	<%' Field��(1)%>
		    
		iArrHeader(0) = "ǰ��׷�"						<%' Header��(0)%>
		iArrHeader(1) = "ǰ��׷��"					<%' Header��(1)%>

		frm1.txtConItemGroupCd.focus

	Case C_PopPlantCd
		iArrParam(1) = "B_PLANT"							<%' TABLE ��Ī %>
		iArrParam(2) = Trim(frm1.txtConPlantCd.value)		<%' Code Condition%>
		iArrParam(3) = ""									<%' Name Cindition%>
		iArrParam(4) = ""									<%' Where Condition%>
		iArrParam(5) = "����"							<%' TextBox ��Ī %>
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "PLANT_CD"	<%' Field��(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "PLANT_NM"	<%' Field��(1)%>
    
	    iArrHeader(0) = "����"							<%' Header��(0)%>
	    iArrHeader(1) = "�����"						<%' Header��(1)%>

		frm1.txtConPlantCd.focus 

	End Select
 
	iArrParam(0) = iArrParam(5)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=======================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopItemGroupCd
		frm1.txtConItemGroupCd.value = pvArrRet(0) 
		frm1.txtConItemGroupNm.value = pvArrRet(1)
		lgBlnItemGroupCdChg = False

	Case C_PopPlantCd
		frm1.txtConPlantCd.value = pvArrRet(0) 
		frm1.txtConPlantNm.value = pvArrRet(1)   
		lgBlnPlantCdChg = False
	End Select

	SetConPopup = True

End Function

'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029											  '��: Load table , B_numeric_format
   
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
    
	Call InitVariables()										  '��: Initializes local global variables
	Call InitSpreadSheet()
	Call InitComboBox()
	Call SetDefaultVal()	
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	If lgStrInitQuery <> "" Then DbQuery()
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

<%'==========================================================================================
'   Event Desc : ǰ��׷� 
'==========================================================================================%>
Function txtConItemGroupCd_OnChange1()
	Dim iStrCode
	
	iStrCode = Trim(frm1.txtConItemGroupCd.value)
	If iStrCode <> "" Then
		iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
		If Not GetCodeName(iStrCode, "" & FilterVar("Y", "''", "S") & " ", "" & FilterVar("N", "''", "S") & " ", "default", "default", "" & FilterVar("IG", "''", "S") & "", C_PopItemGroupCd) Then
			txtConItemGroupCd_OnChange1 = False
			frm1.txtConItemGroupCd.value = ""
			frm1.txtConItemGroupNm.value = ""
		End If
	Else
		frm1.txtConItemGroupNm.value = ""
	End If
End Function

<% '========================================================================================================
'   Event Desc : tag�� '1'�� �ʵ忡(��ȸ����) ���� Data Change ���θ� �����Ѵ�.																						=
'======================================================================================================== %>
' ǰ��׷� 
Function txtConItemGroupCd_OnKeyDown()
	lgBlnItemGroupCdChg = True
	lgBlnFlgConChgValue = True
End Function

' ���� 
Function txtConPlantCd_OnKeyDown()
	lgBlnPlantCdChg = True
	lgBlnFlgConChgValue = True
End Function

<% '====================================================================================================
'	Description : ��ȸ������ ��ȿ���� Check�Ѵ�.
'   ���ǻ��� : ȭ���� tab order ���� ����Ѵ�. 
'==================================================================================================== %>
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True
	If lgBlnItemGroupCdChg Then
		iStrCode = Trim(frm1.txtConItemGroupCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("Y", "''", "S") & " ", "" & FilterVar("N", "''", "S") & " ", "default", "default", "" & FilterVar("IG", "''", "S") & "", C_PopItemGroupCd) Then
				Call DisplayMsgBox("970000", "X", frm1.txtConItemGroupCd.alt, "X")
				frm1.txtConItemGroupCd.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtConItemGroupNm.value = ""
		End If

		lgBlnItemGroupCdChg	= False
	End If

	If lgBlnPlantCdChg Then
		iStrCode = Trim(frm1.txtConPlantCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlantCd) Then
				Call DisplayMsgBox("970000", "X", frm1.txtConPlantCd.alt, "X")
				frm1.txtConPlantCd.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtConPlantNm.value = ""
		End If

		lgBlnPlantCdChg	= False
	End If

End Function

<%'=============================================================================================
'	Description : �ڵ尪�� �ش��ϴ� ���� Display�Ѵ�.
'==================================================================================================== %>
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
		'Item Change�� ���� Fetch�ϴ� ������ ǥ�� ����� Enable ��Ų��.
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
     End If
	
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
    Function vspdData_KeyPress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
			Call OKClick()
		ElseIf KeyAscii = 27 Then
			Call CancelClick()
		End If
    End Function

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If

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
	End If
End Sub

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'=======================================================================================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=======================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field

	' ��ȸ���� ��ȿ�� check
	If 	lgBlnFlgConChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If

    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
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
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'��: �����Ͻ� ó�� ASP�� ���� %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll�� 
			strVal = strVal & "&txtItemCd=" & Trim(.txtHItemCd.value)
			strVal = strVal & "&txtItemNm=" & Trim(.txtHItemNm.value)
			strVal = strVal & "&txtItemGroupCd=" & Trim(.txtHItemGroupCd.value)
			strVal = strVal & "&txtItemAcct=" & Trim(.txtHItemAcct.value)
			strVal = strVal & "&txtItemSpec=" & Trim(.txtHItemSpec.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.txtHPlantCd.value)
		Else
			' ó�� ��ȸ�� 
			strVal = strVal & "&txtItemCd=" & Trim(.txtConItemCd.value)
			strVal = strVal & "&txtItemNm=" & Trim(.txtConItemNm.value)
			strVal = strVal & "&txtItemGroupCd=" & Trim(.txtConItemGroupCd.value)
			strVal = strVal & "&txtItemAcct=" & Trim(.cboConItemAcct.value)
			strVal = strVal & "&txtItemSpec=" & Trim(.txtConItemSpec.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.txtConPlantCd.value)
		End If
		strVal = strVal & "&txtFromDt=" & lgDtFromDt
		strVal = strVal & "&txtToDt=" & lgDtToDt

        strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
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
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
	Else
		frm1.txtConSpPeriod.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>ǰ��</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConItemCd" ALT="ǰ��" SIZE=20 MAXLENGTH=18 TAG="11XXXU"></TD>
						<TD CLASS=TD5>ǰ��׷�</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConItemGroupCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"  ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConItemGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopup(C_PopItemGroupCd)">&nbsp;<INPUT TYPE=TEXT NAME="txtConItemGroupNm" SIZE=25 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5></TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConItemNm" ALT="ǰ���" SIZE=40 MAXLENGTH=40 TAG="11XXXU"></TD>
						<TD CLASS="TD5">ǰ�����</TD>
						<TD CLASS="TD6"><SELECT NAME="cboConItemAcct" tag="11XXXU" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>									
					</TR>
					<TR>
						<TD CLASS=TD5>�԰�</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConItemSpec" ALT="�԰�" SIZE=40 MAXLENGTH=40 TAG="11XXXU"></TD>
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6"><INPUT NAME="txtConPlantCd" ALT="����" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPlantCd)">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s2210pa1_OBJECT1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											  <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG>			</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemNm" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemAcct" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemSpec" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlantCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
