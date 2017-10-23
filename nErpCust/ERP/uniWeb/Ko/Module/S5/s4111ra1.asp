<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S4111RA1.ASP
'*  4. Program Name         : ��ǰó������ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/17
'*  8. Modified date(Last)  : 2002/12/17
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : SON BUM YEOL
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>��ǰó����������</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim gblnWinEvent
Dim arrReturn
Dim lgIsOpenPop

Dim arrParent
Dim PopupParent

arrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Const BIZ_PGM_ID        = "s4111rb1.asp"
Const C_MaxKey          = 8

'=========================================
Sub InitVariables()
         
    lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
    lgPageNo         = ""
	lgIntFlgMode = PopupParent.OPMD_CMODE	
     Redim arrReturn(0)        
     Self.Returnvalue = arrReturn     
     
End Sub

'=========================================
Sub SetDefaultVal()
		Dim arrTemp		
		arrTemp = Split(ArrParent(1), PopupParent.gColSep)			
		frm1.txtShipToParty.value = arrTemp(0)
		frm1.txtReceiver.focus
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "*", "NOCOOKIE", "PA") %>
End Sub

'=========================================
Sub InitSpreadSheet()
    '�׷��� ��� C_GROUP_DBAGENT 
	Call SetZAdoSpreadSheet("s4111ra1","S","A","V20021211", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()	
End Sub

'========================================
Function OKClick()

	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim arrReturn(C_MaxKey - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To C_MaxKey -1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)		
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function
	
'=========================================
Function CancelClick()
	Self.Close()
End Function

'========================================
Function OpenSortPopup()
	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
	End If
End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029			
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables			    
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'=========================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	  '��: ������ üũ	
    	If lgPageNo <> "" Then                   '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess Then Exit Sub
			Call DbQuery
    	End If
	End If
End Sub

'=========================================
Sub txtReceiver_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
	
'=========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
		
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
        Exit Function
	End If
		
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'=========================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=========================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=========================================
Function FncQuery() 
	
	FncQuery = False                                                        
    
    Err.Clear                                                               

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function

'=========================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
       If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							
		    strVal = strVal & "&txtShipToParty=" & Trim(frm1.txtHShipToParty.value)	
			strVal = strVal & "&txtReceiver=" & Trim(frm1.txtHReceiver.value)

	   Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							
		    strVal = strVal & "&txtShipToParty=" & Trim(frm1.txtShipToParty.value)				
			strVal = strVal & "&txtReceiver=" & Trim(frm1.txtReceiver.value)						    

       End if   
			strVal = strVal & "&lgPageNo="       & lgPageNo                '��: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
			
       Call RunMyBizASP(MyBizASP, strVal)										
			
	End With

	DbQuery = True

End Function

'=========================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    
	lgIntFlgMode = PopupParent.OPMD_UMODE
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtReceiver.focus	
    End if

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>��ǰó</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtShipToParty" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="��ǰó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="">&nbsp;<INPUT TYPE=TEXT NAME="txtShipToPartyNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>�μ��ڸ�</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiver" ALT="�μ��ڸ�" TYPE="Text" MAXLENGTH="50" SIZE= 20 tag="11XXXX"></TD>
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
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s4111ra1_vaSpread_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
	                    <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
						</TD>
						<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
						</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0  TABINDEX ="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHShipToParty" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHReceiver" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


