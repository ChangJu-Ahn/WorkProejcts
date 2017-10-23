<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : ����																		*
'*  2. Function Name        : 																*
'*  3. Program ID           : M3212RA1																	*
'*  4. Program Name         : L/C ���� ����																*
'*  5. Program Desc         : L/C Amend ��������� ���� L/C ���� ���� *
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/05/23																*
'*  9. Modifier (First)     : 																*
'* 10. Modifier (Last)      : Kim Jin-Ha
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2002/05/06 : ADO Conv.													*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>���ֳ�������</TITLE>
<!--
'******************************************  1.1 Inc ����   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 ���� Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		
Const BIZ_PGM_ID 		= "m3112rb2.asp"                              '��: Biz Logic ASP Name
Const C_MaxKey          = 20                                           '��: key count of SpreadSheet


<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgCookValue 

Dim IsOpenPop  
Dim gblnWinEvent
Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
				
    gblnWinEvent = False
    ReDim arrReturn(0,0)
    Self.Returnvalue = arrReturn
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
 Sub SetDefaultVal()
 	Dim arrParam
 	arrParam = arrParent(1)
		
 	WITH frm1
			
 		.hdnPoNo.value 		    = arrParam(0)
 		.txtPayMethCd.value 	= arrParam(1)
 		.txtPayMethNm.value 	= arrParam(2)
 		.txtIncotermsCd.value 	= arrParam(3)
 		.txtIncotermsNm.value 	= arrParam(4)
 		.txtCurrency.value 		= arrParam(5)
 		.txtBeneficiaryCd.value	= arrParam(6)
 		.txtBeneficiaryNm.value	= arrParam(7)
 		.txtGrpCd.value 		= arrParam(8)
 		.txtGrpNm.value 		= arrParam(9)
			
 	end with
 End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
'========================================================================================================
 Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 	<% Call loadInfTB19029A("I","*","NOCOOKIE","PA") %>                                '��: 
 	<% Call LoadBNumericFormatA("I","*","NOCOOKIE","PA") %>
 End Sub
'========================================================================================================
' Name : InitSpreadSheet
'========================================================================================================
Sub InitSpreadSheet()
 Call SetZAdoSpreadSheet("M3112RA2","S","A","V20030725",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
 							C_MaxKey, "X","X")
 Call SetSpreadLock
 frm1.vspdData.OperationMode = 5   
End Sub

'========================================================================================================
' Name : SetSpreadLock
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow,RtnValue
	Dim intColPos
	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount, frm1.vspdData.MaxCols)

		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt
			If frm1.vspdData.SelModeSelected Then
			For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					frm1.vspdData.Col = GetKeyPos("A", intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
			Next
			intInsRow = intInsRow + 1
			End IF								
		Next
	End if	
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'++++++++++++++++++++++++++++++++++++++++++++  OpenSoNo()  ++++++++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenSoNo()																					+
'+	Description : SoNo PopUp Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenItem()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If gblnWinEvent = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) then Exit Function
	
	gblnWinEvent = True

	arrParam(0) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "30"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ���							' Header��(1)
	   
	iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "B1B01PA2", "X")
		IsOpenPop = False
		Exit Function
	End If
			
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If	
	
End Function

'==============================================================================================================================
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGrpCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGrpCd.Value)
	
	arrParam(4) = "B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"
    arrHeader(2) = "��������"		
    arrHeader(3) = "����������"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'�������� ã�� �� ���� �����߻���.
	If arrRet(0) = "" Then
		frm1.txtGrpCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtGrpCd.Value= arrRet(0)		
		frm1.txtGrpNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtGrpCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'===============================  OpenTrackingNo()  ============================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = ""	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		frm1.txtTrackingNo.focus
		lgBlnFlgChgValue = True
		Set gActiveElement = document.activeElement
	End If	

End Function
'=========================================  3.1.1 Form_Load()  ==========================================
 Sub Form_Load()
 	Call LoadInfTB19029				                                           '��: Load table , B_numeric_format
 	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.LockField(Document, "N")  
 	Call InitVariables														    '��: Initializes local global variables
 	Call SetDefaultVal	
 	Call InitSpreadSheet()
 	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
 	Call FncQuery()
 End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
 Sub Form_QueryUnload(Cancel, UnloadMode)
 End Sub

'=========================================  3.3.1 vspdData_DblClick()  ==================================
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

'========================================  3.3.2 vspdData_KeyPress()  ===================================
 Function vspdData_KeyPress(KeyAscii)
      On Error Resume Next
      If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
         Call OKClick()
      ElseIf KeyAscii = 27 Then
         Call CancelClick()
      End If
 End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
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

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    
	FncQuery = False                                                 
	   
	Err.Clear                                                        

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData				
	Call InitVariables

	If DbQuery = False Then Exit Function							
	 
	FncQuery = True	
	Set gActiveElement = document.activeElement								
        
End Function

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
   Err.Clear															<%'��: Protect system from crashing%>

   DbQuery = False	

   If LayerShowHide(1) = False Then
   	Exit Function
   End If

   Dim strVal
	
   If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����				
	    strVal = strVal & "&txtItem=" & Trim(frm1.hdnItemCd.value)
		strVal = strVal & "&txtGrp=" & Trim(frm1.txtGrpCd.value)
		strVal = strVal & "&txtBeneficiary=" & Trim(frm1.txtBeneficiaryCd.value)
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		strVal = strVal & "&txtPayMeth=" & Trim(frm1.txtPayMethCd.value)
		strVal = strVal & "&txtIncoterms=" & Trim(frm1.txtIncotermsCd.value)
		strVal = strVal & "&txtPoNo=" & Trim(frm1.hdnPoNo.Value)
	Else
			
    	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtGrp=" & Trim(frm1.txtGrpCd.value)
		strVal = strVal & "&txtBeneficiary=" & Trim(frm1.txtBeneficiaryCd.value)
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		strVal = strVal & "&txtPayMeth=" & Trim(frm1.txtPayMethCd.value)
		strVal = strVal & "&txtIncoterms=" & Trim(frm1.txtIncotermsCd.value)
		strVal = strVal & "&txtPoNo=" & Trim(frm1.hdnPoNo.Value)
	End If
	  strVal = strVal & "&txtTrackingNo=" &Trim(frm1.txtTrackingNo.value)  
	strVal =     strVal & "&lgPageNo="       & lgPageNo                          '��: Next key tag
	strVal =	 strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	strVal =	 strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal =	 strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

	Call RunMyBizASP(MyBizASP, strVal)									<%'��: �����Ͻ� ASP �� ���� %>

	DbQuery = True					
		
End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														<%'��: ��ȸ ������ ������� %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtItemCd.focus
	End If

End Function

'========================================================================================================
' Function Name : OpenOrderBy
'========================================================================================================
Function OpenOrderByPopup()
	 Dim arrRet
		
	 On Error Resume Next
		
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
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
 <TR>
 	<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
 </TR>
 <TR>
 	<TD HEIGHT=20 WIDTH=100%>
 		<FIELDSET CLASS="CLSFLD">
 			<TABLE <%=LR_SPACE_TYPE_40%>>
 					<TR>
 						<TD CLASS="TD5" NOWRAP>ǰ��</TD>
 						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemCd" MAXLENGTH=18 SIZE=10 tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">
 										 	   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
 						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
 						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="���ű׷�" NAME="txtGrpCd" SIZE=10 tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
 										 	   <INPUT TYPE=TEXT ALT="���ű׷�" NAME="txtGrpNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
 					</TR>
 					<TR>
 						<TD CLASS="TD5" NOWRAP>������</TD>
 						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="������" NAME="txtBeneficiaryCd" MAXLENGTH=10 SIZE=10 tag="14NXXU" >&nbsp;&nbsp;&nbsp;&nbsp;
 										 	   <INPUT TYPE=TEXT ALT="������" NAME="txtBeneficiaryNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
 						<TD CLASS="TD5" NOWRAP>ȭ�����</TD>
 						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ȭ�����" NAME="txtCurrency" SIZE=10 tag="14NXXU" ></TD>
 					</TR>
 					<TR>
 						<TD CLASS="TD5" NOWRAP>�������</TD>
 						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="�������" NAME="txtPayMethCd" MAXLENGTH=5 SIZE=10 tag="14NXXU" >&nbsp;&nbsp;&nbsp;&nbsp;
 										 	   <INPUT TYPE=TEXT ALT="�������" NAME="txtPayMethNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
 						<TD CLASS="TD5" NOWRAP>��������</TD>
 						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="��������" NAME="txtIncotermsCd" SIZE=10 tag="14NXXU" >&nbsp;&nbsp;&nbsp;&nbsp;
 										 	   <INPUT TYPE=TEXT ALT="��������" NAME="txtIncotermsNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
 					</TR>
 					<TR>
						<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTrackingNo" ALT="Tracking��ȣ" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
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
 				<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
 				    <OBJECT classid=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"    id=A>
 				        <PARAM NAME="MaxCols" VALUE="0">
 				        <PARAM NAME="MaxRows" VALUE="0">
 				    </OBJECT>
 				</TD>
 			</TR>
 		</TABLE>
 	</TD>
 </TR>
 <tr>
 	<TD <%=HEIGHT_TYPE_01%>></TD>
 </tr>
 <TR HEIGHT="20">
 	<TD WIDTH="100%">
 		<TABLE <%=LR_SPACE_TYPE_30%>>
 			<TR>
 				<TD WIDTH=10>&nbsp;</TD>
 				<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
 									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderByPopup()"   ></IMG></TD>
 				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
 						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
 				<TD WIDTH=10>&nbsp;</TD>
 			</TR>
 		</TABLE>
 	</TD>
 </TR>
 <TR>
 	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
 	</TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHPONo" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" TAG="14">
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>	
</BODY>
</HTML>
