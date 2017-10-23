<%@ LANGUAGE="VBSCRIPT" %>
<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211ra1.asp		(m3221ma1.asp)														*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C Amend ���(m3221ma1.asp)�� ���� L/C Reference ASP															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/05/11																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : KIM JIN HA																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'******************************************************************************************************
%>
-->
<HTML>
<HEAD>
<TITLE>L/C����</TITLE>
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'########################################################################################################
'#									1.  Data Declaration Part
'########################################################################################################

Const BIZ_PGM_ID 		= "M3211rb1_KO441.asp"                              '��: Biz Logic ASP Name

Const C_MaxKey          = 1                                           '��: key count of SpreadSheet
Const C_LC_NO			= 1

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD                                           '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    

Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         
Dim IscookieSplit 

Dim IsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
   													    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
   													    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim arrReturn												'��: Return Parameter Group
Dim arrParam
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


Const gstrPayTermsMajor 	= "B9004"
Const gstrIncotermsMajor 	= "B9006"

'========================================== 2.1.1 InitVariables()  ======================================
Function InitVariables()
		
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
		
	gblnWinEvent = False
        
    arrReturn = ""
    Self.Returnvalue = arrReturn     

End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
		
	frm1.txtFrDt.Text = StartDate 
	frm1.txtToDt.Text = EndDate 
		
	frm1.vspdData.OperationMode = 5
	frm1.txtBeneficiary.focus	  

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGrp, "Q") 
		frm1.txtGrp.Tag = left(frm1.txtGrp.Tag,1) & "4" & mid(frm1.txtGrp.Tag,3,len(frm1.txtGrp.Tag))
        frm1.txtGrp.value = lgPGCd
	End If
	Set gActiveElement = document.activeElement		
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "RA") %>                                '��: 

End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3211RA101","S","A","V20030412",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock 	    
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	

'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col =  GetKeyPos("A",C_LC_NO)
			
	arrReturn = frm1.vspdData.Text
		
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	arrReturn = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
' Function Name : OpenConSItemDC
'========================================================================================================
Function OpenConSItemDC(Byval iWhere)

Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

If gblnWinEvent = True Then Exit Function
gblnWinEvent = True
	
Select Case iWhere
		Case 0
		arrParam(1) = "B_BIZ_PARTNER"							' TABLE ��Ī 
		arrParam(2) = Trim(frm1.txtBeneficiary.value)				' Code Condition
		arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "' Where Condition
		arrParam(5) = "������"								' TextBox ��Ī 
	
		arrField(0) = "BP_CD"									' Field��(0)
		arrField(1) = "BP_NM"									' Field��(1)
    
		arrHeader(0) = "������"								' Header��(0)
		arrHeader(1) = "�����ڸ�"							' Header��(1)
    
	Case 1
		arrParam(1) = "B_Pur_Grp"
		arrParam(2) = Trim(frm1.txtGrp.Value)
		arrParam(4) = ""
		arrParam(5) = "���ű׷�"			
	
		arrField(0) = "PUR_GRP"	
		arrField(1) = "PUR_GRP_NM"	
    
		arrHeader(0) = "���ű׷�"		
		arrHeader(1) = "���ű׷��"
    
	Case 2
			
		arrParam(1) = "b_minor,b_configuration"						' TABLE ��Ī 
		arrParam(2) = Trim(frm1.txtPayTerms.Value)						' Code Condition
		arrParam(4) = "b_minor.Major_Cd= " & FilterVar(gstrPayTermsMajor, "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd AND b_configuration.REFERENCE = " & FilterVar("M", "''", "S") & " "
		arrParam(5) = "�������"								' TextBox ��Ī 

		arrField(0) = "b_minor.Minor_CD"							' Field��(0)
		arrField(1) = "b_minor.Minor_NM"							' Field��(1)

		arrHeader(0) = "�������"								' Header��(0)
		arrHeader(1) = "���������"								' Header��(1)
					
	 Case 3			
		arrParam(1) = "B_Minor"										' TABLE ��Ī 
		arrParam(2) = Trim(frm1.txtIncoterms.Value)						' Code Condition
		arrParam(3) = ""											' Name Cindition
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9006", "''", "S") & ""							' Where Condition
		arrParam(5) = "��������"								' TextBox ��Ī 

		arrField(0) = "Minor_CD"									' Field��(0)
		arrField(1) = "Minor_NM"									' Field��(1)

		arrHeader(0) = "��������"								' Header��(0)
		arrHeader(1) = "�������Ǹ�"								' Header��(1)
End Select
	
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	arrParam(0) = arrParam(5)												' �˾� ��Ī	

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'-------------------------------------------------------------------------------------------------------
'	Name : SetConSItemDC()
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
With frm1
	Select Case iWhere
		Case 0
			.txtBeneficiary.value = arrRet(0)
			.txtBeneficiaryNm.value = arrRet(1)	  
			.txtBeneficiary.focus 
		Case 1
			.txtGrp.Value = arrRet(0)
			.txtGrpNm.Value = arrRet(1)
			.txtGrp.focus
		Case 2
			.txtPayTerms.Value = arrRet(0)
			.txtPayTermsNm.Value = arrRet(1)
			.txtPayTerms.focus		 
		Case 3
			.txtIncoterms.Value = arrRet(0)
			.txtIncotermsNm.Value = arrRet(1) 
			.txtIncoterms.focus	 
	End Select
	Set gActiveElement = document.activeElement
End With
End Function


'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
	Call LoadInfTB19029													'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)		
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
	    
	Call InitVariables											  '��: Initializes local global variables
    Call GetValue_ko441()
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
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
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

'========================================================================================================
'   Event Name : OCX_DbClick()
'========================================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7	
		Call SetFocusToDocument("P")
		frm1.txtFrDt.focus	
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToDt.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'=======================================================================================================
Sub txtFrDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
	   Call CancelClick()
	Elseif KeyAscii = 13 Then
	   Call FncQuery()
	End if
End Sub

Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
	   Call CancelClick()
	Elseif KeyAscii = 13 Then
	   Call FncQuery()
	End if
End Sub

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
Function FncQuery() 
	    
	FncQuery = False                                                        '��: Processing is NG
	    
	Err.Clear                                                               '��: Protect system from crashing

	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToDt.text,gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","������", "X")			
			.txtToDt.Focus
			Exit Function
		End if   
	End with
		
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'��: Initializes local global variables
	    
	If DbQuery = False Then Exit Function									

	FncQuery = True		
    Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Function Name : DbQuery
'========================================================================================================

Function DbQuery() 

	Err.Clear														'��: Protect system from crashing
	DbQuery = False													'��: Processing is NG
		
	If LayerShowHide(1) = False Then
		Exit Function
	End If
		
	Dim strVal
	    
	With frm1
			
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
			strVal = strVal & "&txtBeneficiary=" & Trim(.hdnBeneficiary.value)		<%'��: ��ȸ ���� ����Ÿ %>
			strVal = strVal & "&txtGrp=" & Trim(.hdnGrp.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.hdnPayTerms.value)
			strVal = strVal & "&txtIncoterms=" & Trim(.hdnIncoterms.value)
			strVal = strVal & "&txtFrDt=" & Trim(.hdnFrDt.Value)
			strVal = strVal & "&txtToDt=" & Trim(.hdnToDt.Value)
		else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'��: �����Ͻ� ó�� ASP�� ���� %>
			strVal = strVal & "&txtBeneficiary=" & Trim(.txtBeneficiary.value)		<%'��: ��ȸ ���� ����Ÿ %>
			strVal = strVal & "&txtGrp=" & Trim(.txtGrp.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtPayTerms.value)
			strVal = strVal & "&txtIncoterms=" & Trim(.txtIncoTerms.value)
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.Text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		End if
			
	 	strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	 	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

	    Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
	        
	End With
	    
	DbQuery = True    

End Function

'=========================================================================================================
' Function Name : DbQueryOk
'=========================================================================================================
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.vspdData.focus
	End If
	Set gActiveElement = document.activeElement
End Function
'========================================================================================================
' Function Name : OpenOrderByPopup
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
</HEAD>
<%
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################
%>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
   <TR>
   	<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
   </TR>
   <TR>
   	<TD HEIGHT=20 WIDTH=100%>
   		<FIELDSET CLASS="CLSFLD">
   			<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
   				<TR>
   					<TD CLASS=TD5 NOWRAP>������</TD>
   					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 0">
   										 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
   					<TD CLASS=TD5 NOWRAP>���ű׷�</TD>
   					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT Alt="���ű׷�" NAME="txtGrp" SIZE=10 MAXLENGTH=4  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 1">
   										 <INPUT TYPE=TEXT Alt="���ű׷�" NAME="txtGrpNm" NAME="arrCond" tag="14X"></TD>
   				</TR>
   				<TR>
   					<TD CLASS=TD5 NOWRAP>�������</TD>
   					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="11XXXU" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 2">
   										 <INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
   					<TD CLASS=TD5 NOWRAP>��������</TD>
   					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10  MAXLENGTH=5 TAG="1XNXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoterms" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 3">
   										 <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="14"></TD>
   				</TR>
   				<TR>
   					<TD CLASS="TD5" NOWRAP>������</TD>
   					<TD CLASS="TD6" NOWRAP>
   						<table cellspacing=0 cellpadding=0>
   							<tr>
   								<td>
   									<script language =javascript src='./js/m3211ra1_fpDateTime1_txtFrDt.js'></script>
   								</td>
   								<td>~</td>
   								<td>
   									<script language =javascript src='./js/m3211ra1_fpDateTime2_txtToDt.js'></script>
   								</td>
   							<tr>
   						</table>
   					</TD>
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
   	<TD WIDTH=100% HEIGHT=* valign=top>
   		<TABLE <%=LR_SPACE_TYPE_20%>>
   			<TR>
   				<TD HEIGHT="100%">
   					<script language =javascript src='./js/m3211ra1_vaSpread1_vspdData.js'></script>
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
   				<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
   									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderByPopup()"   ></IMG></TD>
   				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
   						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
   				<TD WIDTH=10>&nbsp;</TD>
   			</TR>
   		</TABLE>
   	</TD>
   </TR>
   <TR>
   	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
   	</TD>
   </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnBeneficiary" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGrp" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayTerms" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIncoterms" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">

<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
   <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
