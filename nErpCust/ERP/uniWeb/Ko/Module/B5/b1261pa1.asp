<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : B1261PA1
'*  4. Program Name         : 거래처팝업 
'*  5. Program Desc         : 거래처정보의 거래처팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : Choinkuk		
'* 10. Modifier (Last)      : Choinkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID 		= "b1261pb1.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 12                                           '☆: key count of SpreadSheet

'========================================================================================================
                   
Dim IscookieSplit 

Dim IsOpenPop  
Dim lgIsOpenPop
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
	Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
                
        gblnWinEvent = False
        Redim arrReturn(0)        
        Self.Returnvalue = arrReturn     
	End Function

'=======================================================================================================
	Sub SetDefaultVal()	
			
		frm1.txtBp_cd.value	 = arrParent(1)				
		frm1.txtRadio2.value = frm1.rdoQueryFlg2_1.value							'거래처구분 
		frm1.txtRadio3.value = frm1.rdoQueryFlg3_2.value							'사용여부	
						
		frm1.txtBp_cd.focus	  
	End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()	
	Call SetZAdoSpreadSheet("B1261PA1","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	Call SetSpreadLock 	 	    			            
End Sub

'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.OperationMode = 3
End Sub	

'========================================================================================================
	Function OKClick()

		Dim intColCnt
		
		If frm1.vspdData.ActiveRow > 0 Then	
		
			Redim arrReturn(frm1.vspdData.MaxCols - 1)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
			For intColCnt = 1 To frm1.vspdData.MaxCols - 1
				frm1.vspdData.Col = GetkeyPos("A", intColCnt)
				arrReturn(intColCnt - 1) = frm1.vspdData.Text
			Next	
					
		End If
		
		Self.Returnvalue = arrReturn
		Self.Close()
	
	End Function

'========================================================================================================
	Function CancelClick()
		Redim arrReturn(0)
		arrReturn(0) = ""
		Self.Returnvalue = arrReturn
		Self.Close()
	End Function

'========================================================================================================
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

'========================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	Select Case iWhere
	Case 0
		arrParam(0) = "영업그룹"
		arrParam(1) = "B_SALES_GRP"											' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtBiz_grp.value)							' Code Condition
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "										' Where Condition
		arrParam(5) = "영업그룹"										' TextBox 명칭 
			
		arrField(0) = "SALES_GRP"											' Field명(0)
		arrField(1) = "SALES_GRP_NM"										' Field명(1)
    
		arrHeader(0) = "영업그룹"										' Header명(0)
		arrHeader(1) = "영업그룹명"										' Header명(1)

		frm1.txtBiz_grp.focus 
				
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case 1
		arrParam(0) = "구매그룹"
		arrParam(1) = "B_PUR_GRP"											' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtPur_grp.value)							' Code Condition
		arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "										' Where Condition
		arrParam(5) = "구매그룹"										' TextBox 명칭 
		
	    arrField(0) = "PUR_GRP"												' Field명(0)
	    arrField(1) = "PUR_GRP_NM"											' Field명(1)
	    
	    arrHeader(0) = "구매그룹"										' Header명(0)
	    arrHeader(1) = "구매그룹명"										' Header명(1)
	    
	    frm1.txtPur_grp.focus 
	    
	    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	End Select

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'========================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtBiz_grp.value = arrRet(0) 
			.txtSales_grp_nm.value = arrRet(1)	   
		Case 1
			.txtPur_grp.value = arrRet(0) 
			.txtPur_grp_nm.value = arrRet(1)	 
		End Select
	End With
End Function


'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
        
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
     
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery()
	
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

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
         If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

'========================================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
		If OldLeft <> NewLeft Then    Exit Sub

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
			If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub

'========================================================================================================

	Sub rdoQueryFlg2_1_OnClick()
		frm1.txtRadio2.value = frm1.rdoQueryFlg2_1.value
	End Sub
	
	Sub rdoQueryFlg2_2_OnClick()
		frm1.txtRadio2.value = frm1.rdoQueryFlg2_2.value
	End Sub
	
	Sub rdoQueryFlg2_3_OnClick()
		frm1.txtRadio2.value = frm1.rdoQueryFlg2_3.value
	End Sub
	
	Sub rdoQueryFlg3_1_OnClick()
		frm1.txtRadio3.value = frm1.rdoQueryFlg3_1.value
	End Sub
	
	Sub rdoQueryFlg3_2_OnClick()
		frm1.txtRadio3.value = frm1.rdoQueryFlg3_2.value
	End Sub
	
	Sub rdoQueryFlg3_3_OnClick()
		frm1.txtRadio3.value = frm1.rdoQueryFlg3_3.value
	End Sub

'========================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
	
	'-----------------------
    'Query function call area
    '-----------------------
	With frm1
		If .rdoQueryFlg2_1.checked = True Then
			.txtRadio2.value = .rdoQueryFlg2_1.value
		ElseIf .rdoQueryFlg2_2.checked = True Then
			.txtRadio2.value = .rdoQueryFlg2_2.value
		ElseIf .rdoQueryFlg2_3.checked = True Then
			.txtRadio2.value = .rdoQueryFlg2_3.value
		ElseIf .rdoQueryFlg3_1.checked = True Then
			.txtRadio3.value = .rdoQueryFlg3_1.value
		ElseIf .rdoQueryFlg3_2.checked = True Then
			.txtRadio3.value = .rdoQueryFlg3_2.value
		ElseIf .rdoQueryFlg3_3.checked = True Then
			.txtRadio3.value = .rdoQueryFlg3_3.value
		End If		
	End With
	
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================================================================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
			strVal = strVal & "&txtBp_cd=" & Trim(.HBp_cd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtBp_nm=" & Trim(.HBp_nm.value)
			strVal = strVal & "&txtBiz_grp=" & Trim(.HBiz_grp.value)
			strVal = strVal & "&txtPur_grp=" & Trim(.HPur_grp.value)
			strVal = strVal & "&txtRadio2=" & Trim(.HRadio2.value)
			strVal = strVal & "&txtRadio3=" & Trim(.HRadio3.value)	
                        'strVal = strVal & "&txtRadio3=" & Trim(.txtRadio3.value)	
			strVal = strVal & "&txtOwnRgstN=" & Trim(.HOwn_Rgst_N.value)		
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtBp_cd=" & Trim(.txtBp_cd.value)
			strVal = strVal & "&txtBp_nm=" & Trim(.txtBp_nm.value)
			strVal = strVal & "&txtBiz_grp=" & Trim(.txtBiz_grp.value)
			strVal = strVal & "&txtPur_grp=" & Trim(.txtPur_grp.value)					
			strVal = strVal & "&txtRadio2=" & Trim(.txtRadio2.value)
			strVal = strVal & "&txtRadio3=" & Trim(.txtRadio3.value)	
			strVal = strVal & "&txtOwnRgstN=" & Trim(.txtOwn_Rgst_N.value)						
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If				
		
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function

'========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else	
		frm1.txtBp_cd.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
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
						<TD CLASS=TD5 NOWRAP>거래처코드</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBp_cd" SIZE=20 TAG="11XXXU" ALT="거래처코드"></TD>
						<TD CLASS=TD5 NOWRAP>거래처구분</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg2" TAG="11X" VALUE="A" CHECKED ID="rdoQueryFlg2_1"><LABEL FOR="rdoQueryFlg2_1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg2" TAG="11X" VALUE="C" ID="rdoQueryFlg2_2"><LABEL FOR="rdoQueryFlg2_2">매출처</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg2" TAG="11X" VALUE="S" ID="rdoQueryFlg2_3"><LABEL FOR="rdoQueryFlg2_3">매입처</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>거래처약칭</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBp_nm" SIZE=30 TAG="11XXXU" ALT="거래처명"></TD>
						<TD CLASS=TD5 NOWRAP>사용여부</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO"   NAME="rdoQueryFlg3" TAG="11X" VALUE="A" ID="rdoQueryFlg3_1"><LABEL FOR="rdoQueryFlg3_1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO"   NAME="rdoQueryFlg3" TAG="11X" VALUE="Y" CHECKED ID="rdoQueryFlg3_2"><LABEL FOR="rdoQueryFlg3_2">사용</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO"   NAME="rdoQueryFlg3" TAG="11X" VALUE="N" ID="rdoQueryFlg3_3"><LABEL FOR="rdoQueryFlg3_3">미사용</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBiz_grp" SIZE=10 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBiz_grp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT TYPE=TEXT NAME="txtSales_grp_nm" SIZE=25 TAG="14"></TD>
						<TD CLASS=TD5 NOWRAP>구매그룹</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPur_grp" SIZE=10 TAG="11XXXU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPur_grp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPur_grp_nm" SIZE=25 TAG="14"></TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>사업자등록번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtOwn_Rgst_N" SIZE=30 TAG="11XXXU" ALT="사업자등록번호"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP>	
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
						<script language =javascript src='./js/b1261pa1_vaSpread_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                     <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtRadio2" tag="14">
<INPUT TYPE=HIDDEN NAME="txtRadio3" tag="14">

<INPUT TYPE=HIDDEN NAME="HBp_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_nm" tag="24">
<INPUT TYPE=HIDDEN NAME="HBiz_grp" tag="24">
<INPUT TYPE=HIDDEN NAME="HPur_grp" tag="24">

<INPUT TYPE=HIDDEN NAME="HRadio1" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadio2" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadio3" tag="24">

<INPUT TYPE=HIDDEN NAME="HOwn_Rgst_N" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
