<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 거래처팝업 
'*  5. Program Desc         : 거래처정보의 거래처팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : 		
'* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID 		= "BpPopUpBiz.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 9                                           '☆: key count of SpreadSheet

'========================================================================================================
                   

Dim IsOpenPop  
Dim lgIsOpenPop
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
Dim lgTableName														    'Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgLabelNm
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

arrParam		= arrParent(1)
Select Case arrParam(5)	
	Case "SUP" 
			top.document.title = "공급처팝업"
	Case "PAYTO"
			top.document.title = "지급처팝업"
	Case "SOL"
			top.document.title = "주문처팝업"
	Case "PAYER"
			top.document.title = "수금처팝업"
	Case "INV"
			top.document.title = "세금계산서발행처팝업"	
	Case Else
			top.document.title = "거래처팝업"					
End Select
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
			
		frm1.txtBp_cd.value		= arrParam(0)				
		lgTableName				=	arrParam(1)
		frm1.hFrDt.value		=	arrParam(2)
		frm1.HToDt.value		=	arrParam(3)
		
		Select Case arrParam(4)				'거래처구분 
			Case "B"	'Bill to Party
				frm1.rdoQueryFlg2_2.checked= true
			Case "S"	'Sold to Party
				frm1.rdoQueryFlg2_3.checked= true	
			Case "T"	'tot
				frm1.rdoQueryFlg2_1.checked= true
		End Select
		
		lgLabelNm = arrParam(5)			 	
		
		frm1.txtRadio3.value = frm1.rdoQueryFlg3_2.value							'사용여부	
	
		frm1.txtBp_cd.focus	  
	End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()	
	Call SetZAdoSpreadSheet("BpPopUp","S","A","V20051106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	ggoSpread.Source = frm1.vspdData
	
	frm1.vspdData.Redraw = False	
	Select Case lgLabelNm		'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 
		Case "SUP" 
			BpCdLabel.innerHTML		="공급처코드"
			BpTypeLabel.innerHTML	="공급처구분"
			BpNmLabel.innerHTML		="공급처약칭"
			ggoSpread.SSSetEdit   GetKeyPos("A",1)   , "공급처" , 10, ,,20,2     
			ggoSpread.SSSetEdit   GetKeyPos("A",2)   , "공급처약칭"  , 20, ,,20,2    
			ggoSpread.SSSetEdit   GetKeyPos("A",6)   , "공급처구분명" , 20, ,,20,2 
			ggoSpread.SSSetEdit   GetKeyPos("A",7)   , "공급처전명" , 30, ,,20,2 
		Case "PAYTO"
			BpCdLabel.innerHTML		="지급처코드"
			BpTypeLabel.innerHTML	="지급처구분"
			BpNmLabel.innerHTML		="지급처약칭"
			ggoSpread.SSSetEdit   GetKeyPos("A",1)   , "지급처" , 10, ,,20,2     
			ggoSpread.SSSetEdit   GetKeyPos("A",2)   , "지급처약칭"  , 20, ,,20,2    
			ggoSpread.SSSetEdit   GetKeyPos("A",6)   , "지급처구분명" , 20, ,,20,2 
			ggoSpread.SSSetEdit   GetKeyPos("A",7)   , "지급처전명" , 30, ,,20,2 
		Case "SOL"
			BpCdLabel.innerHTML		="주문처코드"
			BpTypeLabel.innerHTML	="주문처구분"
			BpNmLabel.innerHTML		="주문처약칭"
			ggoSpread.SSSetEdit   GetKeyPos("A",1)   , "주문처" , 10, ,,20,2     
			ggoSpread.SSSetEdit   GetKeyPos("A",2)   , "주문처약칭"  , 20, ,,20,2    
			ggoSpread.SSSetEdit   GetKeyPos("A",6)   , "주문처구분명" , 20, ,,20,2 
			ggoSpread.SSSetEdit   GetKeyPos("A",7)   , "주문처전명" , 30, ,,20,2 
		Case "PAYER"
			BpCdLabel.innerHTML		="수금처코드"
			BpTypeLabel.innerHTML	="수금처구분"
			BpNmLabel.innerHTML		="수금처약칭"
			ggoSpread.SSSetEdit   GetKeyPos("A",1)   , "수금처" , 10, ,,20,2     
			ggoSpread.SSSetEdit   GetKeyPos("A",2)   , "수금처약칭"  , 20, ,,20,2    
			ggoSpread.SSSetEdit   GetKeyPos("A",6)   , "수금처구분명" , 20, ,,20,2 
			ggoSpread.SSSetEdit   GetKeyPos("A",7)   , "수금처전명" , 30, ,,20,2 
		Case "INV"
			BpCdLabel.innerHTML		="세금계산서발행처코드"
			BpTypeLabel.innerHTML	="세금계산서발행처구분"
			BpNmLabel.innerHTML		="세금계산서발행처약칭"
			ggoSpread.SSSetEdit   GetKeyPos("A",1)   , "세금계산서발행처" , 18, ,,20,2     
			ggoSpread.SSSetEdit   GetKeyPos("A",2)   , "세금계산서발행처약칭"  , 20, ,,20,2    
			ggoSpread.SSSetEdit   GetKeyPos("A",6)   , "세금계산서발행처구분명" , 20, ,,20,2 
			ggoSpread.SSSetEdit   GetKeyPos("A",7)   , "세금계산서발행처전명" , 18, ,,20,2 
		Case Else
	 		BpCdLabel.innerHTML		="거래처코드"
			BpTypeLabel.innerHTML	="거래처구분"
			BpNmLabel.innerHTML		="거래처약칭"
			ggoSpread.SSSetEdit   GetKeyPos("A",1)   , "거래처" , 10, ,,20,2     
			ggoSpread.SSSetEdit   GetKeyPos("A",2)   , "거래처약칭"  , 20, ,,20,2    
			ggoSpread.SSSetEdit   GetKeyPos("A",6)   , "거래처구분명" , 20, ,,20,2 
			ggoSpread.SSSetEdit   GetKeyPos("A",7)   , "거래처전명" , 30, ,,20,2 
		 
	End Select
	
	frm1.vspdData.Redraw =True	
	
	Call SetSpreadLock 	 	    			            
End Sub

'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		.vspdData.ReDraw = True
		.vspdData.OperationMode = 3
    End With
End Sub	

'========================================================================================================
	Function OKClick()

		Dim intColCnt
		
		If frm1.vspdData.ActiveRow > 0 Then	
		
			Redim arrReturn(2)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetkeyPos("A", 1)
			arrReturn(0) = frm1.vspdData.Text
			frm1.vspdData.Col = GetkeyPos("A", 2)
			arrReturn(1) = frm1.vspdData.Text
			
			if lgLabelNm = "PAYTO" then
			frm1.vspdData.Col = GetkeyPos("A", 9)
			arrReturn(2) = frm1.vspdData.Text
			End if				
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

	arrParam(0) = arrParam(5)												' 팝업 명칭	

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


    Call LoadInfTB19029()
     Call ggoOper.LockField(Document, "N")
    
	Call InitVariables()														
	Call SetDefaultVal()	
	Call InitSpreadSheet()
	
	
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
			strVal = strVal & "&txtRadio2=" & Trim(.HRadio2.value)
			strVal = strVal & "&txtRadio3=" & Trim(.HRadio3.value)	
			strVal = strVal & "&txtOwnRgstN=" & Trim(.HOwn_Rgst_N.value)		
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtBp_cd=" & Trim(.txtBp_cd.value)
			strVal = strVal & "&txtBp_nm=" & Trim(.txtBp_nm.value)				
			strVal = strVal & "&txtRadio2=" & Trim(.txtRadio2.value)
			strVal = strVal & "&txtRadio3=" & Trim(.txtRadio3.value)	
			strVal = strVal & "&txtOwnRgstN=" & Trim(.txtOwn_Rgst_N.value)						
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If				
		strVal = strVal & "&lgTableNm="		 & lgTableName	
		strVal = strVal & "&txtFrDt="		 & Trim(.hFrDt.value)	
		strVal = strVal & "&txtToDt="		 & Trim(.HToDt.value)	
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
						<TD CLASS=TD5 ID = "BpCdLabel" NOWRAP>&nbsp;</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBp_cd" SIZE=20 TAG="11XXXU" ALT="거래처코드"></TD>
						<TD CLASS=TD5 ID = "BpTypeLabel"  NOWRAP>&nbsp;</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg2" TAG="11X" VALUE="A"  ID="rdoQueryFlg2_1"><LABEL FOR="rdoQueryFlg2_1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg2" TAG="11X" VALUE="C" ID="rdoQueryFlg2_2"><LABEL FOR="rdoQueryFlg2_2">매출처</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg2" TAG="11X" VALUE="S" ID="rdoQueryFlg2_3"><LABEL FOR="rdoQueryFlg2_3">매입처</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 ID = "BpNmLabel" NOWRAP>&nbsp;</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBp_nm" SIZE=30 TAG="11XXXU" ALT="거래처명"></TD>
						<TD CLASS=TD5 NOWRAP>사용여부</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg3" TAG="11X" VALUE="A"  ID="rdoQueryFlg3_1"><LABEL FOR="rdoQueryFlg3_1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg3" TAG="11X" VALUE="Y" CHECKED ID="rdoQueryFlg3_2"><LABEL FOR="rdoQueryFlg3_2">사용</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg3" TAG="11X" VALUE="N" ID="rdoQueryFlg3_3"><LABEL FOR="rdoQueryFlg3_3">미사용</LABEL>
						</TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtRadio2" tag="14">
<INPUT TYPE=HIDDEN NAME="txtRadio3" tag="14">
<INPUT TYPE=HIDDEN NAME="hToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="HBp_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_nm" tag="24">


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
