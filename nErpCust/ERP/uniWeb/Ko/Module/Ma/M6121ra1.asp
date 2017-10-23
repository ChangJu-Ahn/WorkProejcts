<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m6121ra1
'*  4. Program Name         : 배부내역참조 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/11/15	
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!--'☆: 해당 위치에 따라 달라짐, 상대 경로 -->
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">
Option Explicit					 '☜: indicates that All variables must be declared in advance 
	

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

Const C_ChargeNo 	= 1
Const C_BasNo 		= 2	
Const C_SeqNo		= 3	
Const C_DisbType	= 4
Const C_BaseQty 	= 5											'☆: Spread Sheet의 Column별 상수 
Const C_BaseAmt 	= 6
Const C_DisbQty		= 7
Const C_DisbAmt		= 8
Const C_PlantCd		= 9
Const C_PlantNm 	= 10
Const C_ItemCd 		= 11
Const C_ItemNm 		= 12
Const C_Spec 		= 13	
Const C_MvmtNo		= 14
Const C_PoNo		= 15
Const C_PoSeqNo		= 16	

Const BIZ_PGM_ID 		= "m6121rb1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 16                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop  
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent

Dim lgStrPrevKey1, lgStrPrevKey2, lgStrPrevKey3

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '☆: 
End Sub

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgStrPrevKey1     = "" : lgStrPrevKey2     = "" : lgStrPrevKey3     = ""
		lgPageNo         = ""
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
        lgSortKey        = 1   
        
        lgIntGrpCount = 0										'⊙: Initializes Group View Size

        Redim arrReturn(0,0)        
        Self.Returnvalue = arrReturn     
End Function

Sub SetDefaultVal()
	Dim arrParam
		
	arrParam = arrParent(1)
		
	frm1.vspdData.OperationMode = 5
	frm1.txtPlantCd.value 	= arrParam(0)
	frm1.txtPlantNm.value 		= arrParam(1)
	frm1.txtDisbJobDt.value 	= arrParam(2)
	frm1.txtBatchDt.value 		= arrParam(3)
	frm1.txtDocumentNo.value 	= arrParam(4)
	frm1.txtDistRefNo.value		= arrParam(5)
	
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20041115"			',,PopupParent.gAllowDragDropSpread  
		.ReDraw = false

		.MaxCols = C_PoSeqNo+1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols:    .ColHidden = True
		.MaxRows = 0
		
		'Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit	 C_ChargeNo		, "경비번호",20,,,,2
		ggoSpread.SSSetEdit	 C_BasNo		, "근거번호",20,,,,2
		ggoSpread.SSSetEdit	 C_SeqNo		, "순번",10,,,,2
		ggoSpread.SSSetEdit	 C_DisbType		, "배부유형",10,,,,2
		SetSpreadFloatLocal  C_BaseQty		, "배부기준수량", 20,1,3  
		SetSpreadFloatLocal  C_BaseAmt		, "배부기준금액", 20,1,2
		SetSpreadFloatLocal  C_DisbQty		, "배부수량", 20,1,3
		SetSpreadFloatLocal  C_DisbAmt		, "배부금액", 20,1,2
		ggoSpread.SSSetEdit  C_PlantCd		, "공장",10,,,,2
		ggoSpread.SSSetEdit  C_PlantNm		, "공장명",20
		ggoSpread.SSSetEdit  C_ItemCd		, "품목",18,,,,2
		ggoSpread.SSSetEdit  C_ItemNm		, "품목명",20
		ggoSpread.SSSetEdit  C_Spec			, "규격",20
		ggoSpread.SSSetEdit	 C_MvmtNo		, "입고번호",20,,,,2
		ggoSpread.SSSetEdit	 C_PoNo			, "발주번호",20,,,,2
		ggoSpread.SSSetEdit	 C_PoSeqNo		, "발주순번",10,,,,2
		
		Call SetSpreadLock 
    
		.ReDraw = true
    End With
end sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()	
	With frm1
	ggoSpread.spreadlock -1, -1
    End With
End Sub	

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"P"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"P"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"P"
    End Select
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
'Function CancelClick()
	'Self.Close()
'End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If		

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey1 <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If		 
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
End Function
	

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
				
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value
		strVal = strVal & "&txtBatchJobDt=" & .txtDisbJobDt.value
		strVal = strVal & "&txtDisbDt=" & .txtBatchDt.value
		strVal = strVal & "&txtDocumentNo=" & .txtDocumentNo.value		
		strVal = strVal & "&lgStrPrevKey1="   & lgStrPrevKey1  
		strVal = strVal & "&lgStrPrevKey2="   & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3="   & lgStrPrevKey3 
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows   
   	
        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    
End Function

'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	Call SetFocusToDocument("P")  
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
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
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="14X">
											   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=25 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>배부년월</TD>
						<TD CLASS="TD6"><INPUT NAME="txtDisbJobDt" ALT="배부년월" TYPE="Text" MAXLENGTH=7 SiZE=10 tag="14X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>배부작업일</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="배부작업일" NAME="txtBatchDt" SIZE=10 MAXLENGTH=10 tag="14X">
						<TD CLASS="TD5" NOWRAP>재고처리번호</TD>
						<TD CLASS="TD6"><INPUT NAME="txtDocumentNo" ALT="재고처리번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="14X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>배부참조번호</TD>
						<TD CLASS="TD6"><INPUT NAME="txtDistRefNo" ALT="배부참조번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6"></TD>
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
					<TD HEIGHT="20%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <!--<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>--></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% SRC="../../blank.htm" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     