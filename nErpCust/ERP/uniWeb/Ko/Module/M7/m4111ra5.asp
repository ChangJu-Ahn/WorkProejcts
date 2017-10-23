<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Purchase																	*
'*  2. Function Name		: 																			*
'*  3. Program ID			: m4111ra5.asp																*
'*  4. Program Name			: 자품목출고예상대비재고 														*
'*  5. Program Desc			: Reference Popup															*
'*  6. Comproxy List        : ADO :																		*
'*  7. Modified date(First)	: 2003/06/17																*
'*  8. Modified date(Last)	: 2005/10/27																*
'*  9. Modifier (First)		: KIM JIHYUN																*
'* 10. Modifier (Last)		: KIM DUKHYUN																*
'* 11. Comment 				:																			*
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--'********************************************  1.1 Inc 선언  ************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===============================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ==============================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script LANGUAGE="VBScript">

Option Explicit

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================
Const BIZ_PGM_ID = "m4111rb5.asp"					'☆: 비지니스 로직 ASP명 

Dim C_ParCnt                    '전용여부		
Dim C_ParItemCd                 '모품목코드 
Dim C_ParItemNm                 '모품목명 
Dim C_IssueMthd                 '출고방법 
Dim C_ChildItemCd               '자품목코드 
Dim C_ChildItemNm               '자품목명 
Dim C_ChildItemSpec             '자품목규격 
Dim C_BaseUnit                  '단위 
Dim C_ReservQty                 '필요수량 
Dim C_IssueQty                  '출고수량 
Dim C_IOnhandQty                '현재고수량 
Dim C_OOnHandQty                '외주처재고수량 

Dim C_Seq
Dim C_ParItemCd2                '모품목코드 
Dim C_ParItemNm2				'모품목명 
Dim C_Issuemthd2                '출고방법 
Dim C_ChildItemCd2              '자품목 
Dim C_ChildItemNm2              '자품목명 
Dim C_ChildItemSpec2            '자품목규격 
Dim C_BaseUnit2                 '단위 
Dim C_ReqmtQty                  '필요수량 
Dim C_IOnHandQty2               '현재고수량 
Dim C_OOnHandQty2               '외주처재고수량 
Dim C_IssueQty2                 '출고수량 
'	=== 2005.07.04 사급구분 추가 =====================================================================================
Dim C_SpplType2					'사급구분				
'	=== 2005.07.04 사급구분 추가 =====================================================================================


'============================================  1.2.2 Global 변수 선언  ==================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim arrParent
Dim arrParam
Dim arrReturn
Dim arrRet

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : InitSpreadPosVariables()	
'========================================================================================================
Sub InitSpreadPosVariables()

	C_ParCnt                    =   1
	C_ParItemCd                 =   2
	C_ParItemNm                 =   3
	C_IssueMthd                 =   4	
	C_ChildItemCd               =   5
	C_ChildItemNm               =   6
	C_ChildItemSpec             =   7
	C_BaseUnit                  =   8		
	C_ReservQty                 =   9
	C_IOnhandQty                =   10
	C_OOnHandQty                =   11
	C_IssueQty                  =   12	
	                            
	C_Seq						=	1
	C_ParItemCd2                =   2
	C_ParItemNm2				=   3
	C_Issuemthd2                =   4
	C_ChildItemCd2              =   5
	C_ChildItemNm2              =   6
	C_ChildItemSpec2            =   7
	C_BaseUnit2                 =   8
	C_ReqmtQty                  =   9
	C_IOnHandQty2               =   10
	C_OOnHandQty2               =   11
	C_IssueQty2                 =   12
'	=== 2005.07.04 사급구분 추가 =====================================================================================	
	C_SpplType2					=	13
'	=== 2005.07.04 사급구분 추가 =====================================================================================	
	
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
	lgIntGrpCount = 0
	lgStrPrevKey = ""
'	Self.Returnvalue = Array("")
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
	
'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
'======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA")%>
	
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

    ggoSpread.Source = frm1.vspdData1
	ggoSpread.Spreadinit "V20030710",, PopupParent.gAllowDragDropSpread
   
	frm1.vspdData1.ReDraw = False
    frm1.vspdData1.MaxCols = C_IssueQty + 1
    frm1.vspdData1.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit     C_ParCnt        ,       "전용여부", 10    	
	ggoSpread.SSSetEdit     C_ParItemCd     ,       "모품목코드", 18    
	ggoSpread.SSSetEdit     C_ParItemNm     ,       "모품목명", 20    
	ggoSpread.SSSetEdit     C_IssueMthd     ,       "출고방법", 12    
	ggoSpread.SSSetEdit     C_ChildItemCd   ,       "자품목코드", 18    
	ggoSpread.SSSetEdit     C_ChildItemNm   ,       "자품목명", 20    
	ggoSpread.SSSetEdit     C_ChildItemSpec ,       "자품목규격", 20  
	ggoSpread.SSSetEdit     C_BaseUnit      ,       "단위", 12    	
	ggoSpread.SSSetFloat    C_ReservQty     ,       "필요수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat    C_IOnhandQty    ,       "현재고수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat    C_OOnHandQty    ,       "외주처재고수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat    C_IssueQty      ,       "출고수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"	

    Call ggoSpread.SSSetColHidden(frm1.vspdData1.MaxCols, frm1.vspdData1.MaxCols, True)

	frm1.vspdData1.ReDraw = True

    ggoSpread.Source = frm1.vspdData2
	ggoSpread.Spreadinit "V20030710",, PopupParent.gAllowDragDropSpread

	frm1.vspdData2.ReDraw = False
	        
    frm1.vspdData2.MaxCols = C_SpplType2 + 1
    frm1.vspdData2.MaxRows = 0

	Call GetSpreadColumnPos("B")
	
	ggoSpread.SSSetEDit     C_Seq		    ,       "순번", 10
	ggoSpread.SSSetEDit     C_ParItemCd2    ,       "모품목코드", 18
	ggoSpread.SSSetEDit     C_ParItemNm2	,       "모품목명", 20
	ggoSpread.SSSetEDit     C_Issuemthd2    ,       "출고방법", 12
	ggoSpread.SSSetEDit     C_ChildItemCd2  ,       "자품목코드", 18
	ggoSpread.SSSetEDit     C_ChildItemNm2  ,       "자품목명", 20
	ggoSpread.SSSetEDit     C_ChildItemSpec2,       "자품목규격", 20
	ggoSpread.SSSetEDit     C_BaseUnit2     ,       "단위", 12
	ggoSpread.SSSetFloat    C_ReqmtQty      ,       "필요수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"    
	ggoSpread.SSSetFloat    C_IOnHandQty2   ,       "현재고수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"    
	ggoSpread.SSSetFloat    C_OOnHandQty2   ,       "외주처재고수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"    
	ggoSpread.SSSetFloat    C_IssueQty2     ,       "출고수량", 15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"     
'	=== 2005.07.04 사급구분 추가 =====================================================================================	
	ggoSpread.SSSetEDit     C_SpplType2	    ,       "사급구분", 12
'	=== 2005.07.04 사급구분 추가 =====================================================================================	

    Call ggoSpread.SSSetColHidden(frm1.vspdData2.MaxCols, frm1.vspdData2.MaxCols, True)
    
'    frm1.vspdData2.Col = C_ParItemCd2 : frm1.vspdData2.ColMerge = 1    
    
'    ggoSpread.SSSetSplit2(1)
	frm1.vspdData2.ReDraw = True
	
	Call SetSpreadLock()
End Sub
'============================ 2.2.4 SetSpreadLock() =====================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SpreadLockWithOddEvenRowColor()
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ParCnt                    =   iCurColumnPos(1)            
			C_ParItemCd                 =   iCurColumnPos(2)
			C_ParItemNm                 =   iCurColumnPos(3)
			C_IssueMthd                 =   iCurColumnPos(4)
			C_ChildItemCd               =   iCurColumnPos(5)
			C_ChildItemNm               =   iCurColumnPos(6)
			C_ChildItemSpec             =   iCurColumnPos(7)
			C_BaseUnit                  =   iCurColumnPos(8)
			C_ReservQty                 =   iCurColumnPos(9)
			C_IOnhandQty                =   iCurColumnPos(10)
			C_OOnHandQty                =   iCurColumnPos(11)
			C_IssueQty                  =   iCurColumnPos(12)			
			
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Seq						=	iCurColumnPos(1)
			C_ParItemCd2                =   iCurColumnPos(2)
			C_ParItemNm2				=   iCurColumnPos(3)
			C_Issuemthd2                =   iCurColumnPos(4)
			C_ChildItemCd2              =   iCurColumnPos(5)
			C_ChildItemNm2              =   iCurColumnPos(6)
			C_ChildItemSpec2            =   iCurColumnPos(7)
			C_BaseUnit2                 =   iCurColumnPos(8)
			C_ReqmtQty                  =   iCurColumnPos(9)
			C_IOnHandQty2               =   iCurColumnPos(10)
			C_OOnHandQty2               =   iCurColumnPos(11)
			C_IssueQty2                 =   iCurColumnPos(12) 
'	=== 2005.07.04 사급구분 추가 =====================================================================================			
			C_SpplType2					=	iCurColumnPos(13)
'	=== 2005.07.04 사급구분 추가 =====================================================================================			
    End Select    

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    frm1.vspdData1.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub


Function OKClick()
	Redim arrReturn(1,1)
    arrReturn(0,0) = "OK"
    
	Self.Returnvalue = arrReturn	
	Self.Close()
End Function	

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(1,1)
    arrReturn(0,0) = "CANCEL"
    
    Self.Returnvalue = arrReturn		
	Self.Close()
End Function
'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

Sub vspdData_KeyPress(keyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Sub	



'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call InitVariables
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call SetDefaultVal()
	Call InitSpreadSheet()
	If DbQuery = False Then	
		Exit Sub
	End If

	frm1.vspddata1.focus

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = frm1.vspdData1
    Call SetPopupMenuItemInf("0000111111")
    
    If frm1.vspdData1.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Function vspdData1_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then Exit Function

	If frm1.vspdData1.MaxRows > 0 Then
		If frm1.vspdData1.ActiveRow = Row Or frm1.vspdData1.ActiveRow > 0 Then
'			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
     On Error Resume Next
End Function

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()

	Dim strVal, txtSupplierCd
	        
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear   

	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
	strVal = strVal & "&txtSupplierCd=" & arrParam(1)
	strVal = strVal & "&txtSpread=" & arrParam(0)	
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	Dim LngRow
	Dim iReqmtQty, iOOhnadQty
	Dim iQty

	'-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    
    With frm1
		.vspdData2.ReDraw = False
		If .vspdData2.MaxRows > 0 Then
			For LngRow = 1 To .vspdData2.MaxRows
				.vspdData2.Row = LngRow
				.vspdData2.Col = C_ReqmtQty				
				iReqmtQty = uniConvNum(.vspdData2.Text,0)
				
				.vspdData2.Col = C_OOnHandQty2
				iOOhnadQty = uniConvNum(.vspdData2.Text,0)			
				iQty = cdbl(iOOhnadQty - iReqmtQty)
				If iQty < 0 Then
					.vspdData2.ForeColor = vbRed
					.vspdData2.Col = C_ParItemCd2
					.vspdData2.ForeColor = vbRed
'	=== 2005.07.04 사급구분 추가 =====================================================================================				
					.vspdData2.Col = C_SpplType2
					if (.vspdData2.Text) = "무상" Then
'	=== 2005.07.04 사급구분 추가 =====================================================================================						
						.hdnReceiptflag.value = "F"						
'	=== 2005.07.04 사급구분 추가 =====================================================================================
					End If
'	=== 2005.07.04 사급구분 추가 =====================================================================================								
				End If
		
			Next
		End If
		.vspdData2.ReDraw = True
	End With
	
	
	If frm1.hdnReceiptflag.value = "F" Then
		userview1.style.display = "NONE"
		userview2.style.display = ""	
	Else
		userview1.style.display = ""
		userview2.style.display = "NONE"						
	End If
    
   
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
	<TR>
		<TD HEIGHT=5 WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자품목재고현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>		
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
 					<TR>
 						<TD CLASS = TD5Y2 NOWRAP>전체오더기준</TD>
 					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE WIDTH="100%" HEIGHT="100%">
				<TR HEIGHT="45%">
					<TD WIDTH="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 width="100%" tag="2" TITLE="SPREAD" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
				<TR HEIGHT="10%">
				<TD HEIGHT>
					<FIELDSET CLASS="CLSFLD">
						<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
 						<TR>
 							<TD CLASS = TD5Y2  NOWRAP >현입고기준</TD>
						</TR>
					</TABLE>
					</FIELDSET>
				</TD>
				</TR>
				<TR HEIGHT="45%">
					<TD WIDTH="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" tag="2" TITLE="SPREAD" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=30% ALIGN=RIGHT ID=userview1>
					<IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
				</TD>
				<TD WIDTH=30% ALIGN=RIGHT ID=userview2>
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
				</TD>
				
			</TR>
		</TABLE>	
	</TD></TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnReceiptflag" value="" tag="24" TabIndex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      