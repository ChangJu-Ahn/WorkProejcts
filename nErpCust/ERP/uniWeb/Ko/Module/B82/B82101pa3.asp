<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          :                                                                  *
'*  2. Function Name        :                                                                  *
'*  3. Program ID           :                                                                  *
'*  4. Program Name         :                                                                  *
'*  5. Program Desc         :                                                                  *
'*  7. Modified date(First) :                                                                  *
'*  8. Modified date(Last)  :                                                                  *
'*  9. Modifier (First)     :                                                                  *
'* 10. Modifier (Last)      :                                                                  *
'* 11. Comment              :                                                                  *
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID = "B82101pb3.asp"           <% '☆: 비지니스 로직 ASP명 %>
Const C_SHEETMAXROWS_D = 100				'Sheet Max Rows

Dim C_ITEM_LVL1_CD               '대분류코드 
Dim C_ITEM_LVL1_NM               '대분류코드 
Dim C_ITEM_LVL2_CD               '중분류코드 
Dim C_ITEM_LVL2_NM               '중분류코드 
Dim C_ITEM_LVL3_CD               '소분류코드 
Dim C_ITEM_LVL3_NM               '소분류코드 
Dim C_SPEC_ORDER                 '특성순서 
Dim C_SPEC_NAME                  '특성명 
Dim C_Spec_Value                 '입력값*****
Dim C_Spec_Split                 '구분자*****
Dim C_SPEC_UNIT                  '특성단위 
Dim C_SPEC_LENGTH                '길이 
Dim C_SPEC_EXAMPLE               '예제 
Dim C_REMARK                     '참조  

Dim IsOpenPop                                                                                             
Dim arrReturn
Dim arrParent
Dim arrParam                         
Dim arrField
Dim PopupParent
                    
arrParent = window.dialogArguments

Set PopupParent = arrParent(0)

arrParam = arrParent(1)
arrField = arrParent(2)

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : InitSpreadPosVariables()     
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
    C_ITEM_LVL1_CD = 1              '대분류코드 
    C_ITEM_LVL1_NM = 2              '대분류코드 
    C_ITEM_LVL2_CD = 3              '중분류코드 
    C_ITEM_LVL2_NM = 4              '중분류코드 
    C_ITEM_LVL3_CD = 5              '소분류코드 
    C_ITEM_LVL3_NM = 6              '소분류코드 
    C_SPEC_ORDER   = 7              '특성순서 
    C_SPEC_NAME    = 8              '특성명 
    C_Spec_Value   = 9              '입력값*****
    C_Spec_Split   = 10              '구분자*****
    C_SPEC_UNIT    = 11             '특성단위 
    C_SPEC_LENGTH  = 12             '길이 
    C_SPEC_EXAMPLE = 13             '예제 
    C_REMARK       = 14             '참조 
End Sub

'========================================================================================================
' Name : InitVariables()     
' Desc : Initialize value
'========================================================================================================
Function InitVariables()

     lgIntGrpCount      = 0                                      <%'⊙: Initializes Group View Size%>
     lgStrPrevKey       = ""                           'initializes Previous Key          
     lgStrPrevKeyIndex  = ""
     lgIntFlgMode       = PopupParent.OPMD_CMODE
     Redim arrReturn(0)
     Self.Returnvalue   = arrReturn
          
End Function

'========================================================================================================
' Name : SetDefaultVal()     
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

     frm1.cboItemAcct.value    = arrParam(0)
     frm1.txtItemKind.value    = arrParam(1)
     frm1.txtItemKindNm.value  = arrParam(2)
     frm1.txtItemLvl1.value    = arrParam(3)
     frm1.txtItemLvl1Nm.value  = arrParam(4)
     frm1.txtItemLvl2.value    = arrParam(5)
     frm1.txtItemLvl2Nm.value  = arrParam(6)
     frm1.txtItemLvl3.value    = arrParam(7)
     frm1.txtItemLvl3Nm.value  = arrParam(8)
     
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
     <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
     <%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'========================================================================================================
' Name : InitComboBox()     
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
     Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1001' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
     Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
     Call InitSpreadPosVariables()
     
     ggoSpread.Source = frm1.vspdData
	 ggoSpread.Spreadinit "V20050201", , Popupparent.gAllowDragDropSpread
	    
     frm1.vspdData.ReDraw = False
	            
     frm1.vspdData.MaxCols = C_REMARK + 1
     frm1.vspdData.MaxRows = 0
	
	 Call GetSpreadColumnPos("A")

	 Call AppendNumberPlace("6","3","0")    
         
     ggoSpread.SSSetEdit   C_ITEM_LVL1_CD,   "대분류",         10
     ggoSpread.SSSetEdit   C_ITEM_LVL1_NM,   "대분류",         12
     ggoSpread.SSSetEdit   C_ITEM_LVL2_CD,   "중분류",         10
     ggoSpread.SSSetEdit   C_ITEM_LVL2_NM,   "중분류",         12
     ggoSpread.SSSetEdit   C_ITEM_LVL3_CD,   "소분류",         10
     ggoSpread.SSSetEdit   C_ITEM_LVL3_NM,   "소분류",         12
     ggoSpread.SSSetEdit   C_SPEC_ORDER,     "순서",           5
     ggoSpread.SSSetEdit   C_SPEC_NAME,      "특성명",         20
     ggoSpread.SSSetEdit   C_Spec_Value,     "입력값",         10
     ggoSpread.SSSetEdit   C_Spec_Split,     "구분자",         10
     ggoSpread.SSSetEdit   C_SPEC_UNIT,      "단위",           8
     ggoSpread.SSSetEdit   C_SPEC_LENGTH,    "길이",           8
     ggoSpread.SSSetEdit   C_SPEC_EXAMPLE,   "예제",           10
     ggoSpread.SSSetEdit   C_REMARK,         "참조",           50
                       
     Call ggoSpread.SSSetColHidden(C_ITEM_LVL1_CD, C_ITEM_LVL3_NM,      True)
     
     Call ggoSpread.SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)

     Call SetSpreadLock()

     frm1.vspdData.ReDraw = True
     
     
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
Sub SetSpreadLock()
     
     With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock       C_ITEM_LVL1_CD, -1, C_SPEC_NAME		, -1         
        ggoSpread.SpreadLock       C_SPEC_UNIT	 , -1, C_REMARK		, -1 
        'ggoSpread.SSSetProtected   C_Spec_Value  , -1, -1
        'ggoSpread.SSSetProtected   C_Spec_Split  , -1, -1
		.vspdData.ReDraw = True
    End With
    
    
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)

    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
       
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_ITEM_LVL1_CD = iCurColumnPos(1)              '대분류코드 
			C_ITEM_LVL1_NM = iCurColumnPos(2)              '대분류코드 
			C_ITEM_LVL2_CD = iCurColumnPos(3)              '중분류코드 
			C_ITEM_LVL2_NM = iCurColumnPos(4)              '중분류코드 
			C_ITEM_LVL3_CD = iCurColumnPos(5)              '소분류코드 
			C_ITEM_LVL3_NM = iCurColumnPos(6)              '소분류코드 
			C_SPEC_ORDER   = iCurColumnPos(7)              '특성순서 
			C_SPEC_NAME    = iCurColumnPos(8)              '특성명 
			C_Spec_Value   = iCurColumnPos(9)              '입력값*****
			C_Spec_Split   = iCurColumnPos(10)             '구분자*****
			C_SPEC_UNIT    = iCurColumnPos(11)             '특성단위 
			C_SPEC_LENGTH  = iCurColumnPos(12)             '길이 
			C_SPEC_EXAMPLE = iCurColumnPos(13)             '예제 
			C_REMARK       = iCurColumnPos(14)             '참조 
    End Select    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
     
     gMouseClickStatus = "SPC"                         'SpreadSheet 대상명이 vspdData일경우 
     Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000111111")

    If frm1.vspdData.MaxRows <= 0 Then Exit Sub
            
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
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
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
     If Row = 0 Then 
        Exit Function
     End If

     'If frm1.vspdData.MaxRows > 0 Then
     '     If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
     '        Call OKClick
     '     End If
     'End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)          
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyDown
'   Event Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
     If KeyAscii = 27 Then
        Call CancelClick()
     ElseIf KeyAscii = 13 and frm1.vspdData.ActiveRow > 0 Then
        Call OkClick()
     End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
     With frm1.vspdData
          If Row >= NewRow Then
             Exit Sub
          End If
          If NewRow = frm1.vspdData.MaxRows Then
             If lgStrPrevKeyIndex <> "" Then                                   '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
                 If DbQuery = False Then
                    Exit Sub
                 End If
             End If
          End If
     End With
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
     If OldLeft <> NewLeft Then
         Exit Sub
     End If

     if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
          If lgStrPrevKeyIndex <> "" Then                                   '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
             If DbQuery = False Then
                 Exit Sub
             End If
          End If
     End If
End Sub

'========================================================================================================
'     Name : OKClick()
'     Desc : handle ok icon click event
'========================================================================================================
Function OKClick()
     Dim i, sValue , sSplit,sTemp
     
     If frm1.vspdData.MaxRows > 0 Then
          
          Redim arrReturn(0)

          ggoSpread.Source = frm1.vspdData
          
          For i = 1 To frm1.vspdData.MaxRows
              frm1.vspdData.Row = i
              frm1.vspddata.Col = C_Spec_Value
              sValue = frm1.vspdData.Text
              frm1.vspddata.Col = C_Spec_Split
              sSplit = frm1.vspdData.Text
              sTemp=sTemp & sValue & sSplit
              //If i = frm1.vspdData.MaxRows Then
              //   arrReturn(0) = arrReturn(0) & sValue
              //Else                
              //   arrReturn(0) = arrReturn(0) & sValue & sSplit
              //End If
          Next
          arrReturn(0)=sTemp
          
          Self.Returnvalue = arrReturn
     End If

    Self.Close()
                    
End Function

'========================================================================================================
'     Name : CancelClick()
'     Desc : handle  Cancel click event
'========================================================================================================
Function CancelClick()
     Self.Close()
End Function

'========================================================================================================
'     Name : MousePointer()
'     Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
                    window.document.search.style.cursor = "wait"
            case "POFF"
                    window.document.search.style.cursor = ""
      End Select
End Function


'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()	
     Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")
     Call LoadInfTB19029                                         '⊙: Load table , B_numeric_format
     Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
     Call InitVariables
     Call ggoOper.LockField(Document, "N")                       '⊙: Lock  Suitable  Field
     Call InitComboBox()
     Call SetDefaultVal()
     Call InitSpreadSheet()
     Call FncQuery()
End Sub

'========================================================================================================
'     Name : FncQuery()
'     Desc : 
'========================================================================================================
Function FncQuery()
     FncQuery = False
     Call InitVariables()
          
     frm1.vspdData.MaxRows = 0                              'Grid 초기화 

     lgIntFlgMode = PopupParent.OPMD_CMODE     

     If DbQuery = False Then
        Exit Function
     End If
     
     FncQuery = True
     
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function


'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()

    Dim strVal
    
    If Not chkField(Document, "1") Then                                             
       Exit Function
    End If
    
    lgKeyStream =               Trim(frm1.cboItemAcct.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemKind.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemLvl1.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemLvl2.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemLvl3.value)	& Chr(11)
    
    DbQuery = False                                                                 '⊙: Processing is NG
     
    Call LayerShowHide(1)                                                           '⊙: 작업진행중 표시 
    
    strVal = BIZ_PGM_ID & "?txtMode="        & PopupParent.UID_M0001                '☜: Query
    strVal = strVal     & "&txtKeyStream="   & lgKeyStream                          '☜: Query Key
    strVal = strVal     & "&txtPrevNext="    & ""                                   '☜: Direction
    strVal = strVal     & "&lgStrPrevKeyIndex="   & lgStrPrevKeyIndex               '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="     & Frm1.vspdData.MaxRows                '☜: Max fetched data
    strVal = strVal     & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)               '☜: Max fetched data at a time
    
    Call RunMyBizASP(MyBizASP, strVal)                                              '☜: 비지니스 ASP 를 가동 
     
    DbQuery = True                                                                  '⊙: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
     If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
          Call SetActiveCell(vspdData,1,1,"P","X","X")
          Set gActiveElement = document.activeElement
     End If
    lgIntFlgMode = PopupParent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")                                             '⊙: This function lock the suitable field
End Function
    
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->     
</HEAD>
<!--
'########################################################################################################
'#                              6. TAG 부                                                                                          #
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
							   <TD CLASS=TD5 NOWRAP>품목계정</TD>
							   <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="24" ALT="품목계정"><OPTION VALUE=""></OPTION></SELECT></TD>
							   <TD CLASS=TD5 NOWRAP>품목구분</TD>
							   <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemKind" ALT="품목구분" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
							                        <INPUT NAME="txtItemKindNm" ALT="품목구분명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
							</TR>     
							<TR>
							    <TD CLASS=TD5 NOWRAP>대분류</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl1" ALT="대분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
							                         <INPUT NAME="txtItemLvl1Nm" ALT="대분류명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
							    <TD CLASS=TD5 NOWRAP>중분류</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl2" ALT="중분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
							                         <INPUT NAME="txtItemLvl2Nm" ALT="중분류명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>                                        
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>소분류</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl3" ALT="소분류" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="24XXXU">
							                         <INPUT NAME="txtItemLvl3Nm" ALT="소분류명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>							   
							    <TD CLASS=TD5 NOWRAP>&nbsp;</TD>                     
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
                              <script language =javascript src='./js/b82101pa3_vaSpread1_vspdData.js'></script>
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
                         <TD WIDTH=10>&nbsp;</TD>
                         <TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
                         <TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
                                                   <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>                         </TD>
                         <TD WIDTH=10>&nbsp;</TD>
                    </TR>
               </TABLE>
          </TD>
     </TR>
     <TR>
          <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
          </TD>
     </TR>
</TABLE>

<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
     <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>

<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpread"       TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRdoStatus"    TAG="24" TABINDEX="-1">

</FORM>
</BODY>
</HTML>