<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 생산
*  2. Function Name        : 생산실적조회및출력
*  3. Program ID           : P4422OA1_LKO391
*  4. Program Name         : P4422OA1_LKO391
*  5. Program Desc         : 생산실적조회및출력
*  6. Comproxy List        :
*  7. Modified date(First) : 2007/01/24
*  8. Modified date(Last)  :
*  9. Modifier (First)     : Lim, JaeBon
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--☜:Print Program needs this vbs file-->

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const CookieSplit       = 1233
Const BIZ_PGM_ID        = "P4422OB1_LKO391.asp"                       'Biz Logic ASP
Const C_SHEETMAXROWS    = 21                                          '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgFlgAllSelected        'When Selected All
Dim lgFlgCancelClicked      'Cancel Button Clicked
Dim lgFlgCopyClicked        'Copy Button Clicked
Dim lgFlgBtnSelectAllClicked 'When btnSelectAll Clicked

Dim StartDate
Dim EndDate
Dim strYear, strMonth, strDay

Call ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)

EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)          '☆: 초기화면에 뿌려지는 시작 날짜
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")     '☆: 초기화면에 뿌려지는 마지막 날짜

Dim C_Select
Dim C_Wc_Nm
Dim C_Prodt_Order_No
Dim C_Item_Cd
Dim C_Item_Nm
Dim C_Report_Dt
Dim C_Report_Type
Dim C_Prod_Qty_In_Order_Unit
Dim C_Insp_Good_Qty_In_Order_Unit
Dim C_Insp_Bad_Qty_In_Order_Unit
Dim C_Rcpt_Qty_In_Order_Unit
Dim C_Remark
Dim C_Cur_Cd
Dim C_Seq
Dim C_Opr_No
Dim C_Inst_Dt
Dim C_Report_Type_Cd

'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()
    C_Select                      = 1
    C_Wc_Nm                       = 2
    C_Prodt_Order_No              = 3
    C_Item_Cd                     = 4
    C_Item_Nm                     = 5
    C_Report_Dt                   = 6
    C_Report_Type                 = 7
    C_Prod_Qty_In_Order_Unit      = 8
    C_Insp_Good_Qty_In_Order_Unit = 9
    C_Insp_Bad_Qty_In_Order_Unit  = 10
    C_Rcpt_Qty_In_Order_Unit      = 11
    C_Remark                      = 12
    C_Cur_Cd                      = 13
    C_Seq                         = 14
    C_Opr_No                      = 15
    C_Inst_Dt                     = 16
    C_Report_Type_Cd              = 17
End Sub

'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgIntFlgMode      = Parent.OPMD_CMODE                       '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue  = False                                   '⊙: Indicates that no value changed
    lgIntGrpCount     = 0                                       '⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
    lgFlgAllSelected = False
    lgFlgCancelClicked = False
    lgFlgCopyClicked = False
    lgFlgBtnSelectAllClicked = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtStartDt.Text = StartDate
    frm1.txtEndDt.Text = EndDate
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc :
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream = frm1.hPlantCd.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hStartDt.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hEndDt.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hFromWcCd.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hFromItemCd.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hToItemCd.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hProdtOrderNo1.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hProdtOrderNo2.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hRdoFlg.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hBsItemCd.value & Parent.gColSep
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()

    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData

        ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread

        .ReDraw         = false
        .MaxCols        = C_Report_Type_Cd + 1                                                     ' ☜:☜: Add 1 to Maxcols
        .Col            = .MaxCols                                                              ' ☜:☜: Hide maxcols
        .ColHidden      = True                                                                  ' ☜:☜:
        .MaxRows        = 0

        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A")

        ggoSpread.SSSetCheck    C_Select                     , ""            , 2,,,1
        ggoSpread.SSSetEdit     C_Wc_Nm                      , "작업장"      , 20
        ggoSpread.SSSetEdit     C_Prodt_Order_No             , "제조오더번호", 15
        ggoSpread.SSSetEdit     C_Item_Cd                    , "품목코드"    , 15,,,15,2
        ggoSpread.SSSetEdit     C_Item_Nm                    , "품목명"      , 20
        ggoSpread.SSSetDate     C_Report_Dt                  , "실적일"      , 11, 2, parent.gDateFormat
        ggoSpread.SSSetEdit     C_Report_Type                , "양불구분"    , 10, 2
        ggoSpread.SSSetFloat    C_Prod_Qty_In_Order_Unit     , "실적수량"    , 15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_Insp_Good_Qty_In_Order_Unit, "양품수량"    , 15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_Insp_Bad_Qty_In_Order_Unit , "불량수량"    , 15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_Rcpt_Qty_In_Order_Unit     , "입고수량"    , 15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetEdit     C_Remark                     , "비고"        , 20
        ggoSpread.SSSetEdit     C_Inst_Dt                    , "입력일시"    , 18
        ggoSpread.SSSetEdit     C_Report_Type_Cd             , ""            , 18


        Call ggoSpread.SSSetColHidden(C_Cur_Cd, C_Cur_Cd, True)
        Call ggoSpread.SSSetColHidden(C_Seq   , C_Seq   , True)
        Call ggoSpread.SSSetColHidden(C_Opr_No, C_Opr_No, True)
        Call ggoSpread.SSSetColHidden(C_Insp_Good_Qty_In_Order_Unit, C_Insp_Good_Qty_In_Order_Unit, True)
        Call ggoSpread.SSSetColHidden(C_Insp_Bad_Qty_In_Order_Unit , C_Insp_Bad_Qty_In_Order_Unit, True)
        Call ggoSpread.SSSetColHidden(C_Rcpt_Qty_In_Order_Unit     , C_Rcpt_Qty_In_Order_Unit, True)
        Call ggoSpread.SSSetColHidden(C_Report_Type_Cd             , C_Report_Type_Cd, True)

        .ReDraw = true

        Call SetSpreadLock

    End With
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
        frm1.vspdData.ReDraw = False
        ggoSpread.SpreadLock     C_Wc_Nm                      , -1, C_Wc_Nm
        ggoSpread.SpreadLock     C_Prodt_Order_No             , -1, C_Prodt_Order_No
        ggoSpread.SpreadLock     C_Item_Cd                    , -1, C_Item_Cd
        ggoSpread.SpreadLock     C_Item_Nm                    , -1, C_Item_Nm
        ggoSpread.SpreadLock     C_Report_Dt                  , -1, C_Report_Dt
        ggoSpread.SpreadLock     C_Report_Type                , -1, C_Report_TYPE
        ggoSpread.SpreadLock     C_Prod_Qty_In_Order_Unit     , -1, C_Prod_Qty_In_Order_Unit
        ggoSpread.SpreadLock     C_Insp_Good_Qty_In_Order_Unit, -1, C_Insp_Good_Qty_In_Order_Unit
        ggoSpread.SpreadLock     C_Insp_Bad_Qty_In_Order_Unit , -1, C_Insp_Bad_Qty_In_Order_Unit
        ggoSpread.SpreadLock     C_Rcpt_Qty_In_Order_Unit     , -1, C_Rcpt_Qty_In_Order_Unit
        ggoSpread.SpreadLock     C_Remark                     , -1, C_Remark
        ggoSpread.SpreadLock     C_Inst_Dt                    , -1, C_Inst_Dt

        frm1.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If
       Next
    End If
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

            C_Select                      = iCurColumnPos(1)
            C_Wc_Nm                       = iCurColumnPos(2)
            C_Prodt_Order_No              = iCurColumnPos(3)
            C_Item_Cd                     = iCurColumnPos(4)
            C_Item_Nm                     = iCurColumnPos(5)
            C_Report_Dt                   = iCurColumnPos(6)
            C_Report_Type                 = iCurColumnPos(7)
            C_Prod_Qty_In_Order_Unit      = iCurColumnPos(8)
            C_Insp_Good_Qty_In_Order_Unit = iCurColumnPos(9)
            C_Insp_Bad_Qty_In_Order_Unit  = iCurColumnPos(10)
            C_Rcpt_Qty_In_Order_Unit      = iCurColumnPos(11)
            C_Remark                      = iCurColumnPos(12)
            C_Cur_Cd                      = iCurColumnPos(13)
            C_Seq                         = iCurColumnPos(14)
            C_Opr_No                      = iCurColumnPos(15)
            C_Inst_Dt                     = iCurColumnPos(16)
            C_Report_Type_Cd              = iCurColumnPos(17)
    End Select
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                           '☜: Clear err status

    Call LoadInfTB19029                                                                 '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                               '⊙: Lock Field

    Call SetDefaultVal
    Call InitSpreadSheet                                                                'Setup the Spread sheet
    Call InitVariables                                                                  'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)                      ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetToolbar("1100000000001111")                                                 ' 버튼 툴바 제어

    If parent.gPlant <> "" Then
        frm1.txtPlantCd.value = UCase(parent.gPlant)
        frm1.txtPlantNm.value = parent.gPlantNm
        frm1.txtFromWcCd.focus
        Set gActiveElement = document.activeElement
    Else
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
    End If
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False                                                                    '☜: Processing is NG
    Err.Clear                                                                           '☜: Clear err status

    If  ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
        Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")                    '☜: Data is changed.  Do you want to display it?

        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    ggoSpread.ClearSpreadData

    If Not chkField(Document, "1") Then                                                 '☜: This function check required field
        Exit Function
    End If

    frm1.hPlantCd.value = Trim(frm1.txtPlantCd.value)
    frm1.hStartDt.value = Trim(frm1.txtStartDt.Text)
    frm1.hEndDt.value = Trim(frm1.txtEndDt.Text)
    frm1.hFromWcCd.value = Trim(frm1.txtFromWcCd.value)
    frm1.hFromItemCd.value = Trim(frm1.txtFromItemCd.value)
    frm1.hToItemCd.value = Trim(frm1.txtToItemCd.value)
    frm1.hProdtOrderNo1.value = Trim(frm1.txtProdtOrderNo1.value)
    frm1.hProdtOrderNo2.value = Trim(frm1.txtProdtOrderNo2.value)
    frm1.hBsItemCd.value = Trim(frm1.txtBsItemcd.value)

    If frm1.rdoFlg1.checked = True Then
        frm1.hRdoFlg.value = ""
    ElseIf frm1.rdoFlg2.checked = True Then
        frm1.hRdoFlg.value = "G"
    Else
        frm1.hRdoFlg.value = "B"
    End If

    Call InitVariables                                                                  '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call DisableToolBar(Parent.TBC_QUERY)

    If DbQuery = False Then
        Call RestoreTooBar()
        Exit Function
    End If

    FncQuery = True                                                                     '☜: Processing is OK
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)                                               '☜: 화면 유형
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)                                          '☜:화면 유형, Tab 유무
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc :
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
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
    Call InitData()
End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()
    Dim IntRetCD

    FncExit = False

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                    '⊙: Data is changed.  Do you want to exit?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    DbQuery = False

    Err.Clear                                                                           '☜: Clear err status

    if LayerShowHide(1) = False then
        Exit Function
    end if

    Dim strVal

    With Frm1
        strVal = BIZ_PGM_ID & "?txtMode="       & Parent.UID_M0001
        strVal = strVal     & "&txtKeyStream="  & lgKeyStream                           '☜: Query Key
        strVal = strVal     & "&txtMaxRows="    & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey                          '☜: Next key tag
    End With

    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If
    Call RunMyBizASP(MyBizASP, strVal)                                                  '☜: Run Biz Logic

    DbQuery = True
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
    lgBlnFlgChgValue = False

    Call ggoOper.LockField(Document, "Q")                                               '⊙: Lock field
    Call InitData()
    Call SetToolbar("1100000000011111")

    frm1.vspdData.focus
End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex

    With frm1.vspdData
        .Row = Row
    End With

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim iProdtQty,iRcptqty
    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        If Row < 1 Then Exit Sub

        Select Case Col

            Case C_Select
                .Col = C_Report_Type_Cd
                .Row = Row

				If Trim(.Text) = "B" Then
				
                    .Col = C_Select
                    .Row = Row
                    .Text = "0"
                     Exit Sub
				End If
				 
                      iProdtQty =GetSpreadText(frm1.vspdData, C_Prod_Qty_In_Order_Unit, Row, "X", "X")
                      iRcptqty =GetSpreadText(frm1.vspdData, C_Rcpt_Qty_In_Order_Unit, Row, "X", "X")
                       
				If  unicdbl(iProdtQty) <> unicdbl(iRcptqty)  Then
			
                    .Col = C_Select
                    .Row = Row
                    .Text = "0"
                     Exit Sub
				End If
				
        End Select
    End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
        If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
            Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
        End If
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
    End If

    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData

       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending
           lgSortKey = 1
       End If

       Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

    If Row <= 0 Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKey <> "" Then
            If CheckRunningBizProcess = True Then
                Exit Sub
            End If

            Call DisableToolBar(Parent.TBC_QUERY)

            If DBQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
        End If
    End If
End Sub

'=======================================
'   Event Name : txtStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        Call SetFocusToDocument("M")

        frm1.txtStartDt.Action = 7
        frm1.txtStartDt.focus
    End If
End Sub

'=======================================
'   Event Name : txtEndDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        Call SetFocusToDocument("M")

        frm1.txtEndDt.Action = 7
        frm1.txtEndDt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtStartDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtEndDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Function OpenPlantCd()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "공장팝업"                ' 팝업 명칭
    arrParam(1) = "B_PLANT"                     ' TABLE 명칭
    arrParam(2) = Trim(frm1.txtPlantCd.Value)   ' Code Condition
    arrParam(3) = ""                            ' Name Cindition
    arrParam(4) = ""                            ' Where Condition
    arrParam(5) = "공장"                    ' TextBox 명칭

    arrField(0) = "PLANT_CD"                    ' Field명(0)
    arrField(1) = "PLANT_NM"                    ' Field명(1)

    arrHeader(0) = "공장"                   ' Header명(0)
    arrHeader(1) = "공장명"                 ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetPlantCd(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtPlantCd.focus

End Function

Function SetPlantCd(ByVal arrRet)
    frm1.txtPlantCd.value = arrRet(0)
    frm1.txtPlantNm.value = arrRet(1)
End Function

Function OpenFromWcCd()


    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call displaymsgbox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

   arrParam(0) = "작업장팝업"                                          ' 팝업 명칭
   arrParam(1) = "P_WORK_CENTER A ,  B_MAJOR B , B_MINOR C ,B_CONFIGURATION D "                                            ' TABLE 명칭
   arrParam(2) = Trim(frm1.txtFromWcCd.Value)                                  ' Code Condition
   arrParam(3) = ""'Trim(frm1.txtWCNm.Value)                               ' Name Cindition
   arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")        ' Where Condition

   arrParam(4) = arrParam(4)  & " AND B.MAJOR_CD = C.MAJOR_CD "
   arrParam(4) = arrParam(4)  & " AND C.MAJOR_CD = D.MAJOR_CD AND C.MINOR_CD =  D.MINOR_CD "
   arrParam(4) = arrParam(4)  & " AND A.WC_CD = D.REFERENCE "
   arrParam(4) = arrParam(4)  & " AND B.MAJOR_CD ='Z9001' AND C.MINOR_CD ="& FilterVar(  parent.gUsrId , "''", "S")
   arrParam(5) = "작업장"                                              ' TextBox 명칭

    arrField(0) = "WC_CD"                                                   ' Field명(0)
    arrField(1) = "WC_NM"                                                   ' Field명(1)

    arrHeader(0) = "작업장"                                             ' Header명(0)
    arrHeader(1) = "작업장명"                                           ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SETFromWcCd(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtFromWcCd.focus
End Function

Function SetFromWcCd(ByVal arrRet)
    frm1.txtFromWcCd.value = arrRet(0)
    frm1.txtFromWcNm.value = arrRet(1)
End Function

Function OpenFromItemCd()
    Dim arrRet
    Dim arrParam(5), arrField(6)
    Dim iCalledAspName

    If IsOpenPop = True Then
        IsOpenPop = False
        Exit Function
    End If

    If frm1.txtPlantCd.value= "" Then
        Call parent.DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    iCalledAspName = AskPRAspName("B1B11PA3")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
    arrParam(1) = frm1.txtfromItemCd.Value      ' Item Code
    arrParam(2) = "12!MO"                       ' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분
    arrParam(3) = ""                            ' Default Value

    arrField(0) = 1                             ' Field명(0) :"ITEM_CD"
    arrField(1) = 2                             ' Field명(1) :"ITEM_NM"

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetFromItemCd(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtfromItemCd.focus

End Function

Function SetFromItemCd(ByVal arrRet)
    frm1.txtFromItemCd.value = arrRet(0)
    frm1.txtFromItemNm.value = arrRet(1)
End Function

Function OpenToItemCd()
    Dim arrRet
    Dim arrParam(5), arrField(6)
    Dim iCalledAspName

    If IsOpenPop = True Then
        IsOpenPop = False
        Exit Function
    End If

   If frm1.txtPlantCd.value= "" Then
        Call parent.DisplayMsgBox("971012","X", "공장","X")
        'Call parent.DisplayMsgBox("189220", "x", "x", "x")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    iCalledAspName = AskPRAspName("B1B11PA3")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
    arrParam(1) = frm1.txtToItemCd.value        ' Item Code
    arrParam(2) = "12!MO"                       ' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분
    arrParam(3) = ""                            ' Default Value

    arrField(0) = 1                             ' Field명(0) :"ITEM_CD"
    arrField(1) = 2                             ' Field명(1) :"ITEM_NM"

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam, arrField), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetToItemCd(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtToItemCd.focus

End Function

Function SetToItemCd(ByVal arrRet)
    frm1.txtToItemCd.value = arrRet(0)
    frm1.txtToItemNm.value = arrRet(1)
End Function

Function OpenProdOrderNo1()

    Dim arrRet
    Dim arrParam(8)
    Dim iCalledAspName

    If IsOpenPop = True Or UCase(frm1.txtProdtOrderNo1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call displaymsgbox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    iCalledAspName = AskPRAspName("P4111PA1")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = frm1.txtPlantCd.value
    arrParam(1) = ""
    arrParam(2) = ""
    arrParam(3) = "RL"
    arrParam(4) = "ST"
    arrParam(5) = Trim(frm1.txtProdtOrderNo1.value)
    arrParam(6) = ""
    arrParam(7) = ""
    arrParam(8) = ""

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        frm1.txtProdtOrderNo1.Value = arrRet(0)
    End If

    Call SetFocusToDocument("M")
    frm1.txtProdtOrderNo1.focus

End Function

Function OpenProdOrderNo2()

    Dim arrRet
    Dim arrParam(8)
    Dim iCalledAspName

    If IsOpenPop = True Or UCase(frm1.txtProdtOrderNo2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call displaymsgbox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    iCalledAspName = AskPRAspName("P4111PA1")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = frm1.txtPlantCd.value
    arrParam(1) = ""
    arrParam(2) = ""
    arrParam(3) = "RL"
    arrParam(4) = "ST"
    arrParam(5) = Trim(frm1.txtProdtOrderNo2.value)
    arrParam(6) = ""
    arrParam(7) = ""
    arrParam(8) = ""

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        frm1.txtProdtOrderNo2.Value = arrRet(0)
    End If

    Call SetFocusToDocument("M")
    frm1.txtProdtOrderNo2.focus

End Function
'------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenBsItem()
' Description : OpenItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBsItem()

	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			 
	IsOpenPop = True

	arrParam(0) = "기준품목" 
	arrParam(1) = "B_Item a, B_Item b"    
	 
	arrParam(2) = Trim(frm1.txtBsitemcd.Value)
	arrParam(3) = ""
	 
	arrParam(4) = "a.base_item_cd = b.item_cd"   
	arrParam(5) = "기준품목"   
	 
	arrField(0) = "a.base_item_cd" 
	arrField(1) = "a.item_nm" 
	    
	arrHeader(0) = "기준품목"  
	arrHeader(1) = "기준품목명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtBsItemCd.focus
		Exit Function
	Else
		frm1.txtBsitemcd.Value    = arrRet(0)  
		frm1.txtBsitemNm.Value    = arrRet(1)  
		frm1.txtBsitemcd.focus
		Set gActiveElement = document.activeElement
	End If  
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Dim flag
    flag = frm1.ebFlag.value

    If flag = "VIEW" Then
        Call BtnPreview()
    ElseIf flag = "PRINT" Then
        Call BtnPrint()
    End If
End Function

Function BtnEbrPre(ByVal flag)
    Dim i, query, query2
    query = ""
    query2 = ""

    frm1.ebFlag.value = flag

    query2 = query2 & " plant_cd = '"&frm1.hPlantCd.value&"'"
    query2 = query2 & " AND report_dt >= '"&frm1.hStartDt.value&"'"
    query2 = query2 & " AND report_dt <= '"&frm1.hEndDt.value&"'"

    For i = 1 To  frm1.vspdData.MaxRows
        frm1.vspdData.Row = i

        frm1.vspdData.Col = C_Select

        If frm1.vspdData.value = 1 Then
            If query <> "" Then
                query = query & " or "
            End If

            frm1.vspdData.Col = C_Prodt_Order_No
            query = query & " (prodt_order_no = '" & frm1.vspdData.text & "'"

            frm1.vspdData.Col = C_Opr_No
            query = query & " and opr_no = '" & frm1.vspdData.text & "'"

            frm1.vspdData.Col = C_Seq
            query = query & " and seq = " & frm1.vspdData.text & ") "
        End If
    Next

    If query <> "" Then
        query = "UPDATE p_production_results SET cur_cd = 'Y', updt_user_id = '" & parent.gUsrID & "' , updt_dt = getdate() WHERE " & query

        query2 = " UPDATE p_production_results SET cur_cd = null WHERE " & query2

        With frm1
           .txtMode.value        = parent.UID_M0002
           .txtUpdtUserId.value  = parent.gUsrID
           .txtInsrtUserId.value = parent.gUsrID
           .txtSpread.value      = ""
           .txtQuery1.value      = query
           .txtQuery2.value      = query2
        End With

        Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    Else
        Call displaymsgbox("181216","X", "X","X")
    End If
End Function

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()

    Dim strEbrFile
    Dim objName

    Dim var1
    Dim var2
    Dim var3
    Dim var4
    Dim var5
    dim var6
    dim var7
    Dim var8
    Dim var9

    dim strUrl
    dim arrParam, arrField, arrHeader

    Call BtnDisabled(1)

    If Not chkfield(Document, "x") Then                         '⊙: This function check indispensable field
        Call BtnDisabled(0)
       Exit Function
    End If

    If parent.ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
        Call BtnDisabled(0)
        Exit Function
    End IF

    If frm1.txtPlantCd.value= "" Then
        frm1.txtPlantNm.value = ""
    End If

    If frm1.txtFromWcCd.value = "" Then
        frm1.txtFromWcNm.value = ""
    End If

    If frm1.txtFromItemCd.value = "" Then
        frm1.txtFromItemNm.value = ""
    End If

    If frm1.txtToItemCd.value = "" Then
        frm1.txtToItemNm.value = ""
    End If

    var1 = Trim(frm1.txtPlantCd.value)
    var2 = parent.UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
    var3 = parent.UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)

    If Trim(frm1.txtFromWcCd.value) <> "" Then
        var4 = " AND i_goods_movement_detail.wc_cd = '" & Trim(frm1.txtFromWcCd.value) & "'"
    Else
        var4 = ""
    End if

    If frm1.txtFromItemCd.value = "" Then
        var5 = "0"
    Else
        var5 = Trim(frm1.txtFromItemCd.value)
    End If
    If frm1.txtToItemCd.value = "" Then
        var6 = "zzzzzzzzzzzzzzzzzz"
    Else
        var6 = Trim(frm1.txtToItemCd.value)
    End If
    If frm1.txtProdtOrderNo1.value = "" Then
        var7 = "0"
    Else
        var7 = Trim(frm1.txtProdtOrderNo1.value)
    End If
    If frm1.txtProdtOrderNo2.value = "" Then
        var8 = "zzzzzzzzzzzzzzzzzz"
    Else
        var8 = Trim(frm1.txtProdtOrderNo2.value)
    End If

    If frm1.hRdoFlg.value <> "" Then
        var9 = " AND p_production_results.report_type = '" & frm1.hRdoFlg.value & "'"
    Else
        var9 = ""
    End If

    strUrl = strUrl & "plant_cd|" & var1
    strUrl = strUrl & "|fr_dt|" & var2
    strUrl = strUrl & "|to_dt|" & var3
    strUrl = strUrl & "|fr_item_cd|" & var5
    strUrl = strUrl & "|to_item_cd|" & var6
    strUrl = strUrl & "|fr_pn|" & var7
    strUrl = strUrl & "|to_pn|" & var8
    strUrl = strUrl & "|fr_wc|" & var4
    strUrl = strUrl & "|re_ty|" & var9

    strEbrFile = "P4422OA1_LKO391"
    objName = AskEBDocumentName(strEbrFile,"ebr")

'----------------------------------------------------------------
' Print 함수에서 호출
'----------------------------------------------------------------
    call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------

    Call BtnDisabled(0)

    frm1.btnRun(1).focus
    Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()

    Dim strEbrFile
    Dim objName

    Dim var1
    Dim var2
    Dim var3
    Dim var4
    Dim var5
    dim var6
    dim var7
    Dim var8
    Dim var9

    dim strUrl
    dim arrParam, arrField, arrHeader

    Call BtnDisabled(1)

    If Not chkfield(Document, "x") Then                         '⊙: This function check indispensable field
        Call BtnDisabled(0)
       Exit Function
    End If

    If parent.ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
        Call BtnDisabled(0)
        Exit Function
    End IF

    If frm1.txtPlantCd.value= "" Then
        frm1.txtPlantNm.value = ""
    End If

    If frm1.txtFromWcCd.value = "" Then
        frm1.txtFromWcNm.value = ""
    End If

    If frm1.txtFromItemCd.value = "" Then
        frm1.txtFromItemNm.value = ""
    End If

    If frm1.txtToItemCd.value = "" Then
        frm1.txtToItemNm.value = ""
    End If

    var1 = Trim(frm1.txtPlantCd.value)
    var2 = parent.UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
    var3 = parent.UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)

    If Trim(frm1.txtFromWcCd.value) <> "" Then
        var4 = " AND i_goods_movement_detail.wc_cd = '" & Trim(frm1.txtFromWcCd.value) & "' "
    Else
        var4 = ""
    End if


    If frm1.txtFromItemCd.value = "" Then
        var5 = "0"
    Else
        var5 = Trim(frm1.txtFromItemCd.value)
    End If
    If frm1.txtToItemCd.value = "" Then
        var6 = "zzzzzzzzzzzzzzzzzz"
    Else
        var6 = Trim(frm1.txtToItemCd.value)
    End If
    If frm1.txtProdtOrderNo1.value = "" Then
        var7 = "0"
    Else
        var7 = Trim(frm1.txtProdtOrderNo1.value)
    End If
    If frm1.txtProdtOrderNo2.value = "" Then
        var8 = "zzzzzzzzzzzzzzzzzz"
    Else
        var8 = Trim(frm1.txtProdtOrderNo2.value)
    End If

    If frm1.hRdoFlg.value <> "" Then
        var9 = " AND  p_production_results.report_type  = '" & frm1.hRdoFlg.value & "' "
    Else
        var9 = ""
    End If

    strUrl = strUrl & "plant_cd|" & var1
    strUrl = strUrl & "|fr_dt|" & var2
    strUrl = strUrl & "|to_dt|" & var3
    strUrl = strUrl & "|fr_item_cd|" & var5
    strUrl = strUrl & "|to_item_cd|" & var6
    strUrl = strUrl & "|fr_pn|" & var7
    strUrl = strUrl & "|to_pn|" & var8
    strUrl = strUrl & "|fr_wc|" & var4
    strUrl = strUrl & "|re_ty|" & var9

    strEbrFile = "P4422OA1_LKO391"
    objName = AskEBDocumentName(strEbrFile,"ebr")

    call FncEBRPreview(objName, strUrl)

    Call BtnDisabled(0)

    frm1.btnRun(0).focus
    Set gActiveElement = document.activeElement

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
    <TR>
        <TD <%=HEIGHT_TYPE_00%>></TD>
    </TR>

    <TR HEIGHT=23>
        <TD WIDTH="100%">
            <TABLE <%=LR_SPACE_TYPE_10%>>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>

    <TR HEIGHT=*>
        <TD CLASS="Tab11">
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD <%=HEIGHT_TYPE_02%>></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20>
                        <FIELDSET CLASS="CLSFLD">
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>공장</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="x4" ALT="공장명"></TD>
                                    <TD CLASS=TD5 NOWRAP>생산실적일</TD>
                                    <TD CLASS=TD6 NOWRAP>
                                        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtStartDt CLASSID=<%=gCLSIDFPDT%> tag="12xxxU" ALT="시작일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
                                                        &nbsp;~&nbsp;
                                        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEndDt CLASSID=<%=gCLSIDFPDT%> tag="12xxxU" ALT="종료일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>작업장</TD>
                                    <TD CLASS="TD6" NOWRAP>
                                        <INPUT TYPE=TEXT NAME="txtFromWcCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromWcNm" SIZE=20 MAXLENGTH=20 tag="x4" ALT="작업장명">
                                    </TD>
                                    <TD CLASS="TD5" NOWRAP>양불구분</TD>
                                    <TD CLASS="TD6" NOWRAP>
                                        <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg1" CLASS="RADIO" tag="11" CHECKED><LABEL FOR="rdoFlg1">전체</LABEL>
                                        <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg2" CLASS="RADIO" tag="11"><LABEL FOR="rdoFlg2">양품</LABEL>
                                        <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg3" CLASS="RADIO" tag="11"><LABEL FOR="rdoFlg3">불량</LABEL></TD>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>품목</TD>
                                    <TD CLASS="TD6" NOWRAP colspan="3">
                                        <INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=20 MAXLENGTH=20 tag="x4" ALT="품목명">
                                        &nbsp;~&nbsp;
                                        <INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=20 MAXLENGTH=20 tag="x4" ALT="품목명">
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>제조오더번호</TD>
                                    <TD CLASS="TD6" NOWRAP>
                                        <INPUT TYPE=TEXT NAME="txtProdtOrderNo1" SIZE=20 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo1">
                                        &nbsp;~&nbsp;
                                        <INPUT TYPE=TEXT NAME="txtProdtOrderNo2" SIZE=20 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo2">
                                    </TD>
                                    <TD CLASS="TD5" NOWRAP>기준품목</TD>
                                    <TD CLASS="TD6" NOWRAP>
                                        <INPUT TYPE=TEXT NAME="txtBsItemcd" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="기준품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBsItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtBsitemNm" SIZE=20 MAXLENGTH=20 tag="x4" ALT="품목명">
                                    </TD>
                                </TR>
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>

                <TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>

                <TR>
                    <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_20%> >
                            <TR>
                                <TD HEIGHT="100%">
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
                                </TD>
                            </TR>
                        </TABLE>
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
                     <TD WIDTH = 10 > &nbsp; </TD>
                     <TD>
                       <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnEbrPre('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnEbrPre('PRINT')" Flag=1>인쇄</BUTTON>
                     </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="hPlantCd"       tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt"       tag="24">
<INPUT TYPE=HIDDEN NAME="hEndDt"         tag="24">
<INPUT TYPE=HIDDEN NAME="hFromWcCd"      tag="24">
<INPUT TYPE=HIDDEN NAME="hFromItemCd"    tag="24">
<INPUT TYPE=HIDDEN NAME="hToItemCd"      tag="24">
<INPUT TYPE=HIDDEN NAME="hProdtOrderNo1" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdtOrderNo2" tag="24">
<INPUT TYPE=HIDDEN NAME="hRdoFlg"        tag="24">
<INPUT TYPE=HIDDEN NAME="hBsItemCd"        tag="24">

<INPUT TYPE=HIDDEN NAME="ebFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtQuery1" tag="24">
<INPUT TYPE=HIDDEN NAME="txtQuery2" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
    <input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
