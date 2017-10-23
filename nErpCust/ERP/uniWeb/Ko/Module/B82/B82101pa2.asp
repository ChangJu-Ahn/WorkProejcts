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

Const BIZ_PGM_ID = "B82101pb2.asp"           <% '☆: 비지니스 로직 ASP명 %>
Const C_SHEETMAXROWS_D = 100				'Sheet Max Rows

Dim C_ITEM_CD              '품목코드 
Dim C_ITEM_NM              '품목명 
Dim C_ITEM_SPEC            '규격 
Dim C_ITEM_ACCT            '품목계정 
Dim C_ITEM_ACCT_NM         '품목계정명 
Dim C_ITEM_KIND            '품목구분 
Dim C_ITEM_KIND_NM         '품목구분명 
Dim C_ITEM_LVL1            '대분류 
Dim C_ITEM_LVL1_NM         '대분류명 
Dim C_ITEM_LVL2            '중분류 
Dim C_ITEM_LVL2_NM         '중분류명 
Dim C_ITEM_LVL3            '소분류 
Dim C_ITEM_LVL3_NM         '소분류명 
Dim C_ITEM_SEQNO           'Serial No
Dim C_ITEM_VER             '이슈 
Dim C_ITEM_VER_NM          '이슈 
Dim C_ITEM_NM2             '보조품명 
Dim C_ITEM_SPEC2           '상세규격 
Dim C_ITEM_UNIT            '품목단위 
Dim C_ITEM_GRADE           '품목등급 
Dim C_PUR_TYPE             '조달구분 
Dim C_PUR_TYPE_NM          '조달구분명 
Dim C_BASIC_CODE           '기준품목 
Dim C_BASIC_CODE_NM        '기준품목 
Dim C_PUR_GROUP            '구매그룹 
Dim C_PUR_GROUP_NM         '구매그룹명 
Dim C_PUR_VENDOR           '납품처 
Dim C_PUR_VENDOR_NM        '납품처명 
Dim C_UNIFY_PUR_FLAG       '통합구매구분 
Dim C_UNIT_WEIGHT          'Net중량 
Dim C_UNIT_OF_WEIGHT       '중량단위 
Dim C_GROSS_WEIGHT         'Gross중량 
Dim C_GROSS_UNIT           '중량단위 
Dim C_CBM                  'CBM(부피)
Dim C_CBM_DESCRIPTION      'CBM정보 
Dim C_HS_CODE              'HS코드 
Dim C_HS_CODE_NM           'HS코드 
Dim C_VALID_FROM_DT        'FROM유효일자 
Dim C_VALID_TO_DT          'TO유효일자 
Dim INTERNAL_CD            '내부코드 
Dim C_DOC_NO               '도면번호 
DIM C_R
DIM C_T
DIM C_P
DIM C_Q


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
    C_ITEM_CD              =  1
    C_ITEM_NM              =  2
    C_ITEM_SPEC            =  3
    C_ITEM_ACCT            =  4
    C_ITEM_ACCT_NM         =  5
    C_ITEM_KIND            =  6
    C_ITEM_KIND_NM         =  7
    C_ITEM_LVL1            =  8
    C_ITEM_LVL1_NM         =  9
    C_ITEM_LVL2            = 10
    C_ITEM_LVL2_NM         = 11
    C_ITEM_LVL3            = 12
    C_ITEM_LVL3_NM         = 13
    C_ITEM_SEQNO           = 14
    C_ITEM_VER             = 15
    C_ITEM_VER_NM          = 16
    C_ITEM_NM2             = 17
    C_ITEM_SPEC2           = 18
    C_ITEM_UNIT            = 19
    C_PUR_TYPE             = 20
    C_PUR_TYPE_NM          = 21
    C_BASIC_CODE           = 22
    C_BASIC_CODE_NM        = 23
    C_PUR_GROUP            = 24
    C_PUR_GROUP_NM         = 25
    C_PUR_VENDOR           = 26
    C_PUR_VENDOR_NM        = 27
    C_UNIFY_PUR_FLAG       = 28
    C_UNIT_WEIGHT          = 29
    C_UNIT_OF_WEIGHT       = 30
    C_GROSS_WEIGHT         = 31
    C_GROSS_UNIT           = 32
    C_CBM                  = 33
    C_CBM_DESCRIPTION      = 34
    C_HS_CODE              = 35
    C_HS_CODE_NM           = 36
    C_VALID_FROM_DT        = 37
    C_VALID_TO_DT          = 38
    C_DOC_NO               = 39
    INTERNAL_CD            = 40
    C_R						= 41
    C_T						= 42
    C_P						= 43
    C_Q						= 44
    
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
     frm1.txtNewChange.value = arrParam(0)
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
         
     frm1.vspdData.OperationMode = 3
         
     frm1.vspdData.ReDraw = False
                 
     frm1.vspdData.MaxCols = C_Q + 1
     frm1.vspdData.MaxRows = 0
     
     Call GetSpreadColumnPos("A")     
     
     ggoSpread.SSSetEdit   C_ITEM_CD,            "품목",         12
     ggoSpread.SSSetEdit   C_ITEM_NM,            "품목명",       18
     ggoSpread.SSSetEdit   C_ITEM_SPEC,          "규격",         18
     ggoSpread.SSSetEdit   C_ITEM_ACCT,          "품목계정",      7
     ggoSpread.SSSetEdit   C_ITEM_ACCT_NM,       "품목계정명",   10     
     ggoSpread.SSSetEdit   C_ITEM_KIND,          "품목구분",     10
     ggoSpread.SSSetEdit   C_ITEM_KIND_NM,       "품목구분명",   10
     ggoSpread.SSSetEdit   C_ITEM_LVL1,          "대분류",        7
     ggoSpread.SSSetEdit   C_ITEM_LVL1_NM,       "대분류명",     10
     ggoSpread.SSSetEdit   C_ITEM_LVL2,          "중분류",        7    
     ggoSpread.SSSetEdit   C_ITEM_LVL2_NM,       "중분류명",     10
     ggoSpread.SSSetEdit   C_ITEM_LVL3,          "소분류",        7
     ggoSpread.SSSetEdit   C_ITEM_LVL3_NM,       "소분류명",     10 
     ggoSpread.SSSetEdit   C_ITEM_SEQNO,         "Serial No",        10
     ggoSpread.SSSetEdit   C_ITEM_VER,           "이슈부여",      5
     ggoSpread.SSSetEdit   C_ITEM_VER_NM,        "이슈부여",      5     
     ggoSpread.SSSetEdit   C_ITEM_NM2,           "보조품명",     20
     ggoSpread.SSSetEdit   C_ITEM_SPEC2,         "상세규격",     20
     
     
     
     ggoSpread.SSSetEdit   C_ITEM_UNIT,          "단위",          8          
     ggoSpread.SSSetEdit   C_PUR_TYPE,           "조달구분",     10     
     ggoSpread.SSSetEdit   C_PUR_TYPE_NM,        "조달구분",     12
     ggoSpread.SSSetEdit   C_BASIC_CODE,         "기준품목",     15
     ggoSpread.SSSetEdit   C_BASIC_CODE_NM,      "기준품목명",   20
     ggoSpread.SSSetEdit   C_PUR_GROUP,          "구매그룹",     10
     ggoSpread.SSSetEdit   C_PUR_GROUP_NM,       "구매그룹명",   15     
     ggoSpread.SSSetEdit   C_PUR_VENDOR,         "공급처",       10
     ggoSpread.SSSetEdit   C_PUR_VENDOR_NM,      "공급처명",     15
     ggoSpread.SSSetEdit   C_UNIFY_PUR_FLAG,     "통합구매",     10
     ggoSpread.SSSetFloat  C_UNIT_WEIGHT,        "Net중량",      10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
     ggoSpread.SSSetEdit   C_UNIT_OF_WEIGHT,     "Net단위",      10     
     ggoSpread.SSSetFloat  C_GROSS_WEIGHT,       "Gross중량",    15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
     ggoSpread.SSSetEdit   C_GROSS_UNIT,         "Gross단위",    10
     ggoSpread.SSSetFloat  C_CBM,                "CBM(부피)",    15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
     ggoSpread.SSSetEdit   C_CBM_DESCRIPTION,    "CBM정보",      20
     ggoSpread.SSSetEdit   C_HS_CODE,            "HS코드",       10
     ggoSpread.SSSetEdit   C_HS_CODE_NM,         "HS코드명",     15
     ggoSpread.SSSetDate   C_VALID_FROM_DT,      "시작일자",     10, 2, PopupParent.gDateFormat
     ggoSpread.SSSetDate   C_VALID_TO_DT,        "종료일자",     10, 2, PopupParent.gDateFormat
     ggoSpread.SSSetEdit   C_DOC_NO,             "도면번호",     20
     ggoSpread.SSSetEdit   INTERNAL_CD,          "내부코드",     20
     
     ggoSpread.SSSetEdit   C_R,					 "접수검토",         1 
     ggoSpread.SSSetEdit   C_T,					"기술검토",          1 
     ggoSpread.SSSetEdit   C_P,					"구매검토",          1 
     ggoSpread.SSSetEdit   C_Q,					"품질검토",          1 
  
     
     
                 
     Call ggoSpread.SSSetColHidden(C_ITEM_ACCT, C_ITEM_ACCT_NM,        True)
     Call ggoSpread.SSSetColHidden(C_ITEM_KIND, C_ITEM_KIND,           True)
     Call ggoSpread.SSSetColHidden(C_ITEM_KIND_NM, C_ITEM_KIND_NM,     True)
     Call ggoSpread.SSSetColHidden(C_ITEM_LVL1, C_ITEM_LVL1,           True)
     Call ggoSpread.SSSetColHidden(C_ITEM_LVL2, C_ITEM_LVL2,           True)
     Call ggoSpread.SSSetColHidden(C_ITEM_LVL3, C_ITEM_LVL3,           True)
     Call ggoSpread.SSSetColHidden(C_ITEM_VER,  C_ITEM_VER,            True)
     Call ggoSpread.SSSetColHidden(C_PUR_TYPE,  C_PUR_TYPE,            True)
     Call ggoSpread.SSSetColHidden(C_ITEM_VER_NM,  C_ITEM_VER_NM,      True)
     Call ggoSpread.SSSetColHidden(C_UNIFY_PUR_FLAG, C_UNIFY_PUR_FLAG, True)
     
     Call ggoSpread.SSSetColHidden(C_R, C_R, True)
     Call ggoSpread.SSSetColHidden(C_T, C_T, True) 
     Call ggoSpread.SSSetColHidden(C_P, C_P, True) 
     Call ggoSpread.SSSetColHidden(C_Q, C_Q, True) 
       
         
     Call ggoSpread.SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)

     ggoSpread.SSSetSplit2(2)                                                  'frozen 기능추가 

     frm1.vspdData.ReDraw = True
     
     Call SetSpreadLock()
     
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
Sub SetSpreadLock()
     ggoSpread.Source = frm1.vspdData
     ggoSpread.SpreadLockWithOddEvenRowColor()
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
            
            C_ITEM_CD              = iCurColumnPos(1)
			C_ITEM_NM              = iCurColumnPos(2)
			C_ITEM_SPEC            = iCurColumnPos(3)
			C_ITEM_ACCT            = iCurColumnPos(4)
			C_ITEM_ACCT_NM         = iCurColumnPos(5)
			C_ITEM_KIND            = iCurColumnPos(6)
			C_ITEM_KIND_NM         = iCurColumnPos(7)
			C_ITEM_LVL1            = iCurColumnPos(8)
			C_ITEM_LVL1_NM         = iCurColumnPos(9)
			C_ITEM_LVL2            = iCurColumnPos(10)
			C_ITEM_LVL2_NM         = iCurColumnPos(11)
			C_ITEM_LVL3            = iCurColumnPos(12)
			C_ITEM_LVL3_NM         = iCurColumnPos(13)
			C_ITEM_SEQNO           = iCurColumnPos(14)
			C_ITEM_VER             = iCurColumnPos(15)
			C_ITEM_VER_NM          = iCurColumnPos(16)
			C_ITEM_NM2             = iCurColumnPos(17)
			C_ITEM_SPEC2           = iCurColumnPos(18)
			C_ITEM_UNIT            = iCurColumnPos(19)
			C_PUR_TYPE             = iCurColumnPos(20)
			C_PUR_TYPE_NM          = iCurColumnPos(21)
			C_BASIC_CODE           = iCurColumnPos(22)
			C_BASIC_CODE_NM        = iCurColumnPos(23)
			C_PUR_GROUP            = iCurColumnPos(24)
			C_PUR_GROUP_NM         = iCurColumnPos(25)
			C_PUR_VENDOR           = iCurColumnPos(26)
			C_PUR_VENDOR_NM        = iCurColumnPos(27)
			C_UNIFY_PUR_FLAG       = iCurColumnPos(28)
			C_UNIT_WEIGHT          = iCurColumnPos(29)
			C_UNIT_OF_WEIGHT       = iCurColumnPos(30)
			C_GROSS_WEIGHT         = iCurColumnPos(31)
			C_GROSS_UNIT           = iCurColumnPos(32)
			C_CBM                  = iCurColumnPos(33)
			C_CBM_DESCRIPTION      = iCurColumnPos(34)
			C_HS_CODE              = iCurColumnPos(35)
			C_HS_CODE_NM           = iCurColumnPos(36)
			C_VALID_FROM_DT        = iCurColumnPos(37)
			C_VALID_TO_DT          = iCurColumnPos(38)
			C_DOC_NO               = iCurColumnPos(39)
			INTERNAL_CD			   = iCurColumnPos(40)
			
			C_R					   = iCurColumnPos(41)
			C_T					   = iCurColumnPos(42)
			C_P					   = iCurColumnPos(43)
			C_Q					   = iCurColumnPos(44)
			
			
			
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

     If frm1.vspdData.MaxRows > 0 Then
          If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
             Call OKClick
          End If
     End If
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

'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)

        Dim arrRet
        Dim arrParam(7), arrField(8), arrHeader(8)
        Dim sItemAcct , sItemKind, sItemLvl1, sItemLvl2, sItemLvl3

        If IsOpenPop = True  Then  
           Exit Function
        End If   


        IsOpenPop = True
        Select Case arPopUp
               Case 1 '품목구분 
                                   
                    arrParam(0) = frm1.txtItemKind.Alt
                    arrParam(1) = "B_MINOR A, B_CIS_CONFIG B "
                    arrParam(2) = Trim(frm1.txtItemKind.value)
                    arrParam(4) = "A.MINOR_CD *= B.ITEM_KIND AND MAJOR_CD = 'Y1001' AND B.ITEM_ACCT like "&filtervar(frm1.cboitemacct.value&"%","''","S")
                    
                     
                    arrParam(5) = frm1.txtItemKind.Alt

                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"
    
                    arrHeader(0) = frm1.txtItemKind.Alt
                    arrHeader(1) = frm1.txtItemKind_Nm.Alt  
                    frm1.txtItemKind.focus ()
               Case 2 '대분류                    
                    
                    sItemAcct = Trim(frm1.cboItemAcct.value)
                    If sItemAcct = "" Then sItemAcct = "%"
                    
                    sItemKind = Trim(frm1.txtItemKind.value)
                    If sItemKind = "" Then sItemKind = "%"
                    
                    arrParam(0) = frm1.txtItemLvl1.Alt
                    arrParam(1) = "B_CIS_ITEM_CLASS"
                    arrParam(2) = Trim(frm1.txtItemLvl1.value)
                    arrParam(4) = "ITEM_ACCT like '" & filtervar(sItemAcct,"","SNM")  & "' AND ITEM_KIND like '" & filtervar(sItemKind,"","SNM") & "' AND ITEM_LVL = 'L1' "
                    arrParam(5) = frm1.txtItemLvl1.Alt

                    arrField(0) = "CLASS_CD"
                    arrField(1) = "CLASS_NAME"
    
                    arrHeader(0) = frm1.txtItemLvl1.Alt
                    arrHeader(1) = frm1.txtItemLvl1_Nm.Alt
                     frm1.txtItemLvl1.focus()
               Case 3 '중분류 
                                        
                    sItemLvl1 = Trim(frm1.txtItemLvl1.value)
                    If sItemLvl1 = "" Then sItemLvl1 = "%"
                    sItemAcct = Trim(frm1.cboItemAcct.value)&"%"
                    sItemKind = Trim(frm1.txtItemKind.value)&"%"
                    
                    arrParam(0) = frm1.txtItemLvl2.Alt
                    arrParam(1) = "B_CIS_ITEM_CLASS"
                    arrParam(2) = Trim(frm1.txtItemLvl2.value)
                    //arrParam(4) = "ITEM_ACCT like '" & filtervar(sItemAcct,"","SNM")  & "' AND ITEM_KIND like '" & filtervar(sItemKind,"","SNM") & "' AND ITEM_LVL = 'L2' AND PARENT_CLASS_CD like '" & sItemLvl1 & "' "
                    arrParam(4) = "ITEM_ACCT like '" & filtervar(sItemAcct,"","SNM")  & "' AND ITEM_KIND like '" & filtervar(sItemKind,"","SNM") & "' AND ITEM_LVL = 'L2' AND PARENT_CLASS_CD like '" & sItemLvl1 & "' "
                    arrParam(5) = frm1.txtItemLvl2.Alt

                    arrField(0) = "CLASS_CD"
                    arrField(1) = "CLASS_NAME"
    
                    arrHeader(0) = frm1.txtItemLvl2.Alt
                    arrHeader(1) = frm1.txtItemLvl2_Nm.Alt
                    frm1.txtItemLvl2.focus()
               Case 4 '소분류 
                                        
                    sItemLvl2 = Trim(frm1.txtItemLvl2.value)
                    If sItemLvl2 = "" Then     sItemLvl2 = "%"
                    
                    sItemAcct = Trim(frm1.cboItemAcct.value)&"%"
                    sItemKind = Trim(frm1.txtItemKind.value)&"%"
                    
                    arrParam(0) = frm1.txtItemLvl3.Alt
                    arrParam(1) = "B_CIS_ITEM_CLASS"
                    arrParam(2) = Trim(frm1.txtItemLvl3.value)
                    arrParam(4) = "ITEM_ACCT like '" & filtervar(sItemAcct,"","SNM")  & "' AND ITEM_KIND like '" & filtervar(sItemKind,"","SNM") & "' AND ITEM_LVL = 'L3' AND PARENT_CLASS_CD like '" & sItemLvl2 & "' "
                    arrParam(5) = frm1.txtItemLvl3.Alt

                    arrField(0) = "CLASS_CD"
                    arrField(1) = "CLASS_NAME"
    
                    arrHeader(0) = frm1.txtItemLvl3.Alt
                    arrHeader(1) = frm1.txtItemLvl3_Nm.Alt   
                    frm1.txtItemLvl3.focus()           
               
               Case Else
                    Exit Function
      End Select
        
      arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

      IsOpenPop = False
                
      If arrRet(0) = "" Then
         Exit Function
      Else
         Call SetConPopup(arrRet,arPopUp)
      End If        
        
End Function

'======================================================================================================
Function SetConPopup(Byval arrRet,ByVal arPopUp)

     SetConPopup = False

     Select Case arPopUp
     Case 1
          frm1.txtItemKind.value   = arrRet(0) 
          frm1.txtItemKind_Nm.value = arrRet(1)   
     Case 2
          frm1.txtItemLvl1.value   = arrRet(0) 
          frm1.txtItemLvl1_Nm.value = arrRet(1)   
     Case 3
          frm1.txtItemLvl2.value   = arrRet(0) 
          frm1.txtItemLvl2_Nm.value = arrRet(1)    
     Case 4
          frm1.txtItemLvl3.value   = arrRet(0) 
          frm1.txtItemLvl3_Nm.value = arrRet(1)  
     End Select

     SetConPopup = True

End Function
'========================================================================================================
'     Name : OKClick()
'     Desc : handle ok icon click event
'========================================================================================================
Function OKClick()
     Dim i , iCurColumnPos
     
     If frm1.vspdData.MaxRows > 0 Then
          
          Redim arrReturn(UBound(arrField))
			
          ggoSpread.Source = frm1.vspdData
          Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
          frm1.vspdData.Row = frm1.vspdData.ActiveRow 
          
          For i = 0 To UBound(arrField)
              frm1.vspddata.Col = iCurColumnPos(i + 1)
              arrReturn(i)      = frm1.vspdData.Text
          Next
         
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

'=======================================================================================================
'   Event Name : txtDtFr_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDtFr_DblClick(Button)
    If Button = 1 Then
        txtDtFr.Action = 7
        Call SetFocusToDocument("P")
        txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDtTo_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtDtFr_KeyDown(keycode, shift)
     If keycode = 13 Then
          Call FncQuery()
     End If
End Sub

'=======================================================================================================
'   Event Name : txtDtTo_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDtTo_DblClick(Button)
    If Button = 1 Then
        txtDtTo.Action = 7
        Call SetFocusToDocument("P")
        txtDtTo.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDtTo_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtDtTo_KeyDown(keycode, shift)
     If keycode = 13 Then
          Call FncQuery()
     End If
End Sub

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
     frm1.txtItemCd.focus()  
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
    
    lgKeyStream =               Trim(frm1.txtItemCd.value)      & Chr(11)  
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemNm.value)      & Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.cboItemAcct.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemKind.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemLvl1.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemLvl2.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemLvl3.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemSpec.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtNewChange.value)	& Chr(11)
         
    DbQuery = False                                                                 '⊙: Processing is NG
     
    Call LayerShowHide(1)                                                           '⊙: 작업진행중 표시 
    
    strVal = BIZ_PGM_ID & "?txtMode="             & PopupParent.UID_M0001           '☜: Query
    strVal = strVal     & "&txtKeyStream="        & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="         & ""                              '☜: Direction
    strVal = strVal     & "&lgStrPrevKeyIndex="   & lgStrPrevKeyIndex               '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="          & Frm1.vspdData.MaxRows           '☜: Max fetched data
    strVal = strVal     & "&lgMaxCount="          & CStr(C_SHEETMAXROWS_D)          '☜: Max fetched data at a time
  
    Call RunMyBizASP(MyBizASP, strVal)                                              '☜: 비지니스 ASP 를 가동 
     
    DbQuery = True                                                                  '⊙: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
     If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
          Call SetActiveCell(frm1.vspdData,1,1,"P","X","X")
          Set gActiveElement = document.activeElement
     End If
    lgIntFlgMode = PopupParent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")                                             '⊙: This function lock the suitable field
	frm1.txtItemCd.focus()  
End Function
  
'========================================================================================
' Function Name : txtItem_kind_OnChange
' Function Desc : 
'========================================================================================
Function txtItemkind_OnChange()
    Dim iDx
    Dim IntRetCd
	
    If frm1.txtItemkind.value = "" Then
        frm1.txtItemkind_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='Y1001' and minor_cd="&filterVar(frm1.txtItemkind.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItemkind_nm.value=""
        Else
            frm1.txtItemkind_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

  
'========================================================================================================
Sub txtItemLvl1_OnChange()

    Dim IntRetCD
    Dim sItemAcct , sItemKind
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    sItemAcct = Trim(frm1.cboItemAcct.value)

                    
    sItemKind = Trim(frm1.txtItemKind.value)

    If Trim(frm1.txtItemLvl1.value) = "" Then
       frm1.txtItemLvl1_Nm.Value = ""
    Else
       IntRetCD = CommonQueryRs("CLASS_NAME","B_CIS_ITEM_CLASS","ITEM_ACCT like  '" & sItemAcct  &"%"&  "' AND ITEM_KIND ='" & sItemKind & "' AND ITEM_LVL = 'L1' AND CLASS_CD = '" & TRIM(frm1.txtItemLvl1.value) & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false Then
                //frm1.txtItemLvl1.value = ""
                frm1.txtItemLvl1_Nm.value = ""
       Else   
         frm1.txtItemLvl1_Nm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If
 
End Sub

Sub txtItemLvl2_OnChange()

    Dim IntRetCD
    Dim sItemAcct , sItemKind
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    sItemAcct = Trim(frm1.cboItemAcct.value)

                    
    sItemKind = Trim(frm1.txtItemKind.value)

    If Trim(frm1.txtItemLvl2.value) = "" Then
       frm1.txtItemLvl2_Nm.Value = ""
    Else
       IntRetCD = CommonQueryRs("CLASS_NAME","B_CIS_ITEM_CLASS","ITEM_ACCT like  '" & sItemAcct  &"%"&  "' AND ITEM_KIND ='" & sItemKind & "' AND ITEM_LVL = 'L2' AND PARENT_CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl1.value),"''","S") & " AND CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl2.value),"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false Then
                frm1.txtItemLvl2_Nm.value = ""
       Else   
         frm1.txtItemLvl2_Nm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If
 
End Sub

Sub txtItemLvl3_OnChange()

    Dim IntRetCD
    Dim sItemAcct , sItemKind
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    sItemAcct = Trim(frm1.cboItemAcct.value)

                    
    sItemKind = Trim(frm1.txtItemKind.value)

    If Trim(frm1.txtItemLvl3.value) = "" Then
       frm1.txtItemLvl3_Nm.Value = ""
    Else
       IntRetCD = CommonQueryRs("CLASS_NAME","B_CIS_ITEM_CLASS","ITEM_ACCT like  '" & sItemAcct  &"%"&  "' AND ITEM_KIND ='" & sItemKind & "' AND ITEM_LVL = 'L3' AND PARENT_CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl2.value),"''","S") & " AND CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl3.value),"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false Then
                frm1.txtItemLvl3_Nm.value = ""
       Else   
         frm1.txtItemLvl3_Nm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If
 
End Sub

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
						   <TD CLASS=TD5 NOWRAP>품목코드</TD>
						   <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 tag="11XXXU" ALT="품목코드"></TD>
						   <TD CLASS=TD5 NOWRAP>품목명</TD>
						   <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=100 tag="11" ALT="품목명"></TD>
						</TR>
						<TR>
						   <TD CLASS=TD5 NOWRAP>품목계정</TD>
						   <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="11" ALT="품목계정"><OPTION VALUE=""></OPTION></SELECT></TD>
						   <TD CLASS=TD5 NOWRAP>품목구분</TD>
						   <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemKind" ALT="품목구분" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('1')">
						                        <INPUT NAME="txtItemKind_Nm" ALT="품목구분명" TYPE="Text" SiZE=25   tag="14XXXU"></TD>
						</TR>     
						<TR>
						    <TD CLASS=TD5 NOWRAP>대분류</TD>
						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl1" ALT="대분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('2')">
						                         <INPUT NAME="txtItemLvl1_Nm" ALT="대분류명" TYPE="Text" SiZE=25   tag="14XXXU"></TD>
						    <TD CLASS=TD5 NOWRAP>중분류</TD>
						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl2" ALT="중분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('3')">
						                         <INPUT NAME="txtItemLvl2_Nm" ALT="중분류명" TYPE="Text" SiZE=25   tag="14XXXU"></TD>                                        
						</TR>
						<TR>
						    <TD CLASS=TD5 NOWRAP>소분류</TD>
						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl3" ALT="소분류" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('4')">
						                         <INPUT NAME="txtItemLvl3_Nm" ALT="소분류명" TYPE="Text" SiZE=25   tag="14XXXU"></TD>
						    <TD CLASS=TD5 NOWRAP>규격</TD>
						    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=50 tag="11" ALT="규격">&nbsp;</TD>
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
                              <script language =javascript src='./js/b82101pa2_vaSpread1_vspdData.js'></script>
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
                                                   <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
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
<INPUT TYPE=HIDDEN NAME="txtNewChange"    TAG="24" TABINDEX="-1">

</FORM>
</BODY>
</HTML>