<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : PRERECEIPT
'*  3. Program ID           : A2107MA1
'*  4. Program Name         : 분개형태 등록 
'*  5. Program Desc         : 분개형태 등록 수정 삭제 조회 
'*  6. Component List       : PD1G035
'*  7. Modified date(First) : 2000/09/30
'*  8. Modified date(Last)  : 2003/06/17
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit 

'==========================================================================================
'                                               1.2 Global 변수/상수 선언  
'==========================================================================================

<!-- #Include file="../../inc/lgvariables.inc" --> 

'@PGM_ID

Const BIZ_PGM_QRY_ID1 = "a2107mb1.asp"
Const BIZ_PGM_QRY_ID2 = "a2107mb1.asp"
Const BIZ_PGM_QRY_ID6 = "a2107mb6.asp"


'==========================================================================================

Dim lgStrPrevKeyOne_Seq 
Dim lgStrPrevKeyTwo_CtrlCd
Dim lgStrPrevKeyThree_CtrlCd

Dim intItemCnt
Dim IsOpenPop
Dim gSelframeFlg

Dim lgBlnStartFlag

Dim lgCurrRow1
Dim lgCurrRow2

'==========================================================================================
'Spreadsheet #1

Dim C_JnlCD
Dim C_JnlPopUp
Dim C_JnlNM
Dim C_FormSeq
Dim C_DrCrFgCd
Dim C_DrCrFgNm
Dim C_EventCD
Dim C_EventPopUp
Dim C_EventNm
Dim C_AcctCD
Dim C_AcctPopUp
Dim C_AcctNm
Dim C_MaxColPlus1
Dim C_MaxColPlus2
Dim C_MaxColPlus3

'Spreadsheet #2

Dim C_CtrlCtrlCD
Dim C_CtrlCtrlNm
Dim C_CtrlTransType 
Dim C_CtrlJnlCD 
Dim C_CtrlFormSeq
Dim C_CtrlCtrlCnt
Dim C_CtrlDrCrFgCd
Dim C_CtrlAcctCD
Dim C_CtrlTblId

Dim C_CtrlTblPopUp
Dim C_CtrlDataColmId

Dim C_CtrlColmPopUp1
Dim C_CtrlDataTypeCd
Dim C_CtrlDataTypeNm
Dim C_CtrlKeyColmId1

Dim C_CtrlColmPopUp2
Dim C_CtrlKeyDataType1Cd
Dim C_CtrlKeyDataType1Nm
Dim C_CtrlKeyColmId2

Dim C_CtrlColmPopUp3
Dim C_CtrlKeyDataType2Cd
Dim C_CtrlKeyDataType2Nm
Dim C_CtrlKeyColmId3

Dim C_CtrlColmPopUp4
Dim C_CtrlKeyDataType3Cd
Dim C_CtrlKeyDataType3Nm
Dim C_CtrlKeyColmId4

Dim C_CtrlColmPopUp5
Dim C_CtrlKeyDataType4Cd
Dim C_CtrlKeyDataType4Nm
Dim C_CtrlKeyColmId5

Dim C_CtrlColmPopUp6
Dim C_CtrlKeyDataType5Cd
Dim C_CtrlKeyDataType5Nm

Dim C_CtrlMaxColPlus1 
Dim C_Spread2ColorFg
'==========================================================================================
Sub initSpreadPosVariables()
'Spreadsheet #1

 C_JnlCD       = 1
 C_JnlPopUp    = 2
 C_JnlNM       = 3
 C_FormSeq     = 4
 C_DrCrFgCd    = 5
 C_DrCrFgNm    = 6
 C_EventCD     = 7
 C_EventPopUp  = 8
 C_EventNm     = 9
 C_AcctCD      = 10
 C_AcctPopUp   = 11
 C_AcctNm      = 12
 C_MaxColPlus1 = 13
 C_MaxColPlus2 = 14
 C_MaxColPlus3 = 15


'Spreadsheet #2
 C_CtrlCtrlCD      = 1
 C_CtrlCtrlNm      = 2
 C_CtrlTransType   = 3
 C_CtrlJnlCD       = 4
 C_CtrlFormSeq     = 5
 C_CtrlCtrlCnt     = 6
 C_CtrlDrCrFgCd    = 7
 C_CtrlAcctCD      = 8
 C_CtrlTblId       = 9

 C_CtrlTblPopUp    = 10
 C_CtrlDataColmId  = 11

 C_CtrlColmPopUp1  = 12
 C_CtrlDataTypeCd  = 13
 C_CtrlDataTypeNm  = 14
 C_CtrlKeyColmId1  = 15

 C_CtrlColmPopUp2      = 16
 C_CtrlKeyDataType1Cd  = 17
 C_CtrlKeyDataType1Nm  = 18
 C_CtrlKeyColmId2      = 19

 C_CtrlColmPopUp3      = 20
 C_CtrlKeyDataType2Cd  = 21
 C_CtrlKeyDataType2Nm  = 22
 C_CtrlKeyColmId3      = 23

 C_CtrlColmPopUp4      = 24
 C_CtrlKeyDataType3Cd  = 25
 C_CtrlKeyDataType3Nm  = 26
 C_CtrlKeyColmId4      = 27

 C_CtrlColmPopUp5      = 28
 C_CtrlKeyDataType4Cd  = 29
 C_CtrlKeyDataType4Nm  = 30
 C_CtrlKeyColmId5      = 31

 C_CtrlColmPopUp6      = 32
 C_CtrlKeyDataType5Cd  = 33
 C_CtrlKeyDataType5Nm  = 34

 C_CtrlMaxColPlus1     = 35
 C_Spread2ColorFg      = 36 '20030617

End Sub


'======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKeyOne_Seq = ""

    lgStrPrevKeyTwo_CtrlCd = ""
    lgStrPrevKeyThree_CtrlCd = ""

    lgCurrRow1 = 0
    lgCurrRow2 = 0

    lgSortKey = 1
End Sub


'======================================================================================================
Sub SetDefaultVal()
End Sub


'======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'======================================================================================================
Sub InitSpreadSheet()

    Dim sList
	Call initSpreadPosVariables()

    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021218",,Parent.gAllowDragDropSpread

        .MaxCols = C_MaxColPlus3 + 1
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetEdit      C_JnlCD,    "거래항목코드",  15,  , ,20, 2
        ggoSpread.SSSetButton    C_JnlPopUp
        ggoSpread.SSSetEdit      C_JnlNM,    "거래항목명",30, , , 50
        Call AppendNumberPlace("6","3","0")
        ggoSpread.SSSetFloat     C_FormSeq,  "순번",      5 ,"6" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetCombo     C_DrCrFgCd, "차대구분",  12
        ggoSpread.SSSetCombo     C_DrCrFgNm, "차대구분",  12
        Call InitComboBox("0", "A1012")
        ggoSpread.SSSetEdit      C_EventCd,    "관련거래항목코드",  15,  , ,20, 2
        ggoSpread.SSSetButton    C_EventPopUp
        ggoSpread.SSSetEdit      C_EventNm, "관련거래항목명",30, , , 50
        ggoSpread.SSSetEdit      C_AcctCD,         "계정코드",        25, , , 20, 2
        ggoSpread.SSSetButton    C_AcctPopUp
        ggoSpread.SSSetEdit      C_AcctNm,       "계정명",          30, , , 30
        ggoSpread.SSSetEdit      C_MaxColPlus1,       "",          5, , , 30
        ggoSpread.SSSetEdit      C_MaxColPlus2,       "",          5, , , 30
        ggoSpread.SSSetEdit      C_MaxColPlus3,       "",          5, , , 30

        Call ggoSpread.MakePairsColumn(C_JnlCD,C_JnlPopUp)
        Call ggoSpread.MakePairsColumn(C_DrCrFgCd,C_DrCrFgNm,"1")
        Call ggoSpread.MakePairsColumn(C_EventCd,C_EventPopUp)
        Call ggoSpread.MakePairsColumn(C_AcctCD,C_AcctPopUp)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        Call ggoSpread.SSSetColHidden(C_MaxColPlus1,C_MaxColPlus1,True)
        Call ggoSpread.SSSetColHidden(C_MaxColPlus1,C_MaxColPlus1,True)
        Call ggoSpread.SSSetColHidden(C_MaxColPlus2,C_MaxColPlus2,True)
        Call ggoSpread.SSSetColHidden(C_MaxColPlus3,C_MaxColPlus3,True)
        Call ggoSpread.SSSetColHidden(C_FormSeq,C_FormSeq,True)
        Call ggoSpread.SSSetColHidden(C_DrCrFgCd,C_DrCrFgCd,True)
	End With
	Call SetSpreadLock("Q", 0, 1, "")
End Sub

Sub InitSpreadSheet1()

    With frm1.vspdData2

        ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021130",,parent.gAllowDragDropSpread
        .MaxCols = C_Spread2ColorFg + 1
        .MaxRows = 0

        Call GetSpreadColumnPos("B")
        ggoSpread.SSSetEdit      C_CtrlCtrlCD,      "관리항목코드", 10, , , 3, 2
        ggoSpread.SSSetEdit      C_CtrlCtrlNm,      "관리항목명", 15, , , 30
        ggoSpread.SSSetEdit      C_CtrlTransType,   "거래유형", 5, , , 20, 2
        ggoSpread.SSSetEdit      C_CtrlJnlCD,       "거래항목", 5, , , 50, 2
        ggoSpread.SSSetFloat     C_CtrlFormSeq,    "순번", 3, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat     C_CtrlCtrlCnt,    "순번", 3, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetEdit      C_CtrlDrCrFgCd,    "차대구분", 5, , , 12
        ggoSpread.SSSetEdit      C_CtrlAcctCD,      "계정과목", 5, , , 20, 2
        ggoSpread.SSSetEdit      C_CtrlTblId,       "테이블ID", 15, , , 50, 2

        ggoSpread.SSSetButton    C_CtrlTblPopUp
        ggoSpread.SSSetEdit      C_CtrlDataColmId,  "컬럼ID", 15, , , 255, 2

        ggoSpread.SSSetButton    C_CtrlColmPopUp1
        ggoSpread.SSSetCombo     C_CtrlDataTypeCd, "자료유형", 12
        ggoSpread.SSSetCombo     C_CtrlDataTypeNm, "자료유형", 12
        ggoSpread.SSSetEdit      C_CtrlKeyColmId1, "KEY컬럼ID1", 15, , , 50, 2

        ggoSpread.SSSetButton    C_CtrlColmPopUp2
        ggoSpread.SSSetCombo     C_CtrlKeyDataType1Cd, "자료유형1", 12
        ggoSpread.SSSetCombo     C_CtrlKeyDataType1Nm, "자료유형1", 12

        ggoSpread.SSSetEdit      C_CtrlKeyColmId2, "KEY컬럼ID2", 15, , , 50, 2

        ggoSpread.SSSetButton    C_CtrlColmPopUp3
        ggoSpread.SSSetCombo     C_CtrlKeyDataType2Cd, "자료유형2", 12
        ggoSpread.SSSetCombo     C_CtrlKeyDataType2Nm, "자료유형2", 12

        ggoSpread.SSSetEdit      C_CtrlKeyColmId3, "KEY컬럼ID3", 15, , , 50, 2

        ggoSpread.SSSetButton    C_CtrlColmPopUp4
        ggoSpread.SSSetCombo     C_CtrlKeyDataType3Cd, "자료유형3", 12
        ggoSpread.SSSetCombo     C_CtrlKeyDataType3Nm, "자료유형3", 12


        ggoSpread.SSSetEdit      C_CtrlKeyColmId4, "Key컬럼ID4", 15, , , 50, 2

        ggoSpread.SSSetButton    C_CtrlColmPopUp5
        ggoSpread.SSSetCombo     C_CtrlKeyDataType4Cd, "자료유형4", 12
        ggoSpread.SSSetCombo     C_CtrlKeyDataType4Nm, "자료유형4", 12


        ggoSpread.SSSetEdit      C_CtrlKeyColmId5, "Key컬럼ID5", 15, , , 50, 2

        ggoSpread.SSSetButton    C_CtrlColmPopUp6
        ggoSpread.SSSetCombo     C_CtrlKeyDataType5Cd, "자료유형5", 12
        ggoSpread.SSSetCombo     C_CtrlKeyDataType5Nm, "자료유형5", 12
        ggoSpread.SSSetEdit      C_CtrlMaxColPlus1,       "순번",5
        ggoSpread.SSSetEdit      C_Spread2ColorFg,       " ",5

        Call InitComboBox("1", "A1018")


        Call ggoSpread.MakePairsColumn(C_CtrlCtrlCD,C_CtrlCtrlNm,"1")
        Call ggoSpread.MakePairsColumn(C_CtrlTransType,C_CtrlCtrlNm,"1")
        Call ggoSpread.MakePairsColumn(C_CtrlTblId,C_CtrlTblPopUp)
        Call ggoSpread.MakePairsColumn(C_CtrlDataColmId,C_CtrlColmPopUp1)

        Call ggoSpread.MakePairsColumn(C_CtrlDataTypeCd,C_CtrlDataTypeNm,"1")
        Call ggoSpread.MakePairsColumn(C_CtrlKeyColmId1,C_CtrlColmPopUp2)
        Call ggoSpread.MakePairsColumn(C_CtrlKeyDataType1Cd,C_CtrlKeyDataType1Nm,"1")

        Call ggoSpread.MakePairsColumn(C_CtrlKeyColmId2,C_CtrlColmPopUp3)
        Call ggoSpread.MakePairsColumn(C_CtrlKeyDataType2Cd,C_CtrlKeyDataType2Nm,"1")

        Call ggoSpread.MakePairsColumn(C_CtrlKeyColmId3,C_CtrlColmPopUp4)
        Call ggoSpread.MakePairsColumn(C_CtrlKeyDataType3Cd,C_CtrlKeyDataType3Nm,"1")

        Call ggoSpread.MakePairsColumn(C_CtrlKeyColmId4,C_CtrlColmPopUp5)
        Call ggoSpread.MakePairsColumn(C_CtrlKeyDataType4Cd,C_CtrlKeyDataType4Nm,"1")

        Call ggoSpread.MakePairsColumn(C_CtrlKeyColmId5,C_CtrlColmPopUp6)
        Call ggoSpread.MakePairsColumn(C_CtrlKeyDataType5Cd,C_CtrlKeyDataType5Nm,"1")

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        Call ggoSpread.SSSetColHidden(C_CtrlMaxColPlus1,C_CtrlMaxColPlus1,True)
        Call ggoSpread.SSSetColHidden(C_CtrlTransType,C_CtrlTransType,True)
        Call ggoSpread.SSSetColHidden(C_CtrlJnlCD,C_CtrlJnlCD,True)
        Call ggoSpread.SSSetColHidden(C_CtrlFormSeq,C_CtrlFormSeq,True)
        Call ggoSpread.SSSetColHidden(C_CtrlCtrlCnt,C_CtrlCtrlCnt,True)
        Call ggoSpread.SSSetColHidden(C_CtrlDrCrFgCd,C_CtrlDrCrFgCd,True)
        Call ggoSpread.SSSetColHidden(C_CtrlAcctCD,C_CtrlAcctCD,True)
        Call ggoSpread.SSSetColHidden(C_CtrlDataTypeCd,C_CtrlDataTypeCd,True)
        Call ggoSpread.SSSetColHidden(C_CtrlKeyDataType1Cd,C_CtrlKeyDataType1Cd,True)
        Call ggoSpread.SSSetColHidden(C_CtrlKeyDataType2Cd,C_CtrlKeyDataType2Cd,True)
        Call ggoSpread.SSSetColHidden(C_CtrlKeyDataType3Cd,C_CtrlKeyDataType3Cd,True)
        Call ggoSpread.SSSetColHidden(C_CtrlKeyDataType4Cd,C_CtrlKeyDataType4Cd,True)
        Call ggoSpread.SSSetColHidden(C_CtrlKeyDataType5Cd,C_CtrlKeyDataType5Cd,True)
        Call ggoSpread.SSSetColHidden(C_Spread2ColorFg,C_Spread2ColorFg,True)

        Call SetSpreadLock("I", 1, 1, "")
    End With

    intItemCnt = 0
End Sub

Sub InitSpreadSheet2()

    With frm1.vspdData3

        ggoSpread.Source = frm1.vspdData3
        ggoSpread.Spreadinit "V20021130",,parent.gAllowDragDropSpread
        .MaxCols = C_Spread2ColorFg + 1
        .MaxRows = 0

		ggoSpread.SSSetEdit      C_CtrlDataColmId,  "", 15, , , 255, 2

    End With

End Sub

'==========================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2 )
Dim objSpread
Dim Cnt
Dim indx
    With frm1

    Select Case Index
    Case 0
        ggoSpread.Source = .vspdData
        Set objSpread = .vspdData
    Case 1
        ggoSpread.Source = .vspdData2
        Set objSpread = .vspdData2
    End Select

    If lRow2 = "" Then lRow2 = objSpread.MaxRows

        objSpread.Redraw = False

    Select Case stsFg
         Case "Q"
            Select Case Index
            Case 0

				ggoSpread.SpreadLock C_FormSeq        , -1, C_FormSeq
				ggoSpread.SpreadLock C_JnlCd          , -1, C_JnlCd
				ggoSpread.SpreadLock C_JnlPopUp       , -1, C_JnlPopUp
				ggoSpread.SpreadLock C_JnlNM          , -1, C_JnlNM
				ggoSpread.SpreadLock C_DrCrFgCd       , -1, C_DrCrFgCd
				ggoSpread.SpreadLock C_DrCrFgNm       , -1, C_DrCrFgNm
				ggoSpread.SpreadLock C_EventNm        , -1, C_EventNm
				ggoSpread.SpreadLock C_AcctNm         , -1, C_AcctNm
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			Case 1
				'20030617
				For indx = 1 to frm1.vspdData2.MaxRows
					.vspdData2.Row = indx
					.vspdData2.Col = C_Spread2ColorFg
						If UCase(Trim(.vspddata2.Text)) = "Y" Then
						ggoSpread.SSSetRequired      C_CtrlTblId   , indx, indx
						ggoSpread.SSSetRequired      C_CtrlDataColmId   , indx, indx
						End If
				Next
				ggoSpread.SpreadLock C_CtrlFormSeq    , -1, C_CtrlFormSeq
				ggoSpread.SpreadLock C_CtrlCtrlCnt    , -1, C_CtrlCtrlCnt
				ggoSpread.SpreadLock C_CtrlTransType  , -1, C_CtrlTransType
				ggoSpread.SpreadLock C_CtrlJnlCD      , -1, C_CtrlJnlCD
				ggoSpread.SpreadLock C_CtrlDrCrFgCd   , -1, C_CtrlDrCrFgCd
				ggoSpread.SpreadLock C_CtrlAcctCD     , -1, C_CtrlAcctCD
				ggoSpread.SpreadLock C_CtrlCtrlCD     , -1, C_CtrlCtrlCD
				ggoSpread.SpreadLock C_CtrlCtrlNm     , -1, C_CtrlCtrlNm
				ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1


            End Select
        Case "I"
            Select Case Index
            Case 0
				ggoSpread.SSSetRequired      C_JnlCD         , -1, C_JnlCD
				ggoSpread.SpreadUnLock      C_JnlPopUp      , -1, C_JnlPopUp
				ggoSpread.SpreadLock      C_DrCrFgNm      , -1, C_DrCrFgNm
				ggoSpread.SpreadLock      C_FormSeq       , -1, C_FormSeq
				ggoSpread.SpreadLock      C_JnlNM         , -1, C_JnlNM
				ggoSpread.SpreadLock      C_EventNm       , -1, C_EventNm
				ggoSpread.SpreadLock      C_AcctNm        , -1, C_AcctNm
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1

			Case 1
				ggoSpread.SpreadLock      C_CtrlFormSeq   , -1, C_CtrlFormSeq
				ggoSpread.SpreadLock      C_CtrlCtrlCnt   , -1, C_CtrlCtrlCnt
				ggoSpread.SpreadLock      C_CtrlTransType , -1, C_CtrlTransType
				ggoSpread.SpreadLock      C_CtrlJnlCD     , -1, C_CtrlJnlCD
				ggoSpread.SpreadLock      C_CtrlDrCrFgCd  , -1, C_CtrlDrCrFgCd
				ggoSpread.SpreadLock      C_CtrlAcctCD    , -1, C_CtrlAcctCD
				ggoSpread.SpreadLock      C_CtrlCtrlCD    , -1, C_CtrlCtrlCD
				ggoSpread.SpreadLock      C_CtrlCtrlNm    , -1, C_CtrlCtrlNm
				ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1
            End Select
        End Select
        objSpread.Redraw = True
        Set objSpread = Nothing
    End With    
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal lRow,ByVal lRow2)
  
    With frm1.vspdData 
        .Redraw = False

        ggoSpread.Source = frm1.vspdData
        ggoSpread.SSSetRequired  C_JnlCD,     lRow, lRow2
        ggoSpread.SSSetProtected C_JnlNM,     lRow, lRow2
        ggoSpread.SSSetRequired  C_DrCrFgNm,  lRow, lRow2
        ggoSpread.SSSetProtected C_EventNm,   lRow, lRow2
        ggoSpread.SSSetProtected C_AcctNm,    lRow, lRow2

        .Redraw = True

    End With
End Sub
'======================================================================================================
Sub SetSpreadColor2(ByVal lRow,ByVal lRow2)
    
    With frm1.vspdData2
        .Redraw = False
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.SSSetProtected C_CtrlCtrlNm, lRow, lRow2
        .Col = 1
        .Row = .ActiveRow
        .Action = 0
        .EditMode = True
        .Redraw = True
    End With
End Sub


'======================================================================================================
Function InitComboBox(Byval Index, Byval MajorCd)

	Dim iCodeArr,iNameArr

	Err.Clear
	On Error Resume Next

 Select Case Index
  Case "0"
   Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1012", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

   iCodeArr = lgF0
   iNameArr = lgF1

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DrCrFgCd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DrCrFgNm

  Case "1"
   Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1018", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

   iCodeArr = vbtab & lgF0
   iNameArr = vbtab & lgF1

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CtrlDataTypeCd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CtrlDataTypeNm

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CtrlKeyDataType1Cd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CtrlKeyDataType1Nm

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CtrlKeyDataType2Cd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CtrlKeyDataType2Nm

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CtrlKeyDataType3Cd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CtrlKeyDataType3Nm

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CtrlKeyDataType4Cd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CtrlKeyDataType4Nm

   ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CtrlKeyDataType5Cd
   ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CtrlKeyDataType5Nm

 End Select

End Function


'======================================================================================================
' Name : OpenPopUp()
' Description : PopUp
'======================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
Dim strSelect, strWhere, strFrom
 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 Select Case iWhere
  Case 0
   arrParam(0) = "거래유형 팝업"				' 팝업 명칭 
   arrParam(1) = "A_ACCT_TRANS_TYPE"			' TABLE 명칭 
   arrParam(2) = strCode						' Code Condition
   arrParam(3) = ""								' Name Cindition
   arrParam(4) = ""								' Where Condition
   arrParam(5) = "거래유형"					' 조건필드의 라벨 명칭 

   arrField(0) = "TRANS_TYPE"					' Field명(0)
   arrField(1) = "TRANS_NM"						' Field명(1)

   arrHeader(0) = "거래유형코드"				' Header명(0)
   arrHeader(1) = "거래유형명"				' Header명(1)
  Case 1
   arrParam(0) = "거래항목 팝업"
   arrParam(1) = "A_JNL_ITEM"
   arrParam(2) = strCode
   arrParam(3) = ""
   arrParam(4) = ""
   arrParam(5) = "거래항목"

   arrField(0) = "JNL_CD"
   arrField(1) = "JNL_NM"

   arrHeader(0) = "거래항목코드"
   arrHeader(1) = "거래항목명"
  Case 2
   arrParam(0) = "계정 팝업"
   arrParam(1) = "A_ACCT"
   arrParam(2) = strCode
   arrParam(3) = ""
   arrParam(4) = ""
   arrParam(5) = "계정"

   arrField(0) = "ACCT_CD"
   arrField(1) = "ACCT_NM"

   arrHeader(0) = "계정코드"
   arrHeader(1) = "계정명"
  Case 3
   arrParam(0) = "관련거래항목 팝업"
   arrParam(1) = "A_JNL_ITEM"
   arrParam(2) = strCode
   arrParam(3) = ""
   arrParam(4) = ""
   arrParam(5) = "관련거래항목"

   arrField(0) = "JNL_CD"
   arrField(1) = "JNL_NM"

   arrHeader(0) = "관련거래항목코드"
   arrHeader(1) = "관련거래항목명"

  Case 4


   frm1.vspdData2.Col = C_CtrlTransType
   frm1.vspdData2.Row = frm1.vspdData2.ActiveRow

   strSelect = " Trans_Type"
   strFrom =  " A_JNL_TABLE"
   strWhere = " TRANS_TYPE = " & FilterVar(UCase(frm1.vspdData2.text), "''", "S")

   If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
    arrParam(0) = "테이블명 팝업"
    arrParam(1) = "A_JNL_TABLE"
    arrParam(2) = strCode 
    arrParam(3) = ""
    arrParam(4) = "Trans_Type = " & FilterVar(UCase(frm1.vspdData2.text), "''", "S")
    arrParam(5) = "테이블"

    arrField(0) = "TBL_ID"
    arrHeader(0) = "테이블명"
   Else
    arrParam(0) = "테이블명 팝업"
    arrParam(1) = "SYSOBJECTS"
    arrParam(2) = strCode
    arrParam(3) = ""
    arrParam(4) = "(XTYPE = " & FilterVar("U", "''", "S") & "  OR XTYPE = " & FilterVar("V", "''", "S") & " ) "
    arrParam(5) = "테이블"

    arrField(0) = "NAME"
    arrHeader(0) = "테이블명"
   End If 

  Case 5, 6, 7, 8, 9, 10
   arrParam(0) = "필드명 팝업"
   arrParam(1) = "SYSCOLUMNS A, SYSTYPES B"
   arrParam(2) = strCode
   arrParam(3) = ""

   frm1.vspdData2.Col = C_CtrlTblId
   frm1.vspdData2.row = frm1.vspdData2.ActiveRow
   arrParam(4) = "A.XTYPE = B.XTYPE AND A.ID = (SELECT id FROM SYSOBJECTS  WHERE (XTYPE = " & FilterVar("U", "''", "S") & "  OR XTYPE = " & FilterVar("V", "''", "S") & " )  and name = " & FilterVar(UCase(frm1.vspdData2.text), "''", "S") & ")"   ' Where Condition
   arrParam(5) = "필드명"


   arrField(0) = "UPPER(A.NAME)"
   arrField(1) = "UPPER(B.NAME)"

   arrHeader(0) = "필드명"
   arrHeader(1) = "필드타입"


 End Select

 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
	Select Case iWhere
	Case 0
		frm1.txtTransType.focus
	Case Else
	End Select

  Exit Function
 Else
  Call SetPopUp(arrRet, iWhere)
 End If

End Function

'======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
 dim Row
 Dim colPos

    With frm1
        Row = .vspdData.ActiveRow
        Select Case iWhere
            Case 0
                .txtTransType.focus
                .txtTransType.value = arrRet(0)
                .hTransType.value = arrRet(0)
                .txtTransNM.value = arrRet(1)
            Case 1
                .vspdData.Col = C_JnlCD
                .vspdData.Text = arrRet(0)
                .vspdData.Col = C_JnlNM
                .vspdData.Text = arrRet(1)
                lgBlnFlgChgValue = True
            Case 2
                .vspdData.Col = C_AcctCD
                .vspdData.Text = arrRet(0)
                .vspdData.Col = C_AcctNm
                .vspdData.Text = arrRet(1)

                Call vspdData_Change(C_AcctCD, Row)
                lgBlnFlgChgValue = True
            Case 3
                .vspdData.Col = C_EventCd
                .vspdData.Text = arrRet(0)
                .vspdData.Col = C_EventNm
                .vspdData.Text = arrRet(1)

                Call vspdData_Change(C_EventCD, Row)
                lgBlnFlgChgValue = True
            Case 4
                .vspdData2.Col = C_CtrlTblId
                .vspdData2.Text = arrRet(0)

                Row = .vspdData2.ActiveRow
                Call vspdData2_Change(C_CtrlTblId, Row)
                lgBlnFlgChgValue = True
            Case 5
                colPos = .vspdData2.ActiveCol
                .vspdData2.Col = C_CtrlDataColmId
                .vspdData2.Text = arrRet(0)
                .vspdData2.Col = C_CtrlDataTypeCd
                '// 해당부분 찾아 선택하는 부분이 따로 필요함 
                if instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 THEN

                    .vspdData2.value = "2"
                    .vspdData2.Col = C_CtrlDataTypeNm
                    .vspdData2.value = "2"

                Elseif instr(1,UCase(arrRet(1)),"DATE")>0 THEN
                    .vspdData2.value = "1"
                    .vspdData2.Col = C_CtrlDataTypeNm
                    .vspdData2.value = "1"  

                Elseif instr(1,UCase(arrRet(1)),"CHAR")>0 THEN
                    .vspdData2.value = "3"
                    .vspdData2.Col = C_CtrlDataTypeNm
                    .vspdData2.value = "3"  

                End If
                Row = .vspdData2.Row
                Call vspdData2_Change(colPos, Row)
                lgBlnFlgChgValue = True
            Case 6
                colPos = .vspdData2.ActiveCol
                .vspdData2.Col = C_CtrlKeyColmId1
                .vspdData2.Text = arrRet(0)
                .vspdData2.Col = C_CtrlKeyDataType1Cd

                if instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 THEN

                     .vspdData2.value = "2"
                     .vspdData2.Col = C_CtrlKeyDataType1Nm
                     .vspdData2.value = "2"

                Elseif instr(1,UCase(arrRet(1)),"DATE")>0 THEN
                     .vspdData2.value = "1"
                     .vspdData2.Col = C_CtrlKeyDataType1Nm
                     .vspdData2.value = "1"  

                Elseif instr(1,UCase(arrRet(1)),"CHAR")>0 THEN
                     .vspdData2.value = "3"
                     .vspdData2.Col = C_CtrlKeyDataType1Nm
                     .vspdData2.value = "3"  

                End If
                Row = .vspdData2.Row
                Call vspdData2_Change(colPos, Row)
                lgBlnFlgChgValue = True
            Case 7
                colPos = .vspdData2.ActiveCol
                .vspdData2.Col = C_CtrlKeyColmId2
                .vspdData2.Text = arrRet(0)
                .vspdData2.Col = C_CtrlKeyDataType2Cd

                if instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 THEN

                     .vspdData2.value = "2"
                     .vspdData2.Col = C_CtrlKeyDataType2Nm
                     .vspdData2.value = "2"

                Elseif instr(1,UCase(arrRet(1)),"DATE")>0 THEN
                     .vspdData2.value = "1"
                     .vspdData2.Col = C_CtrlKeyDataType2Nm
                     .vspdData2.value = "1"

                Elseif instr(1,UCase(arrRet(1)),"CHAR")>0 THEN
                     .vspdData2.value = "3"
                     .vspdData2.Col = C_CtrlKeyDataType2Nm
                     .vspdData2.value = "3"

                End If
                Row = .vspdData2.Row
                Call vspdData2_Change(colPos, Row)
                lgBlnFlgChgValue = True

            Case 8
                colPos = .vspdData2.ActiveCol
                .vspdData2.Col = C_CtrlKeyColmId3
                .vspdData2.Text = arrRet(0)
                .vspdData2.Col = C_CtrlKeyDataType3Cd

                if instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 THEN

                     .vspdData2.value = "2"
                     .vspdData2.Col = C_CtrlKeyDataType3Nm
                     .vspdData2.value = "2"

                Elseif instr(1,UCase(arrRet(1)),"DATE")>0 THEN
                     .vspdData2.value = "1"
                     .vspdData2.Col = C_CtrlKeyDataType3Nm
                     .vspdData2.value = "1"

                Elseif instr(1,UCase(arrRet(1)),"CHAR")>0 THEN
                     .vspdData2.value = "3"
                     .vspdData2.Col = C_CtrlKeyDataType3Nm
                     .vspdData2.value = "3"

                End If
                Row = .vspdData2.Row
                Call vspdData2_Change(colPos, Row)
                lgBlnFlgChgValue = True

            Case 9
                colPos = .vspdData2.ActiveCol
                .vspdData2.Col = C_CtrlKeyColmId4
                .vspdData2.Text = arrRet(0)
                .vspdData2.Col = C_CtrlKeyDataType4Cd

                if instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 THEN

                     .vspdData2.value = "2"
                     .vspdData2.Col = C_CtrlKeyDataType4Nm
                     .vspdData2.value = "2"

                Elseif instr(1,UCase(arrRet(1)),"DATE")>0 THEN
                     .vspdData2.value = "1"
                     .vspdData2.Col = C_CtrlKeyDataType4Nm
                     .vspdData2.value = "1"

                Elseif instr(1,UCase(arrRet(1)),"CHAR")>0 THEN
                     .vspdData2.value = "3"
                     .vspdData2.Col = C_CtrlKeyDataType4Nm
                     .vspdData2.value = "3"

                End If
                Row = .vspdData2.Row
                Call vspdData2_Change(colPos, Row)
                lgBlnFlgChgValue = True

            Case 10
                colPos = .vspdData2.ActiveCol
                .vspdData2.Col = C_CtrlKeyColmId5
                .vspdData2.Text = arrRet(0)
                .vspdData2.Col = C_CtrlKeyDataType5Cd

                if instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 THEN

                     .vspdData2.value = "2"
                     .vspdData2.Col = C_CtrlKeyDataType5Nm
                     .vspdData2.value = "2"

                Elseif instr(1,UCase(arrRet(1)),"DATE")>0 THEN
                     .vspdData2.value = "1"
                     .vspdData2.Col = C_CtrlKeyDataType5Nm
                     .vspdData2.value = "1"

                Elseif instr(1,UCase(arrRet(1)),"CHAR")>0 THEN
                     .vspdData2.value = "3"
                     .vspdData2.Col = C_CtrlKeyDataType5Nm
                     .vspdData2.value = "3"

                End If
                Row = .vspdData2.Row

                Call vspdData2_Change(colPos, Row)
                lgBlnFlgChgValue = True

        End Select

    End With
End Function



'=======================================================================================================
'   Function Name : FindNumber
'   Function Desc : 
'=======================================================================================================
Function FindNumber(ByVal objSpread, ByVal intCol)
Dim lngRows
Dim lngPrevNum
Dim lngNextNum

    FindNumber = 0

    lngPrevNum = 0
    lngNextNum = 0

    With frm1

        If objSpread.MaxRows = 0 Then
            Exit Function
        End If

        For lngRows = 1 To objSpread.MaxRows
            objSpread.Row = lngRows
            objSpread.Col = intCol
            lngNextNum = Clng(objSpread.Text)

            If lngNextNum > lngPrevNum Then
                lngPrevNum = lngNextNum
            End If

        Next

    End With

    FindNumber = lngPrevNum

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(Byval pRow)
Dim lRows
Dim CtrlFormCnt2, CtrlCtrlCnt2
Dim CtrlFormCnt3, CtrlCtrlCnt3

 ' 히든Grid에서 Data를 찾는다.

    lgCurrRow2 = 0

    With frm1

  .vspdData2.Row = pRow
        .vspdData2.Col = C_CtrlFormSeq: CtrlFormCnt2 = .vspdData2.Text
        .vspdData2.Col = C_CtrlCtrlCnt: CtrlCtrlCnt2 = .vspdData2.Text

        For lRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lRows
            .vspdData3.Col = 5 'C_CtrlFormSeq
            CtrlFormCnt3 = .vspdData3.Text

            .vspdData3.Row = lRows
            .vspdData3.Col = 6 'C_CtrlCtrlCnt
            CtrlCtrlCnt3 = .vspdData3.Text

            If (CtrlFormCnt2 = CtrlFormCnt3) And (CtrlCtrlCnt2 = CtrlCtrlCnt3) Then
                lgCurrRow2 = lRows
                Exit Function
            End If
        Next
    End With
End Function

'=======================================================================================================
'   Function Name : CopyFromData
'   Function Desc : 
'=======================================================================================================
Function CopyFromData(ByVal pFormCnt)
    Dim lngRows
    Dim boolExist
    Dim iCols
    Dim Cnt
    Dim CtrlFormCnt, CtrlCtrlCnt

    Dim lRows
    Dim CtrlFormCnt2, CtrlCtrlCnt2
    Dim CtrlFormCnt3, CtrlCtrlCnt3

    ' 히든Grid에서 하단Grid로 자료를 복사한다.

    boolExist = False
    CopyFromData = boolExist
    With frm1
        Call SortHSheet()
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows 
            .vspdData3.Col = 5'C_CtrlFormSeq
            CtrlFormCnt = .vspdData3.Text

            If (pFormCnt = CtrlFormCnt) Then
                boolExist = True
                Exit For
            End If
        Next

        '------------------------------------
        ' Show Data
        '------------------------------------ 
        .vspdData3.Row = lngRows

        If boolExist = True Then 

            .vspdData2.Redraw = False

            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows 
                .vspdData3.Col = 5 'C_CtrlFormSeq 
                CtrlFormCnt = .vspdData3.Text

                If (pFormCnt <> CtrlFormCnt) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Col = 5'C_CtrlFormSeq 
                    CtrlFormCnt3 = .vspdData3.Text
                    .vspdData3.Col = 6 'C_CtrlCtrlCnt 
                    CtrlCtrlCnt3 = .vspdData3.Text

                    For lRows = 1 To .vspdData2.MaxRows
                        .vspdData2.Row = lRows 
                        .vspdData2.Col = C_CtrlFormSeq 
                        CtrlFormCnt2 = .vspdData2.Text
                        .vspdData2.Row = lRows 
                        .vspdData2.Col = C_CtrlCtrlCnt 
                        CtrlCtrlCnt2 = .vspdData2.Text

                        If (CtrlFormCnt2 = CtrlFormCnt3) And (CtrlCtrlCnt2 = CtrlCtrlCnt3) Then

                             .vspdData2.Col = 0                     '0
                            .vspdData3.Col  = 0
                            .vspdData2.Text = .vspdData3.Text

                             .vspdData2.Col = C_CtrlCtrlCD           '1
                            .vspdData3.Col  = 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlCtrlNm           '2
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlTransType        '3
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlJnlCD            '4
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlFormSeq          '5
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlCtrlCnt          '6
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlDrCrFgCd         '7
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlAcctCD           '8
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlTblId            '9
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlTblPopUp         '10
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlDataColmId       '11
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col  = C_CtrlColmPopUp1       '12
                            .vspdData3.Col  = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlDataTypeCd       '13
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlDataTypeNm       '14
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyColmId1       '15
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlColmPopUp2       '16
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType1Cd   '17
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType1Nm   '18
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyColmId2       '19
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlColmPopUp3       '20
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType2Cd   '21
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType2Nm   '22
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyColmId3       '23
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlColmPopUp4       '24
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType3Cd   '25
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType3Nm   '26
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyColmId4       '27
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlColmPopUp5       '28
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType4Cd   '29
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType4Nm   '30
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyColmId5       '31
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlColmPopUp6       '32
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType5Cd   '33
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlKeyDataType5Nm   '34
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_CtrlMaxColPlus1      '35
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text

                            .vspdData2.Col = C_Spread2ColorFg      '36
                            .vspdData3.Col = .vspdData3.Col + 1
                            .vspdData2.Text = .vspdData3.Text
   
                        End If
                    Next
                End If
                lngRows = lngRows + 1
            Wend

            ggoSpread.Source = frm1.vspdData2

            Call InitData2

            Call SetSpreadLock("Q", 1, 1, "")

            frm1.vspdData.Row = lgCurrRow1
            frm1.vspdData.Col = frm1.vspdData.MaxCols
            ggoSpread.Source = frm1.vspdData
            frm1.vspdData2.Redraw = True

        End If

    End With
    CopyFromData = boolExist
End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols
Dim intRetCD

 ' 하단Grid에서 수정된내용을 히든Grid로 복사한다.  

    With frm1 

        Call FindData(Row)
        If lgCurrRow2 = 0 Then
            .vspdData3.MaxRows = .vspdData3.MaxRows + 1
            .vspdData3.Row = .vspdData3.MaxRows
        Else
            If lgCurrRow2 > 0 Then
                .vspdData3.Row = lgCurrRow2
            End if
        End If 
        .vspdData2.Row = Row

        .vspdData3.Col = 0                      '0
        .vspdData2.Col = 0
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlCtrlCD           '1
        .vspdData3.Col = 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlCtrlNm           '2
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlTransType        '3
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlJnlCD            '4
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlFormSeq          '5
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlCtrlCnt          '6
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlDrCrFgCd         '7
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlAcctCD           '8
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlTblId            '9
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlTblPopUp         '10
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlDataColmId       '11
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlColmPopUp1       '12
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlDataTypeCd       '13
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlDataTypeNm       '14
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyColmId1       '15
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlColmPopUp2       '16
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType1Cd   '17
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType1Nm   '18
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyColmId2       '19
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlColmPopUp3       '20
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType2Cd   '21
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType2Nm   '22
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyColmId3       '23
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlColmPopUp4       '24
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType3Cd   '25
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType3Nm   '26
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyColmId4       '27
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlColmPopUp5       '28
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType4Cd   '29
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType4Nm   '30
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyColmId5       '31
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlColmPopUp6       '32
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType5Cd   '33
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlKeyDataType5Nm   '34
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_CtrlMaxColPlus1      '35
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

        .vspdData2.Col = C_Spread2ColorFg      '36
        .vspdData3.Col = .vspdData3.Col + 1
        .vspdData3.Text = .vspdData2.Text

 End With
 
 frm1.vspdData.Row = frm1.vspdData.ActiveRow
 
 frm1.vspdData.Col = 0
 If frm1.vspdData.Text <> ggoSpread.InsertFlag and frm1.vspdData.Text <> ggoSpread.DeleteFlag and frm1.vspdData.Text <> ggoSpread.UpdateFlag then
        frm1.vspdData.Text = ggoSpread.UpdateFlag

        frm1.vspdData.Col = C_MaxColPlus2

        If frm1.vspdData.Text = "" Then
      frm1.vspdData.Col = C_MaxColPlus3 '//1
      frm1.vspdData.Text = ggoSpread.UpdateFlag 

  Else
   frm1.vspdData.Text = ""   '//
  End If
   End if

End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal pFormCnt)
Dim boolExist
Dim lngRows
Dim CtrlFormCnt

 ' 상단Grid에서 삭제된내용을 히든Grid에서도 삭제한다.

    DeleteHSheet = False
    boolExist = False

    With frm1

        Call SortHSheet()

        For lngRows = 1 To .vspdData3.MaxRows
           .vspdData3.Row = lngRows
           .vspdData3.Col = 5'C_CtrlFormSeq
           CtrlFormCnt = .vspdData3.Text

            If (pFormCnt = CtrlFormCnt) Then
                boolExist = True
                Exit For
            End If
        Next

      '------------------------------------
        ' Data Delete
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            While lngRows <= .vspdData3.MaxRows
                .vspdData3.Row = lngRows
                .vspdData3.Col = 5'C_CtrlFormSeq
                CtrlFormCnt = .vspdData3.Text

                If (pFormCnt <> CtrlFormCnt) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If
            Wend

            ggoSpread.Source = frm1.vspdData2

            frm1.vspdData.Row = lgCurrRow1
            frm1.vspdData.Col = frm1.vspdData.MaxCols
            ggoSpread.Source = frm1.vspdData

            frm1.vspdData2.Redraw = True

        End If

    End With

    DeleteHSheet = True
End Function    

'======================================================================================================
' Function Name : SortHSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortHSheet()

 ' 히든Grid에 있는 내용을 정렬한다.

 With frm1
     .vspdData3.BlockMode = True
     .vspdData3.BlockMode = False
 End With
    
End Function

'=======================================================================================================
' Function Name : ShowHidden
' Function Desc : 
'=======================================================================================================
Sub ShowHidden()
Dim strHidden
Dim strHidden2
Dim lngRows
Dim lngCols

 ' Test를 위한 내용  

    With frm1.vspdData2
        For lngRows = 1 To .MaxRows
            .Row = lngRows
            For lngCols = 0 To .MaxCols
            .Col = lngCols
                strHidden =  strHidden & " | " & .Text
            Next
            strHidden = strHidden & vbCrLf
        Next
    End With
    
    With frm1.vspdData3
        For lngRows = 1 To .MaxRows
            .Row = lngRows
            For lngCols = 0 To .MaxCols
            .Col = lngCols  
                strHidden2 =  strHidden2 & " | " & .Text
            Next
            strHidden2 = strHidden2 & vbCrLf
        Next
    End With
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            '@Grid_Column
             C_JnlCD       = iCurColumnPos(1)
             C_JnlPopUp    = iCurColumnPos(2)
             C_JnlNM       = iCurColumnPos(3)
             C_FormSeq     = iCurColumnPos(4)
             C_DrCrFgCd    = iCurColumnPos(5)
             C_DrCrFgNm    = iCurColumnPos(6)
             C_EventCD     = iCurColumnPos(7)
             C_EventPopUp  = iCurColumnPos(8)
             C_EventNm     = iCurColumnPos(9)
             C_AcctCD      = iCurColumnPos(10)
             C_AcctPopUp   = iCurColumnPos(11)
             C_AcctNm      = iCurColumnPos(12)
             C_MaxColPlus1 = iCurColumnPos(13)
             C_MaxColPlus2 = iCurColumnPos(14)
             C_MaxColPlus3 = iCurColumnPos(15)

       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

             C_CtrlCtrlCD      = iCurColumnPos(1)
             C_CtrlCtrlNm      = iCurColumnPos(2)
             C_CtrlTransType   = iCurColumnPos(3)
             C_CtrlJnlCD       = iCurColumnPos(4)
             C_CtrlFormSeq     = iCurColumnPos(5)
             C_CtrlCtrlCnt     = iCurColumnPos(6)
             C_CtrlDrCrFgCd    = iCurColumnPos(7)
             C_CtrlAcctCD      = iCurColumnPos(8)
             C_CtrlTblId       = iCurColumnPos(9)

             C_CtrlTblPopUp    = iCurColumnPos(10)
             C_CtrlDataColmId  = iCurColumnPos(11)

             C_CtrlColmPopUp1  = iCurColumnPos(12)
             C_CtrlDataTypeCd   = iCurColumnPos(13)
             C_CtrlDataTypeNm  = iCurColumnPos(14)
             C_CtrlKeyColmId1  = iCurColumnPos(15)

             C_CtrlColmPopUp2      = iCurColumnPos(16)
             C_CtrlKeyDataType1Cd  = iCurColumnPos(17)
             C_CtrlKeyDataType1Nm  = iCurColumnPos(18)
             C_CtrlKeyColmId2      = iCurColumnPos(19)

             C_CtrlColmPopUp3      = iCurColumnPos(20)
             C_CtrlKeyDataType2Cd  = iCurColumnPos(21)
             C_CtrlKeyDataType2Nm  = iCurColumnPos(22)
             C_CtrlKeyColmId3      = iCurColumnPos(23)

             C_CtrlColmPopUp4      = iCurColumnPos(24)
             C_CtrlKeyDataType3Cd  = iCurColumnPos(25)
             C_CtrlKeyDataType3Nm  = iCurColumnPos(26)
             C_CtrlKeyColmId4      = iCurColumnPos(27)

             C_CtrlColmPopUp5      = iCurColumnPos(28)
             C_CtrlKeyDataType4Cd  = iCurColumnPos(29)
             C_CtrlKeyDataType4Nm  = iCurColumnPos(30)
             C_CtrlKeyColmId5      = iCurColumnPos(31)

             C_CtrlColmPopUp6      = iCurColumnPos(32)
             C_CtrlKeyDataType5Cd  = iCurColumnPos(33)
             C_CtrlKeyDataType5Nm  = iCurColumnPos(34)

             C_CtrlMaxColPlus1     = iCurColumnPos(35)
             C_Spread2ColorFg      = iCurColumnPos(36)'20030617

    End Select
End Sub



'==========================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet()
    Call InitSpreadSheet1()
    Call InitSpreadSheet2
    Call InitVariables
    Call SetDefaultVal

	 lgBlnStartFlag = False

    Call SetToolbar("1100110100001111")
    frm1.txtTransType.focus
'	frm1.txtTransType.value ="ya001"
'	fncquery
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
    frm1.vspdData.Row = Row

    Select Case Col
        Case  C_DrCrFgNm
            frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_DrCrFgCd
            frm1.vspdData.Value = intIndex
            Call vspdData_Change(C_AcctCD, Row) '20030617
    End Select
End Sub

'==========================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex

    frm1.vspdData2.Row = Row

    Select Case Col
        Case  C_CtrlDataTypeNm   ' 자료유형 
            frm1.vspdData2.Col = Col
            intIndex = frm1.vspdData2.Value
            frm1.vspdData2.Col = C_CtrlDataTypeCd
            frm1.vspdData2.Value = intIndex
        Case  C_CtrlKeyDataType1Nm  ' 자료유형1
            frm1.vspdData2.Col = Col
            intIndex = frm1.vspdData2.Value
            frm1.vspdData2.Col = C_CtrlKeyDataType1Cd
            frm1.vspdData2.Value = intIndex
        Case  C_CtrlKeyDataType2Nm  ' 자료유형2
            frm1.vspdData2.Col = Col
            intIndex = frm1.vspdData2.Value
            frm1.vspdData2.Col = C_CtrlKeyDataType2Cd
            frm1.vspdData2.Value = intIndex
        Case  C_CtrlKeyDataType3Nm  ' 자료유형3
            frm1.vspdData2.Col = Col
            intIndex = frm1.vspdData2.Value
            frm1.vspdData2.Col = C_CtrlKeyDataType3Cd
            frm1.vspdData2.Value = intIndex
        Case  C_CtrlKeyDataType4Nm ' 자료유형4
            frm1.vspdData2.Col = Col
            intIndex = frm1.vspdData2.Value
            frm1.vspdData2.Col = C_CtrlKeyDataType4Cd
            frm1.vspdData2.Value = intIndex
        Case  C_CtrlKeyDataType5Nm ' 자료유형5
            frm1.vspdData2.Col = Col
            intIndex = frm1.vspdData2.Value
            frm1.vspdData2.Col = C_CtrlKeyDataType5Cd
            frm1.vspdData2.Value = intIndex
    End Select
End Sub

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
    Dim intRow
    Dim intIndex

    For intRow = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row   = intRow
        frm1.vspdData.Col   = C_DrCrFgCd
        intIndex            = frm1.vspdData.value
        frm1.vspdData.col   = C_DrCrFgNm
        frm1.vspdData.value = intindex
    Next
End Sub

'==========================================================================================
Sub InitData2()
    Dim intRow
    Dim intIndex

    For intRow = 1 To frm1.vspdData2.MaxRows

        frm1.vspdData2.Row = intRow

        frm1.vspdData2.Col = C_CtrlDataTypeCd       ' 자료유형 
        intIndex = frm1.vspdData2.value
        frm1.vspdData2.col = C_CtrlDataTypeNm
        frm1.vspdData2.value = intindex

        frm1.vspdData2.Col = C_CtrlKeyDataType1Cd   ' 자료유형1
        intIndex = frm1.vspdData2.value
        frm1.vspdData2.col = C_CtrlKeyDataType1Nm
        frm1.vspdData2.value = intindex

        frm1.vspdData2.Col = C_CtrlKeyDataType2Cd   ' 자료유형2
        intIndex = frm1.vspdData2.value
        frm1.vspdData2.col = C_CtrlKeyDataType2Nm
        frm1.vspdData2.value = intindex

        frm1.vspdData2.Col = C_CtrlKeyDataType3Cd   ' 자료유형3
        intIndex = frm1.vspdData2.value
        frm1.vspdData2.col = C_CtrlKeyDataType3Nm
        frm1.vspdData2.value = intindex

        frm1.vspdData2.Col = C_CtrlKeyDataType4Cd   ' 자료유형4
        intIndex = frm1.vspdData2.value
        frm1.vspdData2.col = C_CtrlKeyDataType4Nm
        frm1.vspdData2.value = intindex

        frm1.vspdData2.Col = C_CtrlKeyDataType5Cd   ' 자료유형5
        intIndex = frm1.vspdData2.value
        frm1.vspdData2.col = C_CtrlKeyDataType5Nm
        frm1.vspdData2.value = intindex
    Next
End Sub



'==========================================================================================
Sub vspdData_onfocus()
 If lgIntFlgMode <> parent.OPMD_UMODE Then
     Call SetToolbar("1100110100001111")
 Else
     Call SetToolbar("1100111100001111")
 End If
End Sub


'==========================================================================================
Sub vspdData2_onfocus()
 Call SetToolbar("1100100000001111")
End Sub


'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

 dim i
 dim RowList
 Dim intRetCD

    If Row <> NewRow And NewRow > 0 Then


        ggoSpread.Source = frm1.vspdData
        lgCurrRow1 = NewRow

        '///////////////////////////////////////////////////////////////////////////////////////////////////
        '//테이블명과 테이블필드가 각각 일치하는지 체크필요   : dbquery가 필요함 
        '//두번째%d 줄의 테이블명과 필드명을 확인하세요 
        '///////////////////////////////////////////////////////////////////////////////////////////////////
'        frm1.vspdData.Row = row
'        frm1.vspdData.Col = C_AcctCd
'        If Len(frm1.vspdData.Text) > 0 Then
'            frm1.vspdData.Col = 0
'            frm1.vspdData.Row = Row

'            If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
'                For i = 1 To frm1.vspdData2.MaxRows
'                    frm1.vspdData2.Col = 0  
'                    frm1.vspdData2.row = i
'                    If (frm1.vspdData2.Text = ggoSpread.UpdateFlag or frm1.vspdData2.Text = ggoSpread.InsertFlag) Then
'                        If DbQuery_Six(i) = False Then
'
'                            If RowList = "" Then
'                                RowList =  i
'                            Else
'                                RowList = RowList & "," & i
'                            End If 
'                        End If
'                   End If
'                Next
'
'              If RowList <> "" Then
'
'                    IntRetCD = DisplayMsgBox("110820", vbokonly, RowList,"x")
'                    frm1.vspdData2.col = C_CtrlMaxColPlus1
'                    frm1.vspddata.row = frm1.vspdData2.text
'                    frm1.vspddata.col = frm1.vspddata.activeCol
'                    frm1.vspdData.Action = 0
'                    Exit sub
'                End If
'            End If 
'        End If
'20030616 jsk
		ggoSpread.Source = frm1.vspdData2
		If Not ggoSpread.SSDefaultCheck Then
			frm1.vspdData2.col = C_CtrlMaxColPlus1
			frm1.vspddata.row = frm1.vspdData2.text
			frm1.vspddata.col = frm1.vspddata.activeCol
			frm1.vspdData.Action = 0
			frm1.vspddata2.focus
			Exit sub
		End If

        frm1.vspdData.Row = NewRow
        frm1.vspdData.Col = C_AcctCd
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
        If Len(frm1.vspdData.Text) > 0 Then
          frm1.vspdData.Col = 0
            If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
            	Call BeforeDbQuery_Two(NewRow)
            End If
        End If

    End If

End Sub


'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim i
    Dim strWhere
    Dim strSelect
    Dim RowList
    Dim intRetCD
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    gMouseClickStatus = "SPC"


	Set gActiveSpdSheet = frm1.vspdData


    ggoSpread.Source = frm1.vspdData

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
        ggoSpread.SSSort ,lgSortKey
        lgSortKey = 1
        End If
        Exit Sub
    End If

    If frm1.vspdData.MaxRows <= 0 Then
        Exit Sub
    End If


    If Row = frm1.vspdData.ActiveRow Then
        Exit Sub
    End If

    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '//테이블명과 테이블필드가 각각 일치하는지 체크필요   : dbquery가 필요함 
    '//두번째%d 줄의 테이블명과 필드명을 확인하세요 
    '///////////////////////////////////////////////////////////////////////////////////////////////////
'    frm1.vspdData.Row = row
'    frm1.vspdData.Col = C_AcctCd
'    If Len(frm1.vspdData.Text) > 0 Then
'        frm1.vspdData.Col = 0
'        frm1.vspdData.Row = Row
'
'        If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
'            For i = 1 To frm1.vspdData2.MaxRows
'                frm1.vspdData2.Col = 0
'                frm1.vspdData2.row = i
'                If (frm1.vspdData2.Text = ggoSpread.UpdateFlag or frm1.vspdData2.Text = ggoSpread.InsertFlag) Then
'                    If DbQuery_Six(i) = False Then
'                        If RowList = "" Then
'                            RowList =  i
'                        Else
'                            RowList = RowList & "," & i
'                        End If 
'                    End If
'                End If
'            Next
'            If RowList <> "" Then
'                IntRetCD = DisplayMsgBox("110820", vbokonly, RowList,"x")
'                Exit sub
'            End If
'        End If
'		ggoSpread.Source = frm1.vspdData2
'		If Not ggoSpread.SSDefaultCheck Then
'			frm1.vspddata2.focus
'			Exit sub
'		End If
'   End If


    frm1.vspdData.Row = Row
    frm1.vspdData.Col = C_AcctCd

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    If Len(frm1.vspdData.Text) > 0 Then
        frm1.vspdData.Col = 0
        If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
            Call BeforeDbQuery_Two(Row)
        End If
    End If
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then

    End If

End Sub

'=======================================================================================================
' Function Name : DbQuery_Six
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery_Six(ByVal Row)
    Dim strSelect
    Dim strFrom
    Dim strWhere, strWhere1
    Dim strTableName
    Dim strField
    Dim strKeyField1
    Dim strKeyField2
    Dim strKeyField3
    Dim strKeyField4
    Dim strKeyField5
    Dim strDataType
    Dim strKeyType1
    Dim strKeyType2
    Dim strKeyType3
    Dim strKeyType4
    Dim strKeyType5
    Dim Rs0,Rs1
    Dim arrFldName
    Dim arrDataType
    Dim i
    Dim strFieldList

    Err.Clear

    With frm1

        .vspdData2.Row = Row
        .vspddata2.col = C_CtrlTblId
        strTableName = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlDataColmId
        strField = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyColmId1
        strKeyField1 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyColmId2
        strKeyField2 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyColmId3
        strKeyField3 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyColmId4
        strKeyField4 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyColmId5
        strKeyField5 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlDataTypeCd
        strDataType = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyDataType1Cd
        strKeyType1 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyDataType2Cd
        strKeyType2 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyDataType3Cd
        strKeyType3 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyDataType4Cd
        strKeyType4 = Trim(.vspddata2.text)

        .vspddata2.col = C_CtrlKeyDataType5Cd
        strKeyType5 = Trim(.vspddata2.text)

    End With 

    If strTableName = "" Then
        If (strField <> "" or strKeyField1 <> "" or strKeyField2 <> "" or strKeyField3 <> "" or strKeyField4 <> "" or strKeyField5 <> "") Then
            DbQuery_Six = False
            Exit function
        End If
    End If
    If strTableName <> "" Then
        If (strField = "" and strKeyField1 = "" and strKeyField2 = "" and strKeyField3 = "" and strKeyField4 = "" and strKeyField5 = "" ) Then
            DbQuery_Six = False
            Exit function
        End If
    End If
    If strTableName = "" and strField = "" and strKeyField1 = "" and strKeyField2 = "" and strKeyField3 = "" and strKeyField4 = "" and strKeyField5 = ""  Then
        DbQuery_Six = true
        Exit function
    End If
 
    DbQuery_Six = False



    strSelect = "UPPER(B.NAME), UPPER(C.NAME) " 
    strFrom = "SYSOBJECTS A, SYSCOLUMNS B, SYSTYPES C "
    strWhere =     " A.ID = B.ID "
    strWhere = strWhere & " AND UPPER(A.NAME) =  " & FilterVar(UCase(strTableName), "''", "S") 
    strWhere = strWhere & " AND  (A.XTYPE = " & FilterVar("U", "''", "S") & "  OR A.XTYPE = " & FilterVar("V", "''", "S") & " ) "
    strWhere = strWhere & " AND B.XTYPE = C.XTYPE"


    If CommonQueryRs2by2(strSelect,strFrom,strWhere,lgF2By2) = False Then
        DbQuery_Six = False
        Exit Function
    Else
        Rs1 = split(lgF2By2,chr(11) & chr(12) )
        i = 0
        Redim arrFldName(ubound(rs1))
        Redim arrDaType(ubound(rs1))
        dim j
        Do While i< ubound(Rs1)
            Rs0 = split(rs1(i), chr(11))
            arrFldName(i) =Rs0(1) 
            if instr(1,UCase(Rs0(2)),"INT")>0 or  instr(1,UCase(Rs0(2)),"NUMERIC") >0 THEN

                arrDaType(i) = "N"
            Elseif instr(1,UCase(Rs0(2)),"DATE")>0 THEN

                arrDaType(i) = "D"
            Elseif instr(1,UCase(Rs0(2)),"CHAR")>0 THEN

                arrDaType(i) = "S"
            End If
            i = i + 1
        Loop

        If FieldCheck(strField, strDataType, arrFldName, arrDaType) = false then
            DbQuery_Six = False
            Exit Function
        END iF

        If FieldCheck(strKeyField1, strKeyType1, arrFldName, arrDaType) = false then
            DbQuery_Six = False
            Exit Function
        END iF
        If FieldCheck(strKeyField2, strKeyType2, arrFldName, arrDaType) = false then
            DbQuery_Six = False
            Exit Function
        END iF
        If FieldCheck(strKeyField3, strKeyType3, arrFldName, arrDaType) = false then
            DbQuery_Six = False
            Exit Function
        END iF
        If FieldCheck(strKeyField4, strKeyType4, arrFldName, arrDaType) = false then
            DbQuery_Six = False
            Exit Function
        END iF
        If FieldCheck(strKeyField5, strKeyType5, arrFldName, arrDaType) = false then
            DbQuery_Six = False
            Exit Function
        END iF
    End If 
    DbQuery_Six = True
End Function



'=======================================================================================================
' Function Name : DbQuery_Seven
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery_Seven(ByVal Row)
    Dim strSelect
    Dim strFrom
    Dim strWhere, strWhere1  
    Dim strTableName
    Dim strField
    Dim strKeyField1
    Dim strKeyField2
    Dim strKeyField3
    Dim strKeyField4
    Dim strKeyField5
    Dim strDataType
    Dim strKeyType1
    Dim strKeyType2
    Dim strKeyType3
    Dim strKeyType4
    Dim strKeyType5
    Dim Rs0,Rs1
    Dim arrFldName
    Dim arrDataType
    Dim i,j
    Dim strFieldList

    Err.Clear

    With frm1

        .vspdData2.Row = Row
        .vspddata2.col = C_CtrlTblId
        strTableName = Trim(.vspddata2.text)

        strSelect = " UPPER(c.name),UPPER(t.name) "
        strFrom  =  " sysindexes i, syscolumns c, sysobjects o, systypes t "
        strWhere =  " o.name = " & FilterVar(UCase(strTableName), "''", "S") 
        strWhere = strWhere & " and o.id = c.id "
        strWhere = strWhere & " and o.id = i.id "
        strWhere = strWhere & " and (i.status & 0x800) = 0x800 "
        strWhere = strWhere & " and c.xtype = t.xtype "
        strWhere = strWhere & " and  (o.XTYPE = " & FilterVar("U", "''", "S") & "  OR o.XTYPE = " & FilterVar("V", "''", "S") & " )  "
        strWhere = strWhere & " and (c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  1) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  2) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  3) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  4) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  5) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  6) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  7) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  8) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  9) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 10) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 11) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 12) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 13) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 14) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 15) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 16)    "
        strWhere = strWhere & "      )                                              "
        strWhere = strWhere & "  order by c.colid                                   "


    If CommonQueryRs2by2(strSelect,strFrom,strWhere,lgF2By2) = False Then

        Exit Function
    Else
        Rs1 = split(lgF2By2,chr(11) & chr(12) )
        i = 0
        j = Ubound(rs1)
        Redim arrFldName(j)
        Redim arrDaType(j)


        Do While i< j
            Rs0 = split(rs1(i), chr(11))
            arrFldName(i) =Rs0(1) 
            if instr(1,UCase(Rs0(2)),"INT")>0 or  instr(1,UCase(Rs0(2)),"NUMERIC") >0 THEN
                arrDaType(i) = "2"
            Elseif instr(1,UCase(Rs0(2)),"DATE")>0 THEN
                arrDaType(i) = "1"
            Elseif instr(1,UCase(Rs0(2)),"CHAR")>0 THEN
                arrDaType(i) = "3"
            End If
            i = i+1
        Loop

            If j>=1 Then
                .vspddata2.col      = C_CtrlKeyColmId1
                .vspddata2.text     = arrFldName(0)
                .vspddata2.col      = C_CtrlKeyDataType1Nm
                .vspddata2.value    = arrDaType(0)
                .vspddata2.col      = C_CtrlKeyDataType1Cd
                .vspddata2.value    = arrDaType(0)
            End If 
                If j>=2 Then
                .vspddata2.col      = C_CtrlKeyColmId2
                .vspddata2.text     = arrFldName(1)
                .vspddata2.col      = C_CtrlKeyDataType2Nm
                .vspddata2.value    = arrDaType(1)
                .vspddata2.col      = C_CtrlKeyDataType2Cd
                .vspddata2.value    = arrDaType(1)
            End If 

            If j>=3 Then
                .vspddata2.col      = C_CtrlKeyColmId3
                .vspddata2.text     = arrFldName(2)
                .vspddata2.col      = C_CtrlKeyDataType3Nm
                .vspddata2.value    = arrDaType(2)
                .vspddata2.col      = C_CtrlKeyDataType3Cd
                .vspddata2.value    = arrDaType(2)
            End If 

            If j>=4 Then
                .vspddata2.col      = C_CtrlKeyColmId4
                .vspddata2.text     = arrFldName(3)
                .vspddata2.col      = C_CtrlKeyDataType4Nm
                .vspddata2.value    = arrDaType(3)
                .vspddata2.col      = C_CtrlKeyDataType4Cd
                .vspddata2.value    = arrDaType(3)
            End If 

            If j>=5 Then
                .vspddata2.col      = C_CtrlKeyColmId5
                .vspddata2.text     = arrFldName(4)
                .vspddata2.col      = C_CtrlKeyDataType5Nm
                .vspddata2.value    = arrDaType(4)
                .vspddata2.col      = C_CtrlKeyDataType5Cd
                .vspddata2.value    = arrDaType(4)
            End If
        End if
    End With
End Function
'=======================================================================================================
' Function Name : FieldCheck
' Function Desc : This function is data fleldname and data type
'=======================================================================================================


Function FieldCheck(fldName, dataType, arrFld, arrDtp)
    Dim exitFlag 
    dim i

    FieldCheck = false
    exitFlag = ""
    If (fldname = "" and dataType <> "" ) or (fldname <> "" and dataType = "" ) Then
        FieldCheck = False
        Exit Function
    End If

    '/// 필드명에 "'" 가 있을경우 
    If fldname <> "" and  left(Trim(fldname),1) = "'" and right(Trim(fldname),1) = "'" Then
        If CommonQueryRs2by2(fldname  ,"","",lgF2By2) = True Then
            FieldCheck = true
            Exit function
        End If
    End If



    If fldName <> "" Then
    for i=0 to ubound(arrFld)
    If fldName = arrFld(i) Then
    If dataType = arrDtp(i) Then
    exitFlag = "1"
    Exit For
    End If
    End If
    Next
    If exitFlag <> "1" Then
    FieldCheck = False
    Exit Function
    End If
    End If
    FieldCheck = true
End Function


'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

 If Button = 2 And gMouseClickStatus = "SPC" Then
  gMouseClickStatus = "SPCR"
 End If

End Sub

'==========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)

 If Button = 2 And gMouseClickStatus = "SP2C" Then
  gMouseClickStatus = "SP2CR"
 End If

End Sub



'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strTemp
    Dim intPos1

    If Row <= 0 Then
        Exit Sub
    End If

    With frm1
    Select Case Col
        Case C_JnlPopUp
            .vspdData.Col = C_JnlCD
            .vspdData.Row = Row
            Call OpenPopUp(.vspdData.Text, 1)

        Case C_EventPopUp
            .vspdData.Col = C_EventCd
            .vspdData.Row = Row
            Call OpenPopUp(.vspdData.Text, 3)
        Case C_AcctPopUp
            .vspdData.Col = C_AcctCD
            .vspdData.Row = Row
            Call OpenPopUp(.vspdData.Text, 2)
        End Select
		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")
    End With

End Sub


'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strTemp
    Dim intPos1

    If Row <= 0 Then
        Exit Sub
    End If

    With frm1
        Select Case Col
            Case C_CtrlTblPopUp
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 4)

            Case C_CtrlColmPopUp1
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 5)
            Case C_CtrlColmPopUp2
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 6)
            Case C_CtrlColmPopUp3
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 7)
            Case C_CtrlColmPopUp4
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 8)
            Case C_CtrlColmPopUp5
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 9)
            Case C_CtrlColmPopUp6
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(.vspdData2.Text, 10)
        End Select
		Call SetActiveCell(.vspdData2,Col-1,.vspdData2.ActiveRow ,"M","X","X")
    End With

End Sub

'==========================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub


'==========================================================================================
Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub


'==========================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )
    Dim DblNetAmt, DblVatAmt, DblNetLocAmt, DblVatLocAmt 

 With frm1.vspdData

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
    Dim LngRowCnt
    Dim intRetCD
    Dim FormSeq
	Dim intIndex
    Dim lngRowCnt2

	On Error Resume Next
	Err.Clear

	lgBlnFlgChgValue = True

    Call CheckMinNumSpread(frm1.vspdData,Col,Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row


    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0

    Select Case Col
		Case  C_DrCrFgNm			' 차대구분 
			frm1.vspdData.Col = Col
			frm1.vspdData.Row = Row
			intIndex = frm1.vspdData.Value
			frm1.vspdData.Col = C_DrCrFgCd
			frm1.vspdData.Value = intIndex
		Case C_AcctCd 
			If	(frm1.vspdData.Text = ggoSpread.InsertFlag Or frm1.vspdData.Text = ggoSpread.UpdateFlag) Then
				frm1.vspdData.Col = C_FormSeq:	FormSeq = frm1.vspdData.Text
			     Call DeleteHSheet(FormSeq)
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData
			    frm1.vspdData.Col = C_AcctCD
			    frm1.txtAcctCD.value = frm1.vspdData.Text

			    If Len(Trim(frm1.vspdData.Text)) > 0 Then
					'//같은것을 수정했을경우 
					frm1.vspdData.row = Row
					frm1.vspdData.col = C_MaxColPlus2'
					frm1.vspdData.text = "ACCT"
					frm1.vspdData.row = Row
					frm1.vspdData.col = C_MaxColPlus1  '//3 
					frm1.vspdData.text = "N"

					Call BeforeDbQuery_Two(Row)
				End If
			End If 
	End Select
End Sub




'==========================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)

    Dim iDx,strText
    lgBlnFlgChgValue = True

    Call CheckMinNumSpread(frm1.vspdData2,Col,Row)

    '//
    frm1.vspdData2.Col = Col
    frm1.vspdData2.Row = Row

    Select Case Col
        Case C_CtrlTblId
            '//imsi
            strText = UCase(Trim(frm1.vspdData2.text))
            With frm1 
                '//테이블이하 필드, 자료유형 을 모두 지워주기 
                .vspdData2.Col = C_CtrlDataColmId
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlDataTypeCd
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlDataTypeNm
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyColmId1
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType1Cd
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType1Nm
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyColmId2
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType2Cd
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType2Nm
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyColmId3
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType3Cd
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType3Nm
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyColmId4
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType4Cd
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType4Nm
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyColmId5
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType5Cd
                .vspdData2.Text = ""
                .vspdData2.Col = C_CtrlKeyDataType5Nm
                .vspdData2.Text = ""

         '//자동으로 키컬럼을 가져와 세팅해주기.
                .vspdData2.Col = C_CtrlTblId
                If Trim(.vspdData2.Text) <> "" Then
                    CALL DBQUERY_SEVEN(Row)
                End If
            End With

        Case  C_CtrlDataTypeNm,C_CtrlKeyDataType1Nm,C_CtrlKeyDataType2Nm,C_CtrlKeyDataType3Nm,C_CtrlKeyDataType4Nm,C_CtrlKeyDataType5Nm
            strText = UCase(Trim(frm1.vspdData2.text))
            If strText = "STRING" Then
                iDx = "3"
            ElseIf strText = "NUMERIC" Then
                iDx = "2"
            ElseIf strText = "DATE" Then
                iDx = "1"
            End If
            Frm1.vspdData2.value = iDx
            Frm1.vspdData2.Col = Col-1
            Frm1.vspdData2.value = iDx
    End Select

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

    Call CopyToHSheet(Row)

End Sub

'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP2C"

    Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

	If Row <= 0 Then
	   ggoSpread.Source = frm1.vspdData2
	   Exit Sub
	End If
End Sub

'==========================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
    End If
End Sub


'==========================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'==========================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)
End Sub


'==========================================================================================
Sub vspddata_KeyPress(KeyAscii )

End Sub


'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt
    Dim LngLastRow
    Dim LngMaxRow

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgPageNo <> "" Then
            Call DisableToolBar(parent.TBC_QUERY)
            If DbQuery_One = False Then
                Call RestoreToolBar()
                Exit Sub
            End if
        End If
    End if
End Sub




'==========================================================================================
Function FncQuery() 
	Dim IntRetCD 
    Dim var1, var2

    FncQuery = False

    Err.Clear


    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange


    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables


	frm1.hTransType.value = UCase(Trim(frm1.txtTransType.value))

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    Call DbQuery_One

    FncQuery = True
   Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncNew()
    On Error Resume Next
End Function

'==========================================================================================
Function FncDelete()
    On Error Resume Next
End Function

'==========================================================================================
Function FncSave()
Dim IntRetCD 
 Dim var1,var2

    FncSave = False

    Err.Clear

    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    If lgBlnFlgChgValue = False And var1 = False And var2 = False  Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
	Exit Function
    End If

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then
		frm1.vspddata.focus
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData2
	If Not ggoSpread.SSDefaultCheck Then
		frm1.vspddata2.focus
		Exit Function
	End If

    Call DbSave_One
    FncSave = True
	Set gActiveElement = document.ActiveElement
End Function


'========================================================================================
Function FncCopy() 
	Dim IntRetCD

	frm1.vspdData.ReDraw = False

	if frm1.vspdData.MaxRows < 1 then Exit Function

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
         If IntRetCD = vbNo Then
             Exit Function
         End If
    End If

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncCancel()
    Dim FormCnt

    if frm1.vspdData.MaxRows < 1 then Exit Function

        frm1.vspdData.Row = frm1.vspdData.ActiveRow
        frm1.vspdData.Col = 0

        If frm1.vspdData.Text = ggoSpread.InsertFlag or frm1.vspdData.Text = ggoSpread.UpdateFlag Then
            frm1.vspdData.Col = C_FormSeq: FormCnt = frm1.vspdData.Text
           Call DeleteHSheet(FormCnt)
        End if

        ggoSpread.Source = frm1.vspdData
        ggoSpread.EditUndo


        Call InitData

        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
        frm1.vspdData.Row = frm1.vspdData.ActiveRow

		If Trim(frm1.hTransType.value) = "" Then
			Exit Function
		End If

        frm1.vspdData.Col = C_AcctCd
        If frm1.vspdData.text <> "" Then
			if frm1.vspdData.MaxRows < 1 then Exit Function
			Call BeforeDbQuery_Two(frm1.vspdData.ActiveRow)
		End If
	Set gActiveElement = document.ActiveElement
    lgBlnFlgChgValue = False
End Function

'==========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim imRow
    Dim lngNum
    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1.vspdData
		intItemCnt = .MaxRows

        lngNum = FindNumber(frm1.vspdData, C_FormSeq) + 1

		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
        ggoSpread.InsertRow ,imRow
        Call SetSpreadColor(.ActiveRow, .ActiveRow + imRow - 1)
		.Col    = C_FormSeq
		.Row    = .ActiveRow
		.Text   = lngNum

        .ReDraw = False
        .focus
        ggoSpread.Source = .vspdData

        .ReDraw = True

    End With

    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

	Call ggoOper.LockField(Document, "Q")

End Function


'==========================================================================================
Function FncDeleteRow()
    Dim lDelRows
    Dim FormSeq

    ggoSpread.Source = frm1.vspdData 

    With frm1.vspdData

    .Row = .ActiveRow
    .Col = 0

    If frm1.vspdData.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function
        .Col = C_FormSeq
        FormSeq = .Text
        lDelRows = ggoSpread.DeleteRow
    End With

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call DeleteHSheet(FormSeq)
    lgBlnFlgChgValue = True

End Function

'==========================================================================================
Function FncPrint()
 Call Parent.FncPrint()
End Function

'==========================================================================================
Function FncExcel()
    Call FncExport(parent.C_MULTI)
End Function

'==========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)
End Function

'==========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'==========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'==========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'==========================================================================================
Sub PopRestoreSpreadColumnInf()

	Dim indx

	On Error Resume Next
	Err.Clear

	If gActiveSpdSheet.Name <> "" Then
		For indx = 0 To frm1.vspdData.MaxRows
			frm1.vspdData.Row = indx
			frm1.vspdData.Col = 0
			Select Case Trim(UCase(gActiveSpdSheet.Name))
				Case "VSPDDATA"
					If frm1.vspdData.Text = ggoSpread.InsertFlag Then
						frm1.vspdData.Col = C_ItemSeq
						Call DeleteHSheet(frm1.vspdData.Text)
					End If
					If frm1.vspdData.Text = ggoSpread.UpdateFlag Then
                        ggoSpread.Source = frm1.vspdData2
                        ggoSpread.ClearSpreadData
                        ggoSpread.Source = frm1.vspdData3
                        ggoSpread.ClearSpreadData
					End If
				Case "VSPDDATA2"

                    If frm1.vspdData.Text = ggoSpread.DeleteFlag Or _
                       frm1.vspdData.Text = ggoSpread.UpdateFlag Then
                        Call FncUndoData(indx)
                    End If
					If frm1.vspdData.Text = ggoSpread.InsertFlag Then
						frm1.vspdData.Col = C_AcctCd
						frm1.vspdData.Text = ""
						frm1.vspdData.Col = C_AcctNm
						frm1.vspdData.Text = ""
						frm1.vspdData.Col = C_ItemSeq
			 			Call DeleteHSheet(frm1.vspdData.Text)
			 		End If
			End Select
		Next
	End If

	ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call InitSpreadSheet()
			Call ggoSpread.ReOrderingSpreadData()
            Call InitData()
            Call InitData2()

		Case "VSPDDATA2"
			Call InitSpreadSheet1()
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()
            Call InitData()
            Call InitData2()

	End Select

    If frm1.vspdData.MaxRows > 0 Then
        If frm1.vspdData.Text <> "" Then
            If frm1.vspdData2.MaxRows <= 0 Then
                Call BeforeDbQuery_Two(frm1.vspdData.ActiveRow)
            End If
        End If
    End If
End Sub


'==========================================================================================
' Function Name : FncUndoData
' Function Desc :
'==========================================================================================
Sub FncUndoData(Byval pRow)
	On Error Resume Next

	Dim indx
	Dim TempSeq

	frm1.vspdData.Row = pRow
	frm1.vspdData.Col = C_ItemSeq
	frm1.vspdData.Action = 0

    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.EditUndo()

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

End Sub

'==========================================================================================
Function FncExit()
    Dim IntRetCD
    Dim var1,var2

    FncExit = False

    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = True or var1 = True or var2 = True Then
        IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    FncExit = True

End Function
'==========================================================================================
' Name : FncMakeHiddenColumn
' Desc : 
'==========================================================================================
Sub FncMakeHiddenColumn(ByVal Index)
    If IsNull(Index) Or Not IsNumeric(Index) Then
       Exit Sub
    End If

    If Not IsObject(gActiveSpdSheet) Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSSetColHidden(Index,Index,True,"FD")
 End Sub
'==========================================================================================
' Name : FncMakeVisibleColumn
' Desc : 
'==========================================================================================
Sub FncMakeVisibleColumn(ByVal Index)
    If IsNull(Index) Or Not IsNumeric(Index) Then
       Exit Sub
    End If

    If Not IsObject(gActiveSpdSheet) Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSSetColHidden(Index,Index,False,"FD")

End Sub

'==========================================================================================
Function FncPasteRepeatedSpreadData()
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.CPasteRepeatedSpreadData
End Function

'==========================================================================================
Sub FncSaveSpreadInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadInf()
End Sub

'==========================================================================================
Sub FncResetSpreadInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ResetSpreadInf()
	Call initSpreadPosVariables()
    Call InitSpreadSheet
	Call InitVariables
	Call InitComboBox
End Sub

'==========================================================================================
Function DbQuery_One() 

    DbQuery_One = False

    Call LayerShowHide(1)

    Dim strVal

    with frm1

  If lgIntFlgMode = parent.OPMD_UMODE Then
   strVal = BIZ_PGM_QRY_ID1 & "?txtMode=" & parent.UID_M0001
   strVal = strVal & "&txtTransType=" & UCase(Trim(.hTransType.value)) 
   strVal = strVal & "&lgStrPrevKeyOne_Seq="    & lgStrPrevKeyOne_Seq
   strVal = strVal & "&txtMaxRows_One="         & .vspdData.MaxRows
  Else
   strVal = BIZ_PGM_QRY_ID1 & "?txtMode="    & parent.UID_M0001
   strVal = strVal & "&txtTransType="        & UCase(Trim(.hTransType.value))
   strVal = strVal & "&lgStrPrevKeyOne_Seq="    & lgStrPrevKeyOne_Seq
   strVal = strVal & "&txtMaxRows_One="         & .vspdData.MaxRows
  End If

    End With

 Call RunMyBizASP(MyBizASP, strVal)

    DbQuery_One = True

End Function

'==========================================================================================
' Function Name : DbQuery_OneOk
' Function Desc : DbQuery_One가 성공적일 경우 MyBizASP 에서 호출되는 Function
'==========================================================================================
Function DbQuery_OneOk()

    Dim intRow

    With frm1
        .vspdData.Col = 1:    intItemCnt = .vspddata.MaxRows
        Call SetSpreadLock("Q", 0, 1, "")

        '-----------------------
        'Reset variables area
        '-----------------------
        Call ggoOper.LockField(Document, "Q")
        Call SetToolbar("1100111100011111")

        Call InitData
        .vspdData2.Redraw = False
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
        ggoSpread.Source = frm1.vspdData3
        ggoSpread.ClearSpreadData

        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            If .vspdData.Text <> "" Then
                Call BeforeDbQuery_Two(1)
            End If
        End If
        frm1.vspdData2.Redraw = True
        lgIntFlgMode = parent.OPMD_UMODE 
        lgBlnFlgChgValue = False
        lgBlnStartFlag = True
    End With
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================
' Function Name : BeforeDbQuery_Two
' Function Desc : This function is data query and display
'==========================================================================================
Function BeforeDbQuery_Two(ByVal Row)
    Dim strVal
    Dim boolExist
    Dim intRetCD

    On Error Resume Next
    Err.Clear

    boolExist = False

    With frm1
        .vspdData.Row = Row

        .vspddata.col = C_FormSeq
        .txtFormSeq.value = .vspddata.text

        .vspddata.col = C_JnlCD
        .txtJnlCD.value = .vspddata.text

        .vspddata.col = C_DrCrFgCd
        .txtDrCrFgCd.value = .vspddata.text

        .vspddata.col = C_AcctCD
        .txtAcctCD.value = .vspddata.text

        If Len(Trim(.hTransType.value)) <= 0 Then 
            intRetCD =  DisplayMsgBox("700103","x","x","x")
            .vspdData.Col = C_AcctCD
            .vspdData.Text = ""
            .vspdData.Col = C_AcctNm
            .vspdData.Text = ""

            Exit Function
        End If

        If Len(Trim(.txtJnlCD.value)) <= 0 Then
            intRetCD =  DisplayMsgBox("700104","x","x","x")

            .vspdData.Col = C_AcctCD
            .vspdData.Text = ""
            .vspdData.Col = C_AcctNm
            .vspdData.Text = ""
            Exit Function
        End If

      If Len(Trim(.txtDrCrFgCd.value)) <= 0 Then
            intRetCD =  DisplayMsgBox("700105","x","x","x")

            .vspdData.Col = C_AcctCD
            .vspdData.Text = ""
            .vspdData.Col = C_AcctNm
            .vspdData.Text = ""
            Exit Function
        End If

        Call LayerShowHide(1)

        BeforeDbQuery_Two = False


        '// A_ACCT_CTRL_ASSN ==>A
        '//A_JNL_CTRL_ASSN   ==>B
        '// A테이블에 있고 B테이블에 없으면 B테이블에 추가 
        '// A테이블에 없고 B테이블에 있으면 B테이블 해당레코드 삭제 
        frm1.vspdData.row = Row
        frm1.vspdData.col = 0
        If frm1.vspddata.text <> ggoSpread.UpdateFlag and frm1.vspddata.text <> ggoSpread.InsertFlag and frm1.vspddata.text <> ggoSpread.DeleteFlag  Then
            strVal = BIZ_PGM_QRY_ID6 & "?txtMode=" & parent.UID_M0002
            strVal = strVal & "&txtTransType=" & UCase(Trim(.hTransType.value))
            strVal = strVal & "&txtJnlCd=" & UCase(Trim(.txtJnlCD.value))
            strVal = strVal & "&txtFormSeq=" & Trim(.txtFormSeq.value)
            strVal = strVal & "&txtDrCrFgCd=" & UCase(Trim(.txtDrCrFgCd.value))
            strVal = strVal & "&txtAcctCd=" & UCase(Trim(.txtAcctCD.value))
            strVal = strVal & "&txtRow=" & Row

           Call RunMyBizASP(MyBizASP, strVal)

        Else
			Call DbQuery_Two(Row)
	    End If
	End With

    Call LayerShowHide(0)
    BeforeDbQuery_Two = True
End Function



'==========================================================================================
' Function Name : BeforeDbQuery_TwoOK
' Function Desc : This function is data query and display
'==========================================================================================
Function BeforeDbQuery_TwoOK(Row)
	Call DbQuery_Two(Row)
End Function


'==========================================================================================
' Function Name : DbQuery_Two
' Function Desc : vspdData2의 조회및 신규입력시 
'==========================================================================================
Function DbQuery_Two(ByVal Row)
    Dim boolExist
    Dim lngRows
    Dim intRetCD
    Dim strSelect
    Dim strFrom
    Dim strWhere
    Dim BasicFlag
    Dim CtrlItemSeq
    Dim BasicFlag2
    Dim strCommand
    Dim Indx1
	Dim arrTemp

    On Error Resume Next

    Err.Clear
    boolExist = False


    With frm1

        .vspdData.Row = Row

        .vspddata.col = C_FormSeq
        .txtFormSeq.value = .vspddata.text

        .vspddata.col = C_JnlCD
        .txtJnlCD.value = .vspddata.text

        .vspddata.col = C_DrCrFgCd
        .txtDrCrFgCd.value = .vspddata.text

        .vspddata.col = C_AcctCD
        .txtAcctCD.value = .vspddata.text

        If Len(Trim(.hTransType.value)) <= 0 Then
            intRetCD =  DisplayMsgBox("700103","x","x","x")
            .vspdData.Col = C_AcctCD
            .vspdData.Text = ""
            .vspdData.Col = C_AcctNm
            .vspdData.Text = ""

            Exit Function
        End If

        .vspdData.Col = C_JnlCD
        .txtJnlCD.value = .vspdData.Text
        If Len(Trim(frm1.vspdData.Text)) <= 0 Then
            intRetCD =  DisplayMsgBox("700104","x","x","x")

            .vspdData.Col = C_AcctCD
            .vspdData.Text = ""
            .vspdData.Col = C_AcctNm
            .vspdData.Text = ""
            Exit Function
        End If

        .vspdData.Col = C_DrCrFgCd
        .txtDrCrFgCd.value = .vspdData.Text
        If Len(Trim(frm1.vspdData.Text)) <= 0 Then
            intRetCD =  DisplayMsgBox("700105","x","x","x")

            .vspdData.Col = C_AcctCD
            .vspdData.Text = ""
            .vspdData.Col = C_AcctNm
            .vspdData.Text = ""
            Exit Function
        End If

        Call LayerShowHide(1)

        DbQuery_Two = False



        frm1.vspdData.row = Row
        frm1.vspdData.col = C_MaxColPlus2 
		'20030617 jsk
        If frm1.vspdData.text <> "ACCT" Then 
            strSelect =    " DISTINCT A.CTRL_CD, A.CTRL_NM, B.TRANS_TYPE, B.JNL_CD, B.SEQ, CHAR(8), "
            strSelect = strSelect & " B.DR_CR_FG, B.ACCT_CD, B.TBL_ID,'', B.DATA_COLM_ID,'', B.DATA_TYPE, '', "
            strSelect = strSelect & " LTrim(ISNULL(B.KEY_COLM_ID1,'')),'', LTrim(ISNULL(B.KEY_DATA_TYPE_1,'')), '', "
            strSelect = strSelect & " LTrim(ISNULL(B.KEY_COLM_ID2,'')),'', LTrim(ISNULL(B.KEY_DATA_TYPE_2,'')), '', "
            strSelect = strSelect & " LTrim(ISNULL(B.KEY_COLM_ID3,'')),'', LTrim(ISNULL(B.KEY_DATA_TYPE_3,'')), '', "
            strSelect = strSelect & " LTrim(ISNULL(B.KEY_COLM_ID4,'')),'', LTrim(ISNULL(B.KEY_DATA_TYPE_4,'')), '', "
            strSelect = strSelect & " LTrim(ISNULL(B.KEY_COLM_ID5,'')),'', LTrim(ISNULL(B.KEY_DATA_TYPE_5,'')), '', "
            strSelect = strSelect & " " & Row & ", "
            strSelect = strSelect & " CASE WHEN  " & FilterVar("DR", "''", "S") & " = " & FilterVar(.txtDrCrFgCd.value, "''", "S") & " THEN ISNULL(C.DR_FG," & FilterVar("N", "''", "S") & " ) ELSE ISNULL(C.CR_FG," & FilterVar("N", "''", "S") & " ) END ,"
			
            strSelect = strSelect & " CHAR(8)  "

            strFrom = " A_CTRL_ITEM A (NOLOCK), A_JNL_CTRL_ASSN B (NOLOCK),  A_ACCT_CTRL_ASSN C (NOLOCK)  "

            strWhere =     " A.CTRL_CD = B.CTRL_CD "
            strWhere = strWhere & " AND B.ACCT_CD =  C.ACCT_CD"			
            strWhere = strWhere & " AND B.CTRL_CD = C.CTRL_CD "			
            strWhere = strWhere & " AND B.TRANS_TYPE =  " & FilterVar(UCase(.hTransType.value), "''", "S")  
            strWhere = strWhere & " AND B.SEQ = " & FilterVar(Trim(.txtFormSeq.value),"0","N")  
            strWhere = strWhere & " AND B.JNL_CD =  " & FilterVar(.txtJnlCD.value, "''", "S")
            strWhere = strWhere & " AND B.DR_CR_FG = " & FilterVar(.txtDrCrFgCd.value, "''", "S") 
            strWhere = strWhere & " AND B.ACCT_CD =  " & FilterVar(.txtAcctCD.value, "''", "S")  

            strWhere = strWhere & " ORDER BY A.CTRL_CD "

        End If

        If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then  
                ggoSpread.Source = frm1.vspdData2
                arrTemp =  Split(lgF2By2,Chr(12))
                For Indx1 = 0 To Ubound(arrTemp) - 1
                    arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
                Next
                lgF2By2 = Join(arrTemp,Chr(12))
                ggoSpread.SSShowData lgF2By2
            BasicFlag = "N"
        ELSE
			'20030617 jsk
            strSelect =    " A.CTRL_CD, A.CTRL_NM, " & FilterVar(UCase(.hTransType.value), "''", "S")  & " , " & FilterVar(UCase(.txtJnlCD.value), "''", "S")  & ", " & FilterVar(Trim(.txtFormSeq.value),"0","N")   & ", CHAR(8), " '20021224 수정 
            strSelect = strSelect & FilterVar(.txtDrCrFgCd.value, "''", "S")  & ", B.ACCT_CD, '','', '','', '','', "
            strSelect = strSelect & " '', '', '', '', "
            strSelect = strSelect & " '', '', '', '', "
            strSelect = strSelect & " '', '', '', '', "
            strSelect = strSelect & " '', '', '', '', "
            strSelect = strSelect & " '', '', '', '', "
            strSelect = strSelect & " " & Row & ", "
            strSelect = strSelect & " CASE WHEN  " & FilterVar("DR", "''", "S") & " = " & FilterVar(.txtDrCrFgCd.value, "''", "S") & " THEN ISNULL(B.DR_FG," & FilterVar("N", "''", "S") & " ) ELSE ISNULL(B.CR_FG," & FilterVar("N", "''", "S") & " ) END ,"
            strSelect = strSelect & " CHAR(8)  " 
            strFrom = " A_CTRL_ITEM A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK) "

            strWhere =     " A.CTRL_CD = B.CTRL_CD "
            strWhere = strWhere & " AND B.ACCT_CD =  " & FilterVar(.txtAcctCD.value, "''", "S")  

            strWhere = strWhere & " ORDER BY B.CTRL_ITEM_SEQ "

            If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
                ggoSpread.Source = frm1.vspdData2
                arrTemp =  Split(lgF2By2,Chr(12))
                For Indx1 = 0 To Ubound(arrTemp) - 1
                    arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
                Next
                lgF2By2 = Join(arrTemp,Chr(12))
                ggoSpread.SSShowData lgF2By2

            END IF

            BasicFlag = "Y"
        END IF

   ' CtrlItemSeq 는 새로운 행마다 순차적으로 증가, 새로 입력되는 행은 입력행으로 표시 
        CtrlItemSeq = 1

        For lngRows = 1 To .vspdData2.Maxrows
            .vspddata2.row = lngRows

            '// 두번째 spreadSheet안에 상위의 row위치를 숨겨둠 
            .vspdData2.col = C_CtrlMaxColPlus1
            .vspdData2.text = row 
            .vspddata2.col = C_CtrlTransType
            if UCase(Trim(.hTransType.value)) = UCase(Trim(.vspddata2.text)) then
                .vspddata2.col = C_CtrlJnlCD
                if UCase(Trim(.txtJnlCD.value)) = UCase(Trim(.vspddata2.text)) then
                .vspddata2.col = C_CtrlFormSeq
                    if UCase(Trim(.txtFormSeq.value)) = UCase(Trim(.vspddata2.text)) then
                        .vspddata2.col = C_CtrlDrCrFgCd
                        if UCase(Trim(.txtDrCrFgCd.value)) = UCase(Trim(.vspddata2.text)) then
                            .vspddata2.col = C_CtrlAcctCD
                            if UCase(Trim(.txtAcctCD.value)) = UCase(Trim(.vspddata2.text)) then
                                .vspddata2.col = C_CtrlCtrlCnt
                                .vspddata2.text = CtrlItemSeq

                                CtrlItemSeq = CtrlItemSeq + 1
                                if BasicFlag <> "N" then
                                    .vspddata2.Col = 0
                                    .vspddata2.Text = ggoSpread.InsertFlag
                                    BasicFlag2 = "Y"
                                end if
                            end if
                        end if
                    end if
                end if
            end if
        NEXT

  '//속도저하로 인해 새로 입력하는 항목인경우는 히든으로 복사하는 함수를 밖으로 뺌 
        frm1.vspdData.row = Row
        frm1.vspdData.col = C_MaxColPlus1 '//3

        If BasicFlag2 = "Y" and frm1.vspdData.text = "N" Then
            For lngRows = 1 To .vspdData2.Maxrows
                call CopyToHSheet(lngRows)
            Next
            frm1.vspdData.row = Row
            frm1.vspdData.col = C_MaxColPlus1 '//3
            frm1.vspdData.text = "Y"
        End If



        if .vspddata3.maxrows > 0 then
            .vspdData.Row = Row
            .vspdData.Col = C_FormSeq      ' 순번 
            .txtFormSeq.value = .vspdData.Text
            If CopyFromData(.txtFormSeq.value) = True Then
                frm1.vspdData2.Redraw = True
            End If
        end if
    End With
    Call LayerShowHide(0)
    Call DbQuery_TwoOk()
    DbQuery_Two = True 
End Function

'==========================================================================================
' Function Name : DbQuery_TwoOk
' Function Desc : DbQuery_Two가 성공적일 경우 MyBizASP 에서 호출되는 Function
'==========================================================================================
Function DbQuery_TwoOk()
    Dim Cnt

    frm1.vspdData2.Redraw = True

    With frm1
        .vspdData.Col = 1:    intItemCnt = .vspddata.MaxRows

        Call InitData2
        ggoSpread.Source = .vspdData2
        Call SetSpreadLock("Q", 1, 1, "")
        ggoSpread.Source = .vspdData
    End With
End Function

'==========================================================================================
' Function Name : DbSave_One
' Function Desc : This function is data query and display
'==========================================================================================
Function DbSave_One()

    Dim lngRows
    Dim lGrpcnt1, lGrpcnt2, lGrpcnt3, lGrpcnt4
    DIM strVal1, strVal2, strVal3, strVal4
    Dim strDel1, strDel2, strDel3, strDel4
    Dim strVspdDataFlag
    Dim row, i, RowList
    Dim intRetCD

    DbSave_One = False
 
 '//테이블명과 테이블필드가 각각 일치하는지 체크필요   : dbquery가 필요함 
 '//두번째%d 줄의 테이블명과 필드명을 확인하세요 

	ggoSpread.Source = frm1.vspdData2
	If Not ggoSpread.SSDefaultCheck Then
		frm1.vspddata2.focus
		Exit Function
	End If

'    row = frm1.vspdData.ActiveRow
'    frm1.vspdData.row = row
'    frm1.vspdData.Col = C_AcctCd
'    If Len(frm1.vspdData.Text) > 0 Then
'        frm1.vspdData.Col = 0
'        frm1.vspdData.Row = Row
'        If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
'            For i = 1 To frm1.vspdData2.MaxRows
'                frm1.vspdData2.Col = 0
'                frm1.vspdData2.row = i
'                If (frm1.vspdData2.Text = ggoSpread.UpdateFlag or frm1.vspdData2.Text = ggoSpread.InsertFlag) Then
'                    If DbQuery_Six(i) = False Then
'                        If RowList = "" Then
'                            RowList =  i
'                        Else
'                            RowList = RowList & "," & i
'                        End If 
'                    End If
'                End If
'            Next
'            If RowList <> "" Then
'                IntRetCD = DisplayMsgBox("110820", vbokonly, RowList,"x")
'                Exit Function
'            End If
'        End If 
'    End If


    '// save 작업 시작 

    Call LayerShowHide(1)

    With frm1
        .txtFlgMode.value = lgIntFlgMode
        .txtMode.value = parent.UID_M0002
    End With

    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

    lGrpCnt1 = 1: lGrpCnt2 = 1: lGrpCnt3 = 1: lGrpCnt4 = 1
    strVal1 = "": strVal2 = "": strVal3 = "": strVal4 = ""
    strDel1 = "": strDel2 = "": strDel3 = "": strDel4 = ""

    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
        For lngRows = 1 To .MaxRows

            .Row = lngRows
            ' sheet1에서 수정되지 않은것도 update 로 인식한다.
            .Col = C_MaxColPlus3 '//1

            if .text = ggoSpread.UpdateFlag then
                strVspdDataFlag = "" 
                .text = ""   '//vspddata1 손대지 않고 vspddata2에서 수정했을경우 
            Else
                .col = 0
                strVspdDataFlag = .text
            end if

            If strVspdDataFlag = ggoSpread.InsertFlag Then
                strVal1 = strVal1 & "C" & parent.gColSep & lngRows & parent.gColSep    'C=Create, Sheet가 2개 이므로 구별 
            ElseIf strVspdDataFlag = ggoSpread.UpdateFlag Then
                strVal1 = strVal1 & "U" & parent.gColSep & lngRows & parent.gColSep    'U=Update
            ElseIf strVspdDataFlag = ggoSpread.DeleteFlag Then
                strDel1 = strDel1 & "D" & parent.gColSep & lngRows & parent.gColSep    'D=Delete
            End If
   
            Select Case strVspdDataFlag
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                    strVal1 = strVal1 & Trim(frm1.hTransType.value) & parent.gColSep ' 거래유형 
                    .Col = C_JnlCD            ' 거래항목 
                    strVal1 = strVal1 & Trim(.Text) & parent.gColSep
                    .Col = C_FormSeq
                    strVal1 = strVal1 & Trim(.Text) & parent.gColSep
                    .Col = C_DrCrFgCd           ' 차대구분 
                    strVal1 = strVal1 & Trim(.Text) & parent.gColSep
                    .Col = C_EventCd
                    strVal1 = strVal1 & Trim(.Text) & parent.gColSep
                    .Col = C_AcctCD            ' 계정과목 
                    '           strVal1 = strVal1 & Trim(.Text) & parent.gColSep
                    '           .Col = C_TransAcctCD          ' 이동계정과목 
                    strVal1 = strVal1 & Trim(.Text) & parent.gRowSep
                    lGrpCnt1 = lGrpCnt1 + 1
                Case ggoSpread.DeleteFlag
                    strDel1 = strDel1 & Trim(frm1.hTransType.value) & parent.gColSep ' 거래유형 
                    .Col = C_JnlCD            ' 거래항목 
                    strDel1 = strDel1 & Trim(.Text) & parent.gColSep
                    .Col = C_FormSeq
                    strDel1 = strDel1 & Trim(.Text) & parent.gColSep
                    .Col = C_DrCrFgCd           ' 차대구분 
                    strDel1 = strDel1 & Trim(.Text) & parent.gColSep
                    .Col = C_AcctCD            ' 계정과목 
                    strDel1 = strDel1 & Trim(.Text) & parent.gRowSep
                    lGrpCnt1 = lGrpCnt1 + 1
            End Select
        Next
 End With

    With frm1.vspdData3
        For lngRows = 1 To .MaxRows

            .Row = lngRows
            ' sheet2에서 모든 필드는 새로 생성한다. delete 는 발생하지 않음 
            .Col = 0
            If .Text = ggoSpread.InsertFlag Then
                strVal2 = strVal2 & "C" & parent.gColSep & lngRows & parent.gColSep    'C=Create, Sheet가 2개 이므로 구별 
            ElseIf .Text = ggoSpread.UpdateFlag Then
                strVal2 = strVal2 & "U" & parent.gColSep & lngRows & parent.gColSep    'U=Update
            ElseIf .Text = ggoSpread.DeleteFlag Then
                strDel2 = strDel2 & "D" & parent.gColSep & lngRows & parent.gColSep    'D=Delete
            End If

            Select Case .Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                    .Col = 3 'C_CtrlTransType
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 4 'C_CtrlJnlCD
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 5 'C_CtrlFormSeq
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 7 'C_CtrlDrCrFgCd
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 8 'C_CtrlAcctCD
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 1 'C_CtrlCtrlCD
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 9 'C_CtrlTblId
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 11 'C_CtrlDataColmId
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 13 'C_CtrlDataTypeCd
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 15 'C_CtrlKeyColmId1
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 17 'C_CtrlKeyDataType1Cd
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 19 'C_CtrlKeyColmId2
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 21 'C_CtrlKeyDataType2Cd
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 23 'C_CtrlKeyColmId3
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 25 'C_CtrlKeyDataType3Cd
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 27 'C_CtrlKeyColmId4
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 29 'C_CtrlKeyDataType4Cd
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 31 'C_CtrlKeyColmId5
                    strVal2 = strVal2 & Trim(.Text) & parent.gColSep
                    .Col = 33 'C_CtrlKeyDataType5Cd
                    strVal2 = strVal2 & Trim(.Text) & parent.gRowSep
                    lGrpCnt2 = lGrpCnt2 + 1
                Case ggoSpread.DeleteFlag
                    .Col = 3 'C_CtrlTransType
                    strDel2 = strDel2 & Trim(.Text) & parent.gColSep
                    .Col = 4 'C_CtrlJnlCD
                    strDel2 = strDel2 & Trim(.Text) & parent.gColSep
                    .Col = 5 'C_CtrlFormSeq
                    strDel2 = strDel2 & Trim(.Text) & parent.gColSep
                    .Col = 7 'C_CtrlDrCrFgCd
                    strDel2 = strDel2 & Trim(.Text) & parent.gColSep
                    .Col = 8 'C_CtrlAcctCD
                    strDel2 = strDel2 & Trim(.Text) & parent.gColSep
                    .Col = 1 'C_CtrlCtrlCD
                    strDel2 = strDel2 & Trim(.Text) & parent.gRowSep
                    lGrpCnt2 = lGrpCnt2 + 1
            End Select
        Next
    End With

    ' 본지사가 없는 Jnl_Form 부분 
    frm1.txtMaxRows_One.value = lGrpCnt1-1          'Spread Sheet의 변경된 최대갯수 
    frm1.txtSpread1.value =  strDel1 & strVal1         'Spread Sheet 내용을 저장 

    ' 본지사가 없는 Jnl_Ctrl_Assn 부분 
    frm1.txtMaxRows_Two.value = lGrpCnt2-1          'Spread Sheet의 변경된 최대갯수 
    frm1.txtSpread2.value =  strDel2 & strVal2        'Spread Sheet 내용을 저장 

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

    ' 본지사가 없는 Jnl_Form의 내용이 있는지 확인한다.
    ' 있으면 저장하고 없으면 본지사가 없는 Jnl_Form 부분을 Skip 한다.
    If frm1.txtMaxRows_One.value > 0 Then
        ' 본지사가 없는 Jnl_Form 부분 
        Call ExecMyBizASP(frm1, BIZ_PGM_QRY_ID1)        '저장 비지니스 ASP 를 가동 
    Else
        If frm1.txtMaxRows_Two.value > 0 Then
            ' 본지사가 없는 Jnl_Ctrl_Assn 부분 
            Call ExecMyBizASP(frm1, BIZ_PGM_QRY_ID1)        '저장 비지니스 ASP 를 가동 
        Else
            Call DbSave_OneOk(UCase(frm1.hTransType.value))
        End If
    End If
    DbSave_One = True
End Function

'==========================================================================================
Function DbSave_Two() 
    ' 본지사가 없는 Jnl_Ctrl_Assn의 내용이 있는지 확인한다.
    ' 있으면 저장하고 없으면 본지사가 없는 Jnl_Ctrl_Assn 부분을 Skip 한다.
    If frm1.txtMaxRows_Two.value > 0 Then
        Call ExecMyBizASP(frm1, BIZ_PGM_QRY_ID1)
    Else
        Call DbSave_OneOk(UCase(frm1.hTransType.value))
    End If
End Function

'==========================================================================================
' Function Name : DbSave_OneOk
' Function Desc : DbSave_One가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'==========================================================================================
Function DbSave_OneOk(ByVal pTransType)
	ggoSpread.SSDeleteFlag 1

	If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.hTransType.value = pTransType
	End If

	lgBlnFlgChgValue = False
	lgBlnStartFlag = True


	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
	Call InitVariables
	Call DbQuery_One
End Function

'==========================================================================================
Function CheckSpread2()

	Dim indx
	Dim tmpDrCrFG

	CheckSpread2 = False

	With frm1
	
	 	For indx = 1 to .vspdData2.MaxRows
		    .vspdData2.Row = indx
		    '.vspdData2.Col = 14

		    'If .vspddata2.Text = "Y" Or .vspddata2.Text = "DC" Then
  			  .vspdData2.Col = C_CtrlTblId
			  If Trim(.vspdData2.Text) = "" Then
				Exit Function
		  	  End If
		    'End If
		Next

        End With

	CheckSpread2 = True

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>분개형태등록</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=*>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR HEIGHT=*>
  <TD WIDTH="100%" CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD HEIGHT=20 WIDTH="100%">
      <FIELDSET CLASS="CLSFLD">
       <TABLE <%=LR_SPACE_TYPE_40%>>
        <TR>
         <TD CLASS="TD5">거래유형</TD>
         <TD CLASS="TD656"><INPUT NAME="txtTransType" MAXLENGTH="20" SIZE = "20" ALT="거래유형" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtTransType.Value, 0)">&nbsp;
             <INPUT NAME="txtTransNM" MAXLENGTH="50" SIZE = "30" ALT="거래유형" tag="14X"></TD>
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
        <TD HEIGHT="50%" NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
        </TD>
       </TR>
       <TR>
        <TD HEIGHT="50%" NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
 <TR>
  <TD WIDTH="100%" HEIGHT=<%=BizSize%>>
   <IFRAME NAME="MyBizASP" WIDTH="100%" src="../../blank.htm"  HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread1 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread2 tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtMaxRows_One" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows_Two" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtMaxRows5" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtFormSeq" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtCtrlCnt" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtJnlCD" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtDrCrFgCd" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtAcctCD" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtCtrlCD" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="hFormCnt" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hCtrlCnt" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hTransType" tag="14" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hJnlCD" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDrCrFgCd" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCD" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="hCtrlCD" tag="24" TABINDEX="-1">

<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100 tag="2" TITLE="SPREAD" id=OBJECT2 TABINDEX="-1"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
