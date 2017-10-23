<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 인사/급여관리 
*  2. Function Name        : 급/상여 공제 관리 
*  3. Program ID           : H5405ma1
*  4. Program Name         : 국민연금 
*  5. Program Desc         : 국민연금 전산 신고 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : Hwang Jeong Won
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h5405mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "h5405mb2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
Const C_SHEETMAXROWS1   = 10                                           '☜: Visble row

Const TAB1 = 1                                                        'Tab Index
Const TAB2 = 2
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow
Dim lgStrPrevKey1

Dim C_KIND1
Dim C_COMP_NO1
Dim C_SEQ_NO1
Dim C_NAME1
Dim C_ANUT_NO1
Dim C_BLANK1
Dim C_TYPE1
Dim C_LOSS_DT1
Dim C_SP_JOB1
Dim C_DCL_DT1

Dim C_KIND
Dim C_COMP_NO
Dim C_SEQ_NO
Dim C_NAME
Dim C_ANUT_NO
Dim C_PAY_AMOUNT
Dim C_GRADE
Dim C_TYPE
Dim C_ENT_DT
Dim C_SP_JOB
Dim C_DCL_DT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_KIND = 1
        C_COMP_NO = 2
        C_SEQ_NO = 3     
        C_NAME  = 4   
        C_ANUT_NO = 5  
        C_PAY_AMOUNT = 6     
        C_GRADE = 7  
        C_TYPE = 8 
        C_ENT_DT = 9 
        C_SP_JOB = 10 
        C_DCL_DT = 11 

    ElseIf pvSpdNo = "B" Then
        C_KIND1 = 1
        C_COMP_NO1 = 2    
        C_SEQ_NO1 = 3     
        C_NAME1  = 4   
        C_ANUT_NO1 = 5  
        C_BLANK1 = 6     
        C_TYPE1 = 7 
        C_LOSS_DT1 = 8 
        C_SP_JOB1 = 9 
        C_DCL_DT1 = 10 
    End If

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
	lgIntGrpCount      = 0										'⊙: Initializes Group View Size
    lgStrPrevKey       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey1	   = ""                                     '⊙: initializes Previous Key Index
    lgSortKey          = 1                                      '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	frm1.txtFr_acq_dt.Text =  UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtTo_acq_dt.Text =  frm1.txtFr_acq_dt.Text
	frm1.txtRprt_dt.Text   =  frm1.txtFr_acq_dt.Text
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function
	
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
   
       lgKeyStream       = Trim(Frm1.txtComp_cd.Value) & Parent.gColSep       'You Must append one character(Parent.gColSep)
       lgKeyStream       = lgKeyStream & Frm1.txtFr_acq_dt.Text & Parent.gColSep
       lgKeyStream       = lgKeyStream & Frm1.txtTo_acq_dt.Text & Parent.gColSep
       lgKeyStream       = lgKeyStream & UNIConvDate(Frm1.txtRprt_dt.Text) & Parent.gColSep
       lgKeyStream       = lgKeyStream & gSelframeFlg & Parent.gColSep
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
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 

	    With Frm1.vspdData
            ggoSpread.Source = Frm1.vspdData
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols = C_DCL_DT + 1                                                      '☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              '☜:☜: Hide maxcols
           .ColHidden = True                                                            '☜:☜:
    
           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("A") 'sbk
	
                ggoSpread.SSSetEdit     C_KIND,         "서식기호", 10,,,,2
                ggoSpread.SSSetEdit     C_COMP_NO,      "사업장기호", 12,,,,2
	            ggoSpread.SSSetEdit     C_SEQ_NO,       "순번", 10,,,,2
                ggoSpread.SSSetEdit     C_NAME,         "성명", 12,,,,2
	            ggoSpread.SSSetEdit	   C_ANUT_NO,	   "국민연금번호", 15,,,,2
	            ggoSpread.SSSetFloat    C_PAY_AMOUNT,   "소득월액", 15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
                ggoSpread.SSSetEdit     C_GRADE,        "등급", 8,,,,2
                ggoSpread.SSSetEdit     C_TYPE,         "부호", 8,,,,2
	            ggoSpread.SSSetEdit     C_ENT_DT,       "취득일", 10,,,,2
                ggoSpread.SSSetEdit     C_SP_JOB,       "직종", 8,,,,2
	            ggoSpread.SSSetEdit     C_DCL_DT,       "신고일", 10,,,,2

	       .ReDraw = true
	
           lgActiveSpd = "M"

           Call SetSpreadLock("A") 
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

	    With Frm1.vspdData1
	
            ggoSpread.Source = Frm1.vspdData1
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols = C_DCL_DT1 + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:
    
           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("B") 'sbk

                ggoSpread.SSSetEdit     C_KIND1,         "서식기호", 12,,,,2
                ggoSpread.SSSetEdit     C_COMP_NO1,      "사업장기호", 13,,,,2
	            ggoSpread.SSSetEdit     C_SEQ_NO1,       "순번", 10,,,,2
                ggoSpread.SSSetEdit     C_NAME1,         "성명", 13,,,,2
	            ggoSpread.SSSetEdit	   C_ANUT_NO1,	    "국민연금번호", 15,,,,2
                ggoSpread.SSSetEdit     C_BLANK1,        "공란", 13,,,,2
                ggoSpread.SSSetEdit     C_TYPE1,         "부호", 10,,,,2
	            ggoSpread.SSSetEdit     C_LOSS_DT1,      "상실일", 10,,,,2
                ggoSpread.SSSetEdit     C_SP_JOB1,       "직종", 10,,,,2
	            ggoSpread.SSSetEdit     C_DCL_DT1,       "신고일", 11,,,,2
	
	       .ReDraw = true
	
           lgActiveSpd = "S"
           
           Call SetSpreadLock("B") 
    
        End With
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "A" Then
        ggoSpread.Source = Frm1.vspdData

        With frm1.vspdData
        	.ReDraw = False

        	ggoSpread.SpreadLock    C_KIND, -1, C_KIND, -1
        	ggoSpread.SpreadLock    C_COMP_NO, -1, C_COMP_NO, -1
        	ggoSpread.SpreadLock    C_SEQ_NO, -1, C_SEQ_NO, -1
        	ggoSpread.SpreadLock    C_NAME, -1, C_NAME, -1
        	ggoSpread.SpreadLock    C_ANUT_NO, -1, C_ANUT_NO, -1
        	ggoSpread.SpreadLock    C_PAY_AMOUNT, -1, C_PAY_AMOUNT, -1
        	ggoSpread.SpreadLock    C_GRADE, -1, C_GRADE, -1
        	ggoSpread.SpreadLock    C_TYPE, -1, C_TYPE, -1
        	ggoSpread.SpreadLock    C_ENT_DT, -1, C_ENT_DT, -1
        	ggoSpread.SpreadLock    C_SP_JOB, -1, C_SP_JOB, -1
        	ggoSpread.SpreadLock    C_DCL_DT, -1, C_DCL_DT, -1
        	ggoSpread.SSSetProtected   .MaxCols   , -1, -1

        	.ReDraw = True
        End With
        
    ElseIf pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData1

        With frm1.vspdData1
        	.ReDraw = False

        	ggoSpread.SpreadLock    C_KIND1, -1, C_KIND1, -1
        	ggoSpread.SpreadLock    C_COMP_NO1, -1, C_COMP_NO1, -1
        	ggoSpread.SpreadLock    C_SEQ_NO1, -1, C_SEQ_NO1, -1
        	ggoSpread.SpreadLock    C_NAME1, -1, C_NAME1, -1
        	ggoSpread.SpreadLock    C_ANUT_NO1, -1, C_ANUT_NO1, -1
        	ggoSpread.SpreadLock    C_BLANK1, -1, C_BLANK1, -1
        	ggoSpread.SpreadLock    C_TYPE1, -1, C_TYPE1, -1
        	ggoSpread.SpreadLock    C_LOSS_DT1, -1, C_LOSS_DT1, -1
        	ggoSpread.SpreadLock    C_SP_JOB1, -1, C_SP_JOB1, -1
        	ggoSpread.SpreadLock    C_DCL_DT1, -1, C_DCL_DT1, -1
        	ggoSpread.SSSetProtected   .MaxCols   , -1, -1

        	.ReDraw = True
        End With
    End If

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
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData1.Col    = iDx
              Frm1.vspdData1.Row    = iRow
              Frm1.vspdData1.Action = 0 ' go to 
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

            C_KIND = iCurColumnPos(1)
            C_COMP_NO = iCurColumnPos(2)
            C_SEQ_NO = iCurColumnPos(3)
            C_NAME  = iCurColumnPos(4)
            C_ANUT_NO = iCurColumnPos(5)
            C_PAY_AMOUNT = iCurColumnPos(6)
            C_GRADE = iCurColumnPos(7)
            C_TYPE = iCurColumnPos(8)
            C_ENT_DT = iCurColumnPos(9)
            C_SP_JOB = iCurColumnPos(10)
            C_DCL_DT = iCurColumnPos(11)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_KIND1 = iCurColumnPos(1)
            C_COMP_NO1 = iCurColumnPos(2)
            C_SEQ_NO1 = iCurColumnPos(3)
            C_NAME1  = iCurColumnPos(4)
            C_ANUT_NO1 = iCurColumnPos(5)
            C_BLANK1 = iCurColumnPos(6)
            C_TYPE1 = iCurColumnPos(7)
            C_LOSS_DT1 = iCurColumnPos(8)
            C_SP_JOB1 = iCurColumnPos(9)
            C_DCL_DT1 = iCurColumnPos(10)
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet("")                                                         'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call SetDefaultVal
    frm1.txtComp_cd.focus 
	gSelframeFlg = TAB1
	Call changeTabs(TAB1)
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
    
	Call CookiePage (0)                                                             '☜: Check Cookie
    
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    If gSelframeFlg = TAB1 Then
	    ggoSpread.Source = frm1.vspdData
    Else
    	ggoSpread.Source = frm1.vspdData1
	End If            

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
   		If IntRetCD = vbNo Then
  			Exit Function
   		End If
	End If

    If Not(ValidDateCheckThisForm(frm1.txtFr_acq_dt, frm1.txtTo_acq_dt)) Then
        Exit Function
    End If

    If Not(txtComp_cd_OnChange()) Then
        Exit Function
    End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
		If gPageNo > 0 Then
			gSelframeFlg = gPageNo
		End If
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    If DbQuery = False Then
        Exit Function
    End If
       
    FncQuery = True																'☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

	If gSelframeFlg <> TAB1 Then
		Call changeTabs(TAB1)	
		gSelframeFlg = TAB1	
	End If	
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    Call SetDefaultVal
    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData1

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = Frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
		 If gPageNo > 0 Then
			gSelframeFlg = gPageNo
		    If gSelframeFlg = TAB1 Then	              
	            Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
		    Else	             
	            Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
		    End If	
         End If
       Exit Function
    End If
    If DbSave = False Then
        Exit Function
    End If
    
    FncSave = True                                                               '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

     If Frm1.vspdData1.MaxRows < 1 Then
        Exit Function
     End If
    
     With Frm1
          If .vspdData1.ActiveRow > 0 Then
             .vspdData1.ReDraw = False
		
              ggoSpread.Source = .vspdData1	
              ggoSpread.CopyRow
              SetSpreadColor   .vspdData1.ActiveRow
             .vspdData1.Col  = 1
              .vspdData1.Text = ""
             .vspdData1.ReDraw = True
             .vspdData1.Focus
         End If
    End With

	If gSelframeFlg = TAB2 Then
    	frm1.vspdData1.ReDraw = False
    	
        ggoSpread.Source = frm1.vspdData1	
        ggoSpread.CopyRow
        SetSpreadColor frm1.vspdData1.ActiveRow
        
    	frm1.vspdData1.ReDraw = True
    ElseIf gSelframeFlg = TAB1 Then 
        Call ggoOper.ClearField(Document, "1")                                  <%'Clear Condition Field%>
    End If 

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 

    ggoSpread.Source = Frm1.vspdData1	
    ggoSpread.EditUndo  

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
  
        
     If gSelframeFlg <> TAB2 Then
    	Call ClickTab2		'sstData.Tab = 1
     End If

     With Frm1
              .vspdData1.ReDraw = False
              .vspdData1.Focus
               ggoSpread.Source = .vspdData1
               ggoSpread.InsertRow
               SetSpreadColor .vspdData1.ActiveRow
              .vspdData1.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    If Frm1.vspdData1.MaxRows < 1 then
       Exit function
    End if	

    With Frm1.vspdData1 
              .Focus
              ggoSpread.Source = frm1.vspdData1 
              lDelRows = ggoSpread.DeleteRow
    End With
                 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, True)
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
    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
	End Select 

	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = Frm1.vspdData1
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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

    Dim strVal 
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
    If LayerShowHide(1) =False Then
       Exit Function
    End If
    

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & gSelframeFlg                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key

    If gSelframeFlg = Tab1 Then    
	   lgCurrentSpd = "M"
       strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data
   Else   
	   lgCurrentSpd = "S"
       strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey1             '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows         '☜: Max fetched data
    End If   
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True                                                                   '☜: Processing is NG

End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If LayerShowHide(1) =False Then
       Exit Function
    End If
		
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '☜: Processing is OK
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

    If gSelframeFlg = TAB1 Then
    	Frm1.vspdData.focus
    Else
    	Frm1.vspdData1.focus
	End If            
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData1
    Frm1.vspdData1.MaxRows = 0
    lgCurrentSpd = "S"

    DBQuery()
   	Call ClickTab1()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function
'===============================================================================
' Function Name : ValidDateCheckThisForm
' Function Desc : Valid Date Check Function
'===============================================================================

Function ValidDateCheckThisForm(ThisObjFromDt, ThisObjToDt)

	ValidDateCheckThisForm = False

	If Len(Trim(ThisObjToDt.Text)) And Len(Trim(ThisObjFromDt.Text)) Then
		If ValidDateCheck(ThisObjFromDt,ThisObjToDt) =False Then
			ThisObjFromDt.Text = ""
            ThisObjToDt.Text = ""
            ThisObjFromDt.focus
            Set gActiveElement = document.activeElement                            			
			Exit Function
		End If
	End If
 
	ValidDateCheckThisForm = True

End Function

'========================================================================================================
' Name : FncOpenCondiPopup()
' Desc : developer describe this line 
'========================================================================================================
Function FncOpenCondiPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	    Case "1"

	        arrParam(0) = "사업장코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "b_biz_area"						    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtComp_cd.value        			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = ""                      		    	' Where Condition
	    	arrParam(5) = "사업장코드" 			            ' TextBox 명칭 
	
	    	arrField(0) = "biz_area_cd"						    	' Field명(0)
	    	arrField(1) = "biz_area_nm"    					    	' Field명(1)
    
	    	arrHeader(0) = "사업장코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "사업장명"	    		            ' Header명(1)
	    Case "2"
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtComp_cd.focus
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtComp_cd.value = arrRet(0)
		        .txtComp_nm.value = arrRet(1)		
		        .txtComp_cd.focus
        End Select
	End With
End Sub

'========================================== Tab Click 처리  =================================================
'	name: Tab Click
'	desc: Tab Click시 필요한 기능을 수행한다.
'===================================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then
	    Exit Function
	End If
	
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then
	     Exit Function
	End If
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
End Function

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
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub    

Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP1C"   

    Set gActiveSpdSheet = frm1.vspdData1
   
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1
       
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
'-----------------------------------------
Sub vspdData1_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SP1C" Then
          gMouseClickStatus = "SP1CR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
          Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then
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

'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change( Col ,  Row)

    Dim iDx
    Frm1.vspdData1.Row = Row
    Frm1.vspdData1.Col = Col

    Select Case Col
    End Select    
             
    If Frm1.vspdData1.CellType = Parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData1.text) < UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
          Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows = 0 Then
        Exit Sub
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub


'=======================================================================================================
'   Event Name : txtYear_lick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFr_acq_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtFr_acq_dt.Action = 7
        frm1.txtFr_acq_dt.focus
    End If
End Sub
Sub txtTo_acq_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtTo_acq_dt.Action = 7
        frm1.txtTo_acq_dt.focus
    End If
End Sub
Sub txtRprt_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtRprt_dt.Action = 7
        frm1.txtRprt_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txttxtFr_acq_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtFr_acq_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtTo_acq_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtRprt_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'======================================================================================================
'   Event Name : txtComp_cd_OnChange
'   Event Desc : 사업장코드가 변경될 경우 
'=======================================================================================================
Function txtComp_cd_OnChange()
    Dim IntRetCd
    Dim strWhere

    If Trim(frm1.txtComp_cd.Value) = "" Then
        frm1.txtComp_nm.Value = ""
        txtComp_cd_OnChange = True
    Else
        strWhere = " biz_area_cd= " & FilterVar(frm1.txtComp_cd.Value, "''", "S") & ""
        IntRetCD = CommonQueryRs(" biz_area_nm "," b_biz_area ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtComp_cd.Value)<>""  Then
            Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
            frm1.txtComp_nm.Value=""
            txtComp_cd_OnChange = False
        Else
            frm1.txtComp_nm.Value=Trim(Replace(lgF0,Chr(11),""))
            txtComp_cd_OnChange = True
        End If
    End If
End Function

'==========================================================================================
'   Event Name : btnCb_autoisrt_OnClick()
'   Event Desc : 파일생성 
'==========================================================================================
Function btnCb_creation_OnClick()
Dim RetFlag
Dim strVal
Dim intRetCD

    Err.Clear                                                                   '☜: Clear err status
    
    If Not chkField(Document, "1") Then                                         ' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
       Exit Function                            
    End If
    
    If (gSelframeFlg = 1 And frm1.vspdData.MaxRows <= 0) OR (gSelframeFlg = 2 And frm1.vspdData1.MaxRows <= 0) Then
		Call DisplayMsgBox("800167", "X","X","X")			 '⊙: Data is changed.  Do you want to exit? 		
		Exit Function		
    End If
 
	If UniConvDateToYYYYMMDD(frm1.txtFr_acq_dt.Text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtTo_acq_dt.Text,Parent.gDateFormat,"") Then
		intRetCD =  DisplayMsgBox("970025","X" , frm1.txtFr_acq_dt.Alt, frm1.txtTo_acq_dt.Alt)		
		frm1.txtFr_acq_dt.focus		
		Exit Function
	End If

	RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분	
	If RetFlag = VBNO Then
		Exit Function
	End IF

    With frm1

	If LayerShowHide(1) =False Then
       Exit Function
    End If					 
	    
	    strVal = BIZ_PGM_ID2 & "?txtMode="	& Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 	    	    
		strVal = strVal & "&txtFr_acq_dt="	& Trim(.txtFr_acq_dt.Text)
		strVal = strVal & "&txtTo_acq_dt="	& Trim(.txtTo_acq_dt.Text)
		strVal = strVal & "&txtComp_cd="	& Trim(.txtComp_cd.value)
		strVal = strVal & "&txtReportDt="	& UNIConvDate(.txtRprt_dt.Text)	 				
		strVal = strVal & "&gSelframeFlg="	& gSelframeFlg
		If lgCurrentSpd = "M" Then
            strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
        Else   
            strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey1             '☜: Next key tag
	    End If   
	        strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)
	
    End With    
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '☜: Protect system from crashing
    If Trim(pFileName) <> "" Then
	    strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
	    strVal = strVal & "&txtFileName=" & pFileName							'☆: 조회 조건 데이타	
	    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자격취득</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자격상실</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
						        <TD CLASS=TD5 NOWRAP>사업장기호</TD>
					            <TD CLASS=TD6 >
								    <INPUT NAME="txtComp_cd" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN:left" ALT ="사업장기호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenCondiPopup('1')">
								    <INPUT NAME="txtComp_nm" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="사업장명" tag="14XXXU">
					            </TD>
							    <TD CLASS=TD5 NOWRAP>해당일</TD>
			                    <TD CLASS=TD6 >
			                        <script language =javascript src='./js/h5405ma1_fr_Acq_dt_txtFr_acq_dt.js'></script>
			                       ~<script language =javascript src='./js/h5405ma1_to_Acq_dt_txtTo_acq_dt.js'></script>
			                    </TD>
			                </TR>
			                <TR>
								<TD CLASS=TD5 NOWRAP>신고일</TD>
			                    <TD CLASS=TD6 >
			                        <script language =javascript src='./js/h5405ma1_rprt_dt_txtRprt_dt.js'></script>
			                    </TD>
								<TD CLASS=TD5 NOWRAP></TD>
			                    <TD CLASS=TD6 ></TD>
							</TR>
						    				
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						    <TABLE <%=LR_SPACE_TYPE_20%> >
						    	<TR>
						    		<TD HEIGHT="100%">
						    			<script language =javascript src='./js/h5405ma1_vaSpread_vspdData.js'></script>
						    		</TD>
						    	</TR>
						    </TABLE>
						</DIV>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						    <TABLE <%=LR_SPACE_TYPE_20%> >
						    	<TR>
						    		<TD HEIGHT="100%">
						    			<script language =javascript src='./js/h5405ma1_vaSpread1_vspdData1.js'></script>
						    		</TD>
						    	</TR>
						    </TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT="20">
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnCb_creation" CLASS="CLSMBTN" Flag=1>파일생성</BUTTON></TD>
				    <TD WIDTH=* Align=RIGHT></TD>
				    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

