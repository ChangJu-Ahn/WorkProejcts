<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 월차수당기준등록 
*  3. Program ID           : H1013ma1
*  4. Program Name         : H1013ma1
*  5. Program Desc         : 기준정보관리/월차수당기준등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/14
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "H1013mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 15	                                      '☜: Visble row

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

Dim C_DILIG 
Dim C_DILIG_POP
Dim C_DILIG_NM
Dim C_DILIG_CNT

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_DILIG		= 1
	 C_DILIG_POP	= 2
	 C_DILIG_NM		= 3
	 C_DILIG_CNT	= 4
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
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
Sub MakeKeyStream(pOpt)
    lgKeyStream = Frm1.txtAllow_cd.Value & parent.gColSep       'You Must append one character( parent.gColSep)
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0092", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtProv_type, lgF0, lgF1, Chr(11))

'   전월/당월/익월 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0093", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtAccum_mm, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0093", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtAccum_mm1, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0093", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtmm_Accum, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0093", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtUse_mm, iCodeArr, iNameArr, Chr(11))


'   전년도/당년도/익년도 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0094", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtcrt_strt_yy, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0094", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtcrt_end_yy, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0094", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtuse_strt_yy, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0094", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call  SetCombo2(frm1.txtuse_end_yy, iCodeArr, iNameArr, Chr(11))

    iCodeArr = "" & Chr(11) & "01" & Chr(11) & "02" & Chr(11) & "03" & Chr(11) & "04" & Chr(11) & "05" & Chr(11) & "06" & Chr(11) & "07" & Chr(11) & "08" & Chr(11) & "09" & Chr(11) & "10" & Chr(11) & "11" & Chr(11) & "12" & Chr(11)
    iNameArr = "" & Chr(11) & "01" & Chr(11) & "02" & Chr(11) & "03" & Chr(11) & "04" & Chr(11) & "05" & Chr(11) & "06" & Chr(11) & "07" & Chr(11) & "08" & Chr(11) & "09" & Chr(11) & "10" & Chr(11) & "11" & Chr(11) & "12" & Chr(11)
    Call  SetCombo2(frm1.txtCrt_strt_mm, iCodeArr, iNameArr, Chr(11))

    Call  SetCombo2(frm1.txtCrt_end_mm, iCodeArr, iNameArr, Chr(11))
    Call  SetCombo2(frm1.txtuse_strt_mm, iCodeArr, iNameArr, Chr(11))
    Call  SetCombo2(frm1.txtuse_end_mm, iCodeArr, iNameArr, Chr(11))

    iCodeArr = "01" & Chr(11) & "02" & Chr(11) & "03" & Chr(11) & "04" & Chr(11) & "05" &_
                      Chr(11) & "06" & Chr(11) & "07" & Chr(11) & "08" & Chr(11) & "09" & Chr(11) & "10" &_
                      Chr(11) & "11" & Chr(11) & "12" & Chr(11) & "13" & Chr(11) & "14" & Chr(11) & "15" &_
                      Chr(11) & "16" & Chr(11) & "17" & Chr(11) & "18" & Chr(11) & "19" & Chr(11) & "20" &_
                      Chr(11) & "21" & Chr(11) & "22" & Chr(11) & "23" & Chr(11) & "24" & Chr(11) & "25" &_
                      Chr(11) & "26" & Chr(11) & "27" & Chr(11) & "28" & Chr(11) & "29" & Chr(11) & "30" &_
                      Chr(11) & "31" & Chr(11)
    iNameArr = "01일" & Chr(11) & "02일" & Chr(11) & "03일" & Chr(11) & "04일" & Chr(11) & "05일" &_
                      Chr(11) & "06일" & Chr(11) & "07일" & Chr(11) & "08일" & Chr(11) & "09일" & Chr(11) & "10일" &_
                      Chr(11) & "11일" & Chr(11) & "12일" & Chr(11) & "13일" & Chr(11) & "14일" & Chr(11) & "15일" &_
                      Chr(11) & "16일" & Chr(11) & "17일" & Chr(11) & "18일" & Chr(11) & "19일" & Chr(11) & "20일" &_
                      Chr(11) & "21일" & Chr(11) & "22일" & Chr(11) & "23일" & Chr(11) & "24일" & Chr(11) & "25일" &_
                      Chr(11) & "26일" & Chr(11) & "27일" & Chr(11) & "28일" & Chr(11) & "29일" & Chr(11) & "30일" &_
                      Chr(11) & "31일" & Chr(11)

    Call  SetCombo2(frm1.txtCrt_strt_dd, iCodeArr, iNameArr, Chr(11))
    Call  SetCombo2(frm1.txtCrt_end_dd, iCodeArr, iNameArr, Chr(11))
    Call  SetCombo2(frm1.txtUse_strt_dd, iCodeArr, iNameArr, Chr(11))
    Call  SetCombo2(frm1.txtUse_end_dd, iCodeArr, iNameArr, Chr(11))
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
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_DILIG_CNT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	
       Call  AppendNumberPlace("6","2","0")
       Call  GetSpreadColumnPos("A") 

        ggoSpread.SSSetEdit   C_DILIG,         "근태코드",    20,,,2,2
        ggoSpread.SSSetButton C_DILIG_POP
        ggoSpread.SSSetEdit   C_DILIG_NM,      "근태코드명",   68,,,68
        ggoSpread.SSSetFloat  C_DILIG_CNT,     "횟수",       25,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0"
        
        Call ggoSpread.MakePairsColumn(C_DILIG,  C_DILIG_POP)
        
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
    
    .vspdData.ReDraw = False
       ggoSpread.SpreadLock      C_DILIG,	 -1, C_DILIG
       ggoSpread.SpreadLock      C_DILIG_NM,  -1, C_DILIG_NM
       ggoSpread.SpreadLock      C_DILIG_POP, -1, C_DILIG_POP
       ggoSpread.SSSetRequired   C_DILIG_CNT, -1, C_DILIG_CNT
       ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
       ggoSpread.SSSetRequired     C_DILIG		, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected    C_DILIG_NM	, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired     C_DILIG_CNT	, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected	.vspdData.MaxCols, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_DILIG		= iCurColumnPos(1)
			C_DILIG_POP	= iCurColumnPos(2)
			C_DILIG_NM	= iCurColumnPos(3)
			C_DILIG_CNT	= iCurColumnPos(4)           
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call  AppendNumberPlace("7","2","0")
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox

    ' 수당코드에 값을 
    Call  CommonQueryRs(" MAX(allow_cd) "," HDA140T ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtAllow_cd.value = Trim(Replace(lgF0, Chr(11),""))

    Call  CommonQueryRs(" allow_nm "," HDA010T "," allow_cd= " & FilterVar(frm1.txtAllow_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtAllow_nm.value = Replace(lgF0, Chr(11), "")


    Call txtProv_type_OnChange()
    Call InitVariables                                                              'Initializes local global variables
	Call SetToolbar("1111110100101111")												'⊙: Set ToolBar
	Call CookiePage (0)                                                             '☜: Check Cookie
    Call MainQuery()			

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

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    If txtAllow_cd_Onchange() Then                                                '☜: enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call MakeKeyStream("X")
    
    Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End if
              
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
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd
    
    FncDelete = False                                                             '☜: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
    If DbDelete = False Then
        Exit Function
    End If
        
    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim lRow
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	
	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

     ggoSpread.Source = frm1.vspdData
	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
				
				.vspdData.Col = C_DILIG_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","근태코드","X")
					.vspdData.focus
       	            exit function
				 end if 

                .vspdData.Col = C_DILIG_CNT
                if  .vspdData.Text = "" then
                    .vspdData.Text = 0
                end if
                if  Cint(.vspdData.Text) > 0 then
                else
                    call  DisplayMsgBox("800443", "x","횟수","0")
                    .vspdData.Action = 0 ' go to
                    exit function
                end if
            end if
        next

    end with

    if  frm1.txtCrt_strt_yy.value <> "" AND frm1.txtCrt_end_yy.value <> "" then
        if   UNICDbl(frm1.txtCrt_strt_yy.value) >  UNICDbl(frm1.txtCrt_end_yy.value) then
            call  DisplayMsgBox("970027", "x","월차지급방법","x")
            frm1.txtCrt_strt_yy.focus
            Set gActiveElement = document.ActiveElement   
            exit function
        end if

        if  frm1.txtCrt_strt_yy.value = frm1.txtCrt_end_yy.value then
            if  frm1.txtCrt_strt_mm.value <> "" AND frm1.txtCrt_end_mm.value <> "" then
                if  frm1.txtCrt_strt_mm.value > frm1.txtCrt_end_mm.value then
                    call  DisplayMsgBox("970027", "x","월차지급방법","x")
                    frm1.txtCrt_strt_mm.focus
                    Set gActiveElement = document.ActiveElement   
                    exit function
                end if

                if  frm1.txtCrt_strt_mm.value = frm1.txtCrt_end_mm.value then
                    if  frm1.txtCrt_end_dd.value = "00" then
                    else
                        if  (frm1.txtCrt_strt_dd.value = "00") OR _
                            (frm1.txtCrt_strt_dd.value > frm1.txtCrt_end_dd.value) then
                            call  DisplayMsgBox("970027", "x","월차지급방법","x")
                            frm1.txtCrt_strt_dd.focus
                            Set gActiveElement = document.ActiveElement   
                            exit function
                        end if
                    end if
                end if
            end if
        END IF
    end if

    if  frm1.txtUse_strt_yy.value <> "" AND frm1.txtUse_end_yy.value <> "" then
        if   UNICDbl(frm1.txtUse_strt_yy.value) >  UNICDbl(frm1.txtUse_end_yy.value) then
            call  DisplayMsgBox("970027", "x","월차지급방법","x")
            frm1.txtUse_strt_yy.focus
            Set gActiveElement = document.ActiveElement   
            exit function
        end if

        if  frm1.txtUse_strt_yy.value = frm1.txtUse_end_yy.value then
            if  frm1.txtUse_strt_mm.value <> "" AND frm1.txtUse_end_mm.value <> "" then
                if  frm1.txtUse_strt_mm.value > frm1.txtUse_end_mm.value then
                    call  DisplayMsgBox("970027", "x","월차지급방법","x")
                    frm1.txtUse_strt_mm.focus
                    Set gActiveElement = document.ActiveElement   
                    exit function
                end if

                if  frm1.txtUse_strt_mm.value = frm1.txtUse_end_mm.value then
                    if  frm1.txtUse_end_dd.value = "00" then
                    else
                        if  (frm1.txtUse_strt_dd.value = "00") OR _
                            (frm1.txtUse_strt_dd.value > frm1.txtCrt_end_dd.value) then
                            call  DisplayMsgBox("970027", "x","월차지급방법","x")
                            frm1.txtUse_strt_dd.focus
                            Set gActiveElement = document.ActiveElement   
                            exit function
                        end if
                    end if
                end if
            end if
        END IF
    end if

    Call MakeKeyStream("X")
    
    If DbSave = False Then
        Exit Function
    End If
            
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
             ggoSpread.CopyRow
			 SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	With Frm1.VspdData
           .Col  = C_DILIG
           .Row  = .ActiveRow
           .Text = ""

           .Col  = C_DILIG_NM
           .Row  = .ActiveRow
           .Text = ""

	End With

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1            
        
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

        If Frm1.vspdData.MaxRows < 1 then
           Exit function
	    End if	
        With Frm1.vspdData 
        	.focus
        	 ggoSpread.Source = frm1.vspdData 
        	lDelRows =  ggoSpread.DeleteRow
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
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	 ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim pP21011
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

	If LayerShowHide(1)	= False Then
		Exit Function
	End If
	
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	    With Frm1
    
           For lRow = 1 To .vspdData.MaxRows
    
               .vspdData.Row = lRow
               .vspdData.Col = 0

               Select Case .vspdData.Text
                   Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                       strVal = strVal & "C" & parent.gColSep
                                                       strVal = strVal & lRow & parent.gColSep
                                                       strVal = strVal & Trim(.txtAllow_cd.Value) & parent.gColSep
                        .vspdData.Col = C_DILIG      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DILIG_CNT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep   
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                       strVal = strVal & "U" & parent.gColSep
                                                       strVal = strVal & lRow & parent.gColSep
                                                       strVal = strVal & Trim(.txtAllow_cd.Value) & parent.gColSep
                        .vspdData.Col = C_DILIG      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DILIG_CNT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep   
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                       strDel = strDel & "D" & parent.gColSep
                                                       strDel = strDel & lRow & parent.gColSep
                                                       strDel = strDel & Trim(.txtAllow_cd.Value) & parent.gColSep
                        .vspdData.Col = C_DILIG 	 : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
	
	    End With

    Frm1.txtMaxRows.value = lGrpCnt-1	
	Frm1.txtSpread.value = strDel & strVal
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If LayerShowHide(1) = False Then
		Exit Function
	End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtKeyStream=" & Trim(frm1.txtAllow_cd.value)             '☜: 

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    if frm1.txtProv_type.value = "" then
        lgIntFlgMode      =  parent.OPMD_CMODE
    else
	    lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	end if
    lgBlnFlgChgValue = False
	Call SetToolbar("1111111100111111")												'⊙: Set ToolBar
    Call InitData()
    Call  ggoOper.LockField(Document, "Q")    
    call txtProv_type_OnChange()
    lgBlnFlgChgValue = false
	frm1.vspdData.focus								
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    lgBlnFlgChgValue = False
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
End Function

'========================================================================================================
' Name : SubOpenCollateralNoPop()
' Desc : developer describe this line Call Master L/C No PopUp
'========================================================================================================
Sub SubOpenCollateralNoPop()
	Dim strRet
	If gblnWinEvent = True Then Exit Sub
	gblnWinEvent = True
		
	strRet = window.showModalDialog("s1413pa1.asp", "", _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
       Exit Sub
	Else
       Call SetCollateralNo(strRet)
	End If	
End Sub

'========================================================================================================
' Name : OpenAllowCd()
' Desc : developer describe this line 
'========================================================================================================
Function OpenAllowCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True

	arrParam(0) = "수당코드 팝업"		' 팝업 명칭 
	arrParam(1) = "HDA010T"				 	' TABLE 명칭 
	arrParam(2) = frm1.txtAllow_cd.value	' Code Condition
	arrParam(3) = ""						' Name Cindition
	arrParam(4) = " pay_cd=" & FilterVar("*", "''", "S") & "  AND code_type=" & FilterVar("1", "''", "S") & " "' Where Condition
	arrParam(5) = "수당코드"			
	
    arrField(0) = "allow_cd"				' Field명(0)
    arrField(1) = "allow_nm"				' Field명(1)
    
    arrHeader(0) = "수당코드"			' Header명(0)
    arrHeader(1) = "수당코드명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Frm1.txtAllow_cd.focus	
		Exit Function
	Else
		Call SubSetAllow(arrRet)
	End If	
	
End Function

'======================================================================================================
'	Name : SetAllow()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetAllow(arrRet)
	With Frm1
		.txtAllow_cd.value = arrRet(0)
		.txtAllow_nm.value = arrRet(1)	
		.txtAllow_cd.focus	
	End With
End Sub

'======================================================================================================
' Name : OpenDiligPopup
' Desc : OpenDiligPopup Reference Popup
'======================================================================================================
Function OpenDiligPopup(Byval strCode, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

     ggoSpread.Source = frm1.vspdData

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		arrParam(1) = "HCA010T"         ' TABLE 명칭 
		arrParam(2) = strCode								           ' Code Condition
		arrParam(3) = ""									           ' Name Cindition
		arrParam(4) = ""		                                       ' Where Condition
		arrParam(5) = "근태코드"	         			           ' TextBox 명칭 
	
		arrField(0) = "dilig_cd"			    			           ' Field명(0)
		arrField(1) = "dilig_nm"		    				           ' Field명(1)
    
		arrHeader(0) = "근태코드"						           ' Header명(0)
		arrHeader(1) = "근태코드명"						           ' Header명(1)

	arrParam(0) = arrParam(5)							               ' 팝업 명칭 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.action = 0	
		Exit Function
	Else
		Call SetDiligPopUp(arrRet)
	End If	
	
End Function

'========================================================================================================
' Name : SetCollateralNo()
' Desc : developer describe this line 
'========================================================================================================
Function SetCollateralNo(arrRet)
	frm1.txtGlNo.Value = arrRet
End Function

'========================================================================================================
' Name : SetBizPartner()
' Desc : developer describe this line 
'========================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtCustomer.Value = arrRet(0)
	frm1.txtCustomerNm.Value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : SetDiligPopUp()
' Desc : OpenSalesPlanPopup에서 Return되는 값 setting
'========================================================================================================
Function SetDiligPopUp(Byval arrRet)

	With frm1
		.vspdData.Col = C_DILIG_NM
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_DILIG
		.vspdData.Text = arrRet(0)
		.vspdData.action =0
	End With

	lgBlnFlgChgValue = True
	
End Function

'========================================================================================================
' Name : SetCurrency
' Desc : developer describe this line 
'========================================================================================================
Function SetCurrency(arrRet)
	frm1.txtCurrency.Value = arrRet(0)
	lgBlnFlgChgValue = True
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row     
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		 ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case C_DILIG_POP
				.Col = Col - 1
				.Row = Row
				Call OpenDiligPopup(.Text, Row)
			End Select
		End If
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        Case C_DILIG
            Frm1.vspdData.Row = Row

            If Trim(Frm1.vspdData.text) = "" Then
                Frm1.vspdData.Col = C_DILIG_NM
                Frm1.vspdData.text = ""
            Else
                iDx =  CommonQueryRs(" dilig_nm "," HCA010T "," dilig_cd =  " & FilterVar(Frm1.vspdData.Text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                If  iDx = true then
                    Frm1.vspdData.Col = C_DILIG_NM
                    Frm1.vspdData.text = Replace(lgF0, Chr(11), "")
                Else
                    Frm1.vspdData.Col = C_DILIG_NM
                    Frm1.vspdData.text = ""
                    Call  DisplayMsgBox("970000", "x","근태코드","x")
                    Frm1.vspdData.Row = Row
                    Frm1.vspdData.Col = C_DILIG
                    Frm1.vspdData.Action = 0 ' go to 
                End if
            End if
    End Select
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
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

Sub txtProv_type_OnChange()

    lgBlnFlgChgValue = True

	If Trim(frm1.txtProv_type.value) <> "" Then 
		Select Case frm1.txtProv_type.value    
		    Case 1
		         ReleaseTag(frm1.txtaccum_mm)
		         Call ggoOper.SetReqAttr(frm1.txtaccum_cnt, "D")
		         ProtectTag(frm1.txtmm_accum)
		         ProtectTag(frm1.txtuse_mm)
		         ProtectTag(frm1.txtcrt_strt_yy)
		         ProtectTag(frm1.txtcrt_strt_mm)
		         ProtectTag(frm1.txtcrt_strt_dd)
		         ProtectTag(frm1.txtcrt_end_yy)
		         ProtectTag(frm1.txtcrt_end_mm)
		         ProtectTag(frm1.txtcrt_end_dd)
		         ProtectTag(frm1.txtuse_strt_yy)
		         ProtectTag(frm1.txtuse_strt_mm)
		         ProtectTag(frm1.txtuse_strt_dd)
		         ProtectTag(frm1.txtuse_end_yy)
		         ProtectTag(frm1.txtuse_end_mm)
		         ProtectTag(frm1.txtuse_end_dd)
		         Call ggoOper.SetReqAttr(frm1.txtDuty_cnt, "Q")
		    Case 2
		         ProtectTag(frm1.txtaccum_mm)
		         Call ggoOper.SetReqAttr(frm1.txtaccum_cnt, "Q")
		         ReleaseTag(frm1.txtmm_accum)
		         ReleaseTag(frm1.txtuse_mm)
		         ProtectTag(frm1.txtcrt_strt_yy)
		         ProtectTag(frm1.txtcrt_strt_mm)
		         ProtectTag(frm1.txtcrt_strt_dd)
		         ProtectTag(frm1.txtcrt_end_yy)
		         ProtectTag(frm1.txtcrt_end_mm)
		         ProtectTag(frm1.txtcrt_end_dd)
		         ProtectTag(frm1.txtuse_strt_yy)
		         ProtectTag(frm1.txtuse_strt_mm)
		         ProtectTag(frm1.txtuse_strt_dd)
		         ProtectTag(frm1.txtuse_end_yy)
		         ProtectTag(frm1.txtuse_end_mm)
		         ProtectTag(frm1.txtuse_end_dd)
		         Call ggoOper.SetReqAttr(frm1.txtDuty_cnt, "Q")
		    Case 3
		         ProtectTag(frm1.txtaccum_mm)
		         Call ggoOper.SetReqAttr(frm1.txtaccum_cnt, "Q")
		         ProtectTag(frm1.txtmm_accum)
		         ProtectTag(frm1.txtuse_mm)
		         ReleaseTag(frm1.txtcrt_strt_yy)
		         ReleaseTag(frm1.txtcrt_strt_mm)
		         ReleaseTag(frm1.txtcrt_strt_dd)
		         ReleaseTag(frm1.txtcrt_end_yy)
		         ReleaseTag(frm1.txtcrt_end_mm)
		         ReleaseTag(frm1.txtcrt_end_dd)
		         ReleaseTag(frm1.txtuse_strt_yy)
		         ReleaseTag(frm1.txtuse_strt_mm)
		         ReleaseTag(frm1.txtuse_strt_dd)
		         ReleaseTag(frm1.txtuse_end_yy)
		         ReleaseTag(frm1.txtuse_end_mm)
		         ReleaseTag(frm1.txtuse_end_dd)
		         Call ggoOper.SetReqAttr(frm1.txtDuty_cnt, "D")
		End Select
    End If
End Sub

Sub txtAccum_mm_OnChange()
    lgBlnFlgChgValue = True
    frm1.txtAccum_mm1.value = frm1.txtAccum_mm.value
End Sub
Sub txtMm_accum_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtCrt_strt_yy_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtCrt_strt_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtCrt_end_yy_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtCrt_end_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_strt_yy_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_strt_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_end_yy_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_end_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtAccum_cnt_Change()
    lgBlnFlgChgValue = True
    frm1.txtAccum_cnt1.text = frm1.txtAccum_cnt.text
End Sub

Sub txtDuty_cnt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCrt_strt_dd_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtCrt_end_dd_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_strt_dd_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtUse_end_dd_OnChange()
    lgBlnFlgChgValue = True
End Sub

function txtAllow_cd_OnChange()

Dim IntRetCd
    
    If frm1.txtAllow_cd.value = "" Then
		frm1.txtAllow_nm.value = ""
    Else
        IntRetCd =  CommonQueryRs(" ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "   AND ALLOW_CD =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call  DisplayMsgBox("800145","X","X","X")  '수당정보에 등록되지 않은 코드입니다.
			frm1.txtAllow_nm.value = ""
            frm1.txtAllow_cd.focus            
			txtAllow_cd_Onchange = true
            Exit Function          
        Else
			frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월차수당기준등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>수당코드</TD>
                                    <TD CLASS=TD6 NOWRAP>
                                    
										<INPUT TYPE=TEXT NAME="txtAllow_cd" MAXLENGTH=3 SIZE=10 MAXLENGTH=8 tag=12XXXU ALT="수당코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWarrentNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAllowCd()">
										<INPUT TYPE=TEXT NAME="txtAllow_nm" tag="14X"></TD>

									<TD CLASS=TDT NOWRAP></TD>
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
					<TD WIDTH=100% HEIGHT=200 VALIGN=TOP>
                        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>월차지급방법</LEGEND>
					        <TABLE WIDTH=100% HEIGHT="100%" CELLSPACING=0>
	 						    <TR>
              				        <TD CLASS="TD5" NOWRAP>지급방법&nbsp;
              				            <SELECT NAME="txtProv_type" ALT="지급방법" STYLE="WIDTH: 120px" TAG="22N"></SELECT>
              				        </TD>
	                   				<TD CLASS="TD656">
                                        <BR>
	                   				    <SELECT NAME="txtAccum_mm" ALT="누계" STYLE="WIDTH:100px" TAG="21N"></SELECT>누계 >&nbsp;
	                   				    <OBJECT classid=<%=gCLSIDFPDS%> ID=fpDoubleSingle2 NAME=txtaccum_cnt STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="21X7" ALT=""></OBJECT>&nbsp;이면 
	                   				    <SELECT NAME="txtAccum_mm1" ALT="" STYLE="WIDTH:100px" TAG="24N"></SELECT> - 
	                   				    <OBJECT classid=<%=gCLSIDFPDS%> ID=fpDoubleSingle2 NAME=txtAccum_cnt1 STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="24X7" ALT=""></OBJECT>&nbsp;= 지급<BR>
                                        <BR>
	                   				    <SELECT NAME="txtMm_accum" ALT="" STYLE="WIDTH:100px" TAG="21N"></SELECT>적치 -&nbsp;<SELECT NAME="txtUse_mm" ALT="" STYLE="WIDTH: 100px" TAG="21N"></SELECT> 사용일수 = 지급<BR>
                                        <BR>

	                   				    월차 의무 사용 갯수&nbsp;<OBJECT classid=<%=gCLSIDFPDS%> ID=fpDoubleSingle2 NAME=txtDuty_cnt STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="21X7Z" ALT=""></OBJECT>&nbsp;일<BR> 
                                        <BR>

	                   				    <SELECT NAME="txtCrt_strt_yy" ALT="" STYLE="WIDTH:100px" TAG="21N"></SELECT>&nbsp;
	                   				    <SELECT NAME="txtCrt_strt_mm" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>월 
                                        <SELECT NAME="txtCrt_strt_dd" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>일&nbsp;~&nbsp;

	                   				    <SELECT NAME="txtCrt_end_yy" ALT="" STYLE="WIDTH:100px" TAG="21N"></SELECT>&nbsp;
	                   				    <SELECT NAME="txtCrt_end_mm" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>월 
                                        <SELECT NAME="txtCrt_end_dd" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>일까지 발생분 - <BR>

	                   				    <SELECT NAME="txtUse_strt_yy" ALT="" STYLE="WIDTH:100px" TAG="21N"></SELECT>&nbsp;
	                   				    <SELECT NAME="txtUse_strt_mm" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>월 
                                        <SELECT NAME="txtUse_strt_dd" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>일&nbsp;~&nbsp;

	                   				    <SELECT NAME="txtUse_end_yy" ALT="" STYLE="WIDTH:100px" TAG="21N"></SELECT>&nbsp;
	                   				    <SELECT NAME="txtUse_end_mm" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>월 
                                        <SELECT NAME="txtUse_end_dd" ALT="" STYLE="WIDTH:70px" TAG="21N"></SELECT>일까지 사용분 = 지급<BR>
                                        <BR>

	                   				</TD>
	                   			</TR>
						    </TABLE>
					    </FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
						    <TR>
						    	<TD HEIGHT="100%" WIDTH="100%">
						    		<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
						    			<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
						    		</OBJECT>
						    	</TD>
						    </TR>
						</TABLE>
                    </TD>
                </TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
