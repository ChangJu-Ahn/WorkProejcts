<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 경력등록 
*  3. Program ID           : H2007ma1
*  4. Program Name         : H2007ma1
*  5. Program Desc         : 인사기본자료관리/경력등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/10
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H2007mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_JUMP_ID = "H2001ma1" 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim C_CAREER_START
Dim C_CAREER_END 
Dim C_COMP_NM 
Dim C_ROLL_PSTN 
Dim C_FUNC_NM
Dim C_CAREER_YY 
Dim C_CAREER_MM 
Dim C_PAYROLL_UPDATE_CD
Dim C_PAYROLL_UPDATE 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	 C_CAREER_START = 1
	 C_CAREER_END = 2
	 C_COMP_NM = 3
	 C_ROLL_PSTN = 4
	 C_FUNC_NM = 5
	 C_CAREER_YY = 6
	 C_CAREER_MM = 7
	 C_PAYROLL_UPDATE_CD  = 8
	 C_PAYROLL_UPDATE  = 9
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)

	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Frm1.txtEmp_no.Value & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtName.Value & parent.gColSep
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    
    iCodeArr = "N" & Chr(11) & "Y" & Chr(11)
    iNameArr = "미포함" & Chr(11) & "포함" & Chr(11)
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAYROLL_UPDATE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAYROLL_UPDATE         ''''''''DB에서 불러 gread에서 
End Sub
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_PAYROLL_UPDATE_CD
			intIndex = .value
			.col = C_PAYROLL_UPDATE
			.value = intindex
		Next	
	End With
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
	    .ReDraw = false
        .MaxCols = C_PAYROLL_UPDATE + 1												'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0        
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData     
        
	    ggoSpread.ClearSpreadData
		Call GetSpreadColumnPos("A")  	
	    Call AppendNumberPlace("6","2","0")

        ggoSpread.SSSetDate     C_CAREER_START, "시작일",     13,2, gDateFormat
        ggoSpread.SSSetDate     C_CAREER_END,   "종료일",     13,2, gDateFormat
        ggoSpread.SSSetEdit     C_COMP_NM,      "회사명",   25,,, 40,1   
        ggoSpread.SSSetEdit     C_ROLL_PSTN,    "직위",     15,,, 10,1   
        ggoSpread.SSSetEdit     C_FUNC_NM,      "담당업무", 25,,, 40,1  
        ggoSpread.SSSetFloat    C_CAREER_YY,    "경력(년)", 12,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","99"
        ggoSpread.SSSetFloat    C_CAREER_MM,    "경력(월)", 12,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, , , ,"0","11"
		ggoSpread.SSSetCombo    C_PAYROLL_UPDATE_CD,   "payrollupdate",  10
		ggoSpread.SSSetCombo    C_PAYROLL_UPDATE,   "인정경력포함여부",  15

        Call ggoSpread.SSSetColHidden(C_PAYROLL_UPDATE_CD,C_PAYROLL_UPDATE_CD,True)
        		
	    .ReDraw = true
        Call SetSpreadLock 
    
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

			 C_CAREER_START = iCurColumnPos(1)
			 C_CAREER_END = iCurColumnPos(2)
			 C_COMP_NM = iCurColumnPos(3)
			 C_ROLL_PSTN = iCurColumnPos(4)
			 C_FUNC_NM = iCurColumnPos(5)
			 C_CAREER_YY = iCurColumnPos(6)
			 C_CAREER_MM = iCurColumnPos(7)
            C_PAYROLL_UPDATE_CD = iCurColumnPos(8)
            C_PAYROLL_UPDATE  = iCurColumnPos(9)
            
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
        ggoSpread.SpreadLock    C_CAREER_START, -1, C_CAREER_START
        ggoSpread.SpreadLock    C_CAREER_END, -1, C_CAREER_END
        ggoSpread.SSSetRequired    C_COMP_NM, -1, C_COMP_NM
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1   
		ggoSpread.SSSetRequired    C_PAYROLL_UPDATE, -1, C_COMP_NM         
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
        ggoSpread.SSSetRequired		C_CAREER_START, pvStartRow, pvEndRow      
        ggoSpread.SSSetRequired		C_CAREER_END, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_COMP_NM, pvStartRow, pvEndRow
      ggoSpread.SSSetRequired   C_PAYROLL_UPDATE_CD, pvStartRow, pvEndRow
      ggoSpread.SSSetRequired   C_PAYROLL_UPDATE, pvStartRow, pvEndRow
      
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
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
    Call InitComboBox    
    frm1.txtEMP_NO.Focus
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

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    if  frm1.txtEmp_no.value = "" AND frm1.txtName.value <> "" then
        OpenEmp()
        exit function
    else
        If  txtEmp_no_Onchange() then
            Exit Function
        End If
    end if

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
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
    Dim strstartDt
    Dim strendDt
    dim strEntr_dt
    Dim lRow
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    if  frm1.txtEmp_no.value = "" then
        Frm1.txtEmp_no.focus
        Set gActiveElement = document.ActiveElement   
        exit function
    end if

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_CAREER_START
                    strstartDt = UniConvDateToYYYYMMDD(.vspdData.text,gDateFormat,"")

   	                .vspdData.Col = C_CAREER_END
                    strendDt = UniConvDateToYYYYMMDD(.vspdData.text,gDateFormat,"")

                    if strstartDt <> "" and strendDt <> "" then
                        if  strstartDt >= strendDt then
                            Call DisplayMsgBox("800033","X","X","X")
        	                .vspdData.Col = C_CAREER_START
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        end if
                    end if

   	                .vspdData.Col = C_CAREER_MM
                    strstartDt = .vspdData.Text
                    if  strstartDt <> "" then
                        if  strstartDt > 11 then
                            Call DisplayMsgBox("970027","X","경력(월)","X")
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        end if
                    end if
                                  
                    strEntr_dt = UniConvDateToYYYYMMDD(frm1.txtEntr_dt.Text,gDateFormat,"")
                    if strEntr_dt <= strendDt then
                            Call DisplayMsgBox("800267","X","X","X")      
                            Exit Function                                          
                    end if
            End Select
        Next
	End With

    Call MakeKeyStream("X")
    
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

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
           .Col  = C_CAREER_START
           .Row  = .ActiveRow
           .Text = ""

           .Col  = C_CAREER_END
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
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
End Function
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
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
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
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
    
    Err.Clear                                                                        '☜: Clear err status

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
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

	Dim strRes_no

    DbSave = False                                                          
    
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                    strVal = strVal & "C" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_CAREER_START: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_END  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COMP_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_YY   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_MM   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAYROLL_UPDATE_CD	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep                           
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_CAREER_START: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_END  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COMP_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_YY   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_MM   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAYROLL_UPDATE_CD	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep  
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                                                    strDel = strDel & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_CAREER_START: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CAREER_END  : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAYROLL_UPDATE_CD	 : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep                           
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete = False Then
		Call RestoreToolBar()
        Exit Function
    End If
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")
	Frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
	call MainQuery()
	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 		
		Set gActiveElement = document.ActiveElement
        call txtEmp_no_Onchange()
		lgBlnFlgChgValue = False
		.txtEmp_no.focus
	End With
	
End Sub

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

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData.Row = Row
	
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
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
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
	Dim fr_dt, to_dt, strSelect, IntRetCd
	Dim strData, strYear, strMonth
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With frm1
        Select Case Col
             Case  C_CAREER_START, C_CAREER_END
					Frm1.vspdData.Col = C_CAREER_START             
					fr_dt = Trim(.vspdData.Text)
					Frm1.vspdData.Col = C_CAREER_END
					to_dt = Trim(.vspdData.Text)
					
                    If fr_dt <> "" and to_dt <> "" Then
                        strSelect = " dbo.ufn_H_GetLongYYMMDD( " & FilterVar(UNIConvDateCompanyToDB(fr_dt,NULL),"NULL","S") & ","  
                        strSelect = strSelect & FilterVar(UNIConvDateCompanyToDB(to_dt,NULL),"NULL","S") & ")"  

                        IntRetCd  = CommonQueryRs(strSelect," hda000t "," 1=1",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						strData = Split(Replace(lgF0,Chr(11),""),"/")

						strYear = strData(0)
						strMonth = strData(1)

						Frm1.vspdData.Col = C_CAREER_YY
						Frm1.vspdData.Text = strYear

						Frm1.vspdData.Col = C_CAREER_MM
						Frm1.vspdData.Text = strMonth

                    End If
			Case  C_PAYROLL_UPDATE     ' 반영여부 
			   iDx = Frm1.vspdData.value
   			   Frm1.vspdData.Col = C_PAYROLL_UPDATE_CD
			   Frm1.vspdData.value = iDx
            
        End Select
    End With

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
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

    Dim strName
    Dim strRel
    Dim strRelCd

	With frm1.vspdData
		.Row = Row
    
        Select Case Col
            Case C_PAYROLL_UPDATE
                .Col = Col
                intIndex = .value
				.Col = C_PAYROLL_UPDATE_CD
				.value = intIndex
		End Select
	End With
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

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
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtName.value = ""
	frm1.txtDept_nm.value = ""
	frm1.txtRoll_pstn.value = ""
	frm1.txtEntr_dt.Text = ""
	frm1.txtPay_grd.value = ""
	Frm1.imgPhoto.src = ""

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtEmp_no.value = ""
		txtEmp_no_Onchange = true
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
			strVal = "../../../CShared/image/default_picture.jpg"
    		Frm1.imgPhoto.src = strVal	    
	    
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			ggoSpread.Source = Frm1.vspdData    
			ggoSpread.ClearSpreadData             
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.Text = UNIDateClientFormat(strEntr_dt)
             'strEntr_dt는 Client Format(parent.gClientDateFormat) 그러므로 Client Format -->Company Format

			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			if   Replace(lgF0, Chr(11), "") > 0  then
				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
			else
				strVal = "../../../CShared/image/default_picture.jpg"
			end if

			Frm1.imgPhoto.src = strVal
        End if 
    End if
    
End Function 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>경력등록</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=10%>
                        <img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=80 HEIGHT=90 HSPACE=10 VSPACE=0 BORDER=1>
    	            </TD>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" SiZE=15 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_nm" ALT="부서명" TYPE="Text" SiZE=15 tag=14></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
							    <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID="txtEntr_dt" NAME="txtEntr_dt" ALT="입사일" CLASS=FPDTYYYYMMDD TITLE=FPDATETIME tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="급호" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">인사마스타</a></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

