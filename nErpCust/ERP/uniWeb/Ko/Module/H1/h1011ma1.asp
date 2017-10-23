<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : ���ϼ�����ص�� 
*  3. Program ID           : H1011ma1
*  4. Program Name         : H1011ma1
*  5. Program Desc         : ������������/���ϼ�����ص�� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/12
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee Sina
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
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H1011mb1.asp"                                      'Biz Logic ASP 
Const C_SHEETMAXROWS    = 21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim C_DEPT_CD
Dim C_DEPT_CD_POP
Dim C_DEPT_CD_NM
Dim C_MAN_AMT
Dim C_WOMAN_AMT

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_DEPT_CD     = 1	
	 C_DEPT_CD_POP = 2
	 C_DEPT_CD_NM  = 3
	 C_MAN_AMT	   = 4
	 C_WOMAN_AMT   = 5	  
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Frm1.txtAllow_cd.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
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
        .MaxCols = C_WOMAN_AMT + 1												<%'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															<%'������Ʈ�� ��� Hidden Column%>
        .ColHidden = True
                
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	       
       Call  GetSpreadColumnPos("A")

		 ggoSpread.SSSetEdit   C_DEPT_CD,       "�μ�", 20,,,13,2
		 ggoSpread.SSSetButton C_DEPT_CD_POP
		 ggoSpread.SSSetEdit   C_DEPT_CD_NM,    "�μ���", 44,,,44,2
		 ggoSpread.SSSetFloat  C_MAN_AMT,       "���ڼ����", 25, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		 ggoSpread.SSSetFloat  C_WOMAN_AMT,     "���ڼ����", 25, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		 
		 Call ggoSpread.MakePairsColumn(C_DEPT_CD,  C_DEPT_CD_POP)
	
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

     ggoSpread.SpreadLock    C_DEPT_CD ,		-1, C_DEPT_CD
     ggoSpread.SpreadLock    C_DEPT_CD_POP , -1, C_DEPT_CD_POP
     ggoSpread.SpreadLock    C_DEPT_CD_NM ,	-1, C_DEPT_CD_NM
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
    
         ggoSpread.SSSetRequired     C_DEPT_CD	 , pvStartRow, pvEndRow
         ggoSpread.SSSetprotected    C_DEPT_CD_NM , pvStartRow, pvEndRow
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
            
            C_DEPT_CD     = iCurColumnPos(1)	
			C_DEPT_CD_POP = iCurColumnPos(2)
			C_DEPT_CD_NM  = iCurColumnPos(3)
			C_MAN_AMT	  = iCurColumnPos(4)
			C_WOMAN_AMT   = iCurColumnPos(5)            
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format

    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'��: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call SetToolbar("1100110100101111")										        '��ư ���� ���� 

    Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")

    ' �����ڵ忡 ���� 
    Call  CommonQueryRs(" MAX(allow_cd) "," hda130t ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtAllow_cd.value = Trim(Replace(lgF0,Chr(11),""))

    Call  CommonQueryRs(" allow_nm "," HDA010T "," allow_cd =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))

    frm1.txtAllow_cd.Focus
    
	Call CookiePage (0)                                                             '��: Check Cookie
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
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field
    ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If
    
    If txtAllow_cd_Onchange() Then                                                '��: enter key �� ��ȸ�� �����ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if

    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")
    
    If DbQuery = False Then
        Exit Function
    End If
              
    FncQuery = True                                                              '��: Processing is OK
    
End Function


'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '��: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '��: Lock  Field
	Call SetToolbar("1110111100111111")							                 '��: Set ToolBar
    Call InitVariables                                                           '��: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '��: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"x","x")                        '��: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
	Call  DisableToolBar( parent.TBC_DELETE)
    If DbDelete = False Then
        Call  RestoreToolBar()
        Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim lRow
    Dim strAllow_cd
    Dim zeroChk
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    strAllow_cd = ""    
    
    IntRetCd =  CommonQueryRs(" distinct ALLOW_CD "," HDA130T "," ALLOW_CD <>  " & FilterVar(frm1.txtAllow_cd.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strAllow_cd = Trim(Replace(lgF0,Chr(11),""))

    If IsNull(strAllow_cd) OR strAllow_cd = "" then
    Else
        Call  DisplayMsgBox("800492","x",UCase(strAllow_cd),"���ϼ���")                        
        Exit Function          
    End if 
    
     ggoSpread.Source = frm1.vspdData
	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
                
                .vspdData.Col = C_DEPT_CD_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","�μ��ڵ�","X")
					.vspdData.focus
       	            exit function
				 end if 

				zeroChk = 0
				
                .vspdData.Col = C_MAN_AMT
                if  .vspdData.Text = "" then
                    .vspdData.Text = 0
                end if
                if   UNICDbl(.vspdData.Text) <= 0 then
					zeroChk = zeroChk + 1
                end if

                .vspdData.Col = C_WOMAN_AMT
                if  .vspdData.Text = "" then
                    .vspdData.Text = 0
                end if
                if   UNICDbl(.vspdData.Text) <= 0 then
					zeroChk = zeroChk + 1
                end if
                if   zeroChk =2 then
                    call  DisplayMsgBox("800410", "x","x","x")
					.vspdData.Col = C_MAN_AMT                    
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if 

            end if
        next

    end with

    Call MakeKeyStream("X")    
    
    If DbSave = False Then
        Exit Function
    End If
            
    FncSave = True                                                              '��: Processing is OK
    
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
           .Col  = C_DEPT_CD
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_DEPT_CD_NM
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
	Dim imRow

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
 
    FncInsertRow = False                                                         '��: Processing is NG

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
       FncInsertRow = True                                                          '��: Processing is OK
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
    Call parent.FncExport( parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
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
    
    Err.Clear                                                                        '��: Clear err status

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
		
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
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
	
    DbSave = False                                                          
    
    If LayerShowHide(1) = False Then
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
 
               Case  ggoSpread.InsertFlag                                      '��: Create
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                   .vspdData.Col = C_DEPT_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                : strVal = strVal & Trim(.txtAllow_cd.value) & parent.gColSep
                   .vspdData.Col = C_MAN_AMT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_WOMAN_AMT	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '��: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                   .vspdData.Col = C_DEPT_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                : strVal = strVal & Trim(.txtAllow_cd.value) & parent.gColSep
                   .vspdData.Col = C_MAN_AMT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_WOMAN_AMT	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '��: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                   .vspdData.Col = C_DEPT_CD	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                                                : strDel = strDel & Trim(.txtAllow_cd.value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
       .txtUpdtUserId.value  =  parent.gUsrID
       .txtInsrtUserId.value =  parent.gUsrID
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
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbDelete = False			                                                 '��: Processing is NG
		
	if LayerShowHide(1) = False Then 
		Exit Function
	End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '��: Delete
	strVal = strVal & "&txtAllow_cd=" & Trim(frm1.txtAllow_cd.value)             '��: 
    strVal = strVal & "&txtPrevNext=" & ""	                             '��: Direction
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run Biz logic
	
	DbDelete = True                                                              '��: Processing is NG

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'��: Lock field
    Call InitData()
	Call SetToolbar("110111110011111")									
	frm1.vspdData.focus								
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call  ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call InitVariables															'��: Initializes local global variables
	Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If
	ggoSpread.ClearSpreadData
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
End Function

'======================================================================================================
'	Name : OpenCode()
'	Description : 
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_DEPT_CD_POP
	        arrParam(0) = "�μ��ڵ� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "H_CURRENT_DEPT"		  			    ' TABLE ��Ī 
	    	arrParam(2) = strCode                            		' Code Condition
	    	arrParam(3) = ""                            		' Name Cindition
	    	arrParam(4) = ""
	    	arrParam(5) = "�μ��ڵ�" 			            ' TextBox ��Ī 
	
	    	arrField(0) = "dept_cd"						    	' Field��(0)
	    	arrField(1) = "dept_nm"    					    	' Field��(1)
    
	    	arrHeader(0) = "�μ��ڵ�"	   		    	    ' Header��(0)
	    	arrHeader(1) = "�μ��ڵ��"	    		        ' Header��(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.action = 0		
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : 
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_DEPT_CD_POP
		    	.vspdData.Col = C_DEPT_CD_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_DEPT_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action = 0		    	
        End Select

	End With

End Function
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

	arrParam(0) = "�����ڵ� �˾�"		' �˾� ��Ī 
	arrParam(1) = "HDA010T"				 	' TABLE ��Ī 
	arrParam(2) = frm1.txtAllow_cd.value	' Code Condition
	arrParam(3) = ""                    	' Name Cindition
	arrParam(4) = " pay_cd=" & FilterVar("*", "''", "S") & "  AND code_type=" & FilterVar("1", "''", "S") & " "' Where Condition
	arrParam(5) = "�����ڵ�"			
	
    arrField(0) = "allow_cd"				' Field��(0)
    arrField(1) = "allow_nm"				' Field��(1)
    
    arrHeader(0) = "�����ڵ�"			' Header��(0)
    arrHeader(1) = "�����ڵ��"			' Header��(1)
    
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetAllow(arrRet)
	With Frm1
		.txtAllow_cd.value = arrRet(0)
		.txtAllow_nm.value = arrRet(1)	
		.txtAllow_cd.focus	
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		 ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			    Case C_DEPT_CD_POP
			    	.Col = Col - 1
			    	.Row = Row
                    Call OpenCode(.text, C_DEPT_CD_POP, Row)
			End Select
		End If
    
	End With
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_DEPT_CD
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_DEPT_CD
    
            If Frm1.vspdData.value = "" Then
   	            Frm1.vspdData.Col = C_DEPT_CD_NM
   	            Frm1.vspdData.value = "" 
            Else
                IntRetCd =  CommonQueryRs(" dept_nm "," H_CURRENT_DEPT "," dept_cd =  " & FilterVar(Frm1.vspdData.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                
                If IntRetCd = false then
			        Call  DisplayMsgBox("800062","X","X","X")	'�μ������� ��ϵǾ� ���� ���� �ڵ��Դϴ�.
  	                Frm1.vspdData.Col = C_DEPT_CD_NM
                    Frm1.vspdData.value = ""
                Else
		       	    Frm1.vspdData.Col = C_DEPT_CD_NM
		       	    Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
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

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
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


function txtAllow_cd_OnChange()
    Dim IntRetCd
    
    If frm1.txtAllow_cd.value = "" Then
		frm1.txtAllow_nm.value = ""
		txtAllow_cd_Onchange = false 
    Else
        IntRetCd =  CommonQueryRs(" ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "   AND ALLOW_CD =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call  DisplayMsgBox("800145","X","X","X")  '���������� ��ϵ��� ���� �ڵ��Դϴ�.
			frm1.txtAllow_nm.value = ""
            frm1.txtAllow_cd.focus
		    Call  ggoOper.ClearField(Document, "2")	            
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ϼ�����ص��</font></td>
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
									<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
                                    <TD CLASS=TD6 NOWRAP>
                                    
										<INPUT TYPE=TEXT NAME="txtAllow_cd" MAXLENGTH=3 SIZE=10 MAXLENGTH=8 tag=12XXXU ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWarrentNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAllowCd()">
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h1011ma1_vaSpread1_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
