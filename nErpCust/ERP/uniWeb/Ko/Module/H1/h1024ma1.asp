<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: ��ȣ�� �ӿ�/��� ���� ��� 
*  3. Program ID           	: H1024ma1
*  4. Program Name         	: H1024ma1
*  5. Program Desc         	: ������������ 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/28
*  8. Modified date(Last)  	: 2003/06/10
*  9. Modifier (First)     	: 
* 10. Modifier (Last)     	: Lee SiNa
* 11. Comment              	:
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H1024mb1.asp"						           '��: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 15	                                      '��: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                 
Dim gblnWinEvent                                                 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop

Dim C_PAY_GRD1
Dim C_PAY_GRD1Popup
Dim C_PAY_GRD1_NM
Dim C_EMP_TYPE
Dim C_EMP_TYPE_CD

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
		C_PAY_GRD1			= 1
		C_PAY_GRD1Popup		= 2		'POPUP
		C_PAY_GRD1_NM		= 3
		C_EMP_TYPE			= 4
		C_EMP_TYPE_CD		= 5		'HIDDEN
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
      ggoSpread.Source = frm1.vspdData

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0122", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_EMP_TYPE_CD
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_EMP_TYPE
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
			.Col = C_EMP_TYPE_CD
			intIndex = .value
			.col = C_EMP_TYPE
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
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_EMP_TYPE_CD + 1												<%'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															<%'������Ʈ�� ��� Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	       
         Call  GetSpreadColumnPos("A")
		
		 ggoSpread.SSSetEdit	 C_PAY_GRD1,          "��ȣ�ڵ�", 20
		 ggoSpread.SSSetButton   C_PAY_GRD1Popup
		 ggoSpread.SSSetEdit	 C_PAY_GRD1_NM,       "��ȣ��", 40,,,50,2 
         ggoSpread.SSSetCombo    C_EMP_TYPE,          "�������", 28
         ggoSpread.SSSetCombo    C_EMP_TYPE_CD,       "��������ڵ�", 14
         
         Call ggoSpread.MakePairsColumn(C_PAY_GRD1,	   C_PAY_GRD1Popup)
         Call ggoSpread.SSSetColHidden(C_EMP_TYPE_CD,  C_EMP_TYPE_CD, True)
				
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
         ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
         ggoSpread.SSSetProtected    C_PAY_GRD1		, -1, C_PAY_GRD1
         ggoSpread.SSSetProtected    C_PAY_GRD1_NM	, -1, C_PAY_GRD1_NM
         ggoSpread.SSSetProtected    C_PAY_GRD1Popup, -1, C_PAY_GRD1Popup
         ggoSpread.SSSetRequired     C_EMP_TYPE		, -1, C_EMP_TYPE
         ggoSpread.SSSetRequired     C_EMP_TYPE_CD	, -1, C_EMP_TYPE_CD
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
         ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
         ggoSpread.SSSetRequired	C_PAY_GRD1		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_PAY_GRD1_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_EMP_TYPE		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_EMP_TYPE_CD	, pvStartRow, pvEndRow
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
            
            C_PAY_GRD1			= iCurColumnPos(1)
			C_PAY_GRD1Popup		= iCurColumnPos(2)
			C_PAY_GRD1_NM		= iCurColumnPos(3)
			C_EMP_TYPE			= iCurColumnPos(4)
			C_EMP_TYPE_CD		= iCurColumnPos(5)
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
    
    Call  ggoOper.LockField(Document, "N")											'��: Lock Field
    
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)                                ' �ڷ����:lgUsrIntCd ("%", "1%")
    Call InitComboBox
    Call SetToolbar("1100110100010111")												'��: Set ToolBar
    
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
    
    FncQuery = False                                                            '��: Processing is NG
    
    Err.Clear                                                                   '��: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If   ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '��: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    ggoSpread.ClearSpreadData
    														'��: Initializes local global variables
    If Not chkField(Document, "1") Then									        '��: This function check indispensable field
       Exit Function
    End If

    Call InitVariables	
    Call MakeKeyStream("X")
    
    Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
              
    FncQuery = True																'��: Processing is OK

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
    
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	Dim IntRetCD 
	Dim CFlag
	Dim strStrtDt
	Dim strEndDt
	Dim strStrtDt1
	Dim strEndDt1
	Dim strStrtDt2
	Dim strEndDt2
	Dim lRow
	Dim strStrtDtType
	Dim strEndDtType

	FncSave = False                                                              '��: Processing is NG    
	Err.Clear                                                                    '��: Clear err status    
	 ggoSpread.Source = frm1.vspdData
	
	If  ggoSpread.SSCheckChange = False Then
	IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
		Exit Function
	End If    
	
	 ggoSpread.Source = frm1.vspdData
	
	  ggoSpread.Source = frm1.vspdData
	 If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	       Exit Function
	 End If
	
	 ggoSpread.Source = frm1.vspdData
	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
				.vspdData.Col = C_PAY_GRD1_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","��ȣ�ڵ�","X")
					.vspdData.focus
					
       	            exit function
				 end if 
            end if
        next

    end with
	    	
	Call MakeKeyStream("X")
	
	Call  DisableToolBar( parent.TBC_SAVE)
	IF DBSAVE =  False Then
	      Call  RestoreToolBar()
	      Exit Function
	End If
	
	FncSave = True               
    
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
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = frm1.vspdData	
     ggoSpread.EditUndo  
    Call Initdata()
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
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
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
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '��: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG

    If   LayerShowHide(1) = False Then
     	Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '��: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '��: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
	
    DbQuery = True                                                               '��: Processing is NG
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
    Dim strCloseDate
    Dim strTemp
	
    Err.Clear                                                                    '��: Clear err status
		
    DbSave = False														         '��: Processing is NG

    If   LayerShowHide(1) = False Then
    	Exit Function
    End If	

    With frm1
	    .txtMode.value        =  parent.UID_M0002                                        '��: Delete
	    .txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
    End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

    With Frm1
        ggoSpread.Source = frm1.vspdData
        
       For lRow = 1 To .vspdData.MaxRows
           
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case  ggoSpread.InsertFlag                                      '��: Update
                                                           strVal = strVal & "C" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_TYPE_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '��: Update
                                                           strVal = strVal & "U" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_TYPE_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.DeleteFlag                                      '��: Delete
                                                           strDel = strDel & "D" & parent.gColSep
                                                           strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1           : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                   
           End Select
       Next
	
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

   End With

   Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '��: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '��:
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
    If IntRetCD = vbNo Then													'------ Delete function call area ------ 
	Exit Function	
    End If    
    
    Call  DisableToolBar( parent.TBC_DELETE)					'Query ��ư�� disable��Ŵ 
    If DBDelete = False Then
  		Call  RestoreToolBar()
		Exit Function
    End If    
    
    FncDelete = True                                                        '��: Processing is OK

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim IRow

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '��: Indicates that current mode is Create mode
	
    IF  Frm1.vspdData.MaxRows > 0 then
        Call SetToolbar("1100111100111111")												'��: Set ToolBar
    else
        Call SetToolbar("1100110100111111")
    end if

    Call  ggoOper.LockField(Document, "Q")

    Set gActiveElement = document.ActiveElement   
    lgBlnFlgChgValue = False
    frm1.vspdData.focus

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
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
' Name : OpenPopUp()        
' Desc : developer describe this line 
'========================================================================================================

Function OpenPopUp(Byval IRow, Byval Part)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim intRetCd
	Dim ArgMt
	
    With Frm1
    	    
	    If IsOpenPop = True Then Exit Function
	    IsOpenPop = True		
	    
	    Select Case Part		                         
            Case "PAY_GRD1"            	 
                arrParam(0) = "��ȣ"		  	            ' �˾� ��Ī 
		        arrParam(1) = "B_MINOR"					    ' TABLE ��Ī 
		    
		        If IRow > 0 Then
		        	.vspdData.Row = IRow
		            .vspdData.Col = C_PAY_GRD1     
		        
		            If Trim(.vspdData.Text) = "" Then    		
		                arrParam(2) = "" 'Trim(Replace(lgF0,Chr(11),""))
		                arrParam(3) = ""
		            Else
		              	arrParam(2) = Trim(.vspdData.Text)
		    	        arrParam(3) = ""
		            End If
		    
		            arrParam(4) = " MAJOR_CD = " & FilterVar("H0001", "''", "S") & ""
		            arrParam(5) = "��ȣ"  			     ' TextBox ��Ī 
		    
		            ArgMt = ""
		        
		            arrField(0) = "MINOR_CD"
	                arrField(1) = "MINOR_NM"    
	                        
	                arrHeader(0) = "��ȣ�ڵ�"
	                arrHeader(1) = "��ȣ��"	        
		        End If 	
        End Select

        arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	   	                  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        	   
	    IsOpenPop = False
	
	    If arrRet(0) = "" Then
			frm1.vspdData.Col = C_PAY_GRD1     			                ' Code Condition        	    	        
			frm1.vspdData.Action = 0 ' go to 
	       	Exit Function
	    Else
	    	Call SetCode(arrRet, IRow, Part)
	    	
	    	If IRow > 0 Then	            
	    	   	 ggoSpread.Source = frm1.vspdData
           		 ggoSpread.UpdateRow IRow
      		End If
	    End If	
     End With
End Function

'======================================================================================================
'	Name : SetCode
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetCode(arrRet, IRow, Part)

	With Frm1	
		If IRow > 0 Then
			SELECT CASE Part
			  Case  "PAY_GRD1"
			    .vspdData.Row = IRow
				.vspdData.Col = C_PAY_GRD1     			                ' Code Condition        
				.vspdData.Value = arrRet(0)
				
				.vspdData.Col = C_PAY_GRD1_NM     			                ' Code Condition        
				.vspdData.Value = arrRet(1)								' Name Cindition      
				
				.vspdData.Row = IRow
				.vspdData.Col = C_PAY_GRD1     			                ' Code Condition        	    	        
				.vspdData.Action = 0 ' go to 
			  End SELECT      
		End If
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)   
    
    frm1.vspdData.Row = Row   	
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
    
	Select Case Col	
		Case C_PAY_GRD1Popup
			Call OpenPopUp(Row, "PAY_GRD1")
	End Select    
   End If         

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
    Dim iDx

    Frm1.vspdData.Row = Row
    Select Case Col
        Case C_EMP_TYPE
            Frm1.vspdData.Row = Row
            Frm1.vspdData.Col = C_EMP_TYPE
            iDx = Frm1.vspdData.value
            Frm1.vspdData.Row = Row
            Frm1.vspdData.Col = C_EMP_TYPE_CD
            Frm1.vspdData.value = iDx
    End Select    
     ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)

    Dim iDx
    Dim IntRetCd
    Dim ArgMt
    Dim iCodeArr 
    Dim iNameArr
    Dim IncurRow
    Dim strBas, strDept_nm
    
     ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    With frm1.vspdData
    		Select Case Col
    			Case C_PAY_GRD1
    				ArgMt = ""
    			 	.Col = C_PAY_GRD1
    				ArgMt = Trim(.Text)
    				
    				IF Trim(ArgMt) = "" Then
						    .Col =  C_PAY_GRD1_NM
						    .Value = ""
						    .Action = 0
                    Else
    					Call  CommonQueryRs(" MINOR_NM"," B_MINOR ","MAJOR_CD = " & FilterVar("H0001", "''", "S") & " AND MINOR_CD =  " & FilterVar(ArgMt , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    					
    					IF Len(Trim(lgF0)) = 0 Then
    						Call  DisplayMsgBox("970000","X","��ȣ�ڵ�","X")
    						'.Value = ""
						    .Col =  C_PAY_GRD1_NM
						    .Value = ""
						    .Action = 0
					    Else
    						.Col =  C_PAY_GRD1_NM
    						.Value = Trim(Replace(lgF0,Chr(11),""))  				
					    End IF
				    End If
    		End Select 
    End With

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
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
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

Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub

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
			<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ȣ��������е��</font></td>
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
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
		         	<TD WIDTH=100% HEIGHT=* valign=top>
		                <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100%> 
					                <script language =javascript src='./js/h1024ma1_vaSpread1_vspdData.js'></script>
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

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
