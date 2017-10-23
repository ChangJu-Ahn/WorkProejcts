<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim BaseDate
Dim StartDate
Dim EndDate
iDBSYSDate = "<%=GetSvrDate%>"
BaseDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)


'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID  = "ZC014MB1_ko441.asp"			 '☆: 비지니스 로직 ASP명 

'==========  1.2.1 Global 상수 선언  ======
'⊙: Grid Columns

Dim OldOrgChangeDesc

Dim C_USR_ID
Dim C_USR_NM
Dim C_BA
Dim C_PL
Dim C_SG
Dim C_SO
Dim C_PG
Dim C_PO
Dim C_NO

'----------------  공통 Global 변수값 정의  -------------------------------------------------------------- 
Dim IsOpenPop


Sub InitVariables()  

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode    
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgPageNo  = 0

End Sub

'============================================================================================================
Sub initSpreadPosVariables()


 C_USR_ID =	1
 C_USR_NM =	2
 C_BA =		3
 C_PL =		4
 C_SG =		5
 C_SO =		6
 C_PG =		7
 C_PO =		8
 C_NO =		9
	 
End Sub

'============================================================================================================
Sub SetDefaultVal()

End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Dim IntRetCD1
	
    on error resume next
    
  
    
End Sub    


'============================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'============================================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()		

    
	    With frm1.vspdData
    	  
            ggoSpread.Source = frm1.vspdData

            ggoSpread.Spreadinit "V20031129",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols   = C_NO + 1                                                      ' ☜:☜: Add 1 to Maxcols
	                                               ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0
            ggoSpread.ClearSpreadData
             Call GetSpreadColumnPos("A")
            
            
            ggoSpread.SSSetEdit		C_USR_ID,		 "사용자",		 18, , , 20
            ggoSpread.SSSetEdit		C_USR_NM,		 "사용자명",		 18, , , 20
            ggoSpread.SSSetCheck 	C_BA,			 "BIZ AREA",		 10, 2, "", True   
            ggoSpread.SSSetCheck 	C_PL,			 "PLANT",		 10, 2, "", True   
            ggoSpread.SSSetCheck 	C_SG,			 "SALES GROUP",		 10, 2, "", True   
            ggoSpread.SSSetCheck 	C_SO,			 "SALES ORG.",		 10, 2, "", True    
            ggoSpread.SSSetCheck 	C_PG,			 "PURCHASE GROUP",		 15, 2, "", True   
            ggoSpread.SSSetCheck 	C_PO,			 "PURCHASE ORG.",		 15, 2, "", True    
            ggoSpread.SSSetEdit		C_NO,			 "사용자",		 18, , , 20
		      	
	        .ReDraw = True
          End With    
        ' ggoSpread.SSSetSplit2(1)
         Call SetSpreadLock("A")
 
   
             
End Sub

'============================================================================================================
Sub SetSpreadLock(ByVal QryFg)     
   With frm1  
            .vspdData.ReDraw = False

    		ggoSpread.SpreadLock C_USR_ID, -1,  C_USR_NM
	         
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
           Frm1.vspdData2.Col = iDx
           Frm1.vspdData2.Row = iRow
           If Frm1.vspdData2.ColHidden <> True And Frm1.vspdData2.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData2.Col = iDx
              Frm1.vspdData2.Row = iRow
              Frm1.vspdData2.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub



'============================================================================================================
Sub Form_Load()
  
    Err.Clear                                                                       '☜: Clear err status
    Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet()                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    'Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, Parent.gDateFormat, 2)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    
    'Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call SetToolbar("1110000000001111")												'⊙: Set ToolBar    								
    
    Call InitComboBox
    Call CookiePage (0)                                                             '☜: Check Cookie
	
End Sub

'============================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		 C_USR_ID =	iCurColumnPos(1)
		 C_USR_NM =	iCurColumnPos(2)
		 C_BA =		iCurColumnPos(3)
		 C_PL =		iCurColumnPos(4)
		 C_SG =		iCurColumnPos(5)
		 C_SO =		iCurColumnPos(6)
		 C_PG =		iCurColumnPos(7)
		 C_PO =		iCurColumnPos(8)
		 C_NO =		iCurColumnPos(9)
		
	  			
    End Select
End Sub

'============================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub



'+++++******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'------------------------------------------  txtFr_Dt_KeyDown ------------------------------------------
'	Name : txtFr_Dt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub


'=======================================================================================================
'   Event Name : txtFr_Dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtDt.Focus
	End If
End Sub



Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Dim iRow
    dIM  IntRetCD

    Dim PrntTransNo
    Dim PrntTransSeq
    Dim i



    
	If lgIntFlgMode = parent.OPMD_CMODE Then
	
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0001111111")
	End If
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

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
	
	With frm1
	     
	    .vspdData.focus
	    
	     Set gActiveElement = document.activeElement 
    
		ggoSpread.Source = .vspdData
    
		.vspdData.ReDraw = False
		iRow =  .vspdData.ActiveRow
		  
		 
		
		     
		     
		    '  For i = 1 to .vspdData.maxCols 
			'	.vspdData.Col = i  
			'	.vspdData.BackColor = RGB(176,234,244) '  14540253
			'	.vspdData.ForeColor = vbBlue
			 'Next 
	     
		 
		  .vspdData.Row = iRow
		   
        
         If .vspdData.Text = "" Then
           Exit Sub 	   	    
         End If         
          .vspdData.ReDraw =True 
	End With 
		     
	
End Sub

'============================================================================================================


'============================================================================================================
Function FncQuery()
	Dim IntRetCD 

    FncQuery = False          '⊙: Processing is NG
    Err.Clear                 '☜: Protect system from crashing
   
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream() 
    Call DisableToolBar(parent.TBC_QUERY)

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery
       
    FncQuery = True
    
End Function

'============================================================================================================
Function FncNew() 

    Dim IntRetCD 

    FncNew = False

    On Error Resume Next
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call SetDefaultVal

    Call InitVariables

	Call SetToolbar("1110110100101111")

    FncNew = True
    
End Function

'============================================================================================================
Function FncDelete() 
	On Error Resume Next
End Function

'============================================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim iRow,ChkChange,ChkChange2
    Dim PrnChangeDesc
   
    FncSave = False
    Err.Clear
    On Error Resume Next
    
    '-----------------------
    'Precheck area
    '-----------------------


    If Not chkField(Document, "1") Then               '⊙: Check required field(Single area)
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then               '⊙: Check required field(Single area)
       Exit Function
    End If

    '-----------------------
    'Update OrgChangeDesc
    '-----------------------         
    
     
 
     
    '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave	
	FncSave = True 
	
End Function

'============================================================================================================
Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData2.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData2

	With frm1.vspdData2
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.Source = frm1.vspdData2	
			ggoSpread.CopyRow
			SetSpreadColor1 .ActiveRow, .ActiveRow    
			
			.ReDraw = true
		End If
    End with

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'============================================================================================================
Function FncCancel() 
	If frm1.vspdData3.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.EditUndo
End Function

'============================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow , x
    Dim StrPhyNo,StrSeqNo
    
    Dim PrntType
    Dim PrntLvlApplyDt
    Dim PrntChangeDesc
    

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

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
		
	 x = .vspdData.Maxrows
	If x < 1 Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtPlantCd.focus
        Exit Function
	End If

	.vspdData3.focus
	 ggoSpread.Source = .vspdData3
				
	.vspdData3.ReDraw = False
     ggoSpread.InsertRow ,imRow      
     SetSpreadColor1 .vspdData3.ActiveRow, .vspdData3.ActiveRow + imRow - 1
    .vspdData3.ReDraw = True
             
    For iRow = .vspdData3.ActiveRow to .vspdData3.ActiveRow + imRow - 1
        '.vspdData3.Row = iRow    
        '.vspdData3.Col = C_SEQ_NO
        '.vspdData3.text = iRow

	Next
  
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    lgBlnFlgChgValue = True     
End Function


Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
	  
	if frm1.vspdData2.maxrows < 1 then exit function
	 
	With frm1.vspdData3 
		If .MaxRows = 0 Then
			Exit Function
		End If

		.focus
		ggoSpread.Source = frm1.vspdData3

		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
	End With
End Function


Function FncPrint()
    parent.FncPrint()
End Function


Function FncPrev() 
    On Error Resume Next
End Function


Function FncNext() 
    On Error Resume Next
End Function


Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function


Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function


Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub


Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub


Sub MakeKeyStream()   
    
	lgKeyStream  = Trim(frm1.txtUsrId1.value) & Parent.gColSep

End Sub        


Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData2	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'============================================================================================================
Function DbQuery() 
	
	Dim strVal
	
    DbQuery = False
    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1) = False then
       Exit Function 
   	End if
	Call MakeKeyStream()
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="       & lgStrPrevKey                 '☜: Next key tag
    End With
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
	   'MSGBOX strVal  	
    DbQuery = True

End Function



'============================================================================================================
Function DbQueryOk()

Dim i

   lgIntFlgMode = Parent.OPMD_UMODE

      
  	ggoSpread.Source = frm1.vspdData
    
		frm1.vspdData.ReDraw = False
		     
		     
		    For i = 1 to frm1.vspdData.maxRows 
				frm1.vspdData.Col = C_NO
				frm1.vspdData.Row = i
				
				IF frm1.vspdData.value = "" THEN
					frm1.vspdData.Col = 0
					frm1.vspdData.Row = i
					frm1.vspdData.text = ggoSpread.InsertFlag
				End If
			 
		    Next 
		frm1.vspdData.ReDraw = TRUE	

    Call ggoOper.LockField(Document, "Q")    	
	Call SetToolbar("1110100000011111")										<%'버튼 툴바 제어 %>
		
    Set gActiveElement = document.ActiveElement
    
    lgBlnFlgChgValue = False        
    
End Function

Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx,strText

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    'Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

    Select Case Col
        
    End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub 

'============================================================================================================
Function DbSave() 

	Dim lRow
	Dim lGrpCnt
	Dim lGrpCnt2
	Dim strVal,strVal2, strDel,strDel2
	Dim iColSep, iRowSep
	
	Err.Clear

    DbSave = False
    
    On Error Resume Next

	Call LayerShowHide(1)


    '-----------------------
    'Data manipulate area
    '-----------------------
	iColSep = parent.gColSep : iRowSep = parent.gRowSep 
		
	With frm1
	
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		lGrpCnt2 = 1
		
		strVal = ""
		strDel = ""
		strDel2 = ""
		strVal2 = ""
		'-----------------------
		'Data manipulate area
		'-----------------------

		For lRow = 1 To .vspdData.MaxRows         
            
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag										'☜: 신규 
					strVal = strVal & "C" & iColSep	& lRow & iColSep					'☜: C=Create
			    Case ggoSpread.UpdateFlag										'☜: 수정 
					strVal = strVal & "U" & iColSep	& lRow & iColSep					'☜: U=Update
			End Select

		    Select Case .vspdData.Text


				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag					'☜: 수정, 신규
				    .vspdData.Col = C_USR_ID
				    strVal = strVal & Trim(.vspdData.Text) & iColSep
				    If CInt(Trim(GetSpreadText(.vspdData, C_BA, lRow, "X", "X"))) = 1 Then
				    	strVal = strVal & "Y"                     & iColSep
				   ElseIf CInt(Trim(GetSpreadText(.vspdData, C_BA, lRow, "X", "X"))) = 0 Then 
				    	strVal = strVal & "N"                     & iColSep
				   End If
				    If CInt(Trim(GetSpreadText(.vspdData, C_PL, lRow, "X", "X"))) = 1 Then
				    	strVal = strVal & "Y"                     & iColSep
				   ElseIf CInt(Trim(GetSpreadText(.vspdData, C_PL, lRow, "X", "X"))) = 0 Then 
				    	strVal = strVal & "N"                     & iColSep
				   End If
				    If CInt(Trim(GetSpreadText(.vspdData, C_SG, lRow, "X", "X"))) = 1 Then
				    	strVal = strVal & "Y"                     & iColSep
				   ElseIf CInt(Trim(GetSpreadText(.vspdData, C_SG, lRow, "X", "X"))) = 0 Then 
				    	strVal = strVal & "N"                     & iColSep
				   End If
				    If CInt(Trim(GetSpreadText(.vspdData, C_SO, lRow, "X", "X"))) = 1 Then
				    	strVal = strVal & "Y"                     & iColSep
				   ElseIf CInt(Trim(GetSpreadText(.vspdData, C_SO, lRow, "X", "X"))) = 0 Then 
				    	strVal = strVal & "N"                     & iColSep
				   End If
				    If CInt(Trim(GetSpreadText(.vspdData, C_PG, lRow, "X", "X"))) = 1 Then
				    	strVal = strVal & "Y"                     & iColSep
				   ElseIf CInt(Trim(GetSpreadText(.vspdData, C_PG, lRow, "X", "X"))) = 0 Then 
				    	strVal = strVal & "N"                     & iColSep
				   End If
				    If CInt(Trim(GetSpreadText(.vspdData, C_PO, lRow, "X", "X"))) = 1 Then
				    	strVal = strVal & "Y"                     & iRowSep
				   ElseIf CInt(Trim(GetSpreadText(.vspdData, C_PO, lRow, "X", "X"))) = 0 Then 
				    	strVal = strVal & "N"                     & iRowSep
				   End If

				   
				   lGrpCnt = lGrpCnt + 1 
				    
				Case ggoSpread.DeleteFlag										'☜: 삭제 

					strDel = strDel & "D" & iColSep						      '☜: D=Delete
			  				 .vspdData.Col = C_USR_ID
				    strDel = strDel & Trim(.vspdData.Text) & iRowSep
  
				    lGrpCnt = lGrpCnt + 1     
	    
		    End Select
		Next
						
		.txtMaxRows.value     = lGrpCnt-1
		.txtSpread.value = strDel & strVal			
		
				
		
		'MSGBOX strDel2 & strVal2
     	Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'☜: 비지니스 ASP 를 가동 

	End With
	
	DbSave = True 
	
End Function

'============================================================================================================
Function DbSaveOk()                                                         <%' 저장 성공후 실행 로직 %>
    Call InitVariables  
    With Frm1 
       ggoSpread.Source= .vspdData
       ggoSpread.ClearSpreadData
    End With
    
    lgBlnFlgChgValue = False
    	    	  
	Call Dbquery
	'Call Dbquery2
End Function



'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


Function OpenUsrId(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자정보 팝업"                                     ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                                          ' TABLE 명칭 
    arrParam(2) = strCode                                                   ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = ""                                                        ' Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                                                  ' Field명(0)
    arrField(1) = "Usr_nm"                                                  ' Field명(1)
    
    arrHeader(0) = "사용자"                                                ' Header명(0)
    arrHeader(1) = "사용자명"                                           ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUsrId(arrRet, iWhere)        'return value setting
    End If    
	frm1.txtUsrId1.focus
	Set gActiveElement = document.activeElement

End Function

'=========================================================================================================
'    Name : SetUsrId()
'    Description : User Master Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetUsrId(Byval arrRet, Byval iWhere)
    With frm1
        If iWhere = 0 Then
            .txtUsrId1.value = arrRet(0)
            .txtUsrNm1.value = arrRet(1)
        End If
    End With
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD  HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100%  CELLSPACING=0>
								<TR>
                                					<TD CLASS="TD5" NOWRAP>사 용 자</TD>
                                					<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtUsrId1" SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="사용자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsrId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId frm1.txtUsrId1.value,0">&nbsp;<INPUT TYPE=TEXT ID="txtUsrNm1" NAME="txtUsrNm1" size=30 tag="14"></TD>
                                					<TD CLASS="TDT"></TD>
                                					<TD CLASS="TD6"></TD>                                                                    
                            					</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD  HEIGHT=2 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE  CLASS="BasicTB" CELLSPACING=0>
							<TR>
								<TD HEIGHT="100%">
								<script language =javascript src='./js/zc014ma1_KO441_vspdData_vspdData.js'></script>
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
				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread2 tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows2" tag="24">
<INPUT TYPE=hidden NAME="hPHY_INV_NO" tag="24">
<INPUT TYPE=hidden NAME="hSeqNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

