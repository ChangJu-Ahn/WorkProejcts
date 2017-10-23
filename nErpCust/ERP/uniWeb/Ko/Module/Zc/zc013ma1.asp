<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID		    : ZC013MA1
'*  4. Program Name         : 화면조회정보등록 
'*  5. Program Desc         : 화면조회정보등록 
'*  6. Component List       : 
'*  7. ModIfied date(First) : 2005/03/03
'*  8. ModIfied date(Last)  : 
'*  9. ModIfier (First)     : SHIN HYUN HO
'* 10. ModIfier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit	

<% 
	Dim StrCo
	StrCo = Request.Cookies("unierp")("gLang")
 %>
 
 Dim StrLang,StrCnt,StrChk
	StrChk = 0
    StrLang = "<%=Strco%>" 


Const BIZ_PGM_ID	= "ZC013MB1.asp"			'☆: List & Manage SCM Orders



Dim	C_KEY_FIELD
Dim	C_KEY_OBJECT_NM
Dim	C_KEY_OBJECT_VALUE
Dim	C_PGM_ID

Dim	C_MOVE_ITEM_CD
Dim	C_MOVE_ITEM_POP
Dim	C_MOVE_ITEM_NM
Dim	C_MOVE_ITEM_QRY
Dim	C_MOVE_ITEM_QRY2
Dim	C_PGM_ID2
Dim	C_MOVE_ITEM_CD2


Dim lgStrPrevKey1
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2
Dim strInsertRow
Dim Strq,StrSqry

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData)
		C_KEY_FIELD			= 1
		C_KEY_OBJECT_NM		= 2
		C_KEY_OBJECT_VALUE	= 3
		C_PGM_ID			= 4
		
	End If	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2)
		C_MOVE_ITEM_CD		= 1
		C_MOVE_ITEM_POP		= 2
		C_MOVE_ITEM_NM		= 3
		C_MOVE_ITEM_QRY		= 4
		C_MOVE_ITEM_QRY2	= 5
		C_PGM_ID2			= 6
		C_MOVE_ITEM_CD2		= 7
				
	End If
End Sub

'========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================
Dim  IsOpenPop

'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0
	
    lgSortKey1 = 1
	lgSortKey2 = 1
    lgPageNo = 0
    strInsertRow = 1
End Sub

'========================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1
			ggoSpread.Source = .vspdData
			ggoSpread.Spreadinit "V20030905", , Parent.gAllowDragDropSpread
			.vspdData.ReDraw = False
	
			.vspdData.MaxCols = C_PGM_ID + 1
			.vspdData.MaxRows = 0
			
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit		C_KEY_FIELD,				"KEY_FIELD", 15,,,15,2
			ggoSpread.SSSetEdit		C_KEY_OBJECT_NM,			"KEY_OBJECT_NM", 20
			ggoSpread.SSSetEdit		C_KEY_OBJECT_VALUE,			"화면표시명칭(ex:frm1.textbox.value)", 50
			ggoSpread.SSSetEdit		C_PGM_ID,					"", 20
			
			Call ggoSpread.SSSetColHidden( C_PGM_ID, C_PGM_ID , True)
			Call ggoSpread.SSSetColHidden( .vspdData.MaxCols, .vspdData.MaxCols , True)
			
			.vspdData.ReDraw = True
   
			ggoSpread.SSSetSplit2(3)
    
			Call SetSpreadLock("A")
			
			.vspdData.ReDraw = true    
    
		End With
	
    End If

    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20040805", , Parent.gAllowDragDropSpread
			.vspdData2.ReDraw = False
	
			.vspdData2.MaxCols = C_MOVE_ITEM_CD2 + 1
			.vspdData2.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			ggoSpread.SSSetEdit		C_MOVE_ITEM_CD,			"이동선택항목", 18,,,20,2
			ggoSpread.SSSetButton	C_MOVE_ITEM_POP
			ggoSpread.SSSetEdit		C_MOVE_ITEM_NM,			"이동선택항목명", 20,,,25,2
			ggoSpread.SSSetEdit		C_MOVE_ITEM_QRY,		"Key Value 추출후 쿼리문", 50
			ggoSpread.SSSetEdit		C_MOVE_ITEM_QRY2,		"", 50
			ggoSpread.SSSetEdit		C_PGM_ID2,				"", 20
			ggoSpread.SSSetEdit		C_MOVE_ITEM_CD2,		"", 20
			
			Call ggoSpread.SSSetColHidden( C_MOVE_ITEM_QRY2, C_MOVE_ITEM_QRY2, True)
			Call ggoSpread.SSSetColHidden( C_PGM_ID2, C_PGM_ID2, True)
			Call ggoSpread.SSSetColHidden( C_MOVE_ITEM_CD2, C_MOVE_ITEM_CD2, True)
			Call ggoSpread.SSSetColHidden( .vspdData2.MaxCols, .vspdData2.MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.vspdData2.ReDraw = true    
    
		End With
    End If

End Sub

'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		With frm1
			ggoSpread.Source = .vspdData
	
			.vspdData.ReDraw = False
   			ggoSpread.SpreadLock	 C_KEY_FIELD, -1,C_KEY_FIELD
   			ggoSpread.SSSetRequired	 C_KEY_OBJECT_NM, -1
   			ggoSpread.SSSetRequired	 C_KEY_OBJECT_value, -1
   		   	.vspdData.ReDraw = True
	
		End With
	End If
		
	If pvSpdNo = "B" Then 
		'--------------------------------
		'Grid 2
		'--------------------------------
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SSSetRequired	 C_MOVE_ITEM_CD, -1
		ggoSpread.SpreadLock	 C_MOVE_ITEM_NM, -1, C_MOVE_ITEM_QRY
	End If	

End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
   
		With frm1.vspdData 
    
			.Redraw = False

			ggoSpread.Source = frm1.vspdData
			ggoSpread.SSSetRequired C_KEY_FIELD,			pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_KEY_OBJECT_NM,		pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_KEY_OBJECT_VALUE,		pvStartRow, pvEndRow
    
			.Col = 1
			.Row = .ActiveRow
			.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
			.EditMode = True
    
			.Redraw = True
    
		End With
 End Sub 
 Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)  
		With frm1.vspdData2 
    
			.Redraw = False

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SSSetRequired C_MOVE_ITEM_CD,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_MOVE_ITEM_NM,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_MOVE_ITEM_QRY,		pvStartRow, pvEndRow
    
			.Col = 1
			.Row = .ActiveRow
			.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
			.EditMode = True
    
			.Redraw = True
    
		End With
  
End Sub

'========================================================================================
Sub InitComboBox()
   
	
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos

 	Select Case Ucase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_KEY_FIELD				= iCurColumnPos(1)
			C_KEY_OBJECT_NM			= iCurColumnPos(2)
			C_KEY_OBJECT_VALUE		= iCurColumnPos(3)
			C_PGM_ID				= iCurColumnPos(4)
						
		Case "B"
			
			ggoSpread.Source = frm1.vspdData2
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_MOVE_ITEM_CD			= iCurColumnPos(1)
			C_MOVE_ITEM_POP			= iCurColumnPos(2)
			C_MOVE_ITEM_NM			= iCurColumnPos(3)
			C_MOVE_ITEM_QRY			= iCurColumnPos(4)
			C_MOVE_ITEM_QRY2		= iCurColumnPos(5)
			C_PGM_ID2				= iCurColumnPos(6)
			C_MOVE_ITEM_CD2			= iCurColumnPos(7)
						
 	End Select
 
End Sub

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopup(Byval iWhere)

	Dim arrRet
	Dim arrParam(7), arrField(8), arrHeader(8)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	   
	   Case 0
			
			
	        arrParam(0) = "PGM_ID"
			arrParam(1) = "             (SELECT A.MNU_ID,A.MNU_TYPE,A.MNU_NM,B.UPPER_MNU_ID,(CASE WHEN B.USE_YN = 1 THEN 'TRUE' ELSE 'FALSE' END) AS USE_YN  "
			arrParam(1) = arrParam(1) & "  FROM Z_LANG_CO_MAST_MNU A,Z_CO_MAST_MNU      B "
			arrParam(1) = arrParam(1) & " WHERE A.MNU_ID = B.MNU_ID AND A.MNU_TYPE = 'P'  AND A.LANG_CD = " & FilterVar(UCASE(Trim(StrLang)),"''","S")
			arrParam(1) = arrParam(1) & "   		 AND EXISTS (SELECT *  "
 			arrParam(1) = arrParam(1) & "                              FROM Z_MOVE_PGM_KEY_OBJECT "
 			arrParam(1) = arrParam(1) & "                             WHERE A.MNU_ID = PGM_ID )) A "
			arrParam(2) = Trim(frm1.txtPGMID.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "PGM_ID"

			arrField(0) = "ED08" & Chr(11) & "A.MNU_ID"
			arrField(1) = "ED18" & Chr(11) & "A.MNU_NM"
			arrField(2) = "ED07" & Chr(11) & "A.UPPER_MNU_ID"
			arrField(3) = "ED07" & Chr(11) & "A.USE_YN"
			
			arrHeader(0) = "NEXT_PGM_ID"
			arrHeader(1) = "NEXT_PGM_ID명"
			arrHeader(2) = "상위메뉴"
			arrHeader(3) = "사용여부"

         
       Case 1
			
			
	        arrParam(0) = "PGM_ID"
			arrParam(1) = "             (SELECT A.MNU_ID,A.MNU_TYPE,A.MNU_NM,B.UPPER_MNU_ID,(CASE WHEN B.USE_YN = 1 THEN 'TRUE' ELSE 'FALSE' END) AS USE_YN  "
			arrParam(1) = arrParam(1) & "  FROM Z_LANG_CO_MAST_MNU A,Z_CO_MAST_MNU      B "
			arrParam(1) = arrParam(1) & " WHERE A.MNU_ID = B.MNU_ID  AND A.MNU_TYPE = 'P'  AND A.LANG_CD = " & FilterVar(UCASE(Trim(StrLang)),"''","S")
			arrParam(1) = arrParam(1) & "   		 AND NOT EXISTS (SELECT *  "
 			arrParam(1) = arrParam(1) & "                              FROM Z_MOVE_PGM_KEY_OBJECT "
 			arrParam(1) = arrParam(1) & "                             WHERE A.MNU_ID = PGM_ID )) A "
			arrParam(2) = Trim(frm1.txtPGMID2.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "NEXT_PGM_ID"

			arrField(0) = "ED08" & Chr(11) & "A.MNU_ID"
			arrField(1) = "ED18" & Chr(11) & "A.MNU_NM"
			arrField(2) = "ED07" & Chr(11) & "A.UPPER_MNU_ID"
			arrField(3) = "ED07" & Chr(11) & "A.USE_YN"
			
			arrHeader(0) = "NEXT_PGM_ID"
			arrHeader(1) = "NEXT_PGM_ID명"
			arrHeader(2) = "상위메뉴"
			arrHeader(3) = "사용여부"

        
       Case 2
			
			frm1.vspdData2.Col = C_MOVE_ITEM_CD
			frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
			
	        arrParam(0) = "이동선택항목"
			arrParam(1) = "             (SELECT A.MOVE_ITEM_CD,B.MOVE_ITEM_NM   "
			arrParam(1) = arrParam(1) & "  FROM Z_MOVE_PGM_INF A,Z_MOVE_ITEM B  "
			arrParam(1) = arrParam(1) & " WHERE A.MOVE_ITEM_CD = B.MOVE_ITEM_CD "
			arrParam(1) = arrParam(1) & "   AND NOT EXISTS (SELECT * "
			arrParam(1) = arrParam(1) & "   				  FROM Z_MOVE_PGM_ITEM_QRY "
			arrParam(1) = arrParam(1) & "					 WHERE MOVE_ITEM_CD = A.MOVE_ITEM_CD AND PGM_ID = " & FilterVar(UCASE(Trim(frm1.txtPGMID2.value)),"''","S")  & ")) A "  		 
			arrParam(2) = Trim(frm1.vspdData2.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "이동선택항목"

			arrField(0) = "ED10" & Chr(11) & "A.MOVE_ITEM_CD"
			arrField(1) = "ED30" & Chr(11) & "A.MOVE_ITEM_NM"

			
			
			arrHeader(0) = "이동선택항목"
			arrHeader(1) = "이동선택항목명"

       			
      End Select
	
	  
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SubSetPopup(arrRet,iWhere)
	End If	
	
End Function

'=======================================================================================================
Function GridsetFocus(Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			'.txtMItemCd.focus
		End Select
	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : MItemItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Sub SubSetPopup(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    
		    CASE 0            
				
				.txtPGMID.value = Trim(arrRet(0))
				.txtPGMNM.value = Trim(arrRet(1))
		       		    
		    CASE 1            
				
				.txtPGMID2.value = Trim(arrRet(0))
				.txtPGMNM2.value = Trim(arrRet(1))

            CASE 2
                
                .vspdData2.Col = C_MOVE_ITEM_CD
				.vspdData2.Text = Trim(arrRet(0))
				.vspdData2.Col = C_MOVE_ITEM_NM
				.vspdData2.Text = Trim(arrRet(1))
				ggoSpread.UpdateRow .vspddata.activerow
				.vspdData2.focus
				
        End Select
	End With
End Sub



'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    Call InitSpreadSheet("*")
    Call InitVariables                            '⊙: Initializes local global Variables
    Call SetToolbar("110011110010111")										'⊙: 버튼 툴바 제어 
    frm1.txtPGMID.focus 
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
            
        End Select

'		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")
    End With
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			
			
		End Select
	End With
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			
			
			
		Next
	End With
	
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    strInsertRow = 1
    Call SetPopupMenuItemInf("1101111111")
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
		Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
        Exit Sub
    End If
End Sub

Sub vspdData2_Click(ByVal Col , ByVal Row)

Dim StrSplit,i,S,T
	
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData2
    strInsertRow = 2
    Call SetPopupMenuItemInf("1101111111")
    If frm1.vspdData2.MaxRows <= 0 Then                                                    'If there is no data.
		Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey2 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If
        Exit Sub
    End If
    
    If Row >= 0  Then
		
		frm1.vspddata2.row = frm1.vspddata2.activerow
		frm1.vspddata2.col = 0
		if frm1.vspddata2.text = ggoSpread.InsertFlag or frm1.vspddata2.text = ggoSpread.updateFlag then
			frm1.txtqry.value = ""
			
			Exit Sub
		Else
			frm1.vspddata2.col = frm1.vspdData2.MaxCols
					
			StrSplit =  frm1.vspdData2.value
			S = "frm1.txtqry.value" & " = " & "frm1.hquery" & StrSplit & ".value" 
					
			execute S
			
			frm1.txtqry.value = REPLACE(frm1.txtqry.value, chr(7), chr(13)&chr(10))
			
		End If
	End If
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)


    If Row <= 0 Then
        Exit Sub
    End If

    With frm1
        Select Case Col
            Case C_MOVE_ITEM_POP '테이블ID
                .vspdData2.Col = Col-1
                .vspdData2.Row = Row
                Call OpenPopUp(2)
           
        End Select

'		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")
    End With
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , ShIft , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'================================================================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    if strInsertRow = 1 then
    Call InitSpreadSheet("A")
    else
    Call InitSpreadSheet("B")
    End if
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()

End Sub 


'========================================================================================================= 
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx,strText

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

    Select Case Col
        
    End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub 

Sub vspdData2_Change(ByVal Col , ByVal Row )
    Dim iDx,strText

    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col

    Call CheckMinNumSpread(frm1.vspdData2, Col, Row) 

    Select Case Col
        
    End Select

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row
End Sub 

'========================================================================================================= 
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
			Exit Sub
		End If
    End With
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End If
End Sub

'========================================================================================
Function FncQuery()
	Dim IntRetCD 
	Dim SSCheck1,SSCheck2
    FncQuery = False
    Err.Clear
	
	
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    SSCheck1 =  ggoSpread.SSCheckChange
    
    ggoSpread.Source = frm1.vspdData2
    SSCheck2 =  ggoSpread.SSCheckChange
    
    
    ggoSpread.Source = frm1.vspdData
    If SSCheck1 = True or  SSCheck2 = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    

    '-----------------------
    'Erase contents area
    '-----------------------
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If

    Call DbQuery

    FncQuery = True
End Function


'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    Dim StrChg1,StrChg2
    FncSave = False
    Err.Clear

    '-----------------------
    'Precheck area
    '-----------------------

    ggoSpread.Source = frm1.vspdData

	StrChg1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2

	StrChg2 = ggoSpread.SSCheckChange
    
    If StrChg1 = False and StrChg2 = False Then   
        IntRetCD = DisplayMsgBox("900001","x","x","x")                          'No data changed!!
        Exit Function
    End If
    
   

    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData

    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area

		Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2

    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area

		Exit Function
    End If
	
    
    If Not chkField(Document, "2") Then 
       		Exit Function
    End If
	
	
    Call DbSave

    FncSave = True
End Function

'========================================================================================
Function FncCopy()
	Dim IntRetCD

	If strInsertRow = 1 Then
		frm1.vspdData.ReDraw = False

		If frm1.vspdData.MaxRows < 1 Then Exit Function

		ggoSpread.Source = frm1.vspdData
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

		'frm1.vspdData.Col = C_MItemCd
		'frm1.vspdData.Text = ""

		frm1.vspdData.ReDraw = True
	End If

	
	If strInsertRow = 2 Then
		frm1.vspdData2.ReDraw = False

		If frm1.vspdData2.MaxRows < 1 Then Exit Function

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.CopyRow
		SetSpreadColor2 frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow

		'frm1.vspdData.Col = C_MItemCd
		'frm1.vspdData.Text = ""

		frm1.vspdData2.ReDraw = True
	End If

End Function

'========================================================================================
Function FncCancel() 
    
    IF strInsertRow = 1 Then
    If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    End If
    
    IF strInsertRow = 2 Then
    If frm1.vspdData2.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.EditUndo
	End If
	
	Call InitData()
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
    imRow = AskSpdSheetAddRowCount()

    If imRow = "" Then
        Exit Function
		End If
    End If
	
    If Not chkField(Document, "2") Then 
       		Exit Function
    End If
	
	If strInsertRow = 1 Then

	With frm1
	        .vspdData.ReDraw = False
	        .vspdData.focus
            ggoSpread.Source = .vspdData
            ggoSpread.InsertRow ,imRow
            SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
			For iDx = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
			
					.vspdData.COL   = C_PGM_ID
					.vspdData.value = .txtPGMID2.value 
					
		    .vspdData.ReDraw = True
		    
		Next
       .vspdData.ReDraw = True
    End With
    End If
    
    If strInsertRow = 2 Then
    
    
   With frm1

	        .vspdData2.ReDraw = False
	        .vspdData2.focus
            ggoSpread.Source = .vspdData2

			if StrChk < 1 then
					ggoSpread.InsertRow ,imRow
					SetSpreadColor2 .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
					StrCnt = .vspdData2.maxrows
			
					For iDx = .vspdData2.ActiveRow To .vspdData2.ActiveRow + imRow - 1
			
							.vspdData2.COL   = C_PGM_ID2
							.vspdData2.value = .txtPGMID2.value 
			
							StrCnt = StrCnt + 1
							StrChk = StrChk + 1	
							.vspdData2.ReDraw = True
		    
					Next
			Else
			'msgbox "입력은 하나씩만 가능합니다"
			End if			
       .vspdData2.ReDraw = True
    End With
    End If
    
     If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   

	If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
Function FncDeleteRow() 
	Dim lDeIRows
	Dim iDeIRowCnt, i
	Dim IntRetCD 
	
	If strInsertRow = 1 Then
		If frm1.vspdData.MaxRows < 1 Then Exit Function

			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 

				lDeIRows = ggoSpread.DeleteRow
			End With
    End If
    If strInsertRow = 2 Then
		If frm1.vspdData2.MaxRows < 1 Then Exit Function

			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2

				lDeIRows = ggoSpread.DeleteRow
			End With
    End If
End Function

'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================


Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


'========================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		 'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		 If IntRetCD = vbNo Then
		     Exit Function
		End If
	End If

    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    Err.Clear

	Call LayerShowHide(1)


    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPGMID=" & Trim(.txtPGMID.value)	'조회 조건 데이타 
			strVal = strVal & "&txtLANG=" & StrLang	'국가 코드 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtMaxRows2=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPGMID=" & Trim(.txtPGMID.value)	'조회 조건 데이타 
			strVal = strVal & "&txtLANG=" & StrLang	'국가 코드 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtMaxRows2=" & .vspdData2.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
    End With

    DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQuery_Seven
' Function Desc : This function is data query and display
'=======================================================================================================

'========================================================================================
Function DbQueryOk()
	
	Dim StrSplit,i,S
	
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
    Call InitData
	Call SetToolbar("110011110011111")
	
	StrCnt = 0
	StrChk = 0
	
	if frm1.hquery.value <> "" then
	
	StrSplit = Split(frm1.hquery.value,chr(11))
	
	For i = 1 to ubound(StrSplit)
	S = "frm1.hquery" & i & ".value = " & chr(34) & StrSplit(i - 1) & chr(34)
	
	execute S 
	Next
	End iF
	
	frm1.vspdData.focus
	
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function DbSave() 
	Dim IRow
	Dim lGrpCnt,lGrpCnt2
	Dim strVal, strDel,strVal2, strDel2,S,T

    DbSave = False
    On Error Resume Next
    Err.Clear 

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For IRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = IRow
		    .vspdData.Col = 0


		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag															'☜: 신규 
					strVal = strVal & "C"  & Parent.gColSep & IRow & Parent.gColSep					'☜: C=Create
		            .vspdData.Col = C_PGM_ID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_FIELD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_OBJECT_NM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_OBJECT_VALUE
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag															'☜: 수정 
					strVal = strVal & "U"  & Parent.gColSep & IRow & Parent.gColSep					'☜: U=Update
		            .vspdData.Col = C_PGM_ID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_FIELD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_OBJECT_NM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_OBJECT_VALUE
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag																'☜: 삭제 
					strDel = strDel & "D" & Parent.gColSep & IRow & Parent.gColSep
		            .vspdData.Col = C_PGM_ID
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KEY_FIELD
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		
		lGrpCnt2 = 1
		strVal2 = ""
		strDel2 = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For IRow = 1 To .vspdData2.MaxRows
		    .vspdData2.Row = IRow
		    .vspdData2.Col = 0


		    Select Case .vspdData2.Text
		        Case ggoSpread.InsertFlag															'☜: 신규 
					strVal2 = strVal2 & "C"  & Parent.gColSep & IRow & Parent.gColSep					'☜: C=Create
		            .vspdData2.Col = C_PGM_ID2
		            strVal2 = strVal2 & Trim(.vspdData2.Text) & Parent.gColSep
		            .vspdData2.Col = C_MOVE_ITEM_CD
		            strVal2 = strVal2 & Trim(.vspdData2.Text) & Parent.gColSep
		            strVal2 = strVal2 & Trim(.hQuery.value)   & Parent.gRowSep
		            
		            
		            lGrpCnt2 = lGrpCnt2 + 1
				Case ggoSpread.UpdateFlag															'☜: 수정 
					strVal2 = strVal2 & "U"  & Parent.gColSep & IRow & Parent.gColSep
					 .vspdData2.Col = C_PGM_ID2
		            strVal2 = strVal2 & Trim(.vspdData2.Text) & Parent.gColSep
		            .vspdData2.Col = C_MOVE_ITEM_CD
		            strVal2 = strVal2 & Trim(.vspdData2.Text) & Parent.gColSep
		            
		            Dim s1
		            s = "s1 = frm1.hquery" & IRow & ".value "
		            
		             execute S
		            
		            strVal2 = strVal2 & s1 & Parent.gColSep
		            .vspdData2.Col = C_MOVE_ITEM_CD2
		            strVal2 = strVal2 & Trim(.vspdData2.Text) & Parent.gRowSep
		            
		            
		            lGrpCnt2 = lGrpCnt2 + 1
		        Case ggoSpread.DeleteFlag																'☜: 삭제 
					strDel2 = strDel2 & "D" & Parent.gColSep & IRow & Parent.gColSep
		            .vspdData2.Col = C_PGM_ID2
		            strDel2 = strDel2 & Trim(.vspdData2.Text) & Parent.gColSep
		            .vspdData2.Col = C_MOVE_ITEM_CD
		            strDel2 = strDel2 & Trim(.vspdData2.Text) & Parent.gRowSep
		            lGrpCnt2 = lGrpCnt2 + 1
		    End Select
		Next
		.txtMaxRows2.value = lGrpCnt2-1
		.txtSpread2.value = strDel2 & strVal2
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)																'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True																						'⊙: Processing is NG
End Function

'========================================================================================
Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	Call DbQuery
End Function

'========================================================================================
Sub SetGridFocus()	
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub

SUB txtQRY_ONCHANGE()
Dim Strqry,I,S,T

Strqry = REPLACE(frm1.txtqry.value	, chr(13)&chr(10)	, chr(7))

ggoSpread.source = frm1.vspdData2

If frm1.vspddata2.maxrows > 0 Then
	frm1.vspddata2.row = frm1.vspddata2.activerow
	frm1.vspddata2.col = frm1.vspddata2.maxcols
	S = frm1.vspddata2.value
	
	frm1.vspddata2.col = C_MOVE_ITEM_QRY
	frm1.vspddata2.value = Strqry
	frm1.vspddata2.col = C_MOVE_ITEM_QRY2
	frm1.vspddata2.value = Strqry
	
	frm1.vspddata2.col = 0
	if frm1.vspddata2.text = ggoSpread.InsertFlag then 
	frm1.hQuery.value = Strqry
	Else
	T = "frm1.hquery" & S & ".value = " & chr(34) & Strqry & chr(34)
	
	execute T
	End If
End If

CALL vspdData2_Change(frm1.vspddata2.activeCol , frm1.vspddata2.activerow )

End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>화면조회정보등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=500>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>PGM ID</TD>
									<TD CLASS=TD656 COLSPAN = 3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPGMID" SIZE=15 MAXLENGTH=15 tag="12xxxU" ALT="PGM ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopUp(0)">&nbsp;<INPUT TYPE=TEXT NAME="txtPGMNM" SIZE=25 tag="14"></TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>PGM ID</TD>
									<TD CLASS=TD656 COLSPAN = 3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPGMID2" SIZE=15 MAXLENGTH=15 tag="22xxxU" ALT="PGM ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopUp(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtPGMNM2" SIZE=25 tag="24"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR HEIGHT="50%" >
								<TD HEIGHT="50%" WIDTH="100%" >
									<script language =javascript src='./js/zc013ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="20" >
								<TD>
									*선택항목별 쿼리문 등록
								</TD>
							</TR>
							<TR HEIGHT="*" >
							    <TD>
								    <script language =javascript src='./js/zc013ma1_vspdData2_vspdData2.js'></script>
								</TD>
							</TR>
							
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
									<TD CLASS=TD656 COLSPAN = 3>*KEY VALUE</TD>

								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
              							    
				            				<TD CLASS="TD656" ColSpan=4><TEXTAREA  NAME="txtQry" tag="21xxx" rows = 8 cols=150  ALT="QRY"></TEXTAREA>
																		<INPUT TYPE=HIDDEN NAME="hQuery"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery1"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery2"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery3"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery4"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery5"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery6"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery7"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery8"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery9"  SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery10" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery11" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery12" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery13" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery14" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery15" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery16" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery17" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery18" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery19" SIZE= 50 MAXLENGTH=100  TAG="24">
																		<INPUT TYPE=HIDDEN NAME="hQuery20" SIZE= 50 MAXLENGTH=100  TAG="24"></TD>
	        					</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>

	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtPGMID" tag="24"><INPUT TYPE=HIDDEN NAME="htxtPGMID" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

