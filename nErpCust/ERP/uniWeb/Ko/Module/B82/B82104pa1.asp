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

Const BIZ_PGM_ID = "B82104pb1.asp"           <% '☆: 비지니스 로직 ASP명 %>

Const C_SHEETMAXROWS_D = 100				'Sheet Max Rows

Dim C_ReqNo         '의뢰번호 
Dim C_ReqIdNm       '의뢰자 
Dim C_ReqDt         '의뢰일자 
Dim C_Status        '상태 
Dim C_itemKind      '품목구분 
Dim C_ItemKindNm    '품목구분명 
Dim C_ItemCd        '품목코드 
Dim C_ItemNm        '품목명 
Dim C_Spec          '규격 
Dim C_AcctR         '접수검토 
Dim C_AcctT         '기술검토 
Dim C_AcctP         '구매검토 
Dim C_AcctQ         '품질검토 
Dim C_TransDt       '이관일자 
Dim C_Remark        '비고 

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

Dim StartDate, EndDate

StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", PopupParent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
EndDate   = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : InitSpreadPosVariables()     
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
    C_ReqNo      = 1    '의뢰번호 
    C_ReqIdNm    = 2    '의뢰자 
    C_ReqDt      = 3    '의뢰일자 
    C_Status     = 4    '상태 
    C_itemKind   = 5    '품목구분 
    C_itemKindNm = 6    '품목구분명 
    C_ItemCd     = 7    '품목코드 
    C_ItemNm     = 8    '품목명 
    C_Spec       = 9    '규격 
    C_AcctR      = 10   '접수검토 
    C_AcctT      = 11   '기술검토 
    C_AcctP      = 12   '구매검토 
    C_AcctQ      = 13   '품질검토 
    C_TransDt    = 14   '이관일자 
    C_Remark     = 15   '비고 
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

     frm1.txtDtFr.Text  = StartDate
     frm1.txtDtTo.Text  = EndDate
     
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
                 
     frm1.vspdData.MaxCols = C_Remark + 1
     frm1.vspdData.MaxRows = 0
     
     Call GetSpreadColumnPos("A")     
     
     ggoSpread.SSSetEdit  C_ReqNo,	    "의뢰번호",	  15
	 ggoSpread.SSSetEdit  C_ReqIdNm,    "의뢰자",	  10
	 ggoSpread.SSSetDate  C_ReqDt,	    "의뢰일자",   10, 2, PopupParent.gDateFormat  
	 ggoSpread.SSSetEdit  C_Status,  	"상태",       10
	 ggoSpread.SSSetEdit  C_ItemKind,  	"품목구분",   10
	 ggoSpread.SSSetEdit  C_ItemKindNm, "품목구분",   10
	 ggoSpread.SSSetEdit  C_ItemCd,  	"품목코드",   15
	 ggoSpread.SSSetEdit  C_ItemNm,	    "품목명",     20
	 ggoSpread.SSSetEdit  C_Spec,	    "규격",	      20
	 ggoSpread.SSSetEdit  C_AcctR,   	"접수검토",    8
	 ggoSpread.SSSetEdit  C_AcctT,    	"기술검토",    8
	 ggoSpread.SSSetEdit  C_AcctP,	    "구매검토",    8
	 ggoSpread.SSSetEdit  C_AcctQ,		"품질검토",    8
	 ggoSpread.SSSetDate  C_TransDt,    "이관일자",   10, 2, PopupParent.gDateFormat  
 	 ggoSpread.SSSetEdit  C_Remark,	    "비고",	      50
 	
 	 Call ggoSpread.SSSetColHidden(C_ItemKind, C_ItemKind, True)	
 	 Call ggoSpread.SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)

     ggoSpread.SSSetSplit2(1)                                                  'frozen 기능추가 

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
            
            C_ReqNo		= iCurColumnPos(1)
	    	C_ReqIdNm   = iCurColumnPos(2)
		    C_ReqDt		= iCurColumnPos(3)
		    C_Status	= iCurColumnPos(4)
		    C_ItemKind	= iCurColumnPos(5)
		    C_ItemKindNm= iCurColumnPos(6)
		    C_ItemCd	= iCurColumnPos(7)
		    C_ItemNm	= iCurColumnPos(8)
		    C_Spec		= iCurColumnPos(9)
		    C_AcctR		= iCurColumnPos(10)
		    C_AcctT		= iCurColumnPos(11)
		    C_AcctP		= iCurColumnPos(12)
		    C_AcctQ		= iCurColumnPos(13)
		    C_TransDt	= iCurColumnPos(14)
		    C_Remark	= iCurColumnPos(15)
    End Select    
End Sub


'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = frm1.vspdData
   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
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

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row <= 0 Then Exit Sub
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : vspdData_KeyPress
'========================================================================================
Function vspdData_KeyPress(KeyAscii)
    On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================================================================
' Function Name : vspdData_TopLeftChange
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
 	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) And lgStrPrevKey <> "" Then
		DbQuery
	End if
End Sub
 

'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)

        Dim arrRet
        Dim arrParam(7), arrField(8), arrHeader(8)
        
        If IsOpenPop = True  Then  
           Exit Function
        End If   

        IsOpenPop = True
        Select Case arPopUp
               Case 1 '품목구분 
                                   
                    arrParam(0) = frm1.txtItemKind.Alt
                    arrParam(1) = "B_MINOR"
                    arrParam(2) = Trim(frm1.txtItemKind.value)
                    arrParam(4) = "MAJOR_CD = 'Y1001' "
                    arrParam(5) = frm1.txtItemKind.Alt

                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"
    
                    arrHeader(0) = frm1.txtItemKind.Alt
                    arrHeader(1) = frm1.txtItemKind_Nm.Alt
                    frm1.txtItemKind.focus()
                    
               Case 2 '의뢰자 
                                   
                    arrParam(0) = frm1.txtreq_user.Alt
                    arrParam(1) = "B_MINOR"
                    arrParam(2) = Trim(frm1.txtreq_user.value)
                    arrParam(4) = "MAJOR_CD = 'Y1006' "
                    arrParam(5) = frm1.txtreq_user.Alt

                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"
    
                    arrHeader(0) = frm1.txtreq_user.Alt
                    arrHeader(1) = frm1.txtreq_user_Nm.Alt
					frm1.txtreq_user.focus()
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
            Case 1 '품목구분 
                 frm1.txtItemKind.value   = arrRet(0) 
                 frm1.txtItemKind_Nm.value = arrRet(1)   
            Case 2 '의뢰자 
                 frm1.txtreq_user.value      = arrRet(0) 
                 frm1.txtreq_user_Nm.value    = arrRet(1)          
     End Select

     SetConPopup = True

End Function

'========================================================================================================
'     Name : OKClick()
'     Desc : handle ok icon click event
'========================================================================================================
Function OKClick()
     Dim i, iCurColumnPos
     
     If frm1.vspdData.MaxRows > 0 Then
          
          Redim arrReturn(UBound(arrField))

          ggoSpread.Source = frm1.vspdData
          Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
          frm1.vspdData.Row = frm1.vspdData.ActiveRow 
               
          For i = 0 To UBound(arrField)
               If arrField(i) <> "" Then
                    frm1.vspddata.Col = iCurColumnPos(CInt(arrField(i)))
                    arrReturn(i)      = frm1.vspdData.Text
               End If
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
        frm1.txtDtFr.Action = 7
        Call SetFocusToDocument("P")
        frm1.txtDtFr.Focus
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
        frm1.txtDtTo.Action = 7
        Call SetFocusToDocument("P")
        frm1.txtDtTo.Focus
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
	If ValidDateCheck(frm1.txtDtFr, frm1.txtDtto)	=	False	Then Exit	Function
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
    
    If frm1.rdoStatus1.checked = True Then
       frm1.txtRdoStatus.value = "1" 
    ElseIf frm1.rdoStatus2.checked = True Then
       frm1.txtRdoStatus.value = "2" 
    ElseIf frm1.rdoStatus3.checked = True Then
       frm1.txtRdoStatus.value = "3" 
    End If
    
    lgKeyStream =               Trim(frm1.txtDtFr.Text)         & Chr(11)  
    lgKeyStream = lgKeyStream & Trim(frm1.txtDtTo.Text)         & Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtRdoStatus.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.cboItemAcct.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemKind.value)	& Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtreq_user.value)	    & Chr(11)
    lgKeyStream = lgKeyStream & Trim(frm1.txtItemSpec.value)	& Chr(11)
    
    DbQuery = False                                                                 '⊙: Processing is NG
     
    Call LayerShowHide(1)                                                           '⊙: 작업진행중 표시 
    
    strVal = BIZ_PGM_ID & "?txtMode="        & PopupParent.UID_M0001                '☜: Query
    strVal = strVal     & "&txtKeyStream="   & lgKeyStream                          '☜: Query Key
    strVal = strVal     & "&txtPrevNext="    & ""                                   '☜: Direction
    strVal = strVal     & "&lgStrPrevKeyIndex="   & lgStrPrevKeyIndex               '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="     & Frm1.vspdData.MaxRows                '☜: Max fetched data
    strVal = strVal     & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)               '☜: Max fetched data at a time
    
    Call RunMyBizASP(MyBizASP, strVal)                                              '☜: 비지니스 ASP 를 가동 
     
    DbQuery = True                                                                  '⊙: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
     If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
          Call SetActiveCell(vspdData,1,1,"P","X","X")
          Set gActiveElement = document.activeElement
     End If
    lgIntFlgMode = PopupParent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")                                             '⊙: This function lock the suitable field
End Function
    
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
                          <TD CLASS=TD5 NOWRAP>의뢰일자</TD>
                          <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82104pa1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
                                               <script language =javascript src='./js/b82104pa1_fpDateTime6_txtDtTo.js'></script>
                        </TD>
                           <TD CLASS=TD5 NOWRAP>Status</TD>
                           <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoStatus" ID="rdoStatus1" Value="1" CLASS="RADIO" tag="1X"><LABEL FOR="rdoStatus1">전체</LABEL>
                                                <INPUT TYPE="RADIO" NAME="rdoStatus" ID="rdoStatus2" Value="2" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoStatus2">진행</LABEL>
                                                <INPUT TYPE="RADIO" NAME="rdoStatus" ID="rdoStatus3" Value="3" CLASS="RADIO" tag="1X"><LABEL FOR="rdoStatus3">완료</LABEL></TD>
                        </TR>
                        <TR>
					       <TD CLASS=TD5 NOWRAP>품목계정</TD>
					       <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="11" ALT="품목계정"><OPTION VALUE=""></OPTION></SELECT></TD>
					       <TD CLASS=TD5 NOWRAP>품목구분</TD>
					       <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemKind" ALT="품목구분" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('1')">
					                            <INPUT NAME="txtItemKind_Nm" ALT="품목구분명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
					    </TR>
                        <TR>
                          <TD CLASS=TD5 NOWRAP>의뢰자</TD>
                          <TD CLASS=TD6 NOWRAP><INPUT NAME="txtreq_user" ALT="의뢰자" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('2')">
                                               <INPUT NAME="txtreq_user_Nm" ALT="의뢰자명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                          <TD CLASS=TD5 NOWRAP>규격</TD>
                          <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemSpec" ALT="규격" TYPE="Text" SiZE=40   tag="11XXXU"></TD>
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
                              <script language =javascript src='./js/b82104pa1_vaSpread1_vspdData.js'></script>
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
                                                   <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>                         </TD>
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

</FORM>
</BODY>
</HTML>