<%@ LANGUAGE="VBSCRIPT"%>
<HTML>
<HEAD>

<!-- #Include file="../inc/IncServer.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="Acctctrl_ko441_1_popup.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js">        </SCRIPT>

<Script Language="VBScript">
Option Explicit                            

Const POPUP_TITLE = 0
Const TABLE_NAME  = 1
Const CODE_CON    = 2
Const NAME_CON    = 3
Const WHERE_CON   = 4
Const TEXT_NAME   = 5
Const TEXT_NAME1   = 6                                                           '--- Index of Textbox Name


Dim lgSortKey

Dim lgStrNextCodeKey
Dim lgStrNextNameKey
Dim lgNameDuplication

Dim arrParent
Dim arrParam			 '--- First Parameter Group
Dim arrTblField			 '--- Second Parameter Group(DB Table Field Name)
Dim arrGridHdr			 '--- Third Parameter Group(Column Captions of the SpreadSheet)
Dim arrReturn			 '--- Return Parameter Group
Dim gintDataCnt			 '--- Data Counts to Query
Dim arrFieldType 
Dim rbo_sort
Dim lgStrNextCondition
		
	arrParent   = window.dialogArguments
	arrParam    = arrParent(0)
	arrTblField = arrParent(1)
	arrGridHdr  = arrParent(2)
		
	top.document.title = arrParam(POPUP_TITLE)


'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim i
    Dim iArr
    Dim iLen
    
    ReDim arrFieldType(gintDataCnt-1)
	    
    vspdData.ReDraw = False
		    
    ggoSpread.Source = vspdData
    vspdData.OperationMode = 3
		
    vspdData.MaxCols = gintDataCnt
    vspdData.MaxRows = 0
	    
    ggoSpread.Spreadinit
		
    ggoSpread.SSSetEdit 1, arrGridHdr(0), 18	' 코드 

    For i = 1 To gintDataCnt - 1
        If InStr(1, UCase(arrTblField(i)), "CONVERT") > 0 And InStr(1, UCase(arrTblField(i)), "CHAR") > 0 Then 
           ggoSpread.SSSetEdit i + 1, arrGridHdr(i), 25, 1,,999
        Else
           ggoSpread.SSSetEdit i + 1, arrGridHdr(i), 30,  ,,999
        End If
    Next

    For i = 0 To gintDataCnt - 1
        If InStr(1, UCase(arrTblField(i)), gColSep) > 0 Then
           iArr = Split(UCase(arrTblField(i)),gColSep)
               
           iLen = 0
               
           If Len(Trim(iArr(0))) > 2 Then
                  iLen = Cint(Mid(iArr(0),3))
           End If
               
           arrFieldType(i) = ""
               
           Select Case Mid(iArr(0),1,2)
                    Case "ED"   '일반문자 
                           If iLen > 0 Then
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), iLen,,,999
                           Else    
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), 30  ,,,999
                           End If   
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)                           
                    Case "DD"   '날짜 
                           If iLen > 0 Then
                              ggoSpread.SSSetDate   i + 1,arrGridHdr(i),iLen,2,gDateFormat
                           Else    
                              ggoSpread.SSSetDate   i + 1,arrGridHdr(i),  12,2,gDateFormat
                           End If   
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)                           
                    Case "F2","F3","F4","F5"
                           If iLen > 0 Then
                              ggoSpread.SSSetFloat  i + 1,arrGridHdr(i),iLen,Mid(iArr(0),2,1),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
                           Else
                              ggoSpread.SSSetFloat  i + 1,arrGridHdr(i),  17,Mid(iArr(0),2,1),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
                           End If
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)                           
                    Case "TT"   ' Time
                           If iLen > 0 Then
                              ggoSpread.SSSetTime   i + 1,arrGridHdr(i),iLen,,1,1
                           Else
                              ggoSpread.SSSetTime   i + 1,arrGridHdr(i),  12,,1,1
                           End If
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)                           
                    Case "HH"  
                           vspdData.Col = i + 1
                           vspdData.ColHidden = True
                           If iLen > 0 Then
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), iLen
                           Else    
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), 30
                           End If   
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)

           End Select
        End If    
    Next

    ggoSpread.Source = vspdData                         '2003/06/26 leejinsoo
    ggoSpread.SpreadLockWithOddEvenRowColor()

    vspdData.ReDraw = True
End Sub

'========================================================================================================
Sub Form_Load()

	Dim intLoopCnt

	<!-- #Include file="LoadInfTB19029.asp" -->
	
	<% Call loadInfTB19029A("Q", "*","NOCOOKIE","COMMONPOPUP") %>
		
	lgSortKey        =  1
	lgStrNextCodeKey = ""
	lgStrNextNameKey = ""
	lgNameDuplication = "F"	
		
	gintDataCnt      = 0

    For intLoopCnt = 0 To Ubound(arrTblField)
        If arrTblField(intLoopCnt) <> "" Then
           gintDataCnt = gintDataCnt + 1
        Else
           Exit For
       End If
    Next

	If gintDataCnt < 2 Then 
	    txtNm.classname = UCN_PROTECTED 
	    txtNm.ReadOnly = True
	    txtNm.TabIndex = "-1"
	End If

	lblTitle.innerHTML = arrParam(TEXT_NAME)
	lblTitle1.innerHTML = "조건"
	txtCd.value = arrParam(CODE_CON)
	txtNm.value = arrParam(NAME_CON)
	
	document.all("txtisCondition").value = ""	
	
	Call InitComboBox()
    Call ClickRadio()
    
	Self.Returnvalue = Array("")

	Call InitSpreadSheet()

	Call FncQuery()
	
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    
    Dim iCodeArr
    Dim iNameArr
    'iCodeArr = "0♂1♂"
    'iNameArr = ">=♂Like♂"
	iCodeArr = "1♂0♂"
    iNameArr = "Like♂>=♂"    
    
    Call SetCombo2(rdoBpLabel ,iCodeArr  ,iNameArr  ,"♂")
        
End Sub

'======================================================================================================
Sub SetCombo2(pCombo, ByVal pCodeArr, ByVal pNameArr,pSeperator)

    Dim iDx

    pCodeArr = Split(pCodeArr,pSeperator)
    pNameArr = Split(pNameArr,pSeperator)
    
    For iDx = 0 To UBound(pCodeArr) - 1
        Call SetCombo(pCombo,pCodeArr(iDx), pNameArr(iDx))
    Next

End Sub

'======================================================================================================
Sub SetCombo(pCombo, ByVal strValue, ByVal strText)
	Dim objEl
			
	Set objEl = Document.CreateElement("OPTION")	
	objEl.Text = strText
	objEl.Value = strValue

	pcombo.Add(objEl)
	Set objEl = Nothing

End Sub


'========================================================================================================
Sub OKClick()
	Dim intColCnt

	If vspdData.MaxRows < 1 Then
	   Call CancelClick()
	   Exit Sub		
	End If
			
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
				
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
		
    Call CancelClick()
End Sub

'========================================================================================================
Sub CancelClick()
	Self.Close()
End Sub

'========================================================================================================
Function FncQuery()

    vspdData.MaxRows = 0
	lgStrNextCodeKey = Trim(txtCd.value)
	lgStrNextNameKey = Trim(txtNm.value)
	lgNameDuplication = "F"	
	
	txtHidden.value = 0
	document.all("txtisCondition").value = ""
	If rbo_sort = 0 Then
	  Call DbQuery()
	Else
	  if lgStrNextCodeKey = "" And lgStrNextNameKey = "" Then 
	    Call DbQuery()
	  Else 
	    Call DbQuery01() 
	  End If
	End If
		
End Function

'========================================================================================================
Sub vspdData_Click(Col, Row)
    
    If Row = 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
    End If
    
End Sub

'========================================================================================================
Function vspdData_DblClick( Col,  Row)

	If Row = 0 Then
	   Exit Function
	End If
	
	If vspdData.MaxRows = 0 Then
	   Exit Function
	End If

    If Row > 0 Then
       Call OKClick()
    End If

End Function

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	
	If vspdData.MaxRows = 0 Then
	   Exit Function
	End If
	
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	'LIKE조건인 경우는 DbQuery01 함수를 실행 아닌경우는 DbQuery 함수를 실행 
	
    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop)  Then
       If lgStrNextCodeKey <> "" Or lgStrNextNameKey <> "" Then
		  If rbo_sort = 0 Then
			Call DbQuery
		  Else

			if Trim(txtCd.value) = "" And Trim(txtNm.value) = "" Then 
				txtHidden.value = "1"
			    Call DbQuery()
			Else 
			    Call DbQuery01() 
			End If
	  
          End If
       End If
    End if

    
End Sub

Sub ClickRadio()

    IF rdoBpLabel.value = "" Or rdoBpLabel.value = 0 THEN
        rbo_sort="0"
    ELSEIF rdoBpLabel.value = 1 THEN
        rbo_sort="1"
    END IF
    
End Sub

'hancFunction DbQuery01()
'hanc    Dim strVal
'hanc    Dim arrStrDT
'hanc    Dim iLoop
'hanc
'hanc	Call LayerShowHide(1)
'hanc
'hanc    arrStrDT = ""
'hanc        
'hanc    For iLoop = 0 To gintDataCnt - 1
'hanc        arrStrDT = arrStrDT & Trim(arrFieldType(iLoop)) & gColSep
'hanc    Next
'hanc    
'hanc    strVal = "adoAcctctrl_ko441_1_popupbiz.asp?txtTable="    & Trim(arrParam(TABLE_NAME)) 
'hanc    strVal =     strVal & "&txtNextCode=" & Trim(txtCd.value)
'hanc    strVal =     strVal & "&txtNextName=" & Trim(txtNm.value)
'hanc    strVal =     strVal & "&NameDuplication=" & lgNameDuplication        
'hanc    strVal =     strVal & "&txtWhere="    & Trim(arrParam(WHERE_CON)) 
'hanc    strVal =     strVal & "&arrField1="   & Trim(arrTblField(0)) 
'hanc    strVal =     strVal & "&arrField2="   & Trim(arrTblField(1)) 
'hanc    strVal =     strVal & "&arrField3="   & Trim(arrTblField(2)) 
'hanc    strVal =     strVal & "&arrField4="   & Trim(arrTblField(3)) 
'hanc    strVal =     strVal & "&arrField5="   & Trim(arrTblField(4)) 
'hanc    strVal =     strVal & "&arrField6="   & Trim(arrTblField(5)) 
'hanc    strVal =     strVal & "&arrField7="   & Trim(arrTblField(gintDataCnt - 1))
'hanc    strVal =     strVal & "&arrStrDT="    & Escape(arrStrDT)
'hanc    strVal =     strVal & "&gintDataCnt=" & gintDataCnt
'hanc    strVal =     strVal & "&isWhere=" & rbo_sort
'hanc
'hanc    Call RunMyBizASP(MyBizASP, strVal)                                      '☜: 비지니스 ASP 를 가동 
'hanc    
'hanc		
'hancEnd Function
	
'hanc
Function DbQuery01()
    Dim strVal
    Dim arrStrDT
    Dim iLoop

	Call LayerShowHide(1)

    arrStrDT = ""
        
    For iLoop = 0 To gintDataCnt - 1
        arrStrDT = arrStrDT & Trim(arrFieldType(iLoop)) & gColSep
    Next
    
    strVal = "adoAcctctrl_ko441_1_popupbiz.asp?txtTable="    & Trim(arrParam(TABLE_NAME)) 
    strVal =     strVal & "&txtNextCode=" & Trim(lgStrNextCodeKey)
    strVal =     strVal & "&txtNextName=" & Trim(lgStrNextNameKey)
    strVal =     strVal & "&NameDuplication=" & Trim(lgNameDuplication)        
    strVal =     strVal & "&txtWhere="    & Trim(arrParam(WHERE_CON)) 
    strVal =     strVal & "&arrField1="   & Trim(arrTblField(0)) 
    strVal =     strVal & "&arrField2="   & Trim(arrTblField(1)) 
    strVal =     strVal & "&arrField3="   & Trim(arrTblField(2)) 
    strVal =     strVal & "&arrField4="   & Trim(arrTblField(3)) 
    strVal =     strVal & "&arrField5="   & Trim(arrTblField(4)) 
    strVal =     strVal & "&arrField6="   & Trim(arrTblField(5)) 
    strVal =     strVal & "&arrField7="   & Trim(arrTblField(gintDataCnt - 1))
    strVal =     strVal & "&arrStrDT="    & arrStrDT
    strVal =     strVal & "&gintDataCnt=" & gintDataCnt
    strVal =     strVal & "&isWhere=" & "1" 'hanc   rbo_sort
    strVal =     strVal & "&isFlag=" & txtHidden.value    

    Call RunMyBizASP(MyBizASP, strVal)                                      '☜: 비지니스 ASP 를 가동 
		
End Function


'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim arrStrDT
    Dim iLoop

	Call LayerShowHide(1)

    arrStrDT = ""
        
    For iLoop = 0 To gintDataCnt - 1
        arrStrDT = arrStrDT & Trim(arrFieldType(iLoop)) & gColSep
    Next
    
    strVal = "adoAcctctrl_ko441_1_popupbiz.asp?txtTable="    & Trim(arrParam(TABLE_NAME)) 
    strVal =     strVal & "&txtNextCode=" & Trim(lgStrNextCodeKey)
    strVal =     strVal & "&txtNextName=" & Trim(lgStrNextNameKey)
    strVal =     strVal & "&NameDuplication=" & Trim(lgNameDuplication)        
    strVal =     strVal & "&txtWhere="    & Trim(arrParam(WHERE_CON)) 
    strVal =     strVal & "&arrField1="   & Trim(arrTblField(0)) 
    strVal =     strVal & "&arrField2="   & Trim(arrTblField(1)) 
    strVal =     strVal & "&arrField3="   & Trim(arrTblField(2)) 
    strVal =     strVal & "&arrField4="   & Trim(arrTblField(3)) 
    strVal =     strVal & "&arrField5="   & Trim(arrTblField(4)) 
    strVal =     strVal & "&arrField6="   & Trim(arrTblField(5)) 
    strVal =     strVal & "&arrField7="   & Trim(arrTblField(gintDataCnt - 1))
    strVal =     strVal & "&arrStrDT="    & arrStrDT
    strVal =     strVal & "&gintDataCnt=" & gintDataCnt
    strVal =     strVal & "&isWhere=" & rbo_sort
    strVal =     strVal & "&isFlag=" & txtHidden.value    

    Call RunMyBizASP(MyBizASP, strVal)                                      '☜: 비지니스 ASP 를 가동 
		
End Function

'========================================================================================================
Function DbQueryOk()
   Dim IntRetCD

   If vspdData.MaxRows = 0 Then
      IntRetCD = DisplayMsgBoxA("900014") 
      If Trim(txtCd.value) > "" Then
         txtCd.Select 
         txtCd.Focus
      Else   
         txtNm.Select 
         txtNm.Focus
     End If
   Else  
         vspdData.Focus
   End If      

End Function
	
</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtCd" SIZE=20 MAXLENGTH=50 tag="12XXXU" ALT="코드" ID="Text1"></TD>
				<TD CLASS="TD6" NOWRAP>&nbsp;<SPAN CLASS="normal" ID="lblTitle1"></SPAN>
				<SELECT NAME="rdoBpLabel" ONCHANGE="ClickRadio()" STYLE="WIDTH: 50px" TAG="11"></SELECT>
				</TD>
			</TR>		
						
			<TR>
				<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtNm" SIZE=30 MAXLENGTH=50 tag="12"   ALT="코드명" ID="Text2"></TD>
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"    onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   ONCLICK="OkClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	
	<INPUT TYPE="HIDDEN" NAME="txtHidden" ID="txtHidden" value=0>
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">  

	<INPUT TYPE="HIDDEN" NAME="txtisCondition" ID="txtisCondition">
		
	<TR>
		<TD WIDTH=100% HEIGHT=10><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=1 FRAMEBORDER=0 SCROLLING=no noresize  framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
	
	</FORM>
	
	
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

