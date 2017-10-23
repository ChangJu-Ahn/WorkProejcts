<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 1999/09/10
*  8. Modified date(Last)  : 1999/09/10
*  9. Modifier (First)     : Lee JaeHoo
* 10. Modifier (Last)      : Lee JaeHoo
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../inc/IncImage.js">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/eventpopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit	


Const BIZ_PGM_ID = "adogrpsortpopupbiz2.asp"

Const C_FieldNm    = 1
Const C_FieldLen   = 2
Const C_DefaultT   = 3
Const C_HIDESHOW   = 4
Const C_SortDirect = 5
Const C_SEQ_NO     = 6
Const C_NEXT_SEQ   = 7
Const C_Update     = 8     '--- 2003-01-04 ����� �߰� spreadUpdae 
Const C_LockFlg    = 9     '--- 2003-01-10 ����� �߰� spreadLockFlag
Const C_PairField  = 10    '--- 2003-01-20 ����� �߰� ���� �ʵ� 

	Dim objDOM
	Dim objDOMNodeList


Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgStrPrevKey
'Dim lgStrPrevKey2

Dim lgLngCurRows

Dim IsOpenPop
Dim lgSortKey

Dim arrParent
Dim arrParam1					 '--- First Parameter Group
Dim arrParam2					 '--- First Parameter Group
Dim arrTblField				 '--- Second Parameter Group(DB Table Field Name)
Dim arrGridHdr				 '--- Third Parameter Group(Column Captions of the SpreadSheet)
Dim arrReturn				 '--- Return Parameter Group
Dim gintDataCnt				 '--- Data Counts to Query
Dim arrFieldType 
Dim arrRowValue              '--- InitValue of Col=3 
Dim arrRowValue2              '--- InitValue of Col=3 
Dim iMessage 	             '--- 2003-01-10 ����� �߰�  

        iMessage         = "�����ֱ������ ��� 0�� �ɼ��� �����ϴ�"
		
		arrParent = window.dialogArguments
		arrParam1 = arrParent(0)
   top.document.title = arrParent(1)
		
Sub InitSpreadSheet()

	With vspdData
		.ReDraw = false
		'.MaxCols = C_NEXT_SEQ
		.MaxCols = C_PairField     ' 2003-01-20 ����� ���� 
		.Col = C_NEXT_SEQ
		.ColHidden = True

		.Col = C_SEQ_NO
		.ColHidden = True
		
		ggoSpread.Source = vspdData

		ggoSpread.Spreadinit
        vspdData.ColWidth(0) = 5		
		
        vspdData.Col = C_FieldNm
        vspdData.Row = 0
        vspdData.Text = "�ʵ��"
        vspdData.ColWidth(C_FieldNm) = 16
		
        vspdData.Col = C_SortDirect
        vspdData.Row = 0
        vspdData.Text = "���Ĺ��"
        vspdData.ColWidth(C_SortDirect) = 14
        vspdData.Row = -1
        vspdData.CellType = 8
        vspdData.TypeComboBoxList = "��������" & vbTab &  "��������"
		
        vspdData.Col = C_FieldLen
        vspdData.Row = 0
        vspdData.Text = "�ʵ����"
        vspdData.ColWidth(C_FieldLen) = 10
        vspdData.Row = -1
        vspdData.CellType = 13  ' Number
        vspdData.TypeNumberDecPlaces = 0
        vspdData.TypeSpin = True
        vspdData.TypeNumberMin = 0		
		
        vspdData.Col = C_DefaultT
        vspdData.Row = 0
        vspdData.Text = "�����ֱ����"
        vspdData.ColWidth(C_DefaultT) = 14
        vspdData.Row = -1
        vspdData.CellType = 13  ' Number
        vspdData.TypeNumberDecPlaces = 0
        vspdData.TypeSpin = True
        vspdData.TypeNumberMin = 0	
        

        vspdData.Col = C_HIDESHOW
        vspdData.Row = 0
        vspdData.Text = "����/����"
        vspdData.ColWidth(C_HIDESHOW) = 10
        vspdData.Row = -1
        vspdData.CellType = 8
        vspdData.TypeComboBoxList = "����" & vbTab &  "����"
        
        vspdData.Col = C_PairField   ' 2003-01-20 ����� �߰� 
        vspdData.Row = 0
        vspdData.Text = "���� �ʵ�"
        vspdData.ColWidth(C_PairField) = 10
        
        vspdData.Col = C_Update   ' 2003-01-04 ����� �߰� 
        vspdData.Row = 0
        .Colhidden = True
        
        vspdData.Col = C_LockFlg   ' 2003-01-10 ����� �߰� 
        vspdData.Row = 0
        .Colhidden = True
        
		ggoSpread.SSSetSplit(1)		
		Call SetSpreadLock("A")
        
        ' 2003-01-20 ����� �߰� 
        vspdData.TextTip = 2
        vspdData.TextTipDelay = 250
        vspdData.SetTextTipAppearance "MS Sans Serif", 12, 0, 0, &HC0FFFF, &H0
        vspdData.ScriptEnhanced = True 
       .ReDraw = true
    End With


	With vspdData1
		.ReDraw = false
		'.MaxCols = C_NEXT_SEQ
		.MaxCols = C_Update     ' 2003-01-04 ����� ���� 
		.Col = C_NEXT_SEQ
		.ColHidden = True
		
		.Col = C_SEQ_NO
		.ColHidden = True
		
		ggoSpread.Source = vspdData1

		ggoSpread.Spreadinit
        vspdData1.ColWidth(0) = 5		
		
        vspdData1.Col = C_FieldNm
        vspdData1.Row = 0
        vspdData1.Text = "�ʵ��"
        vspdData1.ColWidth(C_FieldNm) = 16
		
        vspdData1.Col = C_SortDirect
        vspdData1.Row = 0
        vspdData1.Text = "���Ĺ��"
        vspdData1.ColWidth(C_SortDirect) = 14
        vspdData1.Row = -1
        vspdData1.CellType = 8
        vspdData1.TypeComboBoxList = "��������" & vbTab &  "��������"
        .Colhidden = True
		
        vspdData1.Col = C_FieldLen
        vspdData1.Row = 0
        vspdData1.Text = "�ʵ����"
        vspdData1.ColWidth(C_FieldLen) = 10
        vspdData1.Row = -1
        vspdData1.CellType = 13  ' Number
        vspdData1.TypeNumberDecPlaces = 0
        vspdData1.TypeSpin = True
        vspdData1.TypeNumberMin = 0		
		
        vspdData1.Col = C_DefaultT
        vspdData1.Row = 0
        vspdData1.Text = "�����ֱ����"
        vspdData1.ColWidth(C_DefaultT) = 14
        vspdData1.Row = -1
        vspdData1.CellType = 13  ' Number
        vspdData1.TypeNumberDecPlaces = 0
        vspdData1.TypeSpin = True
        vspdData1.TypeNumberMin = 0	
        .Colhidden = True		

        vspdData1.Col = C_HIDESHOW
        vspdData1.Row = 0
        vspdData1.Text = "����/����"
        vspdData1.ColWidth(C_HIDESHOW) = 10
        vspdData1.Row = -1
        vspdData1.CellType = 8
        vspdData1.TypeComboBoxList = "����" & vbTab &  "����"
        
        vspdData1.Col = C_Update   ' 2003-01-04 ����� �߰� 
        vspdData1.Row = 0
        .Colhidden = True
        
		ggoSpread.SSSetSplit(1)		

		Call SetSpreadLock("B")
    
		.ReDraw = true
    End With
        
End Sub

Sub SetSpreadLock(ByVal pSpdNo)

    Select Case UCase(pSpdNo)
      Case "A"
         ggoSpread.Source = vspdData
         ggoSpread.SSSetProtected C_FieldNm , -1 ,-1
	     'ggoSpread.SSSetRequired C_FieldNm, -1, C_FieldNm
	     ggoSpread.SSSetRequired C_FieldLen,-1,C_FieldLen
	  Case "B"   
         ggoSpread.Source = vspdData1
         ggoSpread.SSSetProtected C_FieldNm , -1 ,-1
         'ggoSpread.SSSetRequired C_FieldNm, -1, C_FieldNm
         ggoSpread.SSSetRequired C_FieldLen,-1,C_FieldLen
   End Select       

End Sub



'=========================================================================================================
Function OKClick()
    Dim iNode
    Dim ii
'	Dim arrReturn
	
    Call objDOM.documentElement.setAttribute("Split", cboOrderBy1.value)	
	
	ggoSpread.Source = vspdData
	If (checkDefaultT(ggoSpread.Source)) then
	 Msgbox iMessage, vbExclamation, gLogoName & "-[Warning]"
	 Exit Function
	End if
    
    For ii = 1 To  vspdData.MaxRows
        vspdData.Row = ii
        'vspdData.Col = 0       
        vspdData.Col = C_Update   '2003-01-04 ����� ����    
        
        If ggoSpread.UpdateFlag = vspdData.Text Then
           vspdData.Col = C_SEQ_NO
           Set iNode = objDOM.selectSingleNode("//DATA[@SEQ = '" & vspdData.text & "']") 
           
           If UCase(TypeName(iNode)) <> "NOTHING" Then
              vspdData.Col = C_FieldNm  : Call iNode.setAttribute("FIELD_NM", vspdData.text)
              vspdData.Col = C_FieldLen : Call iNode.setAttribute("FIELD_LEN", vspdData.text)
              vspdData.Col = C_DefaultT 
        
              If Isnumeric(vspdData.Value) Then
                 If CInt(vspdData.Value) > 0 Then
                    Call iNode.setAttribute("DEFAULT_T", vspdData.Value)
                 Elseif CInt(vspdData.Value) = 0 And iNode.attributes.getNamedItem("DEFAULT_T").nodeValue <> "V" Then
                    Call iNode.setAttribute("DEFAULT_T", "V")
                 End if
              End if
              
              vspdData.Col = C_SortDirect 

              If vspdData.Value = 0 Then
                 Call iNode.setAttribute("SORT_DIR", "ASC")
              Else   
                 Call iNode.setAttribute("SORT_DIR", "DESC")
              End If
        
              vspdData.Col = C_HIDESHOW 

              If vspdData.Value = 0 Then
                 Call iNode.setAttribute("HIDDEN", "F")
              Else   
                 Call iNode.setAttribute("HIDDEN", "T")
             End If
          End If
      End If

    Next
    
	ggoSpread.Source = vspdData1
	
    For ii = 1 To  vspdData1.MaxRows
        vspdData1.Row = ii    
        'vspdData1.Col = 0       
        vspdData1.Col = C_Update   '2003-01-04 ����� ����     
        
        If ggoSpread.UpdateFlag = vspdData1.Text Then
           vspdData1.Col = C_SEQ_NO
           Set iNode = objDOM.selectSingleNode("//DATA[@SEQ = '" & vspdData1.text & "']") 
           If UCase(TypeName(iNode)) <> "NOTHING" Then
           
              vspdData1.Col = C_FieldNm  : Call iNode.setAttribute("FIELD_NM" , vspdData1.text)
              vspdData1.Col = C_FieldLen : Call iNode.setAttribute("FIELD_LEN", vspdData1.text)
              vspdData1.Col = C_DefaultT 
        
              If Isnumeric(vspdData1.Value) Then
                 If CInt(vspdData1.Value) > 0 Then
                    Call iNode.setAttribute("DEFAULT_T", "L" & vspdData1.Value)
                 End if
              End if
        
              If Isnumeric(vspdData1.Value) Then
                 If CInt(vspdData1.Value) = 0 Then
                    Call iNode.setAttribute("DEFAULT_T", "L")
                 End if
              End if
              
'              vspdData1.Col = C_SortDirect 

'              If vspdData1.Value = 0 Then
'                 Call iNode.setAttribute("SORT_DIR", "ASC")
'              Else   
'                 Call iNode.setAttribute("SORT_DIR", "DESC")
'              End If
        
              vspdData1.Col = C_HIDESHOW 

              If vspdData1.Value = 0 Then
                 Call iNode.setAttribute("HIDDEN", "F")
              Else   
                 Call iNode.setAttribute("HIDDEN", "T")
             End If
          End If
      End If

    Next
    
	ReDim arrReturn(1)

    If CHECKBOX1.checked =  True Then
       arrReturn(0) = "T"
    Else
       arrReturn(0) = "F"
    End if   
    
    arrReturn(1) =  objDOM.XML


	Self.Returnvalue = arrReturn
    Self.Close()
					
End Function
'========================== 2003-01-10 ����� �߰� : DefaultT���� ��� 0���� üũ=============
Function checkDefaultT(ByVal obj)
   Dim maxNum
    
   For maxNum=1 to obj.MaxRows
      obj.Row = maxNum
      obj.Col = C_DefaultT 
      IF Clng(obj.value) <> 0 then
       checkDefaultT = False
       Exit Function
      End if
   Next
   checkDefaultT = True
End Function	
'=========================================================================================================
Sub RestoreClick()
'    Dim arrReturn
    
    ReDim arrReturn(1)
    arrReturn(0) = "R"
    Self.Returnvalue = arrReturn
	Self.Close()
End Sub
	
'=========================================================================================================
Function CancelClick()
'    Dim arrReturn
    
    ReDim arrReturn(1)
    arrReturn(0) = "X"
    Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================================================================================
Sub Form_Load()

	Dim IntRetCD
	Dim ii,iii
	Dim jj,kk
	Dim iDefaultTValue
	Dim iColTemp
	Dim iComboLength
	Dim iStrTemp

    ReDim arrReturn(1)
    arrReturn(0) = "X"
    Self.Returnvalue = arrReturn
	
	Set objDOM = CreateObject("Microsoft.XMLDOM")
	objDOM.async = false

	objDOM.loadXML(arrParam1)
	
    'Call LoadInfTB19029  
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet

    cboOrderBy1.length = 2
        
    cboOrderBy1.options(0).value = "0"
    cboOrderBy1.options(0).text  = " "
    
    Set objDOMNodeList = objDOM.selectNodes("//DATA")
    
    vspdData1.MaxRows = 0
    vspdData.MaxRows = 0
    vspdData.ReDraw = false 
    vspdData1.ReDraw = false 
    
    jj = 1
    kk = 1
    iColTemp = 1
    iComboLength = 1
    For ii = 1 To  objDOMNodeList.length
        iDefaultTValue = objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue
'        if objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_TYPE").nodeValue <> "HH" And iDefaultTValue <> "HH" then
        
			Select Case UCase(iDefaultTValue)
			    Case "L","HH"
			        ggoSpread.Source = vspdData1
					vspdData1.MaxRows = vspdData1.MaxRows + 1
					vspdData1.Row = jj
					vspdData1.Col = 0  : vspdData1.text = jj			
					vspdData1.Col = C_FieldNm  : vspdData1.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_NM").nodeValue
					vspdData1.Col = C_FieldLen : vspdData1.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_LEN").nodeValue
			
					vspdData1.Col = C_DefaultT
					ggoSpread.SpreadLock C_DefaultT, jj , C_DefaultT , jj

					vspdData1.Col = C_HIDESHOW
					If objDOMNodeList.item(ii-1).attributes.getNamedItem("HIDDEN").nodeValue = "F" Then
					   vspdData1.Value = 0
					Else
					   vspdData1.Value = 1
					End If
        
					'vspdData1.Col = C_SortDirect 
					'vspdData1.Text = ""
					'ggoSpread.SpreadLock C_SortDirect, Col_Num , C_SortDirect , Col_Num

        
					vspdData1.Col = C_SEQ_NO : vspdData1.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("SEQ").nodeValue
					vspdData1.Col = C_Update : vspdData1.text = ""
			
					If objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_TYPE").nodeValue = "HH" Or objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue = "HH" then			
					'    iColNum2 = iColNum2 - 1
					'    vspdData1.Col = 3   ' C_Default_T = 0
					'    vspdData1.text = 0
					'    vspdData1.RowHidden = True
					     ggoSpread.SSSetProtected -1,jj,jj
					End if
					jj = jj + 1
			
			
			    Case Else
			        ggoSpread.Source = vspdData
			        vspdData.MaxRows = vspdData.MaxRows + 1
			        vspdData.Row = kk
			        vspdData.Col = 0  : vspdData.text = iColTemp			
			        vspdData.Col = C_FieldNm  : vspdData.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_NM").nodeValue
			        vspdData.Col = C_FieldLen : vspdData.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_LEN").nodeValue
			        iStrTemp = objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue
			        vspdData.Col = C_DefaultT : vspdData.text = iStrTemp

			        vspdData.Col = C_HIDESHOW
			        If objDOMNodeList.item(ii-1).attributes.getNamedItem("HIDDEN").nodeValue = "F" Then
			           vspdData.Value = 0
			        Else
			           vspdData.Value = 1
			        End If

			        vspdData.Col = C_SortDirect 
			        If objDOMNodeList.item(ii-1).attributes.getNamedItem("SORT_DIR").nodeValue = "ASC" Then
			           vspdData.Value = 0
			        Else
			           vspdData.Value = 1
			        End If
			        
			        vspdData.Col = C_SEQ_NO : vspdData.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("SEQ").nodeValue
			        vspdData.Col = C_NEXT_SEQ : vspdData.text = objDOMNodeList.item(ii-1).attributes.getNamedItem("NEXT_SEQ").nodeValue
			        
			        ' 2003-01-20 ����� �߰� 
			        If vspdData.text = "" OR vspdData.text="0" then
			          vspdData.Col = C_PairField : vspdData.text = ""
			        Else
			           For iii = 1 To  objDOMNodeList.length
			             If vspdData.text = objDOMNodeList.item(iii-1).attributes.getNamedItem("SEQ").nodeValue then
			               vspdData.Col = C_PairField : vspdData.text = objDOMNodeList.item(iii-1).attributes.getNamedItem("FIELD_NM").nodeValue
			               Exit For
			             End if
			           Next
			        End If
			        
			        ggoSpread.SpreadLock C_PairField, kk , C_PairField , kk  '2003-01-20 ����� �߰� 
                    vspdData.Col = C_Update : vspdData.text = ""  '2003-01-04 ����� �߰� 
                    
                    If objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_TYPE").nodeValue = "HH" Or objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue = "HH" then			
                        'vspdData.Col = 3
                        'vspdData.text = 0
                        'vspdData.RowHidden = True
                        ggoSpread.SpreadLock C_FieldNm, kk , C_FieldLen , kk   '2003-01-10 ����� �߰� 
                        ggoSpread.SpreadLock C_HIDESHOW, kk , C_HIDESHOW , kk  '2003-01-10 ����� �߰� 
                        
                        vspdData.Col = C_LockFlg    '2003-01-10 ����� �߰� 
                        vspdData.text = "L"         '2003-01-10 ����� �߰� 
                    Else
'                        If Isnumeric(iStrTemp) Or iStrTemp = "V" Then
                            If iStrTemp = "V" Then
                                vspdData.Col = C_NEXT_SEQ
                                iStrTemp = vspdData.Text
                                If Isnumeric(iStrTemp) Then
                                    If CLng(iStrTemp) <> 0 Then
                                        vspdData.Row = CLng(iStrTemp)
                                        vspdData.Col = C_DefaultT
                                        iStrTemp = vspdData.Text
                                        If Isnumeric(iStrTemp) Then
                                            If CLng(iStrTemp) <> 0 Then
                                                cboOrderBy1.length = iComboLength + 1
                                                cboOrderBy1.options(iComboLength).value = objDOMNodeList.item(ii-1).attributes.getNamedItem("SEQ").nodeValue
                                                cboOrderBy1.options(iComboLength).text  = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_NM").nodeValue			
                                                iComboLength = iComboLength + 1
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                cboOrderBy1.length = iComboLength + 1
                                cboOrderBy1.options(iComboLength).value = objDOMNodeList.item(ii-1).attributes.getNamedItem("SEQ").nodeValue
                                cboOrderBy1.options(iComboLength).text  = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_NM").nodeValue			
                                iComboLength = iComboLength + 1
                            End If
                         vspdData.Col = C_LockFlg    '2003-01-10 ����� �߰� 
                         vspdData.text = "UL"        '2003-01-10 ����� �߰� 
'                        End If
			            'iColTemp = iColTemp + 1 '2003-01-10 ����� ���� 
			        End If
			        iColTemp = iColTemp + 1 '2003-01-10 ����� ���� 
         			kk = kk + 1
			End Select
			
'       end if
    Next
    iStrTemp = objDOM.documentElement.GetAttribute("Split")
    If Not isnull(iStrTemp) Then
        cboOrderBy1.value = iStrTemp
    End If
    
    Call InitDefaultT(vspdData)
    Call InitRowValue(vspdData) 
    'Call InitRowValue2(vspdData1) 
    vspdData.ReDraw = True 
    vspdData1.ReDraw = True 
    
    
End Sub


Sub vspdData_Change(ByVal Col ,ByVal Row)
    Dim iSeq
    Dim iNode
    Dim iTempValue
    Dim iLoop

    vspdData.Row = Row
    vspdData.Col = Col
    iTempValue = vspdData.Value    
	ggoSpread.Source = vspdData
	
	if vspdData.Col = 3 then  '2002-12-16 ����� �߰� 
      Call ZADOSort(ggoSpread.Source,Row,vspdData.Value)
    end if
    
    Call InitRowValue(ggoSpread.Source)  '2002-12-16 ����� �߰� 
    Call RowLockComboList(ggoSpread.Source) '2002-12-27 ����� �߰� 
    
    'ggoSpread.UpdateRow Row
    Call spreadUpdate(ggoSpread.Source, Row)  '2003-01-04 ����� ���� 

'    If Col = C_HIDESHOW Then
 '       vspdData.Col = C_SEQ_NO
  '      Set iNode = objDOM.selectSingleNode("//DATA[@SEQ = '" & vspdData.text & "']") 
   '     iSeq = CInt(inode.getAttribute("NEXT_SEQ"))
    '    If iSeq <> 0 Then
     '       For iLoop = 1 To vspdData.MaxRows
      '          If iLoop <> Row Then
       '             vspdData.Row = iLoop
        '            If CInt(vspdData.Text) = iSeq Then
         '               vspdData.Col = Col
          '              vspdData.Value = iTempValue
           '             ggoSpread.UpdateRow iLoop
            '            Exit For
             '       End If
              '  End If
'            Next
 '       End If
  '  End If    
  
  
End Sub

Sub vspdData1_Change(ByVal Col ,ByVal Row)

    vspdData1.Row = Row
    vspdData1.Col = Col
    
	ggoSpread.Source = vspdData1
    vspdData.ReDraw = False
    
	if vspdData1.Col = 3 then  '2002-12-16 �߰� 
      Call ZADOSort(ggoSpread.Source,Row,vspdData1.Value)
    end if
    
    Call InitRowValue2(ggoSpread.Source)
    
    'ggoSpread.UpdateRow Row
    Call spreadUpdate(ggoSpread.Source, Row)  '2003-01-04 ����� ���� 
    
    vspdData.ReDraw = True
End Sub

'========================== 2003-01-20 ����� �߰� : ToolTip ��Ÿ����=============
Sub vspdData_ScriptTextTipFetch(ByVal Col,ByVal Row,MultiLine,TipWidth,TipText,ShowTip)

  IF Col = C_PairField then
    ShowTip = True
    TipWidth = 500
    MultiLine = True 
    TipText = "�ش� �ʵ�� �׻� ���� ��Ÿ���� �ʵ���� ��Ÿ���ϴ�"
  End If

End Sub
'========================== 2002-12-27 ����� �߰� : ������ �޺�����Ʈ ����=============
Function RowLockComboList(ByVal obj)
    Dim ii
    Dim iDefaultTValue
    Dim iComboLength
    Dim iRowTemp
    Dim iValueTemp,iValueTemp2,iSEQTemp,iNEXTTemp
    Dim iStrTemp
    
    iStrTemp = cboOrderBy1.value
    
    cboOrderBy1.length = 1
    cboOrderBy1.options(0).value = "0"
    cboOrderBy1.options(0).text  = " "
    
    iRowTemp= 1
    iComboLength = 1

    For ii = 1 To  objDOMNodeList.length
               obj.Row = iRowTemp

               iDefaultTValue = objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue
               If UCase(Mid(iDefaultTValue,1,1)) = "L" Or objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_TYPE").nodeValue = "HH" Or objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue = "HH" then			
              ' ó������ �ѷ����� ���� Row
               iRowTemp = iRowTemp - 1
               else 
                    
               obj.Col = C_DefaultT    ' C_DefaultT = 3
               iValueTemp = obj.value  ' ���� ����� �����ֱ� ������ �� 
                     
               'obj.Col = C_SEQ_NO
               'iSEQTemp = obj.value   ' ���� ���õ� SEQ �� 
                    
               obj.Col = C_NEXT_SEQ   ' C_NEXT_SEQ = 7  
               iNEXTTemp = obj.value  ' ���� ���õ� NextSEQ�� 
               
               'msgbox "���ȣ : " & iRowTemp & vbcrlf & "�����ֱ� �� : " & iValueTemp & vbcrlf & "üũ�� SEQ �� : " & iNEXTTemp
                     
                  If iValueTemp <> 0 Or iNEXTTemp <> 0 then
                     'msgbox "üũ�� �� : " & valCheck(obj, iNEXTTemp)
                    if (iValueTemp <> 0) Or (valCheck(obj, iNEXTTemp) <> 0) then
                           cboOrderBy1.length = iComboLength + 1
                           cboOrderBy1.options(iComboLength).value = objDOMNodeList.item(ii-1).attributes.getNamedItem("SEQ").nodeValue
                           cboOrderBy1.options(iComboLength).text  = objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_NM").nodeValue			
                           iComboLength = iComboLength + 1
                    end if
                         
		          End if  
   	          End If
			        
   	          iRowTemp = iRowTemp + 1
			        
    Next
    cboOrderBy1.value = iStrTemp
End Function

'========================== 2002-12-27 ����� �߰�=============
Function valCheck(ByVal obj ,ByVal chkValue)
 Dim ii
 Dim iRowTemp
 
 iRowTemp = 1
 For ii = 1 To  objDOMNodeList.length
   
   If objDOMNodeList.item(ii-1).attributes.getNamedItem("FIELD_TYPE").nodeValue = "HH" Or objDOMNodeList.item(ii-1).attributes.getNamedItem("DEFAULT_T").nodeValue = "HH" then			
    ' ó������ �ѷ����� ���� Row
    iRowTemp = iRowTemp - 1
   else 
   
    If chkValue = objDOMNodeList.item(ii-1).attributes.getNamedItem("SEQ").nodeValue then
       obj.Row = iRowTemp
       obj.Col = 3
       valCheck = obj.value
       Exit Function
    End IF
    
   end if
   
   iRowTemp = iRowTemp + 1
   
 Next
 
End Function

'========================== 2002-12-16 ����� �߰� : ���������� �ʱⰪ ����=============
Function InitRowValue(ByVal obj)
   Dim maxNum
   Redim arrRowValue(obj.MaxRows)
    
   For maxNum=1 to obj.MaxRows
      obj.Row = maxNum
      obj.Col = 3 
      arrRowValue(maxNum-1) = obj.value
   Next
   
End Function

Function InitRowValue2(ByVal obj)
   Dim maxNum
   Redim arrRowValue2(obj.MaxRows)
    
   For maxNum=1 to obj.MaxRows
      obj.Row = maxNum
      obj.Col = 3 
      arrRowValue2(maxNum-1) = obj.value
   Next
   
End Function

'========================== 2003-01-04 ����� �߰� : ���������� ������Ʈ ��ŷ �ٸ������� =============
Function spreadUpdate(ByVal obj ,ByVal Row)
  
  obj.Row = Row
  obj.Col = C_Update
  obj.text = "����"
  
End Function

'========================== 2003-01-04 ����� �߰� : ���������� C_Defualt_T Init =============
Function InitDefaultT(ByVal obj)
   Dim maxNum
   Dim arr_1
   Dim arr_2
   dim iii,jjj,tmp,matchnum  ' ���� �迭�� �����Ʈ 
   Dim strFValue,strSValue
   
   ReDim arr_1(obj.MaxRows)
   ReDim arr_2(obj.MaxRows)

   For maxNum=1 to obj.MaxRows   '2003-01-10 ����� �߰� 
     obj.Row = maxNum
     obj.Col = C_LockFlg
     
     If obj.value = "UL" then
       obj.Col = C_DefaultT
       strFValue = obj.value
       
        For iii=1 to obj.MaxRows
          obj.Row = iii
          obj.Col = C_DefaultT
          strSValue = obj.value
          
          obj.Col = C_LockFlg
          
          IF obj.value = "L" AND CLng(strSValue)<>0 AND CLng(strFValue) = CLng(strSValue) then
            Call AddOneValue(obj,strFValue)
            Exit Function
          End if
          
        Next
        
     End if
   Next
   
   '���� 
   For maxNum=1 to obj.MaxRows
	obj.Row = maxNum
	obj.Col = C_DefaultT
   
    arr_1(maxNum-1) = obj.value
    arr_2(maxNum-1) = obj.value
    
   Next  '���������� �� ���� 2���� �迭�� ����(�ϳ��� ����, �ϳ��� �ε���)
   
   
   
   
    for iii=0 to ubound(arr_1)-2
	  for jjj=iii+1 to ubound(arr_1)-1
	    if CLng(arr_1(iii)) > CLng(arr_1(jjj)) then
	      tmp = arr_1(jjj)
	      arr_1(jjj)=arr_1(iii)
	      arr_1(iii)=tmp
	    end if
	  next
	next
	

	' �ε����� �ο��� �迭�� �̿��� ���� ��迭 
	matchnum = 0
	for iii=0 to ubound(arr_1)-1
	  if arr_1(iii) <> "" AND arr_1(iii) <> "0" then
	    
	    for jjj=0 to ubound(arr_2)
	    
	      if arr_1(iii) = arr_2(jjj) then
	       matchnum = matchnum + 1
	       
	       obj.Row = jjj + 1
	       obj.Col = C_DefaultT
	       obj.text = matchnum
	       'ggoSpread.UpdateRow jjj + 1
		   Call spreadUpdate(obj, jjj + 1)
	       
	      end if 
	    next
	    
	  end if
	next
   
End Function

'========================== 2003-01-10 ����� �߰� : �ش� ������ ū ���� +1 =============
Function AddOneValue(ByVal obj ,ByVal val)
  Dim maxNum
  Dim strLock
  Dim iValue
  
  For maxNum =1 to obj.maxRows
    obj.Row = maxNum
    obj.Col = C_LockFlg
    strLock = obj.value
    obj.Col = C_DefaultT
    iValue = obj.value
    If strLock = "UL" AND CLng(iValue) >= CLng(val) then
      obj.text = iValue + 1
    End if
  Next
  
  Call InitDefaultT(obj)
  
End Function
'=========================== 2002-12-16 ����� �߰� =============
Function ZADOSort(ByVal obj ,ByVal Row ,ByVal ChangeValue)
   Dim maxNum
   Dim compareFlg
   Dim oneFlg
   Dim arr_1
   Dim arr_2
   
   ReDim arr_1(obj.MaxRows)
   ReDim arr_2(obj.MaxRows)

   compareFlg = false
   oneFlg = false
   
   For maxNum=1 to obj.MaxRows
      obj.Row = maxNum
      obj.Col = 3 
      
      if obj.value="" then obj.text = 0 end if
      if (obj.value = 1) then
         oneFlg = True  ' ���� 1���� ���� 
      end if
      
      if (maxNum <> Row) AND (ChangeValue = obj.value) then
          compareFlg = True
      end if
   Next
   
   'msgbox "����1���� �ִ°�? " & oneFlg & vbcrlf & "�Ȱ��� �񱳰��� �ִ°�? " & compareFlg
   
   if oneFlg then
   
		if compareFlg then
		  
				For maxNum=1 to obj.MaxRows
				   obj.Row = maxNum
				   obj.Col = 3 
				     
				  if (maxNum <> Row) AND (obj.value <> 0) then
                    'msgbox changeValue & " | " & arrRowValue(Row-1)
                    
						if ChangeValue > arrRowValue(Row-1) then  '���������� ū������ ��ȯ 
							
							if arrRowValue(Row-1) = 0 then        ' ������ ���� 0�϶� ��ȯ������ ū�͸� +1
                                if obj.value >= ChangeValue then
								  obj.text = obj.value + 1
								  'ggoSpread.UpdateRow maxNum
								  Call spreadUpdate(obj, maxNum)  '2003-01-04 ����� ���� 
								end if
                            else
								if obj.value >= arrRowValue(Row-1) and obj.value <= ChangeValue then
								  obj.text = obj.value - 1
								  'ggoSpread.UpdateRow maxNum
								  Call spreadUpdate(obj, maxNum)  '2003-01-04 ����� ���� 
								end if
							end if
							
						else                                      'ū������ ���������� ��ȯ 
						    if ChangeValue <= obj.value then
							  obj.text = obj.value + 1
							  'ggoSpread.UpdateRow maxNum
							  Call spreadUpdate(obj, maxNum)  '2003-01-04 ����� ���� 
							end if
						end if
					
				    
				  end if
				  
				Next
				
		end if 
   
   else ' ���� 1���� ������ ���� ���� �����ͺ��� �ϳ��� ������ 
       
               For maxNum=1 to obj.MaxRows
				   obj.Row = maxNum
				   obj.Col = 3 
				     
				  if (maxNum <> Row) AND (obj.value <> 0) then

				    if ChangeValue >= obj.value then
				      obj.text = obj.value - 1
				      'ggoSpread.UpdateRow maxNum
				      Call spreadUpdate(obj, maxNum)  '2003-01-04 ����� ���� 
				    end if
				    
				  end if
				  
				Next
     
   end if
   
   '���� 
   For maxNum=1 to obj.MaxRows
	obj.Row = maxNum
	obj.Col = 3
   
    arr_1(maxNum-1) = obj.value
    arr_2(maxNum-1) = obj.value
    
   Next  '���������� �� ���� 2���� �迭�� ����(�ϳ��� ����, �ϳ��� �ε���)
   
   
   dim iii,jjj,tmp,matchnum  ' ���� �迭�� �����Ʈ 
   
    for iii=0 to ubound(arr_1)-2
	  for jjj=iii+1 to ubound(arr_1)-1
	    if CLng(arr_1(iii)) > CLng(arr_1(jjj)) then
	      tmp = arr_1(jjj)
	      arr_1(jjj)=arr_1(iii)
	      arr_1(iii)=tmp
	    end if
	  next
	next
	

	' �ε����� �ο��� �迭�� �̿��� ���� ��迭 
	matchnum = 0
	for iii=0 to ubound(arr_1)-1
	  if arr_1(iii) <> "" AND arr_1(iii) <> "0" then
	    
	    for jjj=0 to ubound(arr_2)
	    
	      if arr_1(iii) = arr_2(jjj) then
	       matchnum = matchnum + 1
	       
	       obj.Row = jjj + 1
	       obj.Col = 3
	       obj.text = matchnum
	       'ggoSpread.UpdateRow jjj + 1
	       Call spreadUpdate(obj, jjj + 1)  '2003-01-04 ����� ���� 
	      end if 
	    next
	    
	  end if
	next
   
   
   
End Function



'=========================================================================================================

</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCOM.inc" -->


</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<TABLE CLASS="BasicTB" CELLSPACING=0>
	<TR>
		<TD HEIGHT=5></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../CShared/image/table/seltab_up_bg.gif"><img src="../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ȭ��ȯ�漳��</font></td>
								<td background="../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_60%> CELLSPACING=0>
				<TR>
					<TD CLASS=TD5 WIDTH=100% HEIGHT=1% valign=CENTER> ������ </TD>
					<TD CLASS=TD6 WIDTH=100% HEIGHT=1% valign=CENTER>	<SELECT NAME="cboOrderBy1" STYLE="WIDTH: 110px" TAG="1"><OPTION selected></SELECT>
					</TD>
				</TR>
				<TR>
					<TD CLASS=TD5 >���� ����
					<TD CLASS=TD6 >&nbsp;
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=50% valign=top colspan=2>
								<script language =javascript src='./js/zadogrouppopup_vspdData_vspdData.js'></script>
					</TD>
				</TR>
				<TR>
					<TD CLASS=TD5 >���� ����Ÿ
					<TD CLASS=TD6 >&nbsp;
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=50% valign=top colspan=2>
								<script language =javascript src='./js/zadogrouppopup_vspdData1_vspdData1.js'></script>
					</TD>
				</TR>
		</TABLE></TD>
	</TR>
	<TR HEIGHT=30>
		<TD HEIGHT=30>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
							<INPUT TYPE=CHECKBOX CLASS=CHECK id=CHECKBOX1 name=CHECKBOX1>�׸�����������</TD>
						<TD WIDTH=30% ALIGN=RIGHT>
							&nbsp;&nbsp;
							<IMG SRC="../../CShared/image/zpReSet_d.gif"  Style="CURSOR: hand" ALT="ReSet"  NAME="ReSet"  ONCLICK="RestoreClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/zpReSet.gif',1)"      ></IMG>
							<IMG SRC="../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     ONCLICK="OkClick()"      onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"     ></IMG>
							<IMG SRC="../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" ONCLICK="CancelClick()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=10>
		<TD HEIGHT=10> </TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>