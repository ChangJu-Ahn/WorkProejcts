<%'===================================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S3133mb1
'*  4. Program Name         : �����ϻ�����Ȳ��ȸ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Seo Jinkyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2002/09/18 Ado ǥ������ 
'*                            -2002/12/20 : Get��� --> Post������� ���� 
'========================================

                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4	'�� : DBAgent Parameter ���� 
	Dim lgStrData														'�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgMaxCount														'�� : Spread sheet �� visible row �� 
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

'--------------- ������ coding part(��������,Start)----------------------------------------------------
	Dim strSalesGrpNm					'�����׷� 
	Dim strItemCodeNm						'ǰ�� 
	Dim strSoldToPartyNm						'�ŷ�ó 
	Dim strSoTypeNm						'�������� 
	Dim MsgDisplayFlag
   
	MsgDisplayFlag = False
	Dim iFrPoint
    iFrPoint=0
'--------------- ������ coding part(��������,End)------------------------------------------------------

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q","S","NOCOOKIE","QB")
	
    Call HideStatusWnd 
     
if Request("txtFlgMode") = Request("OPMD_UMODE") Then
		txtMode= Request("txtMode")
		txtSalesGrp= Request("HSalesGrp")
		txtSoDtFrom= Request("HSoDtFrom")
		txtSoDtTo= Request("HSoDtTo")
		txtDlvyDtFrom= Request("HDlvyDtFrom")
		txtDlvyDtTo= Request("HDlvyDtTo")
		txtSoldToParty= Request("HSoldToParty")
		txtItemCode= Request("HItemCode")
		txtSoType= Request("HSoType")
		txtTrackingNo = Request("HtxtTrackingNo")
		lgStrPrevKey= Request("txt_lgStrPrevKey")
else
		txtMode= Request("txtMode")
		txtSalesGrp= Request("txtSalesGrp")
		txtSoDtFrom= Request("txtSoDtFrom")
		txtSoDtTo= Request("txtSoDtTo")
		txtDlvyDtFrom= Request("txtDlvyDtFrom")
		txtDlvyDtTo= Request("txtDlvyDtTo")
		txtSoldToParty= Request("txtSoldToParty")
		txtItemCode= Request("txtItemCode")
		txtSoType= Request("txtSoType")
		txtTrackingNo = Request("txtTrackingNo")
		lgStrPrevKey= Request("txt_lgStrPrevKey")
end if

	lgStrPrevKey	 = Request("txt_lgStrPrevKey")
    lgPageNo         = UNICInt(Trim(Request("txt_lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100								                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList     = Request("txt_lgSelectList")
    lgTailList       = Request("txt_lgTailList")
    lgSelectListDT   = Split(Request("txt_lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(lgMaxCount) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()

	SetConditionData = False
	
    On Error Resume Next
	
	If Not(rs1.EOF Or rs1.BOF) Then
       strSalesGrpNm =  rs1(1)
        rs1.Close
        Set rs1 = Nothing       
    Else
        rs1.Close
        Set rs1 = Nothing
            
		If Len(txtSalesGrp) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGrp.focus    
                </Script>
            <%
		End If	
    End If   

    
	
	If Not(rs2.EOF Or rs2.BOF) Then
       strItemCodeNm =  rs2(1)
        rs2.Close
        Set rs2 = Nothing               
    Else
        rs2.Close
        Set rs2 = Nothing        
   		If Len(txtItemCode) And MsgDisplayFlag = False Then
		    Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		    MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtItemCode.focus    
                </Script>
            <%		    
		End If	
    End If   
    
    
    
	If Not(rs3.EOF Or rs3.BOF) Then
	    strSoldToPartyNm =  rs3(1)
		rs3.Close
		Set rs3 = Nothing       
    Else
        rs3.Close
        Set rs3 = Nothing    
   		If Len(txtSoldToParty) And MsgDisplayFlag = False Then
		    Call DisplayMsgBox("970000", vbInformation, "�ֹ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		    MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSoldToParty.focus    
                </Script>
            <%		    
		End If	
    End If   
    
    
    
    If Not(rs4.EOF Or rs4.BOF) Then
       strSoTypeNm =  rs4(1)
        rs4.Close
        Set rs4 = Nothing       
    Else
        rs4.Close
        Set rs4 = Nothing    
   		If Len(txtSoType) And MsgDisplayFlag = False Then
    		Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	    	MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSoType.focus    
                </Script>
            <%	    	
		End If	
    End If   
    
    

	If MsgDisplayFlag = True Then Exit Function

	SetConditionData = True
	
End Function
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
																		  '�Ʒ��� ���� ȭ��ܿ��� �־� �ִ� query�� where�������� �� �� �ִ�.	
    Dim arrVal(3)														  '��: ȭ�鿡�� �˾��Ͽ� query
																		  '�Ʒ��� ���� UNISqlId(1),UNISqlId(2), UNISqlId(3)�� where�������� �� �� �ִ�.
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
																		  '��ȸȭ�鿡�� �ʿ��� query���ǹ����� ����(Statements table�� ����)
    Redim UNIValue(4,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 

    UNISqlId(0) = "S3133QA101"											  ' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(1) = "s0000qa005"											  ' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(2) = "s0000qa001"											  ' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(3) = "s0000qa002"											  ' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(4) = "s0000qa007"											  ' main query(spread sheet�� �ѷ����� query statement)
    
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
	strVal = ""
     '---�߻����� 
    If Len(Trim(txtSoDtFrom)) Then
    	strVal 	= strVal & "AND a.so_dt >=  " & FilterVar(uniConvDate(Trim(txtSoDtFrom)), "''", "S") & " "
    End If
    
    If Len(Trim(txtSoDtTo)) Then
    	strVal 	= strVal & "AND a.so_dt <= " & FilterVar(uniConvDate(Trim(txtSoDtTo)), "''", "S") & " "
    End If    
    
    If Len(Trim(txtDlvyDtFrom)) Then
    	strVal 	= strVal & "AND b.DLVY_DT >=  " & FilterVar(uniConvDate(Trim(txtDlvyDtFrom)), "''", "S") & " "
    End If
    
    If Len(Trim(txtDlvyDtTo)) Then
    	strVal 	= strVal & "AND b.DLVY_DT <= " & FilterVar(uniConvDate(Trim(txtDlvyDtTo)), "''", "S") & " "
    End If    

    '---�����׷� 
    If Len(Trim(txtSalesGrp)) Then
    	strVal 	  = strVal & "AND c.sales_grp =  " & FilterVar(txtSalesGrp, "''", "S") & "  "    	
    End If
    arrVal(0) = FilterVar(Trim(txtSalesGrp), " ", "S")

	'---ǰ�� 
	If Len(Trim(txtItemCode)) Then
    	strVal 	  = strVal & "AND f.item_cd  =  " & FilterVar(txtItemCode, "''", "S") & "  "      	
    End If
    arrVal(1) = FilterVar(Trim(txtItemCode), " ", "S")

	'---�ŷ�ó 
	If Len(Trim(txtSoldToParty)) Then
    	strVal 	  = strVal & "AND d.bp_cd =  " & FilterVar(txtSoldToParty, "''", "S") & "  "       	
    End If
    arrVal(2) = FilterVar(Trim(txtSoldToParty), " ", "S")
    
    '---�������� 
	If Len(Trim(txtSoType)) Then
    	strVal 	  = strVal & "AND g.so_type =  " & FilterVar(txtSoType, "''", "S") & "  "        
    End If
    arrVal(3) = FilterVar(Trim(txtSoType), " ", "S")
    
    If Len(txtTrackingNo) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(Trim(txtTrackingNo), "''" , "S") & ""
	End If

	UNIValue(0,1)  = UCase(Trim(strVal))	
	UNIValue(1,0)  = UCase(arrVal(0))	
	UNIValue(2,0)  = UCase(arrVal(1))	
	UNIValue(3,0)  = UCase(arrVal(2))	
	UNIValue(4,0)  = UCase(arrVal(3))	
	
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg, vbInformation, I_MKSCRIPT)
    End If    
   	
   	If SetConditionData = False Then Exit Sub
        
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
	   MsgDisplayFlag = True
        ' Modify Focus Events    
        %>
            <Script language=vbs>
				Call parent.SetFocusToDocument("M")	
				parent.frm1.txtSoDtFrom.Focus
            </Script>
        <%	   
'       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

%>
<Script Language=vbscript>
    
    With parent
		
		 .frm1.txtSalesGrpNm.value	  = "<%=ConvSPChars(strSalesGrpNm)%>"
		 .frm1.txtItemCodeNm.value	  = "<%=ConvSPChars(strItemCodeNm)%>"
		 .frm1.txtSoldToPartyNm.value = "<%=ConvSPChars(strSoldToPartyNm)%>"
		 .frm1.txtSoTypeNm.value	  = "<%=ConvSPChars(strSoTypeNm)%>"
			
    	 
     If "<%=lgDataExist%>" = "Yes" Then
        'Set condition data to hidden area
        
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.HSalesGrp.value	 = "<%=ConvSPChars(txtSalesGrp)%>"
			.frm1.HSoDtFrom.value	 = "<%=txtSoDtFrom%>"
			.frm1.HSoDtTo.value		 = "<%=txtSoDtTo%>"
			.frm1.HDlvyDtFrom.value	 = "<%=ConvSPChars(txtDlvyDtFrom)%>"
			.frm1.HDlvyDtTo.value	 = "<%=ConvSPChars(txtDlvyDtTo)%>"
			.frm1.HSoldToParty.value = "<%=ConvSPChars(txtSoldToParty)%>"
			.frm1.HItemCode.value	 = "<%=ConvSPChars(txtItemCode)%>"
			.frm1.HSoType.value		 = "<%=ConvSPChars(txtSoType)%>"
			.frm1.HtxtTrackingNo.value = "<%=ConvSPChars(txtTrackingNo)%>"    
		End If
		
		.ggoSpread.Source  = .frm1.vspdData
                
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"

		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")		
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",4),"A", "Q" ,"X","X")		
				
		.lgPageNo	  	   =  "<%=lgPageNo%>"  				  '��: Next next data tag
       	.DbQueryOk
       	
       	.frm1.vspdData.Redraw = True
    End If
       
	End with	
</Script>	

