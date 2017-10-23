<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S4111MA8
'*  4. Program Name         : ������Ȳ��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41118ListDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : RYU KYUNG RAE(1)
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'*                            -2002/04/11 : ADO ��ȯ 
'**********************************************************************************************

'								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
'								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next
'============================================  2002-04-10 ����  =============================================
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4
															'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
DIM strMovType		'�������� 
DIM strSoNo			'���ֹ�ȣ 
DIM strBPartner		'��ǰó 
DIM strTranMeth		'��۹�� 
DIM strPostFlag		'����� 
Dim strSalesGrp		'�����׷� 
DIM strReqFromDate
DIM strReqToxxDate
Dim arrRsVal(7)												'�� : QueryData()����� ���ڵ���� �迭�� ������ ��� 
															'�� : ���� ���ڵ���� ������ŭ �迭 ũ�� ����			
	MsgDisplayFlag = False															
'--------------- ������ coding part(��������,End)----------------------------------------------------------
   
	Call LoadBasisGlobalInf()
	Call HideStatusWnd
	
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100						                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'/////////////////////////////////////////////////////////////////////////////////////////////
'Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
'Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
'Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
'Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
'Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
'Dim lgStrPrevKey                                            '�� : ���� �� 
'Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
'Dim lgTailList
'Dim lgSelectList
'Dim lgSelectListDT
'============================================  2002-04-10 ��  ===============================================

' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
	
'	On Error Resume Next
 
	SetConditionData = False
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strMovType =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtDn_Type")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtDn_Type.focus    
                </Script>
            <%        		    	
		End If
	End If   	    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strBPartner =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtShip_to_party")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "��ǰó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtShip_to_party.focus    
                </Script>
            <%        		    			    	
		End If			
    End If   	

    If Not(rs3.EOF Or rs3.BOF) Then
        strTranMeth =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtTrans_meth")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "��۹��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtTrans_meth.focus    
                </Script>
            <%        		    			    	
		End If				
    End If
    
    If Not(rs4.EOF Or rs4.BOF) Then
        strSalesGrp =  rs4(1)
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtSalesGrp")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    MsgDisplayFlag = True	
		End If				
    End If

	
	If MsgDisplayFlag = True Then Exit Function

	SetConditionData = True

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	Dim arrVal(3)
	Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------	
	Redim UNIValue(4,2)

	UNISqlId(0) = "S4111MA801"
	UNISqlId(1) = "s0000qa000"    ' ��������    'I0001'
	UNISqlId(2) = "s0000qa002"    ' ��ǰó 
	UNISqlId(3) = "s0000qa000"    ' ��۹��	'B9009'
	UNISqlId(4) = "s0000qa005"    ' �����׷� 
	
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strVal = " "
		
    '---�������� 
    If Len(Trim(Request("txtDn_Type"))) Then
    	strVal = strval & " AND B.MOV_TYPE =  " & FilterVar(Request("txtDn_Type"), "''", "S") & " "    			
    End If
    arrVal(0) = FilterVar(Trim(Request("txtDn_Type")), " " , "S")
	
	'---��ǰó 
	If Len(Trim(Request("txtShip_to_party"))) Then
    	strVal = strval & " AND B.SHIP_TO_PARTY =  " & FilterVar(Request("txtShip_to_party"), "''", "S") & " "    	
    End If
    arrVal(1) = FilterVar(Trim(Request("txtShip_to_party")), " " , "S")
    
	'---���ֹ�ȣ 
	If Len(Trim(Request("txtSo_no"))) Then
    	strVal = strval & " AND B.SO_NO =  " & FilterVar(Request("txtSo_no"), "''", "S") & " "
    End If    
    
    '---��۹�� 
	If Len(Trim(Request("txtTrans_meth"))) Then
    	strVal = strval & " AND B.TRANS_METH =  " & FilterVar(Request("txtTrans_meth"), "''", "S") & " "    	
    End If
    arrVal(2) = FilterVar(Trim(Request("txtTrans_meth")), " " , "S")

	'---�����׷� 
	If Len(Trim(Request("txtSalesGrp"))) Then
    	strVal = strval & " AND B.SALES_GRP =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & " "    	
    End If
    arrVal(3) = FilterVar(Trim(Request("txtSalesGrp")), "''", "S")
    
    '---����� 
	If Len(Trim(Request("txtPostGiFlag"))) Then
    	strVal = strval & " AND B.POST_FLAG =  " & FilterVar(Request("txtPostGiFlag"), "''", "S") & ""
    End If

     '---����û�� 
    If Len(Trim(Request("txtReqGiDtFrom"))) Then
    	strVal = strval & " AND B.PROMISE_DT >=  " & FilterVar(uniConvDate(Trim(Request("txtReqGiDtFrom"))), "''", "S") & ""
    End If
    
    If Len(Trim(Request("txtReqGiDtTo"))) Then
    	strVal = strval & " AND B.PROMISE_DT <=  " & FilterVar(uniConvDate(Trim(Request("txtReqGiDtTo"))), "''", "S") & ""
    End If
		   
    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar("I0001", "''", "S")    					'�������� 
    UNIValue(1,1) = arrVal(0)					'�������� 

    UNIValue(2,0) = arrVal(1)				    '��ǰó 
	UNIValue(3,0) = FilterVar("B9009", "''", "S") 					'��۹�� 
	UNIValue(3,1) = arrVal(2)
	UNIValue(4,0) = arrVal(3)					'�����׷�	

	'--------------- ������ coding part(�������,End)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
' Query Data
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
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
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
            Parent.frm1.txtDn_Type.focus    
            </Script>
        <%
	Else    
		Call  MakeSpreadSheetData()
	End If

'	Call  SetConditionData()
    
End Sub

%>
<Script Language=vbscript>

With Parent 	

	.frm1.txtDn_TypeNm.value			= "<%=ConvSPChars(strMovType)%>"
	.frm1.txtShip_to_partyNm.value		= "<%=ConvSPChars(strBPartner)%>"
	.frm1.txtTrans_meth_nm.value		= "<%=ConvSPChars(strTranMeth)%>"
	.frm1.txtSalesGrpNm.value			= "<%=ConvSPChars(strSalesGrp)%>"

	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txtHDn_Type.value			= "<%=ConvSPChars(Request("txtDn_Type"))%>"
			.frm1.txtHSo_no.value			= "<%=ConvSPChars(Request("txtSo_no"))%>"
			.frm1.txtHShip_to_party.value	= "<%=ConvSPChars(Request("txtShip_to_party"))%>"
			.frm1.txtHReqGiDtFrom.value		= "<%=ConvSPChars(Request("txtReqGiDtFrom"))%>"
			.frm1.txtHReqGiDtTo.value		= "<%=ConvSPChars(Request("txtReqGiDtTo"))%>"
			.frm1.txtHTrans_meth.value		= "<%=ConvSPChars(Request("txtTrans_meth"))%>"
			.frm1.txtHPostGiFlag.value		= "<%=ConvSPChars(Request("txtPostGiFlag"))%>"	    
			.frm1.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"	    
		End If

       'Show multi spreadsheet data from this line
       .frm1.vspdData.Redraw = False
       .ggoSpread.Source  = .frm1.vspdData
       .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"          '�� : Display data
       .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       .frm1.vspdData.Redraw = True
       .DbQueryOk
    
    End If  
    
End With     
</Script>
