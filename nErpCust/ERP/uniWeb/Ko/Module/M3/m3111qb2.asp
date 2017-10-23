<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111qa2
'*  4. Program Name         : ���ֻ���ȸ 
'*  5. Program Desc         : ���ֻ���ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/12
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : ByunJiHyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgStrPrevKey                                            '�� : ���� �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim ICount  		                                        '   Count for column index
Dim arrRsVal(11)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
Dim iFrPoint
iFrPoint=0
Dim lgPageNo
Dim lgDataExist
'--------------- ������ coding part(��������,End)------------------------------------------------------

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "QB")

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

     Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
     call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	lgstrData  = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()
	Dim strSQL
    Redim UNISqlId(6)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(6,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "m3111qa201"
     UNISqlId(1) = "M2111QA302"								              '����� 
     UNISqlId(2) = "M3111QA104"								              '���ű׷�� 
     UNISqlId(3) = "M3111QA102"								              '����ó�� 
	 UNISqlId(4) = "M2111QA303"											  'ǰ��� 
	 UNISqlId(5) = "M3111QA103"											  '�������¸�	 
																		  'Reusage is Recommended
'--------------- ������ coding part(�������,End)------------------------------------------------------
	strSQL = ""
     '---���� 
    If Len(Trim(Request("txtPlantCd"))) Then
		strSQL = strSQL & " AND B.PLANT_CD =  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
    End If
     '---���ű׷� 
    If Len(Trim(Request("txtPurGrpCd"))) Then
		strSQL = strSQL & " AND A.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtPurGrpCd"))), " " , "S") & " "
    End If
     '---����ó 
    If Len(Trim(Request("txtBpCd"))) Then
		strSQL = strSQL & " AND A.BP_CD =  " & FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "S") & " "
    End If
     '---������ 
    If Len(Trim(Request("txtPoFrDt"))) Then
		strSQL = strSQL & " AND A.PO_DT >=  " & FilterVar(uniConvDate(Request("txtPoFrDt")), "''", "S") & ""
    Else
		strSQL = strSQL & " AND A.PO_DT >= " & FilterVar("1900/01/01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtPoToDt"))) Then
		strSQL = strSQL & " AND A.PO_DT <=  " & FilterVar(uniConvDate(Request("txtPoToDt")), "''", "S") & ""
    Else
		strSQL = strSQL & " AND A.PO_DT <= " & FilterVar("2999/12/30", "''", "S") & ""
    End If    
    '---ǰ�� 
    If Len(Trim(Request("txtItemCd"))) Then
		strSQL = strSQL & " AND B.ITEM_CD =  " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
    End If

    '---Tracking No
    If Len(Trim(Request("txtTrackNo"))) Then
		strSQL = strSQL & " AND B.TRACKING_NO =  " & FilterVar(Trim(UCase(Request("txtTrackNo"))), " " , "S") & " "
    End If
     '---�������� 
    If Len(Trim(Request("txtPoType"))) Then
		strSQL = strSQL & " AND A.PO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtPoType"))), " " , "S") & " "
    End If
     '---�ܰ����� 
    If Len(Trim(Request("txtPrcFlg"))) Then
		strSQL = strSQL & " AND B.PO_PRC_FLG =  " & FilterVar(Trim(UCase(Request("txtPrcFlg"))), " " , "S") & " "
    End If
    '--�������� 
    If Len(Trim(Request("rdoClsFlg"))) Then
		strSQL = strSQL & " AND B.CLS_FLG =  " & FilterVar(Trim(UCase(Request("rdoClsFlg"))), " " , "S") & " "
    End If
    

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
	 UNIValue(0,1)  = strSQL
     
     UNIValue(1,0)  = " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
     UNIValue(2,0)  = " " & FilterVar(Trim(UCase(Request("txtPurGrpCd"))), " " , "S") & " "
     UNIValue(3,0)  = " " & FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "S") & " "
     UNIValue(4,0)  = " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
     UNIValue(4,1)  = " " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
     UNIValue(5,0)  = " " & FilterVar(Trim(UCase(Request("txtPoType"))), " " , "S") & " "
     
'--------------- ������ coding part(�������,End)----------------------------------------------------
     UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)			
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 
    
    '============================= �߰��� �κ� =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
	       Set rs0 = Nothing
	       Exit Sub
		End If
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
    
    If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtPoType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
    

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         Parent.frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=lgstrData%>", "F"                  '�� : Display data
                  
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",23), Parent.GetKeyPos("A",24),"C", "Q" ,"X","X")	'���ֱݾ� 
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",23), Parent.GetKeyPos("A",25),"A", "Q" ,"X","X")	'���ֱݾ� 
         Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows, Parent.Parent.gCurrency , Parent.GetKeyPos("A",26),"A", "Q" ,"X","X")	'�����ڱ��ݾ�				'�����ڱ��ݾ� 
         
         .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
		 
		 .frm1.hdnPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.hdnPurGrpCd.value	= "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
         .frm1.hdnBpCd.value		= "<%=ConvSPChars(Request("txtBpCd"))%>"
		 .frm1.hdnPoFrDt.value	= "<%=Request("txtPoFrDt")%>"
         .frm1.hdnPoToDt.value	= "<%=Request("txtPoToDt")%>"
         .frm1.hdnItemCd.value	= "<%=ConvSPChars(Request("txtItemCd"))%>"
         .frm1.hdnTrackNo.value	    = "<%=ConvSPChars(Request("txtTrackNo"))%>"
		 .frm1.hdnPoType.value	    = "<%=ConvSPChars(Request("txtPoType"))%>"
		 .frm1.hdncboPrcFlg.value	    = "<%=ConvSPChars(Request("txtPrcFlg"))%>"
		 
		 .frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 
  		 .frm1.txtPurGrpNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		 
  		 .frm1.txtBpNm.value			=  "<%=ConvSPChars(arrRsVal(5))%>" 	
  		 
  		 .frm1.txtItemNm.value			=  "<%=ConvSPChars(arrRsVal(7))%>" 	
  		 
  		 .frm1.txtPoTypeNm.value		=  "<%=ConvSPChars(arrRsVal(9))%>"
         .DbQueryOk
         Parent.frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
