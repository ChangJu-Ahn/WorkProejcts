<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5114qb3
'*  4. Program Name         : �԰��������ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2003-05-20
'*  8. Modifier (First)     : Kim Jin Ha
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : 
'=======================================================================================================
Option Explicit
%>

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
Dim iTotstrData

Dim ICount  		                                        '   Count for column index

Dim lgPageNo
Dim lgDataExist

Dim arrRsVal(4)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
Dim iFrPoint
iFrPoint=0

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()

	Dim iLoopCount                                                                     
	Dim iRowStr
	Dim ColCnt
	Dim PvArr
	Const C_SHEETMAXROWS_D = 100 
	    
	lgDataExist    = "Yes"
	lgstrData      = ""
	  
	If CLng(lgPageNo) > 0 Then
	   rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	   iFrPoint     = C_SHEETMAXROWS_D * CLng(lgPageNo) 
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
	
	iTotstrData = Join(PvArr, "")

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
	Dim iStrSql
	Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(5,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	UNISqlId(0) = "M5114QA301"
    UNISqlId(1) = "M4111QA503"								              '�԰����� 
    UNISqlId(2) = "M5111QA103"								              '�������� 
    UNISqlId(3) = "M3111QA102"								              '����ó�� 
	UNISqlId(4) = "M2111QA302"											  '���� 
    UNISqlId(5) = "M2111QA303"								              'ǰ�� 
    
    UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
    
    iStrSql = "" 
    '---�԰����� 
    If Len(Trim(Request("txtMvmtType"))) Then
	    iStrSql = iStrSql & " AND A.IO_TYPE_CD =  " & FilterVar(UCase(Request("txtMvmtType")), "''", "S") & "  "
	End IF
	
	'---�������� 
    If Len(Trim(Request("txtIvType"))) Then
	    iStrSql = iStrSql & " AND F.IV_TYPE_CD =  " & FilterVar(UCase(Request("txtIvType")), "''", "S") & " "
	End IF
	
	'---����ó 
    If Len(Trim(Request("txtBpCd"))) Then
	    iStrSql = iStrSql & " AND A.BP_CD =  " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
	End IF
	
	'---�԰���(From)
    If Len(Trim(Request("txtMvFrDt"))) Then
	    iStrSql = iStrSql & " AND A.MVMT_DT >=  " & FilterVar(UNIConvDate(Request("txtMvFrDt")), "''", "S") & ""
	Else
		iStrSql = iStrSql & " AND A.MVMT_DT >=  " & FilterVar(UNIConvDate("1900-01-01"), "''", "S") & ""
	End IF
	
	'---�԰���(To)
    If Len(Trim(Request("txtMvToDt"))) Then
	    iStrSql = iStrSql & " AND A.MVMT_DT <=  " & FilterVar(UNIConvDate(Request("txtMvToDt")), "''", "S") & ""
	Else
		iStrSql = iStrSql & " AND A.MVMT_DT <=  " & FilterVar(UNIConvDate("2999-12-30"), "''", "S") & ""
	End IF
	
	'---���ֹ�ȣ 
    If Len(Trim(Request("txtPoNo"))) Then
	    iStrSql = iStrSql & " AND A.PO_NO >=  " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & " "
	End IF
	
	'---�԰��ȣ 
    If Len(Trim(Request("txtMvmtNo"))) Then
	    iStrSql = iStrSql & " AND A.MVMT_RCPT_NO >=  " & FilterVar(UCase(Request("txtMvmtNo")), "''", "S") & " "
	End IF
	
	'---���� 
    If Len(Trim(Request("txtPlantCd"))) Then
	    iStrSql = iStrSql & " AND A.PLANT_CD  =  " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	End IF
	
	'---ǰ�� 
    If Len(Trim(Request("txtItemCd"))) Then
	    iStrSql = iStrSql & " AND A.ITEM_CD =  " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
	End IF
	
	'---����Ȯ������ 
    If Len(Trim(Request("txtCfmFlg"))) Then
	    iStrSql = iStrSql & " AND F.POSTED_FLG =  " & FilterVar(UCase(Request("txtCfmFlg")), "''", "S") & " "
	End IF

     If Request("gPlant") <> "" Then
        iStrSql = iStrSql & " AND a.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        iStrSql = iStrSql & " AND a.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        iStrSql = iStrSql & " AND a.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        iStrSql = iStrSql & " AND a.MVMT_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   
	
	
	UNIValue(0,1)  = iStrSql
    UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 
    
    UNIValue(1,0)  = " " & FilterVar(UCase(Request("txtMvmtType")), "''", "S") & " "
    UNIValue(2,0)  = " " & FilterVar(UCase(Request("txtIvType")), "''", "S") & " "
    UNIValue(3,0)  = " " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    UNIValue(4,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(5,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(5,1)  = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)			
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 
	
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtMvmtType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�԰�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtIvType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(1) = rs2(1)
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
		arrRsVal(2) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(3) = rs4(1)
        rs4.Close
        Set rs2 = Nothing
    End If
    
    If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs5(1)
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
        .frm1.vspdData.Redraw = False
        .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '�� : Display data
        
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",11),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",12),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",13),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",14),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",15),"A","Q","X","X")
		
         .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag

		.frm1.hdnMvmtType.value		= "<%=ConvSPChars(Request("txtMvmtType"))%>"
        .frm1.hdnIvType.value		= "<%=ConvSPChars(Request("txtIvType"))%>"
        .frm1.hdnBpCd.value			= "<%=ConvSPChars(Request("txtBpCd"))%>"
		.frm1.hdnMvFrDt.value		= "<%=Request("txtMvFrDt")%>"
        .frm1.hdnMvToDt.value		= "<%=Request("txtMvToDt")%>"
        .frm1.hdnPoNo.value			= "<%=ConvSPChars(Request("txtPoNo"))%>"
        .frm1.hdnMvmtNo.value	    = "<%=ConvSPChars(Request("txtMvmtNo"))%>"
		.frm1.hdnPlantCd.value	    = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
        .frm1.hdnstrCfmFlg.value	   = "<%=ConvSPChars(Request("txtCfmFlg"))%>"
		
		.frm1.txtMvmtTypeNm.value	=  "<%=ConvSPChars(arrRsVal(0))%>" 	
		.frm1.txtIvTypeNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 	
		.frm1.txtBpNm.value			=  "<%=ConvSPChars(arrRsVal(2))%>" 	
		.frm1.txtPlantNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
		.frm1.txtItemNm.value		=  "<%=ConvSPChars(arrRsVal(4))%>" 	
		
		.DbQueryOk
		.frm1.vspdData.Redraw = True
	End with

</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
