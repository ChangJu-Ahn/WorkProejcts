<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5211qb2
'*  4. Program Name         : B/L����ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jin-hyun Shin
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
<%

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
Dim strIncotermsCd												'�������� 
Dim strIncotermsCdFrom				
Dim strIncotermsLookUp
Dim strPurGrpCd												'	���ű׷� 
Dim strPurGrpCdFrom 										
Dim strBpCd													'   ������ 
Dim strBpCdFrom
Dim strBlFrDt                                               '   Bl������ 
Dim strBlToDt
Dim strLoadingFrDt                                          '   ������ 
Dim strLoadingToDt
Dim strCfmFlg								                '   Ȯ������ 
Dim strCfmFlgFrom	
Dim strItemCd								                '   ǰ�� 
Dim strItemCdFrom	
Dim strPlantCd								                '   ���� 
Dim strPlantCdFrom	
Dim strBlNo
Dim StrBlNoFrom
Dim lgPageNo
Dim lgDataExist

Dim arrRsVal(11)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
Dim iFrPoint
iFrPoint=0

' === 2005.07.13 Tracker No. 9899 ====================================================
'	Const Major_Cd_Incoterms = "B9006"
' === 2005.07.13 Tracker No. 9899 ====================================================

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

    Call  TrimData()                                                     '�� : Parent�� ������ ����Ÿ ���� 
    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

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
    Dim strVal
    Redim UNISqlId(7)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(6,13)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "m5211qa2_KO441"
     UNISqlId(1) = "M3111QA102"								              '����ó�� 
	 UNISqlId(2) = "M2111QA302"											  '���� 
     UNISqlId(3) = "M2111QA303"								              'ǰ�� 
     UNISqlId(4) = "M3111QA104"								              '���ű׷�� 
	 UNISqlId(5) = "S0000QA000"											  '�������� 
	 UNISqlId(6) = "M5211QB2_KO441"											  '�ͼ��� 

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

     UNIValue(0,1)  = strBlFrDt
     UNIValue(0,2)  = strBlToDt
     UNIValue(0,3)  = UCase(Trim(strBpCdFrom))			
     UNIValue(0,4)  = strLoadingFrDt
     UNIValue(0,5)  = strLoadingToDt
	 UNIValue(0,6)  = UCase(Trim(strIncotermsCdFrom))	
	 UNIValue(0,7)  = UCase(Trim(strPurGrpCdFrom))	    
     UNIValue(0,8)  = UCase(Trim(strCfmFlgFrom))
	 UNIValue(0,9)  = UCase(Trim(strPlantCdFrom))	    
     UNIValue(0,10) = UCase(Trim(strItemCdFrom))
     UNIValue(0,11) = UCase(Trim(strBlNoFrom))

     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND B.BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND B.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND B.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gPlant") <> "" Then
        strVal = strVal & " AND A.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If

     UNIValue(0,12) = strVal

     
     UNIValue(1,0)  = UCase(Trim(strBpCd))
     UNIValue(2,0)  = UCase(Trim(strPlantCd))  
     UNIValue(3,0)  = UCase(Trim(strPlantCd))
     UNIValue(3,1)  = UCase(Trim(strItemCd))
     UNIValue(4,0)  = UCase(Trim(strPurGrpCd))  
' === 2005.07.13 Tracker No. 9899 ====================================================
'	 UNIValue(5,0)  = UCase(Trim(Major_Cd_Incoterms))
	 UNIValue(5,0)	= FilterVar("B9006", "''", "S")
' === 2005.07.13 Tracker No. 9899 ====================================================
     UNIValue(5,1)  = UCase(Trim(strIncotermsLookUp))
     
     UNIValue(6,0)  = strBlFrDt
     UNIValue(6,1)  = strBlToDt
     UNIValue(6,2)  = UCase(Trim(strBpCdFrom))			
     UNIValue(6,3)  = strLoadingFrDt
     UNIValue(6,4)  = strLoadingToDt
	 UNIValue(6,5)  = UCase(Trim(strIncotermsCdFrom))	
	 UNIValue(6,6)  = UCase(Trim(strPurGrpCdFrom))	    
     UNIValue(6,7)  = UCase(Trim(strCfmFlgFrom))
	 UNIValue(6,8)  = UCase(Trim(strPlantCdFrom))	    
     UNIValue(6,9) =  UCase(Trim(strItemCdFrom))
     UNIValue(6,10) = UCase(Trim(strBlNoFrom))

     UNIValue(6,11) = strVal
     
     UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 

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
        If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs1(0)
		arrRsVal(5) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
	
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs2(0)
		arrRsVal(1) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
	
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
	       Set rs0 = Nothing
	       Exit Sub
'		   Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
'	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(6) = rs3(0)
		arrRsVal(7) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs4(0)
		arrRsVal(3) = rs4(1)
        rs4.Close
        Set rs2 = Nothing
    End If
    
    If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtIncotermsCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If

    If  rs6.EOF And rs6.BOF Then
        rs6.Close
        Set rs6 = Nothing
    Else    
		arrRsVal(10) = rs6(0)
        rs6.Close
        Set rs6 = Nothing
    End If

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()

    '---ǰ�� 
    If Len(Trim(Request("txtItemCd"))) Then
    	strItemCd	= " " & FilterVar(Request("txtItemCd"), "''", "S") & ""
    	strItemCdFrom = strItemCd
    Else
    	strItemCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strItemCdFrom = "" & FilterVar("%%", "''", "S") & ""
    End If
     '---���� 
    If Len(Trim(Request("txtPlantCd"))) Then
    	strPlantCd	= " " & FilterVar(Request("txtPlantCd"), "''", "S") & ""
    	strPlantCdFrom = strPlantCd
    Else
    	strPlantCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPlantCdFrom = "" & FilterVar("%%", "''", "S") & ""
    End If
     '---�������� 
    If Len(Trim(Request("txtIncotermsCd"))) Then
    	strIncotermsCd	= " " & FilterVar(Request("txtIncotermsCd"), "''", "S") & ""
    	strIncotermsCdFrom = strIncotermsCd
    Else
    	strIncotermsCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strIncotermsCdFrom = "" & FilterVar("%%", "''", "S") & ""
    End If
    If Len(Trim(Request("txtIncotermsCd"))) Then
    	strIncotermsCd	= " " & FilterVar(Trim(Request("txtIncotermsCd")), "''", "S") & ""
    Else
    	strIncotermsCd	= " " & FilterVar("zzzzzzzzz", "''", "S") & ""
    End If
	strIncotermsLookUp = strIncotermsCd
     '---���ű׷� 
    If Len(Trim(Request("txtPurGrpCd"))) Then
    	strPurGrpCd	= " " & FilterVar(Request("txtPurGrpCd"), "''", "S") & ""
    	strPurGrpCdFrom = strPurGrpCd
    Else
    	strPurGrpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPurGrpCdFrom = "" & FilterVar("%%", "''", "S") & ""
    End If
     '---������ 
    If Len(Trim(Request("txtBpCd"))) Then
    	strBpCd	= " " & FilterVar(Request("txtBpCd"), "''", "S") & ""
    	strBpCdFrom = strBpCd
    Else
    	strBpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strBpCdFrom = "" & FilterVar("%%", "''", "S") & ""    	
    End If

     '---bl������ 
    If Len(Trim(Request("txtBlFrDt"))) Then
    	strBlFrDt 	= " " & FilterVar(UNIConvDate(Trim(Request("txtBlFrDt"))), "''", "S") & ""
    Else
    	strBlFrDt	= "" & FilterVar("1900-01-01", "''", "S") & ""
    End If
    If Len(Trim(Request("txtBlToDt"))) Then
    	strBlToDt 	= " " & FilterVar(UNIConvDate(Trim(Request("txtBlToDt"))), "''", "S") & ""
    Else
    	strBlToDt	= "" & FilterVar("2999-12-30", "''", "S") & ""
    End If  
     '---������ 
    If Len(Trim(Request("txtLoadingFrDt"))) Then
    	strLoadingFrDt 	= " " & FilterVar(UNIConvDate(Trim(Request("txtLoadingFrDt"))), "''", "S") & ""
    Else
    	strLoadingFrDt	= "" & FilterVar("1900-01-01", "''", "S") & ""
    End If
	
    If Len(Trim(Request("txtLoadingToDt"))) Then
    	strLoadingToDt 	= " " & FilterVar(UNIConvDate(Trim(Request("txtLoadingToDt"))), "''", "S") & ""
    Else
    	strLoadingToDt	= "" & FilterVar("2999-12-30", "''", "S") & ""
    End If       
     '---Ȯ������ 
    If Len(Trim(Request("txtCfmFlg"))) Then
    	strCfmFlg	= " " & FilterVar(Request("txtCfmFlg"), "''", "S") & ""
    	strCfmFlgFrom = strCfmFlg
    Else
    	strCfmFlg	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strCfmFlgFrom = "" & FilterVar("%%", "''", "S") & ""
    End If
     '---B/L No
    If Len(Trim(Request("txtBlNo"))) Then
    	strBlNo	= " " & FilterVar(Request("txtBlNo"), "''", "S") & ""
    	strBlNoFrom = strBlNo
    Else
    	strBlNo	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strBlNoFrom = "" & FilterVar("%%", "''", "S") & ""
    End If


End Sub

%>

<Script Language=vbscript>
    
    With Parent
        .ggoSpread.Source  = .frm1.vspdData
        Parent.frm1.vspdData.Redraw = False
        .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '�� : Display data
        
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",6),"C","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",7),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",9),"D","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.parent.gCurrency,.GetKeyPos("A",10),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",19),"D","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",20),"D","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",22),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",23),"A","Q","X","X")
        
         .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag

		.frm1.hdnBeneficiaryCd.value	= "<%=ConvSPChars(Request("txtBpCd"))%>"
        .frm1.hdnIncotermsCd.value		= "<%=ConvSPChars(Request("txtIncotermsCd"))%>"
        .frm1.hdnPurGrpCd.value			= "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
		.frm1.hdnBlIssueFrDt.value		= "<%=Request("txtBlFrDt")%>"
        .frm1.hdnBlIssueToDt.value		= "<%=Request("txtBlToDt")%>"
        .frm1.hdnLoadingFrDt.value		= "<%=ConvSPChars(Request("txtLoadingFrDt"))%>"
        .frm1.hdnLoadingToDt.value		= "<%=ConvSPChars(Request("txtLoadingToDt"))%>"
		.frm1.hdnstrCfmFlg.value		= "<%=ConvSPChars(Request("txtCfmFlg"))%>"
		.frm1.hdnItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
        .frm1.hdnPlantCd.value			= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnBlNo.value				= "<%=ConvSPChars(Request("txtBlNo"))%>"
		
		.frm1.txtBeneficiaryNm.value	=  "<%=ConvSPChars(arrRsVal(5))%>" 	
		.frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>"
		.frm1.txtitemNm.value			=  "<%=ConvSPChars(arrRsVal(7))%>"
		.frm1.txtPurGrpNm.value			=  "<%=ConvSPChars(arrRsVal(3))%>"
		.frm1.txtIncotermsNm.value		=  "<%=ConvSPChars(arrRsVal(9))%>"
		.frm1.txtTotQty.Text			=  "<%=UNINumClientFormat(arrRsVal(10), ggQty.DecPoint, 0)%>"

		.DbQueryOk
		Parent.frm1.vspdData.Redraw = True
	End with

</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
