<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/12/12
'*  7. Modified date(Last)  : 2003/06/05
'*  8. Modifier (First)     : ByunJiHyun
'*  9. Modifier (Last)      : Lee Eun hee
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
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

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim iTotstrData

Dim lgDataExist
Dim lgPageNo

Dim ICount  		                                        '   Count for column index
Dim strBizArea												'	����� 
Dim strBizAreaFrom				
Dim strItemCd                                               '   ǰ�� 
Dim strItemCdFrom
Dim strBpCd												    '   �ŷ�ó 
Dim strBpCdFrom
Dim strIvFrDt                                               '   �����ۼ��� 
Dim strIvToDt
Dim strPlantCd												'	���� 
Dim strPlantCdFrom				
Dim strIvType                                               '   �������� 
Dim strIvTypeFrom	
Dim strPurGrpCd												'	���ű׷� 
Dim strPurGrpCdFrom 										
Dim strPstFlg								                '   ȸ��ó������ 
Dim strPstFlgFrom	
Dim arrRsVal(12)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array	

Dim iPrevEndRow
Dim iEndRow


     Call HideStatusWnd 
     lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	 iPrevEndRow = 0
     iEndRow = 0
     
     Call  TrimData()                                                     '�� : Parent�� ������ ����Ÿ ���� 
     Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
     call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100 
        
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 

    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(7)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(7,17)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "M5111QA201"
     
     UNISqlId(1) = "M5111QA102"								              '������ 
     UNISqlId(2) = "M2111QA303"											  'ǰ��� 
     UNISqlId(3) = "M3111QA102"								              '����ó�� 
     UNISqlId(4) = "M2111QA302"								              '����� 
     UNISqlId(5) = "M5111QA103"								              '�������¸�     
     UNISqlId(6) = "M3111QA104"								              '���ű׷��     	
																		  'Reusage is Recommended
     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	 UNIValue(0,1)  = UCase(Trim(strBizAreaFrom))		'---����� 
	 UNIValue(0,2)  = UCase(Trim(strBizArea))
	 UNIValue(0,3)  = UCase(Trim(strItemCdFrom))		'---ǰ�� 
     UNIValue(0,4)	= UCase(Trim(strItemCd))
     UNIValue(0,5)  = UCase(Trim(strBpCdFrom))			'---�ŷ�ó 
     UNIValue(0,6)  = UCase(Trim(strBpCd))
     UNIValue(0,7)  = UCase(Trim(strIvFrDt))			'---�����ۼ��� 
     UNIValue(0,8)  = UCase(Trim(strIvToDt)) 
	 UNIValue(0,9)  = UCase(Trim(strPlantCdFrom))		'---���� 
	 UNIValue(0,10) = UCase(Trim(strPlantCd))	 
     UNIValue(0,11) = UCase(Trim(strIvTypeFrom))		'---�������� 
     UNIValue(0,12) = UCase(Trim(strIvType))
     UNIValue(0,13) = UCase(Trim(strPurGrpCdFrom))	    '---���ű׷� 
     UNIValue(0,14) = UCase(Trim(strPurGrpCd))
     UNIValue(0,15) = UCase(Trim(strPstFlgFrom))		'---�ܰ����� 
     UNIValue(0,16) = UCase(Trim(strPstFlg))
     
     
     UNIValue(1,0)  = UCase(Trim(strBizArea))
     UNIValue(2,0)  = UCase(Trim(strPlantCd)) 
     UNIValue(2,1)  = UCase(Trim(strItemCd)) 
     UNIValue(3,0)  = UCase(Trim(strBpCd))      
     UNIValue(4,0)  = UCase(Trim(strPlantCd))
     UNIValue(5,0)  = UCase(Trim(strIvType))
     UNIValue(6,0)  = UCase(Trim(strPurGrpCd))
     
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
        
    '============================= �߰��� �κ� =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtBizArea")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
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
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
	       Set rs0 = Nothing
	       Exit Sub
'		   Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
'	       FalsechkFlg = True	
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
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
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
        If Len(Request("txtIvType")) And FalsechkFlg = False Then
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
        If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(10) = rs6(0)
		arrRsVal(11) = rs6(1)
        rs6.Close
        Set rs6 = Nothing
    End If
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!

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

     '---����� 
    If Len(Trim(Request("txtBizArea"))) Then
    	strBizArea	= " " & FilterVar(UCase(Request("txtBizArea")), "''", "S") & " "
    	strBizAreaFrom = strBizArea
    Else
    	strBizArea	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strBizAreaFrom = "''"
    End If
    '---ǰ�� 
    If Len(Trim(Request("txtItemCd"))) Then
    	strItemCd	= " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    	strItemCdFrom = strItemCd
    Else
    	strItemCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strItemCdFrom = "''"
    End If
    '---�ŷ�ó 
    If Len(Trim(Request("txtBpCd"))) Then
    	strBpCd	= " " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    	strBpCdFrom = strBpCd
    Else
    	strBpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strBpCdFrom = "''"    	
    End If
     '---�����ۼ��� 
    If Len(Trim(Request("txtIvFrDt"))) Then
    	strIvFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
    	strIvFrDt 	= " " & FilterVar(UNIConvDate(Trim(Request("txtIvFrDt"))), "''", "S") & ""
    Else
    	strIvFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtIvToDt"))) Then
    	strIvToDt 	= " " & FilterVar(UNIConvDate(Trim(Request("txtIvToDt"))), "''", "S") & ""
    Else
    	strIvToDt	= "" & FilterVar("2999/12/30", "''", "S") & ""
    End If        
     '---���� 
    If Len(Trim(Request("txtPlantCd"))) Then
    	strPlantCd	= " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    	strPlantCdFrom = strPlantCd
    Else
    	strPlantCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPlantCdFrom = "''"
    End If
    '---�������� 
    If Len(Trim(Request("txtIvType"))) Then
    	strIvType	= " " & FilterVar(UCase(Request("txtIvType")), "''", "S") & " "
    	strIvTypeFrom = strIvType
    Else
    	strIvType	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strIvTypeFrom = "''"
    End If
     '---���ű׷� 
    If Len(Trim(Request("txtPurGrpCd"))) Then
    	strPurGrpCd	= " " & FilterVar(UCase(Request("txtPurGrpCd")), "''", "S") & " "
    	strPurGrpCdFrom = strPurGrpCd
    Else
    	strPurGrpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPurGrpCdFrom = "''"
    End If
     '---ȸ��ó������ 
    If Len(Trim(Request("txtPstFlg"))) Then
    	strPstFlg	= " " & FilterVar(UCase(Request("txtPstFlg")), "''", "S") & " "
    	strPstFlgFrom = strPstFlg
    Else
    	strPstFlg	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPstFlgFrom = "''"
    End If

End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         Parent.frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>", "F"                  '�� : Display data
         
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,Parent.GetKeyPos("A",12), Parent.GetKeyPos("A",11),"C", "Q" ,"X","X")	'���Դܰ� 
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,Parent.GetKeyPos("A",12), Parent.GetKeyPos("A",13),"A", "Q" ,"X","X")	'���Աݾ� 
         Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>, Parent.Parent.gCurrency , Parent.GetKeyPos("A",14),"A", "Q" ,"X","X")					'�����ڱ��ݾ� 
         
          .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
         
         .frm1.hdnBizArea.value   = "<%=ConvSPChars(Request("txtBizArea"))%>"
         .frm1.hdnItemCd.value    = "<%=ConvSPChars(Request("txtItemCd"))%>"	
         .frm1.hdnBpCd.value      = "<%=ConvSPChars(Request("txtBpCd"))%>"
         .frm1.hdnIvFrDt.value    = "<%=ConvSPChars(Request("txtIvFrDt"))%>"
         .frm1.hdnIvToDt.value	  = "<%=ConvSPChars(Request("txtIvToDt"))%>"
         .frm1.hdnPlantCd.value   = "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.hdnIvType.value	  = "<%=ConvSPChars(Request("txtIvType"))%>"
         .frm1.hdnPurGrpCd.value  = "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
         .frm1.hdncboPstFlg.value = "<%=ConvSPChars(Request("txtPstFlg"))%>"
         
         .frm1.txtBizAreaNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 .frm1.txtItemNm.value				=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		 .frm1.txtBpNm.value				=  "<%=ConvSPChars(arrRsVal(5))%>" 	
  		 .frm1.txtPlantNm.value				=  "<%=ConvSPChars(arrRsVal(7))%>" 	
  		 .frm1.txtIvTypeNm.value			=  "<%=ConvSPChars(arrRsVal(9))%>"
  		 .frm1.txtPurGrpNm.value			=  "<%=ConvSPChars(arrRsVal(11))%>"
         .DbQueryOk
         Parent.frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
