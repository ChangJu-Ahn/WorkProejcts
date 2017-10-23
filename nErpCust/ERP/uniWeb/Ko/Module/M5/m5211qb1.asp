<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5211qb1
'*  4. Program Name         : ������Ȳ��ȸ 
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
<!-- #Include file="../../inc/IncServer.asp" -->
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
Dim arrRsVal(11)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
Dim iFrPoint
iFrPoint=0
Dim lgPageNo
Dim lgDataExist

Const Major_Cd_Incoterms = "B9006"

    Call HideStatusWnd 

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

    Redim UNISqlId(6)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(6,17)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "m5211qa101"
     UNISqlId(1) = "M3111QA104"								              '���ű׷�� 
     UNISqlId(2) = "M3111QA102"								              '����ó�� 
	 UNISqlId(3) = "S0000QA000"											  '�������� 
		
	 '--- 2004-08-19 by Byun Jee Hyun for UNICODE																	  'Reusage is Recommended

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

     UNIValue(0,1)  = " " & FilterVar(UNIConvDate(strBlFrDt), "''", "S") & ""
     UNIValue(0,2)  = " " & FilterVar(UNIConvDate(strBlToDt), "''", "S") & ""
     UNIValue(0,3)  = UCase(Trim(strBpCdFrom))			
     UNIValue(0,4)  = " " & FilterVar(UNIConvDate(strLoadingFrDt), "''", "S") & ""			
     UNIValue(0,5)  = " " & FilterVar(UNIConvDate(strLoadingToDt), "''", "S") & ""
	 UNIValue(0,6)  = UCase(Trim(strIncotermsCdFrom))	
	 UNIValue(0,7)  = UCase(Trim(strPurGrpCdFrom))	    
     UNIValue(0,8)  = UCase(Trim(strCfmFlgFrom))		
     
     UNIValue(1,0)  = UCase(Trim(strPurGrpCd))  
     UNIValue(2,0)  = UCase(Trim(strBpCd))
	 UNIValue(3,0)  = FilterVar(UCase(Trim(Major_Cd_Incoterms)), "''", "S")
     UNIValue(3,1)  = UCase(Trim(strIncotermsLookUp))
     
     
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
        
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs2(0)
		arrRsVal(5) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtIncotermsCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(8) = rs3(0)
		arrRsVal(9) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs1(0)
		arrRsVal(3) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
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

     '---�������� 
    If Len(Trim(Request("txtIncotermsCd"))) Then
    	strIncotermsCd	= " " & FilterVar(Request("txtIncotermsCd"), "''", "S") & ""
    	strIncotermsCdFrom = strIncotermsCd
    Else
    	strIncotermsCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strIncotermsCdFrom = "" & FilterVar("%%", "''", "S") & ""
    End If
    If Len(Trim(Request("txtIncotermsCd"))) Then
    	strIncotermsCd	= FilterVar(Trim(Request("txtIncotermsCd")), "''", "S")
    Else
    	strIncotermsCd	= FilterVar("zzzzzzzzz", "''", "S")
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
    	strBlFrDt 	= "" & (Request("txtBlFrDt")) & ""
    Else
    	strBlFrDt	= unidateClientFormat("1900-01-01")
    End If
    If Len(Trim(Request("txtBlToDt"))) Then
    	strBlToDt 	= "" & (Request("txtBlToDt")) & ""
    Else
    	strBlToDt	= unidateClientFormat("2999-12-30")
    End If  
     '---������ 
    If Len(Trim(Request("txtLoadingFrDt"))) Then
    	strLoadingFrDt 	= "" & (Request("txtLoadingFrDt")) & ""
    Else
    	strLoadingFrDt	= unidateClientFormat("1900-01-01")
    End If
	
    If Len(Trim(Request("txtLoadingToDt"))) Then
    	strLoadingToDt 	= "" & (Request("txtLoadingToDt")) & ""
    Else
    	strLoadingToDt	= unidateClientFormat("2999-12-30")
    End If     
     '---Ȯ������ 
    If Len(Trim(Request("txtCfmFlg"))) Then
    	strCfmFlg	= " " & FilterVar(Request("txtCfmFlg"), "''", "S") & ""
    	strCfmFlgFrom = strCfmFlg
    Else
    	strCfmFlg	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strCfmFlgFrom = "" & FilterVar("%%", "''", "S") & ""
    End If

End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         Parent.frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '�� : Display data
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",11),.GetKeyPos("A",10),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.parent.gCurrency,.GetKeyPos("A",13),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.Parent.gCurrency,.GetKeyPos("A",12),"D","Q","X","X")
         
         .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
  		 
		 .frm1.hdnBeneficiaryCd.value	= "<%=ConvSPChars(Request("txtBpCd"))%>"
         .frm1.hdnIncotermsCd.value		= "<%=ConvSPChars(Request("txtIncotermsCd"))%>"
         .frm1.hdnPurGrpCd.value		= "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
		 .frm1.hdnBlIssueFrDt.value		= "<%=ConvSPChars(Request("txtBlFrDt"))%>"
         .frm1.hdnBlIssueToDt.value		= "<%=ConvSPChars(Request("txtBlToDt"))%>"
         .frm1.hdnLoadingFrDt.value		= "<%=ConvSPChars(Request("txtLoadingFrDt"))%>"
         .frm1.hdnLoadingToDt.value	    = "<%=ConvSPChars(Request("txtLoadingToDt"))%>"
		 .frm1.hdnstrCfmFlg.value	    = "<%=ConvSPChars(Request("txtCfmFlg"))%>"
  		 .frm1.txtPurGrpNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		 .frm1.txtBeneficiaryNm.value	=  "<%=ConvSPChars(arrRsVal(5))%>" 	
		 .frm1.txtIncotermsNm.value		=  "<%=ConvSPChars(arrRsVal(9))%>" 	
  		 
         .DbQueryOk
         Parent.frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
