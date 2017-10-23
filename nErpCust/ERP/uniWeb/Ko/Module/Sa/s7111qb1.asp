<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S7111QB1
'*  4. Program Name         : NEGO ��Ȳ��ȸ 
'*  5. Program Desc         : NEGO ��Ȳ��ȸ 
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 12. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3               '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim arrRsVal(5)
Dim FalsechkFlg
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call HideStatusWnd 


    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(100)                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Private Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Private Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
	Dim MajorCd
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "S7111QA101"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA000"


    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtconBp_cd")) Then
		strVal = "AND A.APPLICANT = " & FilterVar(Trim(Request("txtconBp_cd")), "" , "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(Request("txtconBp_cd")), "" , "S")

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Trim(Request("txtSalesGroup")), "" , "S") & " "				
	End If		
    arrVal(1) = FilterVar(Trim(Request("txtSalesGroup")), "" , "S")
    
 	If Len(Request("txtPayTerms")) Then
		strVal = strVal & " AND A.PAY_METH = " & FilterVar(Trim(Request("txtPayTerms")), "" , "S") & " "			
		MajorCd = "B9004"
	End If		
	arrVal(2) = FilterVar(Trim(Request("txtPayTerms")), "" , "S")
    
    If Len(Request("txtNegoFrDt")) Then
		strVal = strVal & " AND A.NEGO_DT >= " & FilterVar(UNIConvDate(Trim(Request("txtNegoFrDt"))), "''", "S") & ""		
	End If		
	
	If Len(Request("txtNegoToDt")) Then
		strVal = strVal & " AND A.NEGO_DT <= " & FilterVar(UNIConvDate(Trim(Request("txtNegoToDt"))), "''", "S") & ""		
	End If

	If Request("txtRadio") = "Y" Then
		strVal = strVal & "AND A.FLAW_EXIST = " & FilterVar("Y", "''", "S") & " "
	ElseIf Request("txtRadio") = "N" Then
		strVal = strVal & "AND A.FLAW_EXIST = " & FilterVar("N", "''", "S") & " "			
	End If

    UNIValue(0,1) = strVal   '---������ 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
    UNIValue(3,0) =  FilterVar(MajorCd, "''", "S")
    UNIValue(3,1) = arrVal(2)
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Private Sub QueryData()
    Dim iStr
	FalsechkFlg = False 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
 
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

		If Len(Request("txtSalesGroup")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			%>
			<Script Language=vbscript>
			parent.frm1.txtSalesGroup.focus
		    </Script>	
		    <%
	        FalsechkFlg = True		
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If 
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
		If FalsechkFlg = False Then
			If Len(Request("txtconBp_cd")) And FalsechkFlg = False Then
				Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				parent.frm1.txtconBp_cd.focus
				</Script>	
				<%
				FalsechkFlg = True	
			End If	
		End if
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
		If FalsechkFlg = False Then
			If Len(Request("txtPayTerms")) And FalsechkFlg = False Then
				Call DisplayMsgBox("970000", vbInformation, "�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				parent.frm1.txtPayTerms.focus
				</Script>	
				<%
				FalsechkFlg = True		
			End if
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False  Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Private Sub TrimData()
End Sub


%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
          .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'�� : Display data
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",2),.GetKeyPos("A",3),"A","Q","X","X")
                  
         .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '��: set next data tag
         .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
         .frm1.txtconBp_nm.value		= "<%=ConvSPChars(arrRsVal(1))%>"
         .frm1.txtPaytermsNm.value		= "<%=ConvSPChars(arrRsVal(5))%>"
         .frm1.txtHconBp_cd.value	    = "<%=ConvSPChars(Request("txtconBp_cd"))%>"
 		 .frm1.txtHSalesGroup.value	    = "<%=ConvSPChars(Request("txtSalesGroup"))%>"
 		 .frm1.txtHPayTerms.value	    = "<%=ConvSPChars(Request("txtPayTerms"))%>"
 		 .frm1.txtHNegoFrDt.value	    = "<%=Request("txtNegoFrDt")%>"
         .frm1.txtHNegoToDt.value	    = "<%=Request("txtNegoToDt")%>" 
         .frm1.txtHRadio.value	    = "<%=ConvSPChars(Request("txtRadio"))%>"
 
         If "<%=FalsechkFlg%>" = False Then
         .DbQueryOk
         End If
		  .frm1.vspdData.Redraw = True
		
	End with
</Script>	

