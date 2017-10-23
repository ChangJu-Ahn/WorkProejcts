<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : s3134mb1(ADO)
'*  4. Program Name         : ��� ��Ȳ��ȸ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/09
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Cho inkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date ǥ������ 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgArrData                                                              '�� : data for spreadsheet data

Dim lgPageNo                                                           '�� : ���� �� 
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
Dim arrRsVal(9)
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q","S","NOCOOKIE","QB")
    Call HideStatusWnd 
		
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = 100							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim iArrRow
    Dim iRowCnt
    Dim iColCnt
	Dim iLngStartRow
     call svrmsgbox("1", vbinformation, i_mkscript)
    ReDim iArrRow(UBound(lgSelectListDT) - 1)
	
	iLngStartRow = CLng(lgMaxCount) * CLng(lgPageNo)
	
	' Scroll ��ȸ�� Client�� ���� ù ���� Row�� �̵��Ѵ�.
    If CLng(lgPageNo) > 0 Then
       rs0.Move = iLngStartRow
    End If
     call svrmsgbox("2", vbinformation, i_mkscript)
    ' Client�� ������ ��ȸ����� �� Page�� �Ѿ �� 
    If rs0.RecordCount > CLng(lgMaxCount) * (CLng(lgPageNo) + 1) Then
        lgPageNo = lgPageNo + 1
	    Redim lgArrData(lgMaxCount - 1)

    ' Client�� ������ ��ȸ����� �� Page�� ���� ���� ��, �� ������ �ڷ��� ��� 
    Else
		Redim lgArrData(rs0.RecordCount - (iLngStartRow + 1))
		lgPageNo = ""
    End If

    For iRowCnt = 0 To UBound(lgArrData)
		For iColCnt = 0 To UBound(lgSelectListDT) - 1 
            iArrRow(iColCnt) = FormatRsString(lgSelectListDT(iColCnt),rs0(iColCnt))
		Next
		
		lgArrData(iRowCnt) = Chr(11) & Join(iArrRow, Chr(11))
		
        rs0.MoveNext
    Next
   

    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,5)

    UNISqlId(0) = "s4121ma1_ko441"
    UNISqlId(1) = "s0000qa002"
    UNISqlId(2) = "s0000qa009"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "

  '  If Len(Request("txtSoDtFrom")) Then
'		strVal = strVal & " AND DLVY_DT >= " & FilterVar(UNIConvDate(Request("txtSoDtFrom")), "''", "S") & ""		
	'End If		


    UNIValue(0,1) = FilterVar(Request("txtPlantCode"), "''", "S")
	UNIValue(0,2) = FilterVar(Request("txtShipToParty"), "''", "S")
	UNIValue(0,3) = FilterVar(Request("txtSoDtFrom"), "''", "S")
	UNIValue(0,4) = FilterVar(Request("txtSoDtTo"), "''", "S")

    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,5) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub QueryData()
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    FalsechkFlg = False
   
   
    If Len(Request("txtShipToParty")) Then
		If  rs1.EOF And rs1.BOF Then
		    rs1.Close
		    Set rs1 = Nothing

			If Len(Request("txtShipToParty")) And FalsechkFlg = False Then
			   Call DisplayMsgBox("970000", vbInformation, "��ü", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		       FalsechkFlg = True	
		        ' Modify Focus Events    
		        %>
		            <Script language=vbs>
		            Parent.frm1.txtShipToParty.focus    
		            </Script>
		        <%        	       	       
			End If	
			Exit Sub
		Else    
			arrRsVal(2) = rs1(0)
			arrRsVal(3) = rs1(1)
		    rs1.Close
		    Set rs1 = Nothing
		End If
	End If
    call svrmsgbox(Request("txtPlantCode"), vbinformation, i_mkscript)
   If Len(Request("txtPlantCode")) Then
		If  rs2.EOF And rs2.BOF Then
		    rs2.Close
		    Set rs2 = Nothing

			If Len(Request("txtPlantCode")) And FalsechkFlg = False Then
			   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		       FalsechkFlg = True	
		        ' Modify Focus Events    
		        %>
		            <Script language=vbs>
		            Parent.frm1.txtPlantCode.focus    
		            </Script>
		        <%        	       	       
			End If	
			Exit Sub
		Else    
			arrRsVal(6) = rs2(0)
			arrRsVal(7) = rs2(1)		
		    rs2.Close
		    Set rs2 = Nothing
		End If
	End If
    call svrmsgbox(Request("txtPlantCode"), vbinformation, i_mkscript)
    
     iStr = Split(lgstrRetMsg,gColSep)
     call svrmsgbox(iStr(0), vbinformation, i_mkscript)
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        ' Modify Focus Events    
        %>
            <Script language=vbs>
				Call parent.SetFocusToDocument("M")	
				parent.frm1.txtSoDtFrom.Focus
            </Script>
        <%        	               
    Else    
        Call  MakeSpreadSheetData()
    End If

   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub


%>
<Script Language=vbscript>
    With parent
		.ggoSpread.Source    = .frm1.vspdData 
                
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip  "<%=Join(lgArrData, Chr(11) & Chr(12)) & Chr(11) & Chr(12)%>", "F"
	
        .lgPageNo        =  "<%=lgPageNo%>"                       '��: set next data tag
<%If UNICInt(Trim(Request("lgPageNo")),0) = 0 Then %>        
        .frm1.txtShipToPartyNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
		.frm1.txtPlantName.value		= "<%=ConvSPChars(arrRsVal(7))%>"

		<%If Trim(lgPageNo) <> "" Then %>
		.frm1.HShipToParty.value	= "<%=ConvSPChars(Request("txtShipToParty"))%>"
		.frm1.HPlantCode.value		= "<%=ConvSPChars(Request("txtPlantCode"))%>"
		.frm1.HSoDtFrom.value		= "<%=Request("txtSoDtFrom")%>"
		.frm1.HSoDtTo.value			= "<%=Request("txtSoDtTo")%>"
		<%End If%>
<%End If%>

        .DbQueryOk
        .frm1.vspdData.Redraw = True
	End with
</Script>	
<%
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
