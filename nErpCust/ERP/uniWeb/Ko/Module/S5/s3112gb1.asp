<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S3112GA1
'*  4. Program Name         : �����������ȸ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Dateǥ������ 
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
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim arrRsVal(7)
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = 100						                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"    

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
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
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(3)
    Dim MajorCd
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,2)

    UNISqlId(0) = "S3112GA101"
    UNISqlId(1) = "S0000qA002"
    UNISqlId(2) = "S0000qA005"
    UNISqlId(3) = "S0000qA001"
    UNISqlId(4) = "s0000qa009"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = "(SELECT B.DN_NO , B.DN_SEQ , B.TRACKING_NO, B.ITEM_CD , C.ITEM_NM , C.SPEC, A.SHIP_TO_PARTY , D.BP_NM AS SHIP_TO_PARTY_NM,"
	strVal = strVal & " J.SOLD_TO_PARTY, G.BP_NM AS SOLD_TO_PARTY_NM , NULL AS SO_NO , NULL AS SO_SEQ, NULL AS SO_DT,"
	strVal = strVal & " A.DLVY_DT, NULL AS SO_UNIT, A.SALES_GRP, H.SALES_GRP_NM, B.PLANT_CD, I.PLANT_NM,"
	strVal = strVal & " CASE WHEN J.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -B.REQ_QTY ELSE B.REQ_QTY END AS REQ_QTY,"
	strVal = strVal & " CASE WHEN J.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -B.GI_QTY ELSE B.GI_QTY END AS GI_QTY "
	strVal = strVal & " FROM S_DN_HDR A INNER JOIN S_DN_DTL B ON (A.dn_no = B.dn_no) "
	strVal = strVal & " INNER JOIN B_ITEM C ON (B.item_cd = C.item_cd) "
	strVal = strVal & "	INNER JOIN B_BIZ_PARTNER D ON (A.SHIP_TO_PARTY = D.BP_CD) "
	strVal = strVal & " INNER JOIN B_SALES_GRP H ON (A.SALES_GRP = H.SALES_GRP) "
	strVal = strVal & " INNER JOIN B_PLANT I ON (B.PLANT_CD = I.PLANT_CD) "
	strVal = strVal & " INNER JOIN S_DN_SALES J ON (A.DN_NO = J.DN_NO) "
	strVal = strVal & " LEFT OUTER JOIN B_BIZ_PARTNER G ON (J.SOLD_TO_PARTY = G.BP_CD) "
	strVal = strVal & " WHERE A.POST_FLAG <> " & FilterVar("Y", "''", "S") & "  "
	strVal = strVal & " UNION ALL "
	strVal = strVal & " SELECT B.DN_NO, B.DN_SEQ, B.TRACKING_NO, B.ITEM_CD, C.ITEM_NM, C.SPEC, A.SHIP_TO_PARTY, D.BP_NM, E.SOLD_TO_PARTY, "
	strVal = strVal & " G.BP_NM, E.SO_NO, F.SO_SEQ, E.SO_DT, A.DLVY_DT, F.SO_UNIT, A.SALES_GRP, H.SALES_GRP_NM,"
	strVal = strVal & " B.PLANT_CD,  I.PLANT_NM, CASE WHEN E.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -B.REQ_QTY ELSE B.REQ_QTY END, "
	strVal = strVal & " CASE WHEN E.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -B.GI_QTY ELSE B.GI_QTY END "
	strVal = strVal & " FROM S_DN_HDR A, S_DN_DTL B, B_ITEM C, B_BIZ_PARTNER D, S_SO_HDR E, S_SO_DTL F, B_BIZ_PARTNER G,"
	strVal = strVal & " B_SALES_GRP H, B_PLANT I "
	strVal = strVal & " WHERE A.DN_NO = B.DN_NO AND B.ITEM_CD = C.ITEM_CD AND A.SHIP_TO_PARTY = D.BP_CD "
	strVal = strVal & " AND E.SOLD_TO_PARTY = G.BP_CD AND E.SO_NO = F.SO_NO AND B.SO_NO = F.SO_NO AND B.SO_SEQ = F.SO_SEQ "
	strVal = strVal & " AND A.SALES_GRP = H.SALES_GRP AND B.PLANT_CD = I.PLANT_CD AND A.POST_FLAG <> " & FilterVar("Y", "''", "S") & "  ) T " 
	strVal = strVal & " WHERE DN_NO IS NOT NULL"

	Dim lgDate 
	lgDate = GetSvrDate
	
	Dim ChangeFlg
	ChangeFlg =False

	If Len(Request("txtconBp_cd")) Then
		strVal = strVal & " AND SHIP_TO_PARTY = " & FilterVar(Request("txtconBp_cd"), "''", "S") & " "		
		ChangeFlg = True
'	Else
'		strVal = ""
	End If
	arrVal(0) = FilterVar(Request("txtconBp_cd"), " ", "S")

 	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
	End If		
	arrVal(1) = FilterVar(Request("txtSalesGroup"), " ", "S")
    
	If Len(Request("txtItem_cd")) Then
		strVal = strVal & " AND ITEM_CD = " & FilterVar(Request("txtItem_cd"), "''", "S") & " "			
	End If		
	arrVal(2) = FilterVar(Request("txtItem_cd"), " ", "S")

	If Len(Request("txtPlant")) Then
		strVal = strVal & " AND PLANT_CD = " & FilterVar(Request("txtPlant"), "''", "S") & " "	
	End If		
	arrVal(3) = FilterVar(Request("txtPlant"), " ", "S")

    If Len(Request("txtDlvyFromDt")) Then
		strVal = strVal & " AND DLVY_DT >= " & FilterVar(UNIConvDate(Request("txtDlvyFromDt")), "''", "S") & ""		
	End If		

    If Len(Request("txtDlvyToDt")) Then
		strVal = strVal & " AND DLVY_DT <= " & FilterVar(UNIConvDate(Request("txtDlvyToDt")), "''", "S") & ""		
	End If		
 
	If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND TRACKING_NO = " & FilterVar(Trim(Request("txtTrackingNo")), "''" , "S") & ""
	End If
	
    UNIValue(0,1) = strVal   '---������ 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
    UNIValue(3,0) = arrVal(2)
    UNIValue(4,0) = arrVal(3)
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

   		If Len(Trim(Request("txtconBp_cd"))) Then
		   Call DisplayMsgBox("970000", vbInformation, "��ǰó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconBp_cd.focus    
                </Script>
            <%        
	       	
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
		If FalsechkFlg = False Then
   			If Len(Trim(Request("txtSalesGroup"))) Then
			   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			   FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGroup.focus    
                </Script>
            <%        
			   
			End If	
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

		If FalsechkFlg = False Then
   			If Len(Trim(Request("txtItem_cd"))) Then
			   Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			   FalsechkFlg = True	
                ' Modify Focus Events    
                %>
                    <Script language=vbs>
                    Parent.frm1.txtItem_cd.focus    
                    </Script>
                <%        
			   
			End If	
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

		If FalsechkFlg = False Then
	   		If Len(Trim(Request("txtPlant"))) Then
			   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			   FalsechkFlg = True
                ' Modify Focus Events    
                %>
                    <Script language=vbs>
                    Parent.frm1.txtPlant.focus    
                    </Script>
                <%        
			   
			End If	
		End If
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtSalesGroup.focus    
            </Script>
        <%
                
    Else    
        Call  MakeSpreadSheetData()
    End If
  
End Sub
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub
%>
<Script Language=vbscript>

With parent

    .frm1.txtconBp_nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
    .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
    .frm1.txtItem_Nm.value		= "<%=ConvSPChars(arrRsVal(5))%>" 
    .frm1.txtPlantNm.value		= "<%=ConvSPChars(arrRsVal(7))%>" 

	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists

			parent.frm1.HtxtconBp_cd.value = "<%=ConvSPChars(Request("txtconBp_cd"))%>"
			parent.frm1.HtxtSalesGroup.value = "<%=ConvSPChars(Request("txtSalesGroup"))%>"
			parent.frm1.HtxtItem_cd.value = "<%=ConvSPChars(Request("txtItem_cd"))%>"
			parent.frm1.HtxtPlant.value = "<%=ConvSPChars(Request("txtPlant"))%>"
			parent.frm1.HtxtDlvyFromDt.value = "<%=Request("txtDlvyFromDt")%>"
			parent.frm1.HtxtDlvyToDt.value = "<%=Request("txtDlvyToDt")%>"
			parent.frm1.HtxtTrackingNo.value  = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
        
        End if
        
		.frm1.vspdData.Redraw = False
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                            '��: Display data 
        .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.frm1.vspdData.Redraw = True
        .DbQueryOk
    End if        

End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

