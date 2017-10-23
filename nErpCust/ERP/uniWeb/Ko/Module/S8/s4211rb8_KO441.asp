<%
'********************************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : B/L���� 
'*  3. Program ID           : S4211RB8
'*  4. Program Name         : ������� 
'*  5. Program Desc         : B/L��� ����ȭ�� 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/07/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2002/07/30 : ADO��ȯ 
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4                           '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList													       '�� : select ����� 
Dim lgSelectListDT														   '�� : �� �ʵ��� ����Ÿ Ÿ��	
Dim lgDataExist
Dim lgPageNo
   
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
Call HideStatusWnd 

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = CInt(30)                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgTailList     = Request("lgTailList")                                 '�� : Order by value
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
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
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
Sub SetConditionData()
End Sub
'----------------------------------------------------------------------------------------------------------
	Sub FixUNISQLData()

	    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
																			  '�Ʒ��� ���� ȭ��ܿ��� �־� �ִ� query�� where�������� �� �� �ִ�.	
	    Redim UNISqlId(0)                                                        '��: SQL ID ������ ���� ����Ȯ�� 
	    Redim UNIValue(0,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
		If Request("txtSubcontractFlg") = "N" Then
			UNISqlId(0) = "S4211RA801"  ' main query(spread sheet�� �ѷ����� query statement)
		Else
			UNISqlId(0) = "S4211RA802"  ' main query(spread sheet�� �ѷ����� query statement)
		End If
	    
	    '--------------- ������ coding part(�������,End)------------------------------------------------------
	    UNIValue(0,0) = lgSelectList                                          '��: Select list
																			  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
	    '--------------- ������ coding part(�������,Start)----------------------------------------------------

		strVal = ""

		If Len(Trim(Request("txtApplicant"))) Then
			strVal = " AND CH.applicant  =  " & FilterVar(Request("txtApplicant"), "''", "S") & ""
		End If
		
		If Len(Trim(Request("txtSalesGrp"))) Then
			strVal = strVal & " AND CH.sales_grp =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""
		End If
		
		If Len(Trim(Request("txtPayTerms"))) Then
			strVal = strVal & " AND CH.pay_meth =  " & FilterVar(Request("txtPayTerms"), "''", "S") & ""
		End If
		
		If Len(Trim(Request("txtFromDt"))) Then
			strVal = strVal & " AND CH.iv_dt >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
		End If
		
		If Len(Trim(Request("txtToDt"))) Then
			strVal = strVal & " AND CH.iv_dt <=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
		End If

	If Len(Request("gBizArea")) Then
		strVal = strVal & " AND CH.BIZ_AREA = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If

	If Len(Request("gPlant")) Then
		strVal = strVal & " AND CD.PLANT_CD = " & FilterVar(Request("gPlant"), "''", "S") & " "			
	End If

	If Len(Request("gSalesGrp")) Then
		strVal = strVal & " AND CH.SALES_GRP = " & FilterVar(Request("gSalesGrp"), "''", "S") & " "			
	End If

	If Len(Request("gSalesOrg")) Then
		strVal = strVal & " AND CH.SALES_ORG = " & FilterVar(Request("gSalesOrg"), "''", "S") & " "			
	End If
		
		 UNIValue(0,1) = strVal				'UNISqlId(0)�� �ι�° ?�� �Էµ�	
	  
	     '--------------- ������ coding part(�������,End)------------------------------------------------------
	    UNIValue(0,UBound(UNIValue, 2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
	    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	 
	End Sub
'----------------------------------------------------------------------------------------------------------
	Sub QueryData()
	    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
		Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	    Dim iStr
		
		Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'��:ADO ��ü�� ���� 
	    
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4)

	    Set lgADF   = Nothing    
	    iStr = Split(lgstrRetMsg,gColSep)
		
		
	    If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If    
	    
	    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
	        rs0.Close
	        Set rs0 = Nothing
	        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)	'No Data Found!!
			%>
			<Script Language=vbscript>
			parent.frm1.txtApplicant.focus
			</Script>	
			<%
	    Else   
	        Call  MakeSpreadSheetData()
			Call  SetConditionData()
	    End If
	   
	End Sub
%>   

<Script Language=vbscript>
	With parent.frm1
	
		<%If lgDataExist = "Yes" Then%>
		'Set condition data to hidden area
		' "1" means that this query is first and next data exists	
			<%If lgPageNo = "1" Then %>
			.txtHApplicant.value	= "<%=Request("txtApplicant")%>"
			.txtHFromDt.value		= "<%=Request("txtFromDT")%>"
			.txtHToDt.value			= "<%=Request("txtToDT")%>"
			.txtHSalesGrp.value		= "<%=Request("txtSalesGrp")%>"
			.txtHPayTerms.value		= "<%=Request("txtPayTerms")%>"
			.txtHSubcontractFlg.value = "<%=Request("txtSubcontractFlg")%>"
			<%End If%>
	    
	    'Show multi spreadsheet data from this line
	    If "<%=Request("txtSubcontractFlg")%>" = "N" Then
			parent.ggoSpread.Source = .vspdData 
		Else
			parent.ggoSpread.Source = .vspdData1
		End If	 
		
		parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>"
		parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
        parent.DbQueryOk
		<%End If%>
	End With
</Script>	
<%
Response.End
%>
