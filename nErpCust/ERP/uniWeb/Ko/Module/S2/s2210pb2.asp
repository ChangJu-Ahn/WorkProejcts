<%
'********************************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2210PB2
'*  4. Program Name         : ����� �̵�� ��Ȳ Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/01/16
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
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
   
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
'--------------- ������ coding part(��������,End)----------------------------------------------------------
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
Call HideStatusWnd 

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)				'��: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = CInt(30)											'�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgTailList     = Request("lgTailList")                                 '�� : Order by value
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgDataExist      = "No"

Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
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
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
	Sub FixUNISQLData()

	    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
																			  '�Ʒ��� ���� ȭ��ܿ��� �־� �ִ� query�� where�������� �� �� �ִ�.	
	    Redim UNISqlId(0)                                                        '��: SQL ID ������ ���� ����Ȯ�� 
	    Redim UNIValue(0,4)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 

	    UNISqlId(0) = "S2210PA201"  ' main query(spread sheet�� �ѷ����� query statement)
	    
	    '--------------- ������ coding part(�������,End)------------------------------------------------------
	    UNIValue(0,0) = lgSelectList                                          '��: Select list
																			  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
	    '--------------- ������ coding part(�������,Start)----------------------------------------------------
		UNIValue(0,1) = Trim(Request("txtFrSpPeriod"))
		UNIValue(0,2) = Trim(Request("txtToSpPeriod"))
		
		If Len(Trim(Request("txtSalesGrp"))) Then
			UNIValue(0,3) = " AND SALES_GRP =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""
		End If
	  
	     '--------------- ������ coding part(�������,End)------------------------------------------------------
	    UNIValue(0,UBound(UNIValue, 2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
	    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	 
	End Sub
 
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
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
	    Else   
	        Call  MakeSpreadSheetData()
			Call  SetConditionData()
	    End If
	   
	End Sub
%>   

<Script Language=vbscript>
	With parent.frm1
	
		<%If lgDataExist = "Yes" Then%>
	    'Show multi spreadsheet data from this line
		parent.ggoSpread.Source = .vspdData 
		parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>"
		parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
        parent.DbQueryOk
		<%End If%>
	End With
</Script>	
<%
Response.End
%>
