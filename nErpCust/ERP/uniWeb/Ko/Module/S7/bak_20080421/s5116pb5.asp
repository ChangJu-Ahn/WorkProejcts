<%'======================================================
'********************************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5116pa5
'*  4. Program Name         : ����ä�ǻ� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/05/03
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Kwak Eunkyoung
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

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'�� : DBAgent Parameter ���� 
   Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
   Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")

'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim lgFromDt			'��ȸ�Ⱓ���� 
    Dim lgToDt				'��ȸ�Ⱓ�� 
    Dim lgBizArea			'����� 
	Dim lgBillTypeCd		'����ä������ 
	Dim lgBpCd				'�ŷ�ó 
    Dim lgRdoFlag			'����ä��Ȯ������ 
   
    lgFromDt		= Trim(Request("txtHConFromDt"))    
    lgToDt			= Trim(Request("txtHConToDt"))
    lgBizArea		= Trim(Request("txtHConBizArea"))
    lgBillTypeCd	= Trim(Request("txtHConBillType"))
    lgBpCd			= Trim(Request("txtHConBpCd"))
    lgRdoFlag		= Trim(Request("txtHConRdoFlag"))
            
'--------------- ������ coding part(��������,End)----------------------------------------------------------

    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value

    lgMaxCount       = 50							                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 

    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim iStrVal
	Dim arrVal(0)
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 

    Redim UNIValue(1,6)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 

    iStrVal = "WHERE"    				
	iStrVal = iStrVal & " BILL_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	If Len(lgToDt) Then
		iStrVal = iStrVal & " AND BILL_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""			
	End If				

	'�����=========================================================================================
	If Len(lgBizArea) Then
		UNIValue(0,2)	=  " " & FilterVar(lgBizArea, "''", "S") & ""
	Else
		UNIValue(0,2)	= "NULL"
	End If		

	'����ä�����¸�=============================================================================================    	
	If Len(lgBillTypeCd) Then
		UNIValue(0,3)	=  " " & FilterVar(lgBillTypeCd, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
	End If

	'�ŷ�ó=========================================================================================
	If Len(lgBpCd) Then
		UNIValue(0,4)	=  " " & FilterVar(lgBpCd, "''", "S") & ""
	Else
		UNIValue(0,4)	= "NULL"
	End If		

	'Ȯ������===========================================================================================	
	If lgRdoFlag <> "%" Then
		UNIValue(0,5)	= " " & FilterVar(lgRdoFlag, "''", "S") & ""
	Else
		UNIValue(0,5)	= "NULL"
	End If

	UNISqlId(0) = "S5116PA501"
	UNISqlId(1) = "S5116PA501"											
    UNIValue(0,0) = Trim(lgSelectList)                                      
	UNIValue(0,1) = iStrVal	         

	UNIValue(1,0) = " SUM(ISNULL(BH.BILL_AMT_LOC,0) + ISNULL(BH.VAT_AMT_LOC,0)) AS TOTAL_AMT, SUM(ISNULL(BH.BILL_AMT_LOC,0)) AS BILL_AMT, SUM(ISNULL(BH.VAT_AMT_LOC,0)) AS VAT_AMT, SUM(ISNULL(BH.COLLECT_AMT_LOC,0)) AS COLLECT_AMT, SUM(ISNULL(BH.DEPOSIT_AMT_LOC,0)) AS DEPOSIT_AMT "
	Dim iLoop
	For iLoop = 1 To 5
		UNIValue(1,iLoop) = UNIValue(0,iLoop)	
	Next

'================================================   
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
       'Show multi spreadsheet data from this line
		With parent       
			.ggoSpread.Source  = .frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'�� : Display data
			.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag

			.frm1.txtTotalAmt.text		= "<%=rs1(0)%>"
			.frm1.txtBillAmt.text		= "<%=rs1(1)%>"
			.frm1.txtVatAmt.text		= "<%=rs1(2)%>"
			.frm1.txtCollectAmt.text	= "<%=rs1(3)%>"
			.frm1.txtDepositAmt.text	= "<%=rs1(4)%>"

 			.DbQueryOk
			.frm1.vspdData.Redraw = True        
		End with
    End If   

</Script>	
