<%
'********************************************************************************************************
'*  1. Module Name          : ��������																	*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3212rb1.asp																*
'*  4. Program Name         : L/C ��������(L/C Amend ������Ͽ���)										*
'*  5. Program Desc         : L/C ��������(L/C Amend ������Ͽ���)										*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/07																*
'*  8. Modified date(Last)  : 2002/04/26																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : ȭ�� design												*
'*                            2. Ado ��ȯ																*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")   
Call LoadBNumericFormatB("I","*","NOCOOKIE","RB")

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           '�� : DBAgent Parameter ���� 
   Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ����   
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim strItemNm 'item Name
   Dim BlankchkFlg
'--------------- ������ coding part(��������,Start)----------------------------------------------------
   Dim iFrPoint
   iFrPoint=0
   Const C_SHEETMAXROWS_D  = 30                                          '��: Fetch max count at once
'--------------- ������ coding part(��������,End)------------------------------------------------------

    Call HideStatusWnd 
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)                  
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                               '�� : DB-Agent�� ���� parameter ����Ÿ set    	
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
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
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
    On Error Resume Next

    If Not(rs1.EOF Or rs1.BOF) Then
        strItemNm =  rs1(1)
        rs1.Close
        Set rs1 = Nothing
    Else
		rs1.Close
		Set rs1 = Nothing
		If Len(Request("txtItem")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			BlankchkFlg  =  True
		 %>
            <Script language=vbs>
            parent.frm1.txtItem.focus    
            </Script>
         <%		 				
		End If
	End If   	

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim strVal1
    Dim strVal2
	Dim arrVal(0)
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,3)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "S3212RA101" 
     UNISqlId(1) = "s0000qa001"
     
'--------------- ������ coding part(�������,End)------------------------------------------------------

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strVal = " "

	If Len(Request("txtLCNo")) Then
		strVal = " " & FilterVar(Request("txtLCNo"), "''", "S") & " "	
	End If
	

	If Len(Request("txtHLCAmdNo")) Then
		strVal2 = " " & FilterVar(Request("txtHLCAmdNo"), "''", "S") & " "	
	End If

	If Len(Request("txtItem")) Then
		strVal1 = " AND a.item_cd = " & FilterVar(Request("txtItem"), "''", "S") & " "
		arrVal(0) = Trim(Request("txtItem"))
	End If	
	
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " and b.TRACKING_NO = " & FilterVar(Request("txtTrackingNo"), "''", "S") & " "	
	End If



'--------------- ������ coding part(�������,End)----------------------------------------------------

    UNIValue(0,1) = strVal   'UNISqlId(0)�� ù��° Where������ ���� 
    UNIValue(0,2) = strVal2   'UNISqlId(0)�� ù��° Where������ ���� 
    UNIValue(0,3) = strVal1   'UNISqlId(0)�� �ι�° Where������ ���� 
    UNIValue(1,0) = FilterVar(Trim(Request("txtItem")), " " , "S")  'UNISqlId(1)�� ù��° Where������ ���� 
    
'================================================================================================================   
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    Call  SetConditionData()
    
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then         
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
		 %>
            <Script language=vbs>
            parent.frm1.txtItem.focus    
            </Script>
         <%		 		    
		Else    
		    Call  MakeSpreadSheetData()	    
		End If
	End If

End Sub
%>
<Script Language=vbscript>

    parent.frm1.txtItemNm.Value	= "<%=ConvSPChars(strItemNm)%>"

    If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			
			parent.frm1.txtHLCNo.value = "<%=ConvSPChars(Request("txtLCNo"))%>"
			parent.frm1.txtHItem.value = "<%=ConvSPChars(Request("txtItem"))%>"
			parent.frm1.txtHLCAmdNo.value = "<%=ConvSPChars(Request("txtHLCAmdNo"))%>"
                        parent.frm1.txtHTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"						
			
       End If
       'Show multi spreadsheet data from this line
       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.frm1.vspdData.Redraw = False
	   parent.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
	   
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,parent.frm1.txtCurrency.value,Parent.GetKeyPos("A",6),"C", "Q" ,"X","X")		
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,parent.frm1.txtCurrency.value,Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")		
	     	            
       parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
 
       parent.DbQueryOk
       parent.frm1.vspdData.Redraw = True
    End If   
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>