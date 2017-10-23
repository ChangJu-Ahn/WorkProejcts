<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211QB1
'*  4. Program Name         : ǰ�񺰰���ó��ȸ 
'*  5. Program Desc         : ǰ�񺰰���ó��ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/01/08
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
On Error Resume Next														'��: 

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'�� : DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgStrPrevKey                                            '�� : ���� �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim arrRsVal(11)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
'--------------- ������ coding part(��������,End)------------------------------------------------------	
Dim lgPageNo

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")
    
	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    
    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query
   
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
	
	lgstrData  = Join(PvArr, "")

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
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,8)

    UNISqlId(0) = "M1211QA101"
    UNISqlId(1) = "M2111QA302"								              '����� 
    UNISqlId(2) = "M2111QA303"											  'ǰ���   
    UNISqlId(3) = "M3111QA102"								              '�ŷ�ó�� 
    	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     
	If Len(Request("txtPlantCd")) Then
		UNIValue(0,1)	=  " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
		UNIValue(0,2)	=  " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
	else
		UNIValue(0,1)	=  "''"
		UNIValue(0,2)	=  "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If

    If Len(Request("txtItemCd")) Then
		UNIValue(0,3)	= " " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
		UNIValue(0,4)	= " " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
	else
		UNIValue(0,3)	=  "''"
		UNIValue(0,4)	=  "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If	
    
    If Len(Request("txtSupplierCd")) Then
		UNIValue(0,5)	= " " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & " "
		UNIValue(0,6)	= " " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & " "
	else
		UNIValue(0,5)	=  "''"
		UNIValue(0,6)	=  "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If		
	
    If Request("rdoUseflg") = "A"then
	    UNIValue(0,7)	= ""
    elseif Request("rdoUseflg") = "Y"then
        UNIValue(0,7)	=" AND C.USAGE_FLG = " & FilterVar("Y", "''", "S") & " "
    else  
	    UNIValue(0,7)	= " AND C.USAGE_FLG = " & FilterVar("N", "''", "S") & " "
	end if	 
    
    UNIValue(1,0)  = UNIValue(0,1)
    UNIValue(2,0)  = UNIValue(0,1)
    UNIValue(2,1)  = UNIValue(0,3)
    UNIValue(3,0)  = UNIValue(0,5)      
        
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        exit sub
    End If    
        
    FalsechkFlg = False 
    
    '====================================== �߰��� �κ� (�̸� ����)    =====================================
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
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
		   Call DisplayMsgBox("122700", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtItemCd.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
	       FalsechkFlg = True	
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
        If Len(Request("txtSupplierCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
        
End Sub

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
         .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
         
		.frm1.hdnPlant.Value 	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnItem.Value 	= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hdnSupplier.Value = "<%=ConvSPChars(Request("txtSupplierCd"))%>"
		.frm1.hdnflg.Value 		= "<%=ConvSPChars(Request("rdoUseflg"))%>"
			
		.frm1.txtPlantNm.value	= "<%=ConvSPChars(arrRsVal(1))%>" 	
  		.frm1.txtItemNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtSupplierNm.value = "<%=ConvSPChars(arrRsVal(5))%>" 	  		 
  		
		.DbQueryOk
		          
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>

