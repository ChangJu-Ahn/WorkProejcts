<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/12
'*  7. Modified date(Last)  : 2001/01/12
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "QB")

Dim pPB6S101
Dim lgADF, ADF
Dim lgstrRetMsg, strRetMsg, iStr
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPlantCd	                                                           '�� : �����ڵ� 
Dim strItemCd	                                                           '�� : ǰ���ڵ� 
Dim strRoutingNo	                                                           '�� : ����ù�ȣ 
Dim TmpBuffer
Dim iTotalStr
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
'-----------------------
'Com action area
'-----------------------
Set pPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If



Call pPB6S101.B_LOOK_UP_PLANT(gStrGlobalCollection, strPlantCd)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB6S101 = Nothing															'��: Unload Component
	Response.End
End If

Set pPB6S101 = Nothing															'��: Unload Component
	
'======================================================================================================
'	ǰ���̸� ó�����ִ� �κ� 
'======================================================================================================
Redim UNISqlId(2)
Redim UNIValue(2, 0)
	
UNISqlId(0) = "122600sac"
UNISqlId(1) = "122700sab"
UNISqlId(2) = "181300sac"
	
strPlantCd = Trim(UCase(Request("txtPlantCd") ))
strItemCd = Request("txtItemCd")
strRoutingNo = Request("txtRoutingNo") 
	
UNIValue(0, 0) = FilterVar(strItemCd,"''","S")
UNIValue(1, 0) = FilterVar(strPlantCd,"''","S")
UNIValue(2, 0) = FilterVar(strRoutingNo,"''","S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

Response.Write "<Script Language=VBScript>" & vbCrLf

If rs0.EOF And rs0.BOF Then
	Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf						'��: ȭ�� ó�� ASP �� ��Ī�� 
Else
	Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf		'��: ȭ�� ó�� ASP �� ��Ī�� 
End If
	
If rs1.EOF And rs1.BOF Then
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf					'��: ȭ�� ó�� ASP �� ��Ī�� 
Else
	Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf				'��: ȭ�� ó�� ASP �� ��Ī�� 
End If
	
If rs2.EOF And rs2.BOF Then
	Response.Write "parent.frm1.txtRoutingNm.value = """"" & vbCrLf				'��: ȭ�� ó�� ASP �� ��Ī�� 
Else
	Response.Write "parent.frm1.txtRoutingNm.value = """ & ConvSPChars(rs2(0)) & """" & vbCrLf				'��: ȭ�� ó�� ASP �� ��Ī�� 

End If
Response.Write "</Script>" & vbCrLf

rs0.Close
rs1.Close
rs2.Close
		
Set rs0 = Nothing
Set rs1 = Nothing
Set rs2 = Nothing

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
									'��: ActiveX Data Factory Object Nothing

lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
    Dim iCnt
    Dim iRCnt
    Dim iStr
    Dim ColCnt
     
    iCnt = 0
    lgstrData = ""
	
	ReDim TmpBuffer(0)
	
    If Len(Trim(lgStrPrevKey)) Then                                              '�� : Chnage str into int
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                         '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '��¥ 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' �ݾ� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
               Case "F3"  '���� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
               Case "F4"  '�ܰ� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
               Case "F5"   'ȯ�� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)
               Case Else
                    iStr = iStr & Chr(11) & rs0(ColCnt) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
			ReDim Preserve TmpBuffer(iRCnt)
            TmpBuffer(iRCnt) = iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotalStr = Join(TmpBuffer, "")
	
    If  iRCnt < lgMaxCount Then                                     '��: Check if next data exists
        lgStrPrevKey = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
    Set lgADF = Nothing                                             '��: Release ADF

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,4)

    UNISqlId(0) = "181501saa"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = UCase(Trim(strPlantCd))    '---������ 
    UNIValue(0,2) = UCase(Trim(strItemCd))
    UNIValue(0,3) = UCase(Trim(strRoutingNo))
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("181500", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strPlantCd     = FilterVar(Request("txtPlantCd"),"' '","S")                     '�����ڵ� 
     strItemCd     = FilterVar(Request("txtItemCd"),"' '","S")            'ǰ���ڵ� 
     strRoutingNo     = FilterVar(Request("txtRoutingNo"),"' '","S")            '����ù�ȣ 
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub


%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowDataByClip "<%=ConvSPChars(iTotalStr)%>"                            '��: Display data 
         .lgStrPrevKey        =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
         .DbQueryOk
	End with
</Script>	

