<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : ADO Template (Query)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/16
'*  7. Modified date(Last)  : 2002/12/16
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%  
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "*", "NOCOOKIE", "QB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strShiftCd																'�� : Shift
Dim strShiftPlantCd
Dim strResourceCd
Dim strResourcePlantCd
Dim PosBreakFlg
Dim iOpt
Dim TmpBuffer
Dim iTotalStr

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
Call HideStatusWnd 


lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgMaxCount     = 30							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Orderby value

'On Error Resume Next
Err.Clear
    
Call GetColPos(lgSelectList,"WORK_FLG")
Call TrimData()
Call FixUNISQLData()
Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
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
	
	ReDim TmpBuffer(0)
	
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
                    'iStr = iStr & Chr(11) & rs0(ColCnt) 
                    If PosBreakFlg = ColCnt Then
						If UCase(Trim(rs0(ColCnt))) = "Y" Then
							iStr = iStr & Chr(11) & "OverTime"
						ElseIf UCase(Trim(rs0(ColCnt))) = "N" Then
							iStr = iStr & Chr(11) & "DownTime"
						Else
							iStr = iStr & Chr(11) & rs0(ColCnt)
						End If
					Else
						iStr = iStr & Chr(11) & rs0(ColCnt)
					End If
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
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,4)

    UNISqlId(0) = "181900sab"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strShiftPlantCd))
    UNIValue(0,2) = UCase(Trim(strResourcePlantCd))
    UNIValue(0,3) = UCase(Trim(strShiftCd))
    UNIValue(0,4) = UCase(Trim(strResourceCd))
    
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
        Call DisplayMsgBox(181900, vbOKOnly, "", "", I_MKSCRIPT)
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
     strShiftCd     = FilterVar(Request("txtShiftCd"),"' '","S")                   'Shift
     strResourcePlantCd     = FilterVar(Request("txtResourcePlantCd"),"' '","S")                   'ResourcePlantCd
     strShiftPlantCd     = FilterVar(Request("txtShiftPlantCd"),"' '","S")                   'ShiftPlantCd
     strResourceCd     = FilterVar(Request("txtResourceCd"),"' '","S")                   'ResourceCd
	 iOpt		   = Request("iOpt")
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

Sub GetColPos(ByVal sVal1,ByVal sVal2)
	Dim aCol
	Dim i
	Dim iCnt
	Dim iPos
	Dim sStr
	
	aCol = Split(sVal1,",")
	iCnt = Len(sVal2)

	For i = 0 To UBound(aCol)

		iPos = InStr(1,aCol(i), ".")
		
		If iPos = 0 Then
			sStr = aCol(i)
		Else
			sStr = Mid(aCol(i),iPos+1,Len(aCol(i)))
		End If
		
		If UCase(Trim(sStr)) = UCase(Trim(sVal2)) Then
			PosBreakFlg = i
			Exit Sub
		End If
	Next
	
	PosBreakFlg = -1
	
End Sub

%>

<Script Language=vbscript>

    With parent
         .ggoSpread.Source    = .frm1.vspdData2 
         .ggoSpread.SSShowDataByClip "<%=ConvSPChars(iTotalStr)%>"                            '��: Display data 
         .lgStrPrevKey_B      =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
         .DbQueryOk("<%=iOpt%>")
	End with
</Script>	

