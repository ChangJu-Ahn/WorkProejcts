<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4212mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2001.11.16
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : Park, BumSoo
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                  '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ�
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next                                '��:

Dim ADF                                                     'ActiveX Data Factory ���� ��������
Dim strRetMsg                                               'Record Set Return Message ��������
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1          'DBAgent Parameter ����

Const C_SHEETMAXROWS = 50

Dim strMode                                         '��: ���� MyBiz.asp �� ������¸� ��Ÿ��

'=======================================================================================================
'   �Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'   ����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ����
'   uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ����
'   ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgNextFlg
Dim i

Call HideStatusWnd

On Error Resume Next

Dim strPlantCd
Dim strProdOrdNo

    '=======================================================================================================
    '   ����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
    '=======================================================================================================
    Redim UNISqlId(0)
    Redim UNIValue(0, 2)

    UNISqlId(0) = "p4112mb6_lko391"

    strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
    strProdOrdNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")

    UNIValue(0, 0) = "^"
    UNIValue(0, 1) = strPlantCd
    UNIValue(0, 2) = strProdOrdNo

    UNILock = DISCONNREAD : UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    '-------------------------------------------
    ' Display Spread 2
    '-------------------------------------------
    If rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing
        Response.End
    End If

%>
<Script Language=vbscript>
    Dim LngMaxRow
    Dim strData
    Dim TmpBuffer
    Dim iTotalStr

    Dim tempWcCd
    tempWcCd = ""

    With parent                                                                 '��: ȭ�� ó�� ASP �� ��Ī��
        LngMaxRow = .frm1.vspdData2.MaxRows                                     'Save previous Maxrow
        ReDim TmpBuffer(<%=rs0.RecordCount-1%>)
<%
        For i=0 to rs0.RecordCount-1

%>
            strData = ""
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Cd"))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Nm"))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MACHINE_CD"))%>"     '2008-03-25 1:45���� :: hanc
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MACHINE_NM"))%>"     '2008-03-25 1:45���� :: hanc
            strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("Item_Cd"))%>")
            strData = strData & Chr(11) & ""
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
            strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Req_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"            '��: Required Quantity
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"                                  '��: Base Unit
            strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Issued_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"         '��: Required Quantity
            strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Req_Dt"))%>"                             '��: Required Date
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"                                    '��: Tracking No.
'20080118::hanc            strData = strData & Chr(11) & "<%=ConvSPChars(mid(rs0("Wc_cd"),1,6))%>"                                      '��: Storage Location Code
            strData = strData & Chr(11) & "<%=ConvSPChars(mid(rs0("sl_cd"),1,6))%>"                                      '��: Storage Location Code
            strData = strData & Chr(11) & ""                                                                        '��: Storage Location Popup
'20080118::hanc            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Nm"))%>"                                      '��: Storage Location Name
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_Nm"))%>"                                      '��: Storage Location Name
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Resv_Status"))%>"                                    '��: Reserve Status
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Resv_Desc"))%>"                                  '��: Reserve Status
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd"))%>"                                 '��: Issue Method
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd_Desc"))%>"                                    '��: Issue Method
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("req_no"))%>"                                    '��: Issue Method
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("seq"))%>"                                    '��: Issue Method
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"                                    '��: Issue Method
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("order_status"))%>"                                    '��: Issue Method
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("inside_flg"))%>"                                    '��: Issue Method

            strData = strData & Chr(11) & LngMaxRow + <%=i%>
            strData = strData & Chr(11) & Chr(12)

            TmpBuffer(<%=i%>) = strData
<%
            rs0.MoveNext
        Next
%>
        iTotalStr= Join(TmpBuffer, "")
        .ggoSpread.Source = .frm1.vspdData2
        .ggoSpread.SSShowDataByClip iTotalStr

<%
        rs0.Close
        Set rs0 = Nothing

%>

    .DbDtlQueryOk(LngMaxRow+1)

    End With
</Script>
<%

    Set ADF = Nothing                                               '��: ActiveX Data Factory Object Nothing
%>
