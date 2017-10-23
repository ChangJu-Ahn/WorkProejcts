<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@page import="java.beans.XMLEncoder"%>
<%@ page import="org.w3c.dom.*"%>
<%@ page import="org.xml.sax.*"%>
<%@ page import="java.util.*"%>
<%@ page import="java.io.*"%>
<%@ page import="java.net.*"%>
<%@ page import="javax.xml.parsers.*"%>
<%@ page import="javax.servlet.http.HttpServletResponse.*"%>
<%
	String geturl = "http://asp1.krx.co.kr/servlet/krx.asp.XMLJemu?code=033640";
	String gettime = "";

	String DaeCha_ym[] = new String[6];
	String SonIk_ym[] = new String[6];
	String CashFlow_ym[] = new String[6];

	String DaeCha_data[][] = new String[10][10];
	String SonIk_data[][] = new String[6][10];
	String CashFlow_data[][] = new String[6][7];

	int daeCha_data_length = 0;
	int sonIk_data_length = 0;
	int cashFlow_data_length = 0;

	String line = "";
	String xml = "";

	try {
		URL url = new URL(geturl);
		URLConnection conn = url.openConnection();
		HttpURLConnection httpConnection = (HttpURLConnection) conn;
		InputStream is = null;
		InputStreamReader isr = null;

		is = new URL(geturl).openStream();
		isr = new InputStreamReader(is, "euc-kr");

		BufferedReader rd = new BufferedReader(isr, 400);

		StringBuffer strbuf = new StringBuffer();

		while ((line = rd.readLine()) != null) {

			strbuf.append(line);
		}

		//System.out.println("재무정보");
		//System.out.println(strbuf.toString());

		DocumentBuilderFactory docFact = DocumentBuilderFactory.newInstance();
		docFact.setNamespaceAware(true);
		DocumentBuilder docBuild = docFact.newDocumentBuilder();

		Document doc = docBuild.parse(new InputSource(new StringReader(strbuf.toString())));
		doc.getDocumentElement().normalize();

		Element root = doc.getDocumentElement();

		/*조회시각*/
		NodeList financialTotal = doc.getElementsByTagName("financialTotal");

		NamedNodeMap stockinfo = financialTotal.item(0).getAttributes();
		gettime = stockinfo.getNamedItem("querytime").getNodeValue();

		/*대차대조표*/

		NodeList TBL_DaeCha_ym = doc.getElementsByTagName("TBL_DaeCha");
		NamedNodeMap daeCha_ym = TBL_DaeCha_ym.item(0).getAttributes();

		DaeCha_ym[0] = daeCha_ym.getNamedItem("year0").getNodeValue();
		DaeCha_ym[1] = daeCha_ym.getNamedItem("month0").getNodeValue();
		DaeCha_ym[2] = daeCha_ym.getNamedItem("year1").getNodeValue();
		DaeCha_ym[3] = daeCha_ym.getNamedItem("month1").getNodeValue();
		DaeCha_ym[4] = daeCha_ym.getNamedItem("year2").getNodeValue();
		DaeCha_ym[5] = daeCha_ym.getNamedItem("month2").getNodeValue();

		NodeList TBL_DaeCha_data = doc.getElementsByTagName("TBL_DaeCha_data");

		daeCha_data_length = TBL_DaeCha_data.getLength();

		for (int i = 0; i < daeCha_data_length; i++) {

			NamedNodeMap daeCha_data = TBL_DaeCha_data.item(i).getAttributes();

			DaeCha_data[i][0] = daeCha_data.getNamedItem("hangMok" + i).getNodeValue();
			DaeCha_data[i][1] = daeCha_data.getNamedItem("year1Money" + i).getNodeValue();
			DaeCha_data[i][2] = daeCha_data.getNamedItem("year1GuSungRate" + i).getNodeValue();
			DaeCha_data[i][3] = daeCha_data.getNamedItem("year1JungGamRate" + i).getNodeValue();
			DaeCha_data[i][4] = daeCha_data.getNamedItem("year2Money" + i).getNodeValue();
			DaeCha_data[i][5] = daeCha_data.getNamedItem("year2GuSungRate" + i).getNodeValue();
			DaeCha_data[i][6] = daeCha_data.getNamedItem("year2JungGamRate" + i).getNodeValue();
			DaeCha_data[i][7] = daeCha_data.getNamedItem("year3Money" + i).getNodeValue();
			DaeCha_data[i][8] = daeCha_data.getNamedItem("year3GuSungRate" + i).getNodeValue();
			DaeCha_data[i][9] = daeCha_data.getNamedItem("year3JungGamRate" + i).getNodeValue();

		}

		/*손익계산서*/

		NodeList TBL_SonIk_ym = doc.getElementsByTagName("TBL_SonIk");
		NamedNodeMap sonIk_ym = TBL_SonIk_ym.item(0).getAttributes();

		SonIk_ym[0] = daeCha_ym.getNamedItem("year0").getNodeValue();
		SonIk_ym[1] = daeCha_ym.getNamedItem("month0").getNodeValue();
		SonIk_ym[2] = daeCha_ym.getNamedItem("year1").getNodeValue();
		SonIk_ym[3] = daeCha_ym.getNamedItem("month1").getNodeValue();
		SonIk_ym[4] = daeCha_ym.getNamedItem("year2").getNodeValue();
		SonIk_ym[5] = daeCha_ym.getNamedItem("month2").getNodeValue();

		NodeList TBL_SonIk_data = doc.getElementsByTagName("TBL_SonIk_data");

		sonIk_data_length = TBL_SonIk_data.getLength();

		for (int i = 0; i < sonIk_data_length; i++) {

			NamedNodeMap sonIk_data = TBL_SonIk_data.item(i).getAttributes();

			SonIk_data[i][0] = sonIk_data.getNamedItem("hangMok" + i).getNodeValue();
			SonIk_data[i][1] = sonIk_data.getNamedItem("year1Money" + i).getNodeValue();
			SonIk_data[i][2] = sonIk_data.getNamedItem("year1GuSungRate" + i).getNodeValue();
			SonIk_data[i][3] = sonIk_data.getNamedItem("year1JungGamRate" + i).getNodeValue();
			SonIk_data[i][4] = sonIk_data.getNamedItem("year2Money" + i).getNodeValue();
			SonIk_data[i][5] = sonIk_data.getNamedItem("year2GuSungRate" + i).getNodeValue();
			SonIk_data[i][6] = sonIk_data.getNamedItem("year2JungGamRate" + i).getNodeValue();
			SonIk_data[i][7] = sonIk_data.getNamedItem("year3Money" + i).getNodeValue();
			SonIk_data[i][8] = sonIk_data.getNamedItem("year3GuSungRate" + i).getNodeValue();
			SonIk_data[i][9] = sonIk_data.getNamedItem("year3JungGamRate" + i).getNodeValue();

		}

		/*현금흐름표*/

		NodeList TBL_CashFlow = doc.getElementsByTagName("TBL_CashFlow");
		NamedNodeMap cashFlow_ym = TBL_CashFlow.item(0).getAttributes();

		CashFlow_ym[0] = cashFlow_ym.getNamedItem("year0").getNodeValue();
		CashFlow_ym[1] = cashFlow_ym.getNamedItem("month0").getNodeValue();
		CashFlow_ym[2] = cashFlow_ym.getNamedItem("year1").getNodeValue();
		CashFlow_ym[3] = cashFlow_ym.getNamedItem("month1").getNodeValue();
		CashFlow_ym[4] = cashFlow_ym.getNamedItem("year2").getNodeValue();
		CashFlow_ym[5] = cashFlow_ym.getNamedItem("month2").getNodeValue();

		NodeList TBL_CashFlow_data = doc.getElementsByTagName("TBL_CashFlow_data");

		cashFlow_data_length = TBL_CashFlow_data.getLength();

		for (int i = 0; i < cashFlow_data_length; i++) {

			NamedNodeMap cashFlow_data = TBL_CashFlow_data.item(i).getAttributes();

			CashFlow_data[i][0] = cashFlow_data.getNamedItem("hangMok" + i).getNodeValue();
			CashFlow_data[i][1] = cashFlow_data.getNamedItem("year1Money" + i).getNodeValue();
			CashFlow_data[i][2] = cashFlow_data.getNamedItem("year1JungGamRate" + i).getNodeValue();
			CashFlow_data[i][3] = cashFlow_data.getNamedItem("year2Money" + i).getNodeValue();
			CashFlow_data[i][4] = cashFlow_data.getNamedItem("year2JungGamRate" + i).getNodeValue();
			CashFlow_data[i][5] = cashFlow_data.getNamedItem("year3Money" + i).getNodeValue();
			CashFlow_data[i][6] = cashFlow_data.getNamedItem("year3JungGamRate" + i).getNodeValue();

		}

	} catch (Exception e) {

	}
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script type="text/javascript" src="js/common.js"></script>
<link rel="stylesheet" type="text/css" href="css/financialInfo.css" />
<title>nepes 재무정보</title>
</head>
<body>
	<div class="header-wrap">
		재무정보<span><span class="time_img"></span><%=gettime%> 기준</span>
	</div>
	<div class="body-wrap">
		<div class="data-lists">
			<dl>
				<dt>
					<span></span>대차대조표
				</dt>
				<dd>
					<table id="balance">
						<tr>
							<th rowspan="2">재무항목</th>
							<%
								if (!DaeCha_ym[0].equals("")) {
							%>
							<th colspan="3"><%=DaeCha_ym[0]%>년 <%=DaeCha_ym[1]%>월</th>
							<th colspan="3"><%=DaeCha_ym[2]%>년 <%=DaeCha_ym[3]%>월</th>
							<th colspan="3"><%=DaeCha_ym[4]%>년 <%=DaeCha_ym[5]%>월</th>
							<%
								}
							%>
						</tr>
						<tr>
							<th>금액</th>
							<th>구성비</th>
							<th>증감율</th>
							<th>금액</th>
							<th>구성비</th>
							<th>증감율</th>
							<th>금액</th>
							<th>구성비</th>
							<th>증감율</th>
						</tr>
						<%
							if (daeCha_data_length > 0) {
						%>
						<%
							for (int i = 0; i < daeCha_data_length; i++) {
						%>
						<tr>
							<td><%=DaeCha_data[i][0]%></td>
							<td><%=DaeCha_data[i][1]%></td>
							<td><%=DaeCha_data[i][2]%></td>
							<td><%=DaeCha_data[i][3]%></td>
							<td><%=DaeCha_data[i][4]%></td>
							<td><%=DaeCha_data[i][5]%></td>
							<td><%=DaeCha_data[i][6]%></td>
							<td><%=DaeCha_data[i][7]%></td>
							<td><%=DaeCha_data[i][8]%></td>
							<td><%=DaeCha_data[i][9]%></td>
						</tr>
						<%
							}
						%>
						<%
							}
						%>
					</table>
				</dd>
				<dt>
					<span></span>손익계산서
				</dt>
				<dd>
					<table id="income_statement">
						<tr>
							<th rowspan="2">재무항목</th>
							<%
								if (!SonIk_ym[0].equals("")) {
							%>
							<th colspan="3"><%=SonIk_ym[0]%>년 <%=SonIk_ym[1]%>월</th>
							<th colspan="3"><%=SonIk_ym[2]%>년 <%=SonIk_ym[3]%>월</th>
							<th colspan="3"><%=SonIk_ym[4]%>년 <%=SonIk_ym[5]%>월</th>
							<%
								}
							%>
						</tr>
						<tr>
							<th>금액</th>
							<th>구성비</th>
							<th>증감율</th>
							<th>금액</th>
							<th>구성비</th>
							<th>증감율</th>
							<th>금액</th>
							<th>구성비</th>
							<th>증감율</th>
						</tr>
						<%
							if (sonIk_data_length > 0) {
						%>
						<%
							for (int i = 0; i < sonIk_data_length; i++) {
						%>
						<tr>
							<td><%=SonIk_data[i][0]%></td>
							<td><%=SonIk_data[i][1]%></td>
							<td><%=SonIk_data[i][2]%></td>
							<td><%=SonIk_data[i][3]%></td>
							<td><%=SonIk_data[i][4]%></td>
							<td><%=SonIk_data[i][5]%></td>
							<td><%=SonIk_data[i][6]%></td>
							<td><%=SonIk_data[i][7]%></td>
							<td><%=SonIk_data[i][8]%></td>
							<td><%=SonIk_data[i][9]%></td>
						</tr>
						<%
							}
						%>
						<%
							}
						%>
					</table>
				</dd>
				<dt>
					<span></span>현금흐름표
				</dt>
				<dd>
					<table id="cash_flow">
						<tr>
							<th rowspan="2">재무항목</th>
							<%
								if (!CashFlow_ym[0].equals("")) {
							%>
							<th colspan="2"><%=CashFlow_ym[0]%>년 <%=CashFlow_ym[1]%>월</th>
							<th colspan="2"><%=CashFlow_ym[2]%>년 <%=CashFlow_ym[3]%>월</th>
							<th colspan="2"><%=CashFlow_ym[4]%>년 <%=CashFlow_ym[5]%>월</th>
							<%
								}
							%>
						</tr>
						<tr>
							<th>금액</th>
							<th>증감액</th>
							<th>금액</th>
							<th>증감액</th>
							<th>금액</th>
							<th>증감액</th>
						</tr>
						<%
							if (cashFlow_data_length > 0) {
						%>
						<%
							for (int i = 0; i < cashFlow_data_length; i++) {
						%>
						<tr>
							<td><%=CashFlow_data[i][0]%></td>
							<td><%=CashFlow_data[i][1]%></td>
							<td><%=CashFlow_data[i][2]%></td>
							<td><%=CashFlow_data[i][3]%></td>
							<td><%=CashFlow_data[i][4]%></td>
							<td><%=CashFlow_data[i][5]%></td>
							<td><%=CashFlow_data[i][6]%></td>
						</tr>
						<%
							}
						%>
						<%
							}
						%>
					</table>
				</dd>
			</dl>
		</div>
	</div>
	<div class="footer-wrap"></div>
</body>
</html>