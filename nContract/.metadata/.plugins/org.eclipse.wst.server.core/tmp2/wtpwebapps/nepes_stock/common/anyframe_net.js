/*  anyFrame JavaScript Network, version 1.0.0
 *  (c) 2007 CyberImagination <http://www.cyber-i.com>
 *
/*--------------------------------------------------------------------------*/

//var contextPath = "/anyframe4";
var contextPath = "";
var WebServerURL = "http://" + window.location.host + contextPath;


function xc() {

	var C = null;
	try {
		C = new ActiveXObject("Msxml2.XMLHTTP")
	} catch (e) {
		try {
			C = new ActiveXObject("Microsoft.XMLHTTP")
		} catch (sc) {
			C = null
		}
	}

	if (!C && typeof XMLHttpRequest != "undefined") {
		C = new XMLHttpRequest()
	}
	return C
}

var l = null;

function sndHttp(url, output) {
	try {
		if (l && l.readyState != 0) {
			l.abort()
		}
	} catch (e) {
	}
	l = xc();
	if (l) {
		l.open("GET", url, false);
		l.send(null);

		var doc = l.responseXML;

		var nodeList = doc.getElementsByTagName("result");
		var chNdLst = nodeList.item(0).childNodes;
		var idx = 0;
		var preNodeName = '';

		for (i = 0; i < chNdLst.length; i++) {
			var nd = chNdLst.item(i);
			var ndNm = nd.nodeName;
			if (preNodeName != ndNm && ndNm != '#text') {
				idx = 0;
				preNodeName = ndNm;
			}

			if (nd.nodeType != nd.TEXT_NODE) {

				var cchNdLst = nd.childNodes;
				for (y = 0; y < cchNdLst.length; y++) {
					var elmnt = cchNdLst.item(y)
					var key = elmnt.nodeName;
					var val = (elmnt.firstChild == null) ? ""
							: elmnt.firstChild.nodeValue;
					output.put(key, val, idx);
					if ('city_name' == key && i < 10)
						window.status = 'key=' + key + ' val=' + val + ' idx='
								+ idx;
				}
				idx++;

			}
		}
	}
}

function sndHttpPost(url, val, output) {
	try {
		if (l && l.readyState != 0) {
			l.abort()
		}
	} catch (e) {
	}
	l = xc();
	if (l) {
		l.open("POST", url, false);
		l.setRequestHeader("Accept-Language", "ko");
		l.setRequestHeader("Content-type:", "text/html; charset=euc_kr");
		l.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		l.setRequestHeader("Content-length", val.length);
		l.send(val);

		var doc = l.responseXML;
		var nodeList = doc.getElementsByTagName("result");
		var chNdLst = nodeList.item(0).childNodes;
		var idx = 0;
		for (i = 0; i < chNdLst.length; i++) {
			var nd = chNdLst.item(i);
			if (nd.nodeType != nd.TEXT_NODE) {
				var cchNdLst = chNdLst.item(i).childNodes;
				for (y = 0; y < cchNdLst.length; y++) {
					var nN = cchNdLst.item(y).nodeName;
					var nV = (cchNdLst.item(y).firstChild == null) ? ""
							: cchNdLst.item(y).firstChild.nodeValue;
					output.put(nN, nV, idx);
				}
				idx++;
			}
		}
	}
}

var errMsg;

function InteractionBean() {
	function execute(tr, input, output) {

		var param = input.getParam();
		if (output == null)
			output = new DataSet();
		var xmlhttp;
		var url = WebServerURL + "/anylogic/process/" + tr + ".xml?"
				+ Math.random() + "&" + param;
		try {

			// ///////////////////////////////////////////////////////////////////////////////////////
			sndHttp(url, output);
			return output;
		} catch (e) {
			alert('URL : ' + url);
			errMsg = 'Interaction execute error.' + '\n* TR:' + tr + ' '
					+ e.message;
			if (xmlhttp != null)
				errMsg = errMsg + "\n\n* HTTP RESPONSE\n"
						+ xmlhttp.ResponseText;
			window.setTimeout("alertMsg();", 500);
			// window.status=errMsg;
		}
		window.status = 'communication complete.';
		return output;
	}

	function executePost(tr, input, output) {

		var param = input.getParam();
		alert(param);
		if (output == null)
			var output = new DataSet();
		var xmlhttp;
		var url = WebServerURL + "/anylogic/process/" + tr + ".xml";
		var val = param + "&" + Math.random();
		try {
			// ///////////////////////////////////////////////////////////////////////////////////////
			// alert("#url======>"+url+"\n#val======>"+val);
			sndHttpPost(url, val, output);
			return output;
		} catch (e) {
			alert('URL : ' + url);
			errMsg = 'Interaction execute error.' + '\n* TR:' + tr + ' '
					+ e.message;
			if (xmlhttp != null)
				errMsg = errMsg + "\n\n* HTTP RESPONSE\n"
						+ xmlhttp.ResponseText;
			window.setTimeout("alertMsg();", 500);
			// window.status=errMsg;
		}
		window.status = 'communication complete.';
		return output;
	}

	function executeEncodePost(tr, input, output) {

		var param = input.getEncodeParam();
		alert(param);
		if (output == null)
			output = new DataSet();
		var xmlhttp;
		var url = WebServerURL + "/anylogic/process/" + tr + ".xml";
		var val = param + "&" + Math.random();
		try {
			// ///////////////////////////////////////////////////////////////////////////////////////
			// alert("#url======>"+url+"\n#val======>"+val);
			sndHttpPost(url, val, output);
			return output;
		} catch (e) {
			alert('URL : ' + url);
			errMsg = 'Interaction execute error.' + '\n* TR:' + tr + ' '
					+ e.message;
			if (xmlhttp != null)
				errMsg = errMsg + "\n\n* HTTP RESPONSE\n"
						+ xmlhttp.ResponseText;
			window.setTimeout("alertMsg();", 500);
			// window.status=errMsg;
		}
		window.status = 'communication complete.';
		return output;
	}

	this.execute = execute;
	this.executePost = executePost;
}

function alertMsg() {
	alert(errMsg);
}