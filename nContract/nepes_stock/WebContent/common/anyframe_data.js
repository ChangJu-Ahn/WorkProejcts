/*  anyFrame JavaScript Data, version 1.0.0
 *  (c) 2007 CyberImagination <http://www.cyber-i.com>
 *
/*--------------------------------------------------------------------------*/

/**
 * anyFRAME 의 DataSet 과 유사
 * 
 * 
 * @constructor
 * @author 강상미, 수정:김성권
 */
function DataSet() {

	this.mapObj = new Object();
	this.keyArray = new Array();

	function put(key, value, seq) {
		if (!this.mapObj[key]) {
			this.mapObj[key] = new Array();
			this.keyArray[this.keyArray.length] = key;

		}
		if (!seq) {
			seq = 0;
		}

		this.mapObj[key][seq] = value;

	}

	function get(key, seq) {
		if (!this.mapObj[key])
			return '';
		if (!seq) {
			seq = 0;
		}
		var val = this.mapObj[key][seq];
		if (!val)
			return '';

		return val;

	}
	
	function getCount(key) {
		if (!this.mapObj[key])
			return 0;

		var array = this.mapObj[key];
		if (!array) {
			return 0;
		}

		return array.length;
	}

	function getDataSetCount() {
		return this.mapObj[keyArray[keyArray.length - 1]].length;
	}

	function getKeyCount() {
	    return this.keyArray.length;
	}

	function getKey(seq) {
	    return this.keyArray[seq];
	}	
	
	function getParam() {
		var str = '';
		var j = 0;
		for ( var i = 0; i < this.keyArray.length; i++) {
			var key = this.keyArray[i];
			var valArr = this.mapObj[key];
			var valCount = valArr.length;
			for (j = 0; j < valCount; j++) {

				if (str.length > 0)
					str += '&';
				str += key + '=' + valArr[j];

			}
		}
		// alert(str);
		return str;
	}
	
	function clear() {
		this.mapObj = new Object();
	}

	this.put = put;
	this.get = get;
	this.getCount = getCount;
	this.getParam = getParam;
	this.getDataSetCount = getDataSetCount;
	this.getKeyCount = getKeyCount;		
	this.getKey = getKey;		
}