'use strict'

Object.prototype.extend = function(obj) {
   for (var i in obj) {
	  if (obj.hasOwnProperty(i)) {
		 this[i] = obj[i];
	  }
   }
};


function ShImportInteract()
{
		
	this._getDocument = function(){
		var frame = window.frames['iframe'];
		if(frame)
			return frame.contentDocument.getElementsByTagName("frame")[0].contentDocument;
		else{
			return  document;
			}
	};
	
	
}

var constsExtentions = {
	_selectSelector : "select.standardText",
	_tbodySelector : ".defTable > tbody:nth-child(1) > tr:nth-child(11) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1)",
	_tdSelector : "tr[style] > td:nth-child(2)"
	
}


var selectElExtentions = {
	
		
		_selectEl : function(selector){
			return this._getDocument().querySelectorAll(selector);
		},
		_selectEl2 :function(el,selector)
		{
			return el.querySelectorAll(selector);
		},
		_selectElIndex : function(selector,index){
			return this._selectEl()[index];
		},
	
}


var selectShElExtentions={
	getImportSelectList:function ()
	{
		return this._selectEl(this._selectSelector);
	},
	getCaptionsForSelect:function()
	{
		var tbodySelector = ".defTable > tbody:nth-child(1) > tr:nth-child(11) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1)";
		var tdSelector = "tr[style] > td:nth-child(2)";
		var tbody = this._selectEl(this._tbodySelector)[0];
		var tdLabels = this._selectEl2(tbody,this._tdSelector);
		var labels =[];
		for (var i = 1; i < tdLabels.length; i++) {
			labels.push(tdLabels[i].innerText);
		}
		
		return labels;
		
		
	}
	
}

var GetSetSelectExtentions={
	_getSelIndex : function(el)
	{
		return el.selectedIndex;
	},
	_setSelIndex : function(el,val)
	{
		el.selectedIndex = val;
	},
	
}

var mainOperationsExtentions = {
	// увеличить значения в указанном промежутке на единицу
		
	
	_incrementN : function(start,end, incr=1 )
	{
		var selects = this.getImportSelectList();
		for (var i = 0; i < selects.length; i++) {
			 var el = selects[i];
			 var val = this._getSelIndex(el);
		     if((val>=start)&&(val<=end))
			 {
				 this._setSelIndex(el,val+=incr);
			 }
		}
	
	},
	
	_decrementN : function(start,end, incr=1 )
	{
		var selects = this.getImportSelectList();
		for (var i = 0; i < selects.length; i++) {
			 var el = selects[i];
			 var val = this._getSelIndex(el);
			 if(val>=0)
			 {
				 if((val>=start)&&(val<=end))
				 {
					 this._setSelIndex(el,val-=incr);
				 }
			 }
		}
	
	},
	increment : function(start,end, incr=1 )
	{
		var sRow= this._columnToNumber(start);
		var eRow = this._columnToNumber(end);
		this._incrementN(sRow,eRow);
		this.log();
	},
	decrement : function(start,end, incr=1 )
	{
		var sRow= this._columnToNumber(start);
		var eRow = this._columnToNumber(end);
		this._decrementN(sRow,eRow);
		this.log();
	},
	clear : function(){
		var selects = this.getImportSelectList();
		for (var i = 0; i < selects.length; i++) {
			var el = selects[i];
			this._setSelIndex(el,-1);
		}
		
	}
	
}

var excelColumnMapExtentions = {
	_columnToNumber : function(string)
	{
	string = string.toUpperCase();
    var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', sum = 0, i;
    for (i = 0; i < string.length; i++) {
        sum += Math.pow(letters.length, i) * (letters.indexOf(string.substr(((i + 1) * -1), 1)) + 1);
    }
    return --sum; // align to select
	},
	_numberToColumn : function(num)
	{
		num++; // aling to select 
		for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
			ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
		}
		return ret;
	}
}

var logExtention = {
	log:function(){
		var selects = this.getImportSelectList();
		var labels = this.getCaptionsForSelect();
		var result = [];
		 for (var i = 0; i < selects.length; i++) {
			var select = selects[i];
			var selectVal = this._getSelIndex(select);
			var selectCol = this._numberToColumn(selectVal);
			var label = labels[i];
			var log = {del1:'|', label:label, del2:'|', column:selectCol, del3:'|', index:selectVal};
			result.push(log);
			
		 }
		 console.table(result);
	},
	
	log2:function(){
		var selects = this.getImportSelectList();
		var labels = this.getCaptionsForSelect();
		var result = [];
		 for (var i = 0; i < selects.length; i++) {
			var select = selects[i];
			var selectVal = this._getSelIndex(select);
			var selectCol = this._numberToColumn(selectVal);
			var label = labels[i];
			var log = {label:label, column:selectCol};
			console.log(log);
			
		 }
		
	}
	
}

var interact = new ShImportInteract();
interact.extend(constsExtentions);
interact.extend(selectElExtentions);
interact.extend(selectShElExtentions);
interact.extend(GetSetSelectExtentions);
interact.extend(excelColumnMapExtentions);
interact.extend(logExtention);
interact.extend(mainOperationsExtentions);

//interact.increment(1,7);
// interact._numberToColumn(27); //AA
// interact._columnToNumber("AA"); //27
// interact.increment('L','M');
//interact.getCaptionsForSelect();
interact.log();






