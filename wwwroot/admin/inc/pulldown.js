<!--
var timeOn = null;
var stack = new Array();
var stackint = 0;

function myPush(i) {
	stack[stackint] = i;
	stackint = stackint + 1;
}

function myPop() {
	stackint = stackint - 1;
	return stack[stackint];
}

if (document.layers != null || is_nav6 || is_nav6up) {
	layerStyleRef="layer.";
	layerRef="document.layers";
	styleSwitch="";
} 
	
if (document.all != null) {
	layerStyleRef="layer.style.";
	layerRef="document.all";
	styleSwitch=".style";
}

function showLayer(layerName) {

	var w,l = 0;
	
	if (is_ie5 || is_ie4)
		myPush(layerName);
	else
		stack.push(layerName);
	
	if (timeOn != null) {
		clearTimeout(timeOn);
	}
	
	if (is_nav4) {
		w = window.outerWidth;
		eval('l = ((w - 781) / 2) + posx.' + layerName + ';');
		eval('document.' + layerName + '.left = l - 12;');
		eval('document.' + layerName + '.top = posy.' + layerName + ';');
		eval('document.' + layerName + '.visibility="visible";');
	}
	
	if (is_ie4 || is_ie5 || is_ie5up || is_opera) {
		w = document.body.clientWidth;
			
		eval('l = ((w - 781) / 2) + posx.' + layerName + ';');
		eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.left=' + l);
		eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.top=posy.'+layerName);
		eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="visible"');
	}
	
	if (is_nav6 || is_nav6up || is_gecko) {
		var elm
		elm = document.getElementById(layerName);
		elm.style.visibility='hidden';
		w = window.outerWidth;
		eval('l = ((w - 781) / 2) + posx.' + layerName + ';');

		elm.style.left = l - 12;
		
		elm.style.top=posy[layerName] - 8;
		elm.style.visibility='visible';
	}
	
}

//Het verbergen van de submenu-dropdown-layer..
function hideLayer(layerName){
	
	if (layerName=='') {
		return;
	}
	
	if (is_nav4) {
		eval('document.' + layerName + '.visibility="hidden";');	
	} 
	
	if (is_ie4 || is_ie5 || is_ie5up || is_opera) {
		eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="hidden"');
	}
	
	if (is_nav6 || is_nav6up) {
		document.getElementById(layerName).style.visibility='hidden';
	}
}

function hideAllLayers() {
	if (is_ie4 || is_ie5) 
		x = myPop();
	else
		x = stack.pop();

	while (x != null) {
		hideLayer(x);
		if (is_ie4 || is_ie5) 
			x = myPop();
		else
			x = stack.pop();
	}
}

function hideChildLayers(layerName, except_child) {
	var j;
	eval ('j = child' + layerName + '.length;')
	for (var i = 0; i < j; i++) {
		var x;
		eval ('x = child' + layerName + '[' + i + ']')
		if (x != except_child) {
			hideLayer(x);
			// check children
			var k;
			eval ('k = child' + x + '.length;')
			if (k > 0)
				hideChildLayers(x, except_child);
		}
	}
}

//Timer voor het verlaten van de submenu-dropdown-layer met de muiscursor..
function btnTimer() {
	timeOn = setTimeout("btnOut()",1000)
}

//Het verlaten van de hoofdmenu-knop..
function btnOut(layerName) {
	hideAllLayers();
}

//Het onmouseover-event van de submenu-dropdown-layer..
function menuOver() {
	clearTimeout(timeOn)
}

//Het onmouseout-event van de submenu-dropdown-layer..
function menuOut(itemName) {
	timeOn = setTimeout("hideAllLayers()", 400)
}
 
function swap(imagename, swapname) {
	imagename.src=swapname.src;
}

// -->
