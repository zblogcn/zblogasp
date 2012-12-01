<%
Function AntiXSS_VbsTrim(s)
	AntiXSS_VbsTrim=Trim(s)
End Function

%>
<script language="javascript" runat="server">
//原GITHUB：https://github.com/leizongmin/js-xss/blob/master/index.js
//过滤XSS攻击 @author 老雷<leizongmin@gmail.com>
//转换到ASP by zsx(http://www.zsxsoft.com)
function AntiXSS_run(html){
	String.prototype.trim=function(){return AntiXSS_VbsTrim(this)};
	return AntiXSS(html,AntiXSS_config);
}
var AntiXSS_noTag = function(text) {
	return text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
};
function AntiXSS(html, options) {
	var whiteList = options.whiteList;
	var onTagAttr = options.onTagAttr;
	var onIgnoreTag = options.onIgnoreTag;
	var rethtml = '';
	var lastPos = 0;
	var tagStart = false;
	var quoteStart = false;
	var currentPos = 0;
	var filterAttributes = function(tagName, attrs) {
		tagName = tagName.toLowerCase();
		var whites = whiteList[tagName];
		var lastPos = 0;
		var _attrs = [];
		var tmpName = false;
		var hasSprit = false;
		var addAttr = function(name, value) {
			name = name.trim();
			if (!hasSprit && name === '/') {
				hasSprit = true;
				return;
			};
			name = name.replace(/[^a-zA-Z0-9_:\.\-]/img, '').toLowerCase();
			if (name.length < 1) return;
			if (whites.join().indexOf(name) !== -1) {
				if (value) {
					value = value.trim().replace(/"/g, '&quote;');
					value = value.replace(/&#([a-zA-Z0-9]*);?/img,
					function(str, code) {
						code = parseInt(code);
						return String.fromCharCode(code);
					});
					var _value = '';
					for (var i = 0, len = value.length; i < len; i++) {
						_value += value.charCodeAt(i) < 32 ? ' ': value.split("")[i];
					}
					value = _value.trim();
					var newValue = onTagAttr(tagName, name, value);
					if (typeof(newValue) !== 'undefined') {
						value = newValue;
					}
				}
				_attrs.push(name + (value ? '="' + value + '"': ''));
			}
		};
		for (var i = 0, len = attrs.length; i < len; i++) {
			var c = attrs.split("")[i];
			if (tmpName === false && c === '=') {
				tmpName = attrs.slice(lastPos, i);
				lastPos = i + 1;
				continue;
			}
			if (tmpName !== false) {
				if (i === lastPos && (c === '"' || c === "'")) {
					var j = attrs.indexOf(c, i + 1);
					if (j === -1) {
						break;
					} else {
						var v = attrs.slice(lastPos + 1, j).trim();
						addAttr(tmpName, v);
						tmpName = false;
						i = j;
						lastPos = i + 1;
						continue;
					}
				}
			}
			if (c === ' ') {
				var v = attrs.slice(lastPos, i).trim();
				if (tmpName === false) {
					addAttr(v);
				} else {
					addAttr(tmpName, v);
				}
				tmpName = false;
				lastPos = i + 1;
				continue;
			}
		}
		if (lastPos < attrs.length) {
			if (tmpName === false) {
				addAttr(attrs.slice(lastPos));
			} else {
				addAttr(tmpName, attrs.slice(lastPos));
			}
		}
		if (hasSprit) _attrs.push('/');
		return _attrs.join(' ');
	};
	var addNewTag = function(tag, end) {
		rethtml += AntiXSS_noTag(html.slice(lastPos, tagStart));
		lastPos = end + 1;
		var spos = tag.slice(0, 2) === '</' ? 2: 1;
		var i = tag.indexOf(' ');
		if (i === -1) {
			var tagName = tag.slice(spos, tag.length - 1).trim();
		} else {
			var tagName = tag.slice(spos, i + 1).trim();
		}
		tagName = tagName.toLowerCase();
		if (tagName in whiteList) {
			if (i === -1) {
				rethtml += tag.slice(0, spos) + tagName + '>';
			} else {
				var attrs = filterAttributes(tagName, tag.slice(i + 1, tag.length - 1).trim());
				rethtml += tag.slice(0, spos) + tagName + (attrs.length > 0 ? ' ' + attrs: '') + '>';
			}
		} else {
			var options = {
				isClosing: (spos === 2),
				position: rethtml.length,
				originalPosition: currentPos - tag.length + 1
			};
			var tagHtml = onIgnoreTag(tagName, tag, options);
			if (typeof(tagHtml) === 'undefined') {
				tagHtml = AntiXSS_noTag(tag);
			}
			rethtml += tagHtml;
		}
	};
	for (var currentPos = 0, len = html.length; currentPos < len; currentPos++) {
		var c = html.split("")[currentPos];
		if (tagStart === false) {
			if (c === '<') {
				tagStart = currentPos;
				continue;
			}
		} else {
			if (quoteStart === false) {
				if (c === '<') {
					rethtml += AntiXSS_noTag(html.slice(lastPos, currentPos));
					tagStart = currentPos;
					lastPos = currentPos;
					continue;
				}
				if (c === '>') {
					addNewTag(html.slice(tagStart, currentPos + 1), currentPos);
					tagStart = false;
					continue;
				}
				if (c === '"' || c === "'") {
					quoteStart = c;
					continue;
				}
			} else {
				if (c === quoteStart) {
					quoteStart = false;
					continue;
				}
			}
		}
	}
	if (lastPos < html.length) {
		rethtml += AntiXSS_noTag(html.substr(lastPos));
	}
	return rethtml;
};
var AntiXSS_config = {
	"whiteList": {
//		h1: ['style', 'class'],
//		h2: ['style', 'class'],
//		h3: ['style', 'class'],
//		h4: ['style', 'class'],
//		h5: ['style', 'class'],
//		h6: ['style', 'class'],
		hr: ['style', 'class'],
		span: ['style', 'class'],
		strong: ['style', 'class'],
		b: ['style', 'class'],
		i: ['style', 'class'],
		br: [],
		p: ['style', 'class'],
		pre: ['style', 'class'],
		code: ['style', 'class'],
		a: ['style', 'class', 'target', 'href', 'title' ,'rel'],
		img: ['style', 'class', 'src', 'alt', 'title'],
		div: ['style', 'class'],
		table: ['style', 'class', 'width', 'border'],
		tr: ['style', 'class'],
		td: ['style', 'class', 'width', 'colspan'],
		th: ['style', 'class', 'width', 'colspan'],
		tbody: ['style', 'class'],
		ul: ['style', 'class'],
		li: ['style', 'class'],
		ol: ['style', 'class'],
		dl: ['style', 'class'],
		dt: ['style', 'class'],
		em: ['style'],
//		cite: ['style'],
//		section: ['style', 'class'],
//		header: ['style', 'class'],
//		footer: ['style', 'class'],
		blockquote: ['style', 'class']//,
//		audio: ['autoplay', 'controls', 'loop', 'preload', 'src'],
//		video: ['autoplay', 'controls', 'loop', 'preload', 'src', 'height', 'width']
	},
	"onTagAttr": function(tag, attr, value) {
		if (attr === 'href' || attr === 'src') {
			if (/\/\*|\*\//mg.test(value)) {
				return '#';
			}
			if (/^[\s"'`]*((j\s*a\s*v\s*a|v\s*b|l\s*i\s*v\s*e)\s*s\s*c\s*r\s*i\s*p\s*t\s*|m\s*o\s*c\s*h\s*a):/ig.test(value)) {
				return '#';
			}
		} else if (attr === 'style') {
			if (/\/\*|\*\//mg.test(value)) {
				return '#';
			}
			if (/((j\s*a\s*v\s*a|v\s*b|l\s*i\s*v\s*e)\s*s\s*c\s*r\s*i\s*p\s*t\s*|m\s*o\s*c\s*h\s*a):/ig.test(value)) {
				return '';
			}
		}
	},
	"onIgnoreTag": function(tag, html, options) {
		return AntiXSS_noTag(html);
	}
};



</script>