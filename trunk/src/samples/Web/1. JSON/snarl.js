Snarl = function(){
	var SNARL_WEB_BRIDGE = "http://localhost:9889/?d={0}&u={1}";
	var REGISTRATION = 0;
	var NOTIFICATION = 1;
	var iframe = null;
	var form = null;
	var field = null;
	var isInitialized = false;
	var isRegistered = false;
	var snarlsRunning = true; //TODO:
	var appName = null;

	function send(data){
		if(!data) data = "";
		var u = new Date().getTime();
		var url = SNARL_WEB_BRIDGE.replace(/\{0\}/,data).replace(/\{1\}/,u);
		//window.prompt("url", url);
		iframe.contentWindow.location.replace(url);
	}

	return {
		init : function(){
			if(!isInitialized){
				// set up iframe for communication
				iframe = document.createElement("iframe");
				iframe.id = "SnarlCommunicator";
				iframe.name = iframe.id;
				iframe.src = "about:blank";
				iframe.style.border = "0px";
				iframe.style.width = "0px";
				iframe.style.height = "0px";
				document.body.appendChild(iframe);
				if(!iframe.contentWindow)
					iframe.document = iframe.contentDocument;

				isInitialized = true;
		    }
		},

		register : function(applicationName, notificationTypes){
			Snarl.init();
		//	if(snarlIsRunning){
			    appName = applicationName;
				var data = {"action" : REGISTRATION,
							"applicationName" : applicationName,
							"notificationTypes" : notificationTypes};
				var json = Snarl.utils.json.stringify(data);
				send(json);
				isRegistered = true;
		//	}
		},

		notify : function(notificationType, title, description, priority, sticky){
			Snarl.init();
			if(isRegistered){
				var data = {"action" : NOTIFICATION,
							"applicationName" : appName,
							"notificationType" : notificationType,
							"title" : title,
							"description" : description,
							"priority" : priority,
							"sticky" : sticky};
				var json = Snarl.utils.json.stringify(data);
				send(json);
			}
		}
	}
}();

Snarl.Priority = function(){
	return {
		Emergency : 2,
		High : 1,
		Normal : 0,
		Moderate : -1,
		VeryLow : -2
    }
}();

Snarl.NotificationType = function(name, enabled){
	this.name = name;
	this.enabled = enabled;
}

/* *************************************************
Snarl.utils.json

- the json module includes some regexs that look like unterminated strings, so some minification tools might choke on it

- if you would prefer to use another json library, you can override the following methods:
	Snarl.utils.json.stringify(obj) - where obj is the object to convert. returns a json-formatted string
	Snarl.utils.json.parse(json) - where json is the json-formatted string to parse. returns a javascript object

************************************************* */
Snarl.utils = function(){
	return {
		json : function() {
			// Test for modern browser (any except IE5).
			var JS13 = ('1'.replace(/1/, function() { return ''; }) == '');

			// CHARS array stores special strings for encodeString() function.
			var CHARS = {
				'\b': '\\b',
				'\t': '\\t',
				'\n': '\\n',
				'\f': '\\f',
				'\r': '\\r',
				'\\': '\\\\',
				'"' : '\\"'
			};

			for (var i = 0; i < 32; i++) {
				var c = String.fromCharCode(i);
				if (!CHARS[c]) CHARS[c] = ((i < 16) ? '\\u000' : '\\u00') + i.toString(16);
			};

			function encodeString(str) {
				if (!/[\x00-\x1f\\"]/.test(str)) {
				  return str;
				} else if (JS13) {
				  return str.replace(/([\x00-\x1f\\"])/g, function($0, $1) {
					return CHARS[$1];
				});
				} else {
				var out = new Array(str.length);
				for (var i = 0; i < str.length; i++) {
					var c = str.charAt(i);
					out[i] = CHARS[c] || c;
				}
				return out.join('');
				}
			};

			return {
				stringify : function (arg) {
					switch (typeof arg) {
						//case 'string'   : return '"' + encodeString(arg) + '"'; // break command is redundant here and below.
						case 'string'   : return '"' + arg + '"'; // break command is redundant here and below.
						case 'number'   : return String(arg);
						case 'object':
						if (arg) {
							var out = [];
							if (arg instanceof Array) {
							for (var i = 0; i < arg.length; i++) {
								var json = this.stringify(arg[i]);
								if (json != null) out[out.length] = json;
							}
							return '[' + out.join(',') + ']';
							} else {
							for (var p in arg) {
								var json = this.stringify(arg[p]);
								if (json != null) out[out.length] = '"' + encodeString(p) + '":' + json;
							}
							return '{' + out.join(',') + '}';
							}
						}
						return 'null'; // if execution reaches here, arg is null.
						case 'boolean'  : return String(arg);
						// cases function & undefined return null implicitly.
					}
				},

				parse: function (text) {
					try {
						return !(/[^,:{}\[\]0-9.\-+Eaeflnr-u \n\r\t]/.test(text.replace(/"(\\.|[^"\\])*"/g, ''))) && eval('(' + text + ')');
					} catch (e) {
						return false;
					}
				}
			}
		}()
	}
}();
