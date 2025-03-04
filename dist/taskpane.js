/*! For license information please see taskpane.js.LICENSE.txt */
!function(){"use strict";var e={58394:function(e,t,r){e.exports=r.p+"1fda685b81e1123773f6.css"},98362:function(e,t,r){e.exports=r.p+"assets/logo-filled.png"}},t={};function r(n){var o=t[n];if(void 0!==o)return o.exports;var a=t[n]={exports:{}};return e[n](a,a.exports,r),a.exports}r.m=e,r.d=function(e,t){for(var n in t)r.o(t,n)&&!r.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),r.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;r.g.importScripts&&(e=r.g.location+"");var t=r.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/^blob:/,"").replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=e}(),r.b=document.baseURI||self.location.href,function(){function e(t){return e="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},e(t)}function t(e,t){(null==t||t>e.length)&&(t=e.length);for(var r=0,n=Array(t);r<t;r++)n[r]=e[r];return n}function r(){r=function(){return n};var t,n={},o=Object.prototype,a=o.hasOwnProperty,c=Object.defineProperty||function(e,t,r){e[t]=r.value},s="function"==typeof Symbol?Symbol:{},i=s.iterator||"@@iterator",l=s.asyncIterator||"@@asyncIterator",u=s.toStringTag||"@@toStringTag";function f(e,t,r){return Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{f({},"")}catch(t){f=function(e,t,r){return e[t]=r}}function p(e,t,r,n){var o=t&&t.prototype instanceof w?t:w,a=Object.create(o.prototype),s=new P(n||[]);return c(a,"_invoke",{value:O(e,r,s)}),a}function d(e,t,r){try{return{type:"normal",arg:e.call(t,r)}}catch(e){return{type:"throw",arg:e}}}n.wrap=p;var h="suspendedStart",y="suspendedYield",m="executing",g="completed",v={};function w(){}function x(){}function b(){}var k={};f(k,i,(function(){return this}));var E=Object.getPrototypeOf,C=E&&E(E(j([])));C&&C!==o&&a.call(C,i)&&(k=C);var I=b.prototype=w.prototype=Object.create(k);function A(e){["next","throw","return"].forEach((function(t){f(e,t,(function(e){return this._invoke(t,e)}))}))}function S(t,r){function n(o,c,s,i){var l=d(t[o],t,c);if("throw"!==l.type){var u=l.arg,f=u.value;return f&&"object"==e(f)&&a.call(f,"__await")?r.resolve(f.__await).then((function(e){n("next",e,s,i)}),(function(e){n("throw",e,s,i)})):r.resolve(f).then((function(e){u.value=e,s(u)}),(function(e){return n("throw",e,s,i)}))}i(l.arg)}var o;c(this,"_invoke",{value:function(e,t){function a(){return new r((function(r,o){n(e,t,r,o)}))}return o=o?o.then(a,a):a()}})}function O(e,r,n){var o=h;return function(a,c){if(o===m)throw Error("Generator is already running");if(o===g){if("throw"===a)throw c;return{value:t,done:!0}}for(n.method=a,n.arg=c;;){var s=n.delegate;if(s){var i=R(s,n);if(i){if(i===v)continue;return i}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===h)throw o=g,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=m;var l=d(e,r,n);if("normal"===l.type){if(o=n.done?g:y,l.arg===v)continue;return{value:l.arg,done:n.done}}"throw"===l.type&&(o=g,n.method="throw",n.arg=l.arg)}}}function R(e,r){var n=r.method,o=e.iterator[n];if(o===t)return r.delegate=null,"throw"===n&&e.iterator.return&&(r.method="return",r.arg=t,R(e,r),"throw"===r.method)||"return"!==n&&(r.method="throw",r.arg=new TypeError("The iterator does not provide a '"+n+"' method")),v;var a=d(o,e.iterator,r.arg);if("throw"===a.type)return r.method="throw",r.arg=a.arg,r.delegate=null,v;var c=a.arg;return c?c.done?(r[e.resultName]=c.value,r.next=e.nextLoc,"return"!==r.method&&(r.method="next",r.arg=t),r.delegate=null,v):c:(r.method="throw",r.arg=new TypeError("iterator result is not an object"),r.delegate=null,v)}function L(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function T(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function P(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(L,this),this.reset(!0)}function j(r){if(r||""===r){var n=r[i];if(n)return n.call(r);if("function"==typeof r.next)return r;if(!isNaN(r.length)){var o=-1,c=function e(){for(;++o<r.length;)if(a.call(r,o))return e.value=r[o],e.done=!1,e;return e.value=t,e.done=!0,e};return c.next=c}}throw new TypeError(e(r)+" is not iterable")}return x.prototype=b,c(I,"constructor",{value:b,configurable:!0}),c(b,"constructor",{value:x,configurable:!0}),x.displayName=f(b,u,"GeneratorFunction"),n.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===x||"GeneratorFunction"===(t.displayName||t.name))},n.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,b):(e.__proto__=b,f(e,u,"GeneratorFunction")),e.prototype=Object.create(I),e},n.awrap=function(e){return{__await:e}},A(S.prototype),f(S.prototype,l,(function(){return this})),n.AsyncIterator=S,n.async=function(e,t,r,o,a){void 0===a&&(a=Promise);var c=new S(p(e,t,r,o),a);return n.isGeneratorFunction(t)?c:c.next().then((function(e){return e.done?e.value:c.next()}))},A(I),f(I,u,"Generator"),f(I,i,(function(){return this})),f(I,"toString",(function(){return"[object Generator]"})),n.keys=function(e){var t=Object(e),r=[];for(var n in t)r.push(n);return r.reverse(),function e(){for(;r.length;){var n=r.pop();if(n in t)return e.value=n,e.done=!1,e}return e.done=!0,e}},n.values=j,P.prototype={constructor:P,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=t,this.done=!1,this.delegate=null,this.method="next",this.arg=t,this.tryEntries.forEach(T),!e)for(var r in this)"t"===r.charAt(0)&&a.call(this,r)&&!isNaN(+r.slice(1))&&(this[r]=t)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var r=this;function n(n,o){return s.type="throw",s.arg=e,r.next=n,o&&(r.method="next",r.arg=t),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var c=this.tryEntries[o],s=c.completion;if("root"===c.tryLoc)return n("end");if(c.tryLoc<=this.prev){var i=a.call(c,"catchLoc"),l=a.call(c,"finallyLoc");if(i&&l){if(this.prev<c.catchLoc)return n(c.catchLoc,!0);if(this.prev<c.finallyLoc)return n(c.finallyLoc)}else if(i){if(this.prev<c.catchLoc)return n(c.catchLoc,!0)}else{if(!l)throw Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return n(c.finallyLoc)}}}},abrupt:function(e,t){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&a.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var o=n;break}}o&&("break"===e||"continue"===e)&&o.tryLoc<=t&&t<=o.finallyLoc&&(o=null);var c=o?o.completion:{};return c.type=e,c.arg=t,o?(this.method="next",this.next=o.finallyLoc,v):this.complete(c)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),v},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var r=this.tryEntries[t];if(r.finallyLoc===e)return this.complete(r.completion,r.afterLoc),T(r),v}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var r=this.tryEntries[t];if(r.tryLoc===e){var n=r.completion;if("throw"===n.type){var o=n.arg;T(r)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,r,n){return this.delegate={iterator:j(e),resultName:r,nextLoc:n},"next"===this.method&&(this.arg=t),v}},n}function n(e,t,r,n,o,a,c){try{var s=e[a](c),i=s.value}catch(e){return void r(e)}s.done?t(i):Promise.resolve(i).then(n,o)}function o(e){return function(){var t=this,r=arguments;return new Promise((function(o,a){var c=e.apply(t,r);function s(e){n(c,o,a,s,i,"next",e)}function i(e){n(c,o,a,s,i,"throw",e)}s(void 0)}))}}function a(){return c.apply(this,arguments)}function c(){return c=o(r().mark((function e(){return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(r().mark((function e(t){var n;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(n=t.workbook.getSelectedRange()).load(["columnIndex","values"]),e.next=4,t.sync();case 4:if(void 0!==n.columnIndex){e.next=6;break}throw new Error("Please select a column");case 6:document.getElementById("passwordDialog").showModal();case 7:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=9;break;case 5:e.prev=5,e.t0=e.catch(0),console.error(e.t0),Office.context.ui.displayDialogAsync("Error: ".concat(e.t0.message));case 9:case"end":return e.stop()}}),e,null,[[0,5]])}))),c.apply(this,arguments)}function s(){return i.apply(this,arguments)}function i(){return i=o(r().mark((function e(){return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(r().mark((function e(t){var n,o,a,c,s,i,l,u,f;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(n=t.workbook.getSelectedRange()).load(["columnIndex","values"]),e.next=4,t.sync();case 4:if(void 0!==n.columnIndex){e.next=6;break}throw new Error("Please select a column");case 6:return o=t.workbook.worksheets.getActiveWorksheet(),(a=o.getUsedRange()).load(["rowCount"]),e.next=11,t.sync();case 11:return c=n.columnIndex,(s=o.getRangeByIndexes(0,c,a.rowCount,1)).load(["values","rowCount"]),(i=s.getCell(0,0)).format.load("fill"),e.next=18,t.sync();case 18:if(l=s.values[0][0],console.log("Header content before decryption:",l),console.log("Header cell background color:",i.format.fill.color),(u=i.format.fill.color)&&"#c8e6c9"===u.toLowerCase()){e.next=27;break}return console.log("This column is not encrypted"),e.next=26,Office.context.ui.displayDialogAsync("This column is not encrypted",{height:30,width:30});case 26:return e.abrupt("return");case 27:(f=document.getElementById("passwordDialog")).dataset.action="decrypt",f.showModal();case 30:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=9;break;case 5:e.prev=5,e.t0=e.catch(0),console.error(e.t0),Office.context.ui.displayDialogAsync("Error: ".concat(e.t0.message));case 9:case"end":return e.stop()}}),e,null,[[0,5]])}))),i.apply(this,arguments)}function l(){return l=o(r().mark((function e(n){var a;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(r().mark((function e(o){var a,c,s,i,l,u,f,p,d,h,y,m,g,v,w,x,b,k,E,C,I,A,S,O,R;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(a=o.workbook.getSelectedRange()).load(["columnCount","rowCount","values","columnIndex"]),e.next=4,o.sync();case 4:if(1===a.columnCount){e.next=6;break}throw new Error("Please select only one column");case 6:return console.log("Selected range data:",a.values),console.log("Selected column index:",a.columnIndex),c=o.workbook.worksheets.getActiveWorksheet(),(s=c.getUsedRange()).load(["rowCount","address"]),e.next=13,o.sync();case 13:return console.log("Used range row count:",s.rowCount),i=a.columnIndex,(l=c.getRangeByIndexes(0,i,s.rowCount,1)).load(["values","rowCount"]),(u=l.getCell(0,0)).format.load("fill"),e.next=21,o.sync();case 21:if(console.log("Header cell background color:",u.format.fill.color),"#C8E6C9"!==u.format.fill.color){e.next=26;break}return e.next=25,Office.context.ui.displayDialogAsync("This column is already encrypted",{height:30,width:30});case 25:return e.abrupt("return");case 26:if(l.values&&0!==l.values.length){e.next=28;break}throw new Error("Unable to read column data");case 28:if(console.log("Entire column data:",l.values),0!==(f=l.values.slice(1).map((function(e){return e[0]})).filter((function(e){return""!==e&&null!=e}))).length){e.next=32;break}throw new Error("No valid data to encrypt in selected column");case 32:return console.log("Data to encrypt:",f),p=l.values.slice(1).map((function(e,t){return{value:e[0],rowIndex:t+1}})).filter((function(e){return""!==e.value&&null!==e.value&&void 0!==e.value})),console.log("Valid data rows:",p),d=btoa(n),console.log("Base64 encoded password:",d),h='VSAuth vsauth_method="sharedSecret",vsauth_data="'.concat(d,'",vsauth_identity_ascii="demo@voltage.com",vsauth_version="200"'),console.log("Complete auth header:",h),y={format:"AUTO",data:f},console.log("=== API Request Start ==="),m="https://voltage-pp-0000.dataprotection.voltage.com/vibesimple/rest/v1/protect",console.log("Request URL:",m),console.log("Request method:","POST"),console.log("Request headers:",{"Content-Type":"application/json",Authorization:h}),console.log("Request body:",JSON.stringify(y,null,2)),g=new Date,e.prev=47,v=new AbortController,w=setTimeout((function(){return v.abort()}),3e4),e.next=52,fetch(m,{method:"POST",headers:{"Content-Type":"application/json",Authorization:h},body:JSON.stringify(y),signal:v.signal,redirect:"follow",referrerPolicy:"no-referrer"});case 52:if(x=e.sent,clearTimeout(w),b=new Date,k=b-g,console.log("=== API Response Info ==="),console.log("Response status:",x.status,x.statusText),console.log("Response headers:",Object.fromEntries(function(e){if(Array.isArray(e))return t(e)}(r=x.headers)||function(e){if("undefined"!=typeof Symbol&&null!=e[Symbol.iterator]||null!=e["@@iterator"])return Array.from(e)}(r)||function(e,r){if(e){if("string"==typeof e)return t(e,r);var n={}.toString.call(e).slice(8,-1);return"Object"===n&&e.constructor&&(n=e.constructor.name),"Map"===n||"Set"===n?Array.from(e):"Arguments"===n||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)?t(e,r):void 0}}(r)||function(){throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}())),console.log("Request duration:",k,"ms"),x.ok){e.next=66;break}return e.next=63,x.text();case 63:throw E=e.sent,console.error("Error response body:",E),new Error("Encryption service request failed: ".concat(x.status," ").concat(x.statusText," - ").concat(E));case 66:return e.next=68,x.json();case 68:if(C=e.sent,console.log("Response body:",JSON.stringify(C,null,2)),C&&C.data&&Array.isArray(C.data)){e.next=72;break}throw new Error("Invalid server response format");case 72:if(C.data.length===f.length){e.next=74;break}throw new Error("Encrypted data length does not match original data");case 74:return console.log("=== API Request End ==="),p.forEach((function(e,t){var r=l.getCell(e.rowIndex,0);r.values=[[C.data[t]]],r.format.fill.color="#E8F5E9"})),(I=l.getCell(0,0)).format.fill.color="#C8E6C9",(A=o.workbook.worksheets.getActiveWorksheet()).protection.load("protected"),e.next=82,o.sync();case 82:if(!A.protection.protected){e.next=86;break}return A.protection.unprotect(),e.next=86,o.sync();case 86:return(S=A.getUsedRange()).load("columnCount"),e.next=90,o.sync();case 90:return S.format.protection.locked=!1,e.next=93,o.sync();case 93:for(O=0;O<l.rowCount;O++)l.getCell(O,0).format.protection.locked=!0;return e.next=96,o.sync();case 96:return A.protection.protect({allowInsertRows:!0,allowInsertColumns:!0,allowDeleteRows:!0,allowDeleteColumns:!0,allowSort:!0,allowFilter:!0,allowEditObjects:!0,allowEditScenarios:!0}),e.next=99,o.sync();case 99:R="🔒 Encryption time: ".concat((new Date).toLocaleString(),"\nEncrypted data count: ").concat(C.data.length),I.worksheet.comments.add(I,R),e.next=117;break;case 103:if(e.prev=103,e.t0=e.catch(47),console.error("=== Network Request Error ==="),console.error("Error type:",e.t0.name),console.error("Error message:",e.t0.message),"AbortError"!==e.t0.name){e.next=112;break}throw new Error("Request timeout, please check your network connection");case 112:if(!e.t0.message.includes("Failed to fetch")){e.next=116;break}throw new Error("Unable to connect to encryption server, please check:\n1. Company network connection\n2. VPN status");case 116:throw e.t0;case 117:case"end":return e.stop()}var r}),e,null,[[47,103]])})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=13;break;case 5:e.prev=5,e.t0=e.catch(0),console.error("=== API Error ==="),console.error("Error details:",e.t0),console.error("Error stack:",e.t0.stack),"Error during encryption",a=e.t0.message.includes("Cannot read properties of null")?"Invalid server response format":e.t0.message.includes("Failed to fetch")?e.t0.message:e.t0.message.includes("timeout")?"Request timeout, please check network connection":"Encryption failed: ".concat(e.t0.message),Office.context.ui.displayDialogAsync(a);case 13:case"end":return e.stop()}}),e,null,[[0,5]])}))),l.apply(this,arguments)}function u(){return u=o(r().mark((function e(t){var n;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(r().mark((function e(n){var o,a,c,s,i,l,u,f,p,d,h,y,m,g,v,w,x,b,k,E;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(o=n.workbook.getSelectedRange()).load(["columnCount","rowCount","values","columnIndex"]),e.next=4,n.sync();case 4:if(1===o.columnCount){e.next=6;break}throw new Error("Please select only one column");case 6:return a=n.workbook.worksheets.getActiveWorksheet(),(c=a.getUsedRange()).load(["rowCount"]),e.next=11,n.sync();case 11:return s=o.columnIndex,(i=a.getRangeByIndexes(0,s,c.rowCount,1)).load(["values","rowCount"]),(l=i.getCell(0,0)).format.load("fill"),e.next=18,n.sync();case 18:if(u=i.values[0][0],console.log("Header content before decryption:",u),console.log("Header cell background color:",l.format.fill.color),(f=l.format.fill.color)&&"#c8e6c9"===f.toLowerCase()){e.next=27;break}return console.log("This column is not encrypted"),e.next=26,Office.context.ui.displayDialogAsync("This column is not encrypted",{height:30,width:30});case 26:return e.abrupt("return");case 27:if(0!==(p=i.values.slice(1).map((function(e){return e[0]})).filter((function(e){return""!==e&&null!=e}))).length){e.next=30;break}throw new Error("No valid data to decrypt in selected column");case 30:return d=i.values.slice(1).map((function(e,t){return{value:e[0],rowIndex:t+1}})).filter((function(e){return""!==e.value&&null!==e.value&&void 0!==e.value})),h=btoa(t),y='VSAuth vsauth_method="sharedSecret",vsauth_data="'.concat(h,'",vsauth_identity_ascii="demo@voltage.com",vsauth_version="200"'),m={format:"AUTO",data:p},console.log("=== Decryption API Request Start ==="),e.prev=36,g=new AbortController,v=setTimeout((function(){return g.abort()}),3e4),e.next=41,fetch("https://voltage-pp-0000.dataprotection.voltage.com/vibesimple/rest/v1/access",{method:"POST",headers:{"Content-Type":"application/json",Authorization:y},body:JSON.stringify(m),signal:g.signal});case 41:if(w=e.sent,clearTimeout(v),w.ok){e.next=48;break}return e.next=46,w.text();case 46:throw x=e.sent,new Error("Decryption service request failed: ".concat(w.status," ").concat(w.statusText," - ").concat(x));case 48:return e.next=50,w.json();case 50:if((b=e.sent)&&b.data&&Array.isArray(b.data)){e.next=53;break}throw new Error("Invalid server response format");case 53:return(k=n.workbook.worksheets.getActiveWorksheet()).protection.load("protected"),e.next=57,n.sync();case 57:if(!k.protection.protected){e.next=61;break}return k.protection.unprotect(),e.next=61,n.sync();case 61:return d.forEach((function(e,t){var r=i.getCell(e.rowIndex,0);r.values=[[b.data[t]]],r.format.fill.clear()})),(E=i.getCell(0,0)).format.fill.clear(),E.clear(Excel.ClearApplyTo.comments),e.next=67,n.sync();case 67:return E.values=[[u]],e.next=70,n.sync();case 70:return console.log("Header content after decryption:",u),e.next=73,Office.context.ui.displayDialogAsync("Decryption completed!",{height:30,width:30});case 73:e.next=87;break;case 75:if(e.prev=75,e.t0=e.catch(36),console.error("=== Network Request Error ==="),"AbortError"!==e.t0.name){e.next=82;break}throw new Error("Request timeout, please check your network connection");case 82:if(!e.t0.message.includes("Failed to fetch")){e.next=86;break}throw new Error("Unable to connect to decryption server, please check:\n1. Company network connection\n2. VPN status");case 86:throw e.t0;case 87:case"end":return e.stop()}}),e,null,[[36,75]])})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=12;break;case 5:e.prev=5,e.t0=e.catch(0),console.error("=== Decryption Error ==="),console.error("Error details:",e.t0),"Error during decryption",n=e.t0.message.includes("Cannot read properties of null")?"Invalid server response format":e.t0.message.includes("Failed to fetch")?e.t0.message:e.t0.message.includes("timeout")?"Request timeout, please check network connection":"Decryption failed: ".concat(e.t0.message),Office.context.ui.displayDialogAsync(n);case 12:case"end":return e.stop()}}),e,null,[[0,5]])}))),u.apply(this,arguments)}Office.onReady((function(e){if(e.host===Office.HostType.Excel){document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex";var t=document.getElementById("encrypt"),r=document.getElementById("decrypt");t&&(t.onclick=a),r&&(r.onclick=s);var n=document.getElementById("passwordDialog"),o=document.getElementById("confirmPassword"),c=document.getElementById("cancelPassword");o&&(o.onclick=function(){var e=document.getElementById("password").value;e&&("decrypt"===n.dataset.action?function(e){u.apply(this,arguments)}(e):function(e){l.apply(this,arguments)}(e),n.close())}),c&&(c.onclick=function(){n.close()})}}))}(),new URL(r(58394),r.b),new URL(r(98362),r.b)}();
//# sourceMappingURL=taskpane.js.map