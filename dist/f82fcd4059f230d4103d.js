function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return e; }; var t, e = {}, r = Object.prototype, n = r.hasOwnProperty, o = Object.defineProperty || function (t, e, r) { t[e] = r.value; }, i = "function" == typeof Symbol ? Symbol : {}, a = i.iterator || "@@iterator", c = i.asyncIterator || "@@asyncIterator", u = i.toStringTag || "@@toStringTag"; function define(t, e, r) { return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e]; } try { define({}, ""); } catch (t) { define = function define(t, e, r) { return t[e] = r; }; } function wrap(t, e, r, n) { var i = e && e.prototype instanceof Generator ? e : Generator, a = Object.create(i.prototype), c = new Context(n || []); return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a; } function tryCatch(t, e, r) { try { return { type: "normal", arg: t.call(e, r) }; } catch (t) { return { type: "throw", arg: t }; } } e.wrap = wrap; var h = "suspendedStart", l = "suspendedYield", f = "executing", s = "completed", y = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var p = {}; define(p, a, function () { return this; }); var d = Object.getPrototypeOf, v = d && d(d(values([]))); v && v !== r && n.call(v, a) && (p = v); var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p); function defineIteratorMethods(t) { ["next", "throw", "return"].forEach(function (e) { define(t, e, function (t) { return this._invoke(e, t); }); }); } function AsyncIterator(t, e) { function invoke(r, o, i, a) { var c = tryCatch(t[r], t, o); if ("throw" !== c.type) { var u = c.arg, h = u.value; return h && "object" == _typeof(h) && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) { invoke("next", t, i, a); }, function (t) { invoke("throw", t, i, a); }) : e.resolve(h).then(function (t) { u.value = t, i(u); }, function (t) { return invoke("throw", t, i, a); }); } a(c.arg); } var r; o(this, "_invoke", { value: function value(t, n) { function callInvokeWithMethodAndArg() { return new e(function (e, r) { invoke(t, n, e, r); }); } return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(e, r, n) { var o = h; return function (i, a) { if (o === f) throw Error("Generator is already running"); if (o === s) { if ("throw" === i) throw a; return { value: t, done: !0 }; } for (n.method = i, n.arg = a;;) { var c = n.delegate; if (c) { var u = maybeInvokeDelegate(c, n); if (u) { if (u === y) continue; return u; } } if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) { if (o === h) throw o = s, n.arg; n.dispatchException(n.arg); } else "return" === n.method && n.abrupt("return", n.arg); o = f; var p = tryCatch(e, r, n); if ("normal" === p.type) { if (o = n.done ? s : l, p.arg === y) continue; return { value: p.arg, done: n.done }; } "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg); } }; } function maybeInvokeDelegate(e, r) { var n = r.method, o = e.iterator[n]; if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y; var i = tryCatch(o, e.iterator, r.arg); if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y; var a = i.arg; return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y); } function pushTryEntry(t) { var e = { tryLoc: t[0] }; 1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e); } function resetTryEntry(t) { var e = t.completion || {}; e.type = "normal", delete e.arg, t.completion = e; } function Context(t) { this.tryEntries = [{ tryLoc: "root" }], t.forEach(pushTryEntry, this), this.reset(!0); } function values(e) { if (e || "" === e) { var r = e[a]; if (r) return r.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) { var o = -1, i = function next() { for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next; return next.value = t, next.done = !0, next; }; return i.next = i; } } throw new TypeError(_typeof(e) + " is not iterable"); } return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), o(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) { var e = "function" == typeof t && t.constructor; return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name)); }, e.mark = function (t) { return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t; }, e.awrap = function (t) { return { __await: t }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () { return this; }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) { void 0 === i && (i = Promise); var a = new AsyncIterator(wrap(t, r, n, o), i); return e.isGeneratorFunction(r) ? a : a.next().then(function (t) { return t.done ? t.value : a.next(); }); }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () { return this; }), define(g, "toString", function () { return "[object Generator]"; }), e.keys = function (t) { var e = Object(t), r = []; for (var n in e) r.push(n); return r.reverse(), function next() { for (; r.length;) { var t = r.pop(); if (t in e) return next.value = t, next.done = !1, next; } return next.done = !0, next; }; }, e.values = values, Context.prototype = { constructor: Context, reset: function reset(e) { if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t); }, stop: function stop() { this.done = !0; var t = this.tryEntries[0].completion; if ("throw" === t.type) throw t.arg; return this.rval; }, dispatchException: function dispatchException(e) { if (this.done) throw e; var r = this; function handle(n, o) { return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o; } for (var o = this.tryEntries.length - 1; o >= 0; --o) { var i = this.tryEntries[o], a = i.completion; if ("root" === i.tryLoc) return handle("end"); if (i.tryLoc <= this.prev) { var c = n.call(i, "catchLoc"), u = n.call(i, "finallyLoc"); if (c && u) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } else if (c) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); } else { if (!u) throw Error("try statement without catch or finally"); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } } } }, abrupt: function abrupt(t, e) { for (var r = this.tryEntries.length - 1; r >= 0; --r) { var o = this.tryEntries[r]; if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) { var i = o; break; } } i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null); var a = i ? i.completion : {}; return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a); }, complete: function complete(t, e) { if ("throw" === t.type) throw t.arg; return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y; }, finish: function finish(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y; } }, catch: function _catch(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.tryLoc === t) { var n = r.completion; if ("throw" === n.type) { var o = n.arg; resetTryEntry(r); } return o; } } throw Error("illegal catch attempt"); }, delegateYield: function delegateYield(e, r, n) { return this.delegate = { iterator: values(e), resultName: r, nextLoc: n }, "next" === this.method && (this.arg = t), y; } }, e; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
var storageKey = null;
var emailSent = false;
Office.initialize = function (reason) {
  var _Office$context$mailb;
  mailboxItem = (_Office$context$mailb = Office.context.mailbox) === null || _Office$context$mailb === void 0 ? void 0 : _Office$context$mailb.item;
};
function onNewComposeWindow() {
  emailSent = false;
  currentComposeId = Office.context.mailbox.item.itemId;
  storageKey = "attachments";
  if (!getFromLocalStorage("attachments")) {
    setInLocalStorage("attachments", JSON.stringify([]));
  }
}
function uploadFileToRoom(_x, _x2, _x3, _x4, _x5, _x6) {
  return _uploadFileToRoom.apply(this, arguments);
} // Entry point for AGORA add-in before send is allowed.
function _uploadFileToRoom() {
  _uploadFileToRoom = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(guestRID, event, roomId, userId, file, accessToken) {
    var url, formData, metadata, byteCharacters, byteNumbers, i, byteArray, blob, response, data, urlGuest, formDataGuest, raw, responseGuest, responseGuestJSON, errorData, _errorData;
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          url = globalConfig.backendBaseUrl + "/agora-api/api/v1/rooms/".concat(roomId, "/files");
          formData = new FormData();
          metadata = {
            type: 5,
            subtype: 2,
            permissions: {
              items: [{
                rid: userId,
                type: 2,
                subtype: 2,
                permission: 5,
                level: 5,
                inheritance: 0
              }]
            },
            name: file.name,
            expireDate: file.expire
          };
          formData.append("Resource", JSON.stringify(metadata));
          try {
            byteCharacters = atob(file.data);
            byteNumbers = new Array(byteCharacters.length);
            for (i = 0; i < byteCharacters.length; i++) {
              byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            byteArray = new Uint8Array(byteNumbers);
            blob = new Blob([byteArray], {
              type: file.type
            });
            formData.append("file", blob, file.name);
          } catch (error) {
            Office.context.mailbox.item.body.setSelectedDataAsync(error.toString(), {
              coercionType: Office.CoercionType.Text
            }, function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
              }
            });
          }
          _context.prev = 5;
          _context.next = 8;
          return fetch(url, {
            method: "POST",
            headers: {
              Authorization: "Bearer ".concat(accessToken),
              Accept: "application/json, text/plain, */*"
            },
            body: formData
          });
        case 8:
          response = _context.sent;
          if (!response.ok) {
            _context.next = 35;
            break;
          }
          _context.next = 12;
          return response.json();
        case 12:
          data = _context.sent;
          //return data;
          urlGuest = globalConfig.backendBaseUrl + "/agora-api/api/v1/files/".concat(data.rid, "/guest");
          formDataGuest = new FormData();
          if (guestRID == 0) {
            raw = JSON.stringify({
              expireDate: "".concat(file.expire),
              type: 2,
              subtype: 3,
              authenticationMethod: 0,
              mobile: "",
              language: "en_US",
              permission: "".concat(file.permission),
              country: "ch",
              email: "".concat(data.rid, "@").concat(data.rid, ".com"),
              guestAccessPropagate: true,
              loginInfo: true,
              token: true
            });
          } else {
            raw = JSON.stringify({
              rid: "".concat(guestRID),
              expireDate: "".concat(file.expire),
              type: 2,
              subtype: 3,
              authenticationMethod: 0,
              mobile: "",
              language: "en_US",
              permission: "".concat(file.permission),
              country: "ch",
              guestAccessPropagate: true
            });
          }
          _context.next = 18;
          return fetch(urlGuest, {
            method: "PUT",
            headers: {
              Authorization: "Bearer ".concat(accessToken),
              Accept: "application/json, text/plain, */*",
              "Content-Type": "application/json"
            },
            body: raw,
            redirect: "follow"
          });
        case 18:
          responseGuest = _context.sent;
          if (!responseGuest.ok) {
            _context.next = 29;
            break;
          }
          _context.next = 22;
          return responseGuest.json();
        case 22:
          responseGuestJSON = _context.sent;
          _context.next = 25;
          return new Promise(function (r) {
            return setTimeout(r, 3000);
          });
        case 25:
          responseGuestJSON.fid = data.rid;
          return _context.abrupt("return", responseGuestJSON);
        case 29:
          _context.next = 31;
          return responseGuest.text();
        case 31:
          errorData = _context.sent;
          throw new Error("File upload failed: ".concat(errorData));
        case 33:
          _context.next = 41;
          break;
        case 35:
          _context.next = 37;
          return response.text();
        case 37:
          _errorData = _context.sent;
          Office.context.mailbox.item.body.setSelectedDataAsync(_errorData, {
            coercionType: Office.CoercionType.Text
          }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log(asyncResult.error.message);
            }
          });
          console.error("File upload error response:", _errorData);
          // event.completed({ allowEvent: false });
          throw new Error("File upload failed: ".concat(_errorData));
        case 41:
          _context.next = 49;
          break;
        case 43:
          _context.prev = 43;
          _context.t0 = _context["catch"](5);
          console.error("An error occurred during file upload:", _context.t0);
          _context.next = 48;
          return addBodyOnSend(event, _context.t0.message);
        case 48:
          throw _context.t0;
        case 49:
        case "end":
          return _context.stop();
      }
    }, _callee, null, [[5, 43]]);
  }));
  return _uploadFileToRoom.apply(this, arguments);
}
function Upload(event) {
  mailboxItem.body.getAsync("text", {
    asyncContext: event
  }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
    }
    UploadOnlyOnSendCallBack(asyncResult);
  });
}
function addBodyOnSend(event, variable) {
  return new Promise(function (resolve, reject) {
    mailboxItem.body.setSelectedDataAsync(variable, {
      coercionType: Office.CoercionType.Html
    }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}
function addLinkOnSend(event, variable, displayText) {
  var hyperlink = '<br><a href="' + variable + '">' + displayText + "</a><br>";
  return new Promise(function (resolve, reject) {
    mailboxItem.body.setSelectedDataAsync(hyperlink, {
      coercionType: Office.CoercionType.Html
    }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}
function UploadOnlyOnSendCallBack(_x7) {
  return _UploadOnlyOnSendCallBack.apply(this, arguments);
}
function _UploadOnlyOnSendCallBack() {
  _UploadOnlyOnSendCallBack = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2(event) {
    var globalAttachments, message, attachments, final, guestRID, firstlink, link, found, _iterator, _step, attachment, i, attachment_internal, roomId, access, userId, regex, result;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          globalAttachments = [];
          _context2.next = 3;
          return getAttachments();
        case 3:
          globalAttachments = _context2.sent;
          message = getFromLocalStorage("myKey1");
          currentComposeId = Office.context.mailbox.item.conversationId;
          storageKey = "attachments-".concat(currentComposeId);
          attachments = JSON.parse(getFromLocalStorage(storageKey) || "[]");
          console.log("".concat(attachments.length));
          console.log("".concat(globalAttachments.length));
          //  let attachments_internal = Office.context.mailbox.item.attachments;
          if (attachments.length == 0) {
            try {
              console.log("Email body updated successfully.");
            } catch (error) {
              console.error("Failed to update email body: " + error.message);
              // Handle the error here
            }
            event.asyncContext.completed({
              allowEvent: true
            });
          }
          if (!(attachments.length > 0)) {
            _context2.next = 70;
            break;
          }
          final = event.value.toString();
          guestRID = 0;
          firstlink = "";
          link = "";
          found = false;
          _iterator = _createForOfIteratorHelper(attachments);
          _context2.prev = 18;
          _iterator.s();
        case 20:
          if ((_step = _iterator.n()).done) {
            _context2.next = 62;
            break;
          }
          attachment = _step.value;
          console.log(attachment.name);
          if (!(globalAttachments.length > 0)) {
            _context2.next = 37;
            break;
          }
          i = 0;
        case 25:
          if (!(i < globalAttachments.length)) {
            _context2.next = 37;
            break;
          }
          console.log("locals:".concat(attachment.name));
          console.log("global:".concat(globalAttachments[i].name));
          attachment_internal = globalAttachments[i];
          if (!(attachment.name == attachment_internal.name)) {
            _context2.next = 34;
            break;
          }
          found = true;
          _context2.next = 33;
          return Office.context.mailbox.item.removeAttachmentAsync(attachment_internal.id);
        case 33:
          return _context2.abrupt("break", 37);
        case 34:
          i++;
          _context2.next = 25;
          break;
        case 37:
          if (!(found == false)) {
            _context2.next = 39;
            break;
          }
          return _context2.abrupt("continue", 60);
        case 39:
          found = false;
          roomId = getFromLocalStorage("access-roomId");
          access = getFromLocalStorage("access");
          userId = getFromLocalStorage("access-userId");
          try {
            console.log("Email body updated successfully.");
          } catch (error) {
            console.error("Failed to update email body: " + error.message);
          }
          _context2.prev = 44;
          regex = /file=[0-9]+/i;
          _context2.next = 48;
          return uploadFileToRoom(guestRID, event.asyncContext, roomId, userId, attachment, access);
        case 48:
          result = _context2.sent;
          if (guestRID == 0) {
            firstlink = result.password;
            link = firstlink;
            guestRID = result.rid;
            console.log("first link  " + link);
          } else {
            decodedlink = decodeURIComponent(firstlink);
            console.log("encoded link  " + firstlink);
            console.log("decoded link  " + decodedlink);
            link = decodedlink.replace(regex, "file=".concat(result.fid));
            console.log("second link  " + link);
          }
          _context2.next = 52;
          return addLinkOnSend(event.asyncContext, link, attachment.name);
        case 52:
          _context2.next = 60;
          break;
        case 54:
          _context2.prev = 54;
          _context2.t0 = _context2["catch"](44);
          console.error("Failed to update email body: " + _context2.t0.message);
          _context2.next = 59;
          return addBodyOnSend(event.asyncContext, _context2.t0.message);
        case 59:
          event.asyncContext.completed({
            allowEvent: false
          });
        case 60:
          _context2.next = 20;
          break;
        case 62:
          _context2.next = 67;
          break;
        case 64:
          _context2.prev = 64;
          _context2.t1 = _context2["catch"](18);
          _iterator.e(_context2.t1);
        case 67:
          _context2.prev = 67;
          _iterator.f();
          return _context2.finish(67);
        case 70:
          //clean the attachments

          cleanupAttachments(storageKey);
          // Allow send.
          event.asyncContext.completed({
            allowEvent: true
          });
        case 72:
        case "end":
          return _context2.stop();
      }
    }, _callee2, null, [[18, 64, 67, 70], [44, 54]]);
  }));
  return _UploadOnlyOnSendCallBack.apply(this, arguments);
}
function openDialog(_x8) {
  return _openDialog.apply(this, arguments);
}
function _openDialog() {
  _openDialog = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee3(event) {
    var ready;
    return _regeneratorRuntime().wrap(function _callee3$(_context3) {
      while (1) switch (_context3.prev = _context3.next) {
        case 0:
          _context3.next = 2;
          return environmentReady();
        case 2:
          ready = _context3.sent;
          if (!ready) {
            Office.context.ui.displayDialogAsync(globalConfig.serverBaseUrl + "/taskpane/taskpane.html", {
              width: 50,
              height: 50,
              displayInIframe: true
            }, function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                var _asyncResult$error;
                console.error("Error: ".concat(asyncResult.error.message));
                Office.context.mailbox.item.notificationMessages.addAsync("openTaskPaneWarning", {
                  type: "informationalMessage",
                  message: "Error: ".concat((_asyncResult$error = asyncResult.error) === null || _asyncResult$error === void 0 ? void 0 : _asyncResult$error.message),
                  icon: "icon16",
                  persistent: false
                }, function (asyncResult) {
                  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Warning message displayed successfully.");
                  } else {
                    console.error("Error displaying warning message: ".concat(asyncResult.error.message));
                  }
                });
              }
              var dialog = asyncResult.value;
              dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                console.log("Dialog closed.");
                event.completed();
              });
            });
          } else {
            Office.context.ui.displayDialogAsync(globalConfig.serverBaseUrl + "/commands/dialog.html", {
              width: 50,
              height: 50,
              displayInIframe: true
            }, function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error: ".concat(asyncResult.error.message));
              }
              var dialog = asyncResult.value;
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (message) {
                addAttachment(event, JSON.parse(message.message));
              });
              dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                console.log("Dialog closed.");
                event.completed();
              });
            });
          }
        case 4:
        case "end":
          return _context3.stop();
      }
    }, _callee3);
  }));
  return _openDialog.apply(this, arguments);
}
function addAttachment(event, file) {
  var attachment = file;
  currentComposeId = Office.context.mailbox.item.conversationId;
  storageKey = "attachments-".concat(currentComposeId);
  var aa = JSON.parse(getFromLocalStorage(storageKey));
  if (!aa) {
    aa = [];
  }
  aa.push(attachment);
  try {
    setInLocalStorage(storageKey, JSON.stringify(aa));
  } catch (error) {
    addBodyOnSend(event, error.message);
    localStorage.clear();
  }
  var attachmentFile = {
    type: Office.MailboxEnums.AttachmentType.File,
    name: attachment.name,
    content: "This is an example attachment content."
  };
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(file.data, attachmentFile.name);
  console.log("Attachments stored in sessionStorage:");
  return;
}
function setInLocalStorage(key, value) {
  var myPartitionKey = Office.context.partitionKey;
  // Check if local storage is partitioned.
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}
function getFromLocalStorage(key) {
  var myPartitionKey = Office.context.partitionKey;
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}
function cleanupAttachments(key) {
  var myPartitionKey = Office.context.partitionKey;
  if (myPartitionKey) {
    return localStorage.removeItem(myPartitionKey + key);
  } else {
    return localStorage.removeItem(key);
  }
}
function environmentReady() {
  return _environmentReady.apply(this, arguments);
}
function _environmentReady() {
  _environmentReady = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee4() {
    var accessToken, roomId, finalResult;
    return _regeneratorRuntime().wrap(function _callee4$(_context4) {
      while (1) switch (_context4.prev = _context4.next) {
        case 0:
          accessToken = getFromLocalStorage("access");
          if (!(accessToken == null)) {
            _context4.next = 3;
            break;
          }
          return _context4.abrupt("return", false);
        case 3:
          roomId = getFromLocalStorage("access-roomId");
          if (!(roomId == null)) {
            _context4.next = 7;
            break;
          }
          Office.context.mailbox.item.notificationMessages.addAsync("openTaskPaneWarning", {
            type: "informationalMessage",
            message: "Your Personal Room is not created or deleted\nPlease check and create a Personal Room if needed then login again.",
            icon: "icon16",
            persistent: false
          }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Warning message displayed successfully.");
            } else {
              console.error("Error displaying warning message: ".concat(asyncResult.error.message));
            }
          });
          return _context4.abrupt("return", false);
        case 7:
          _context4.next = 9;
          return checkRoom(accessToken);
        case 9:
          finalResult = _context4.sent;
          return _context4.abrupt("return", finalResult);
        case 11:
        case "end":
          return _context4.stop();
      }
    }, _callee4);
  }));
  return _environmentReady.apply(this, arguments);
}
function checkRoom(_x9) {
  return _checkRoom.apply(this, arguments);
}
function _checkRoom() {
  _checkRoom = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee5(accessToken) {
    var response, data, roomId;
    return _regeneratorRuntime().wrap(function _callee5$(_context5) {
      while (1) switch (_context5.prev = _context5.next) {
        case 0:
          _context5.prev = 0;
          _context5.next = 3;
          return fetch("".concat(globalConfig.backendBaseUrl, "/agora-api/api/v1/user/settings/"), {
            method: "GET",
            headers: {
              Authorization: "Bearer ".concat(accessToken),
              Accept: "application/json, text/plain, */*"
            }
          });
        case 3:
          response = _context5.sent;
          _context5.next = 6;
          return response.json();
        case 6:
          data = _context5.sent;
          if (!(response.ok && data.roomId && data.roomPermission >= 4)) {
            _context5.next = 13;
            break;
          }
          roomId = data.roomId;
          setInLocalStorage("access-roomId", roomId);
          return _context5.abrupt("return", true);
        case 13:
          Office.context.mailbox.item.notificationMessages.addAsync("openTaskPaneWarning", {
            type: "informationalMessage",
            message: "Your Personal Room is not created or deleted\nPlease check and create a Personal Room if needed then login again.",
            icon: "icon16",
            persistent: false
          }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Warning message displayed successfully.");
            } else {
              console.error("Error displaying warning message: ".concat(asyncResult.error.message));
            }
          });
          return _context5.abrupt("return", false);
        case 15:
          _context5.next = 21;
          break;
        case 17:
          _context5.prev = 17;
          _context5.t0 = _context5["catch"](0);
          mailboxItem.body.setSelectedDataAsync(_context5.t0.message, {
            coercionType: Office.CoercionType.Text
          }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log(asyncResult.error.message);
            }
          });
          return _context5.abrupt("return", false);
        case 21:
        case "end":
          return _context5.stop();
      }
    }, _callee5, null, [[0, 17]]);
  }));
  return _checkRoom.apply(this, arguments);
}
function getAttachments() {
  return new Promise(function (resolve, reject) {
    var Attachments = [];
    Office.context.mailbox.item.getAttachmentsAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Update the variable with the attachments
        Attachments = asyncResult.value;
        resolve(Attachments);
      } else {
        reject("Error getting attachments: " + asyncResult.error.message);
      }
    });
  });
}
Office.actions.associate("openDialog", openDialog);