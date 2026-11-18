/*
 * SCORM run-time API for SCORM 1.2 (window.API) and SCORM 2004 (window.API_1484_11).
 *
 * The SCO runs in an iframe that is same-origin with this page, so it finds the
 * API by walking window.parent as the SCORM spec describes. All calls are answered
 * synchronously from an in-memory CMI store; changes are pushed to the SharePoint
 * web part asynchronously (postMessage) on Commit / Finish / page unload.
 *
 * Data is kept as a flat map: { "cmi.core.lesson_status": "incomplete", ... }
 */
(function (global) {
  'use strict';

  // ------------------------------------------------------------------ helpers
  var RO = 'ro', RW = 'rw', WO = 'wo';

  function template(element) {
    // "cmi.interactions.3.objectives.0.id" -> "cmi.interactions.n.objectives.n.id"
    return element.replace(/\.\d+(?=\.|$)/g, '.n');
  }

  function isDecimal(v) { return /^-?\d+(\.\d+)?$/.test(String(v)); }
  function inRange(min, max) { return function (v) { return isDecimal(v) && +v >= min && +v <= max; }; }
  function vocab(list) { return function (v) { return list.indexOf(v) >= 0; }; }
  function maxLen(n) { return function (v) { return String(v).length <= n; }; }
  function any() { return true; }

  var timespan12 = function (v) { return /^\d{2,4}:\d{2}:\d{2}(\.\d{1,2})?$/.test(v); };
  var duration04 = function (v) { return /^P(?!$)(\d+Y)?(\d+M)?(\d+D)?(T(?=\d)(\d+H)?(\d+M)?(\d+(\.\d{1,2})?S)?)?$/.test(v); };
  var time04 = function (v) { return /^\d{4}(-\d{2}(-\d{2}(T\d{2}(:\d{2}(:\d{2}(\.\d{1,2})?)?)?(Z|[+-]\d{2}(:\d{2})?)?)?)?)?$/.test(v); };

  // Seconds <-> SCORM time formats
  function parse12(t) {
    var m = /^(\d+):(\d+):(\d+(?:\.\d+)?)$/.exec(t || '');
    return m ? (+m[1]) * 3600 + (+m[2]) * 60 + parseFloat(m[3]) : 0;
  }
  function format12(sec) {
    var h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec - h * 3600 - m * 60;
    var pad = function (n, l) { n = String(n); while (n.length < l) { n = '0' + n; } return n; };
    return pad(h, 4) + ':' + pad(m, 2) + ':' + pad(s.toFixed(2), 5);
  }
  function parse04(t) {
    var m = /^P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(t || '');
    if (!m) { return 0; }
    return (+(m[1] || 0)) * 31536000 + (+(m[2] || 0)) * 2592000 + (+(m[3] || 0)) * 86400 +
      (+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + parseFloat(m[6] || 0);
  }
  function format04(sec) {
    var h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = Math.round((sec % 60) * 100) / 100;
    return 'PT' + h + 'H' + m + 'M' + s + 'S';
  }

  // ------------------------------------------------------------ data models
  var MODEL_12 = {
    'cmi._children': [RO], 'cmi._version': [RO],
    'cmi.core._children': [RO],
    'cmi.core.student_id': [RO], 'cmi.core.student_name': [RO],
    'cmi.core.lesson_location': [RW, maxLen(255)],
    'cmi.core.credit': [RO],
    'cmi.core.lesson_status': [RW, vocab(['passed', 'completed', 'failed', 'incomplete', 'browsed'])],
    'cmi.core.entry': [RO],
    'cmi.core.score._children': [RO],
    'cmi.core.score.raw': [RW, function (v) { return v === '' || inRange(0, 100)(v); }],
    'cmi.core.score.min': [RW, function (v) { return v === '' || inRange(0, 100)(v); }],
    'cmi.core.score.max': [RW, function (v) { return v === '' || inRange(0, 100)(v); }],
    'cmi.core.total_time': [RO],
    'cmi.core.lesson_mode': [RO],
    'cmi.core.exit': [WO, vocab(['time-out', 'suspend', 'logout', ''])],
    'cmi.core.session_time': [WO, timespan12],
    'cmi.suspend_data': [RW, maxLen(4096)],
    'cmi.launch_data': [RO],
    'cmi.comments': [RW, maxLen(4096)],
    'cmi.comments_from_lms': [RO],
    'cmi.objectives._children': [RO], 'cmi.objectives._count': [RO],
    'cmi.objectives.n.id': [RW, maxLen(255)],
    'cmi.objectives.n.score._children': [RO],
    'cmi.objectives.n.score.raw': [RW, function (v) { return v === '' || isDecimal(v); }],
    'cmi.objectives.n.score.min': [RW, function (v) { return v === '' || isDecimal(v); }],
    'cmi.objectives.n.score.max': [RW, function (v) { return v === '' || isDecimal(v); }],
    'cmi.objectives.n.status': [RW, vocab(['passed', 'completed', 'failed', 'incomplete', 'browsed', 'not attempted'])],
    'cmi.student_data._children': [RO],
    'cmi.student_data.mastery_score': [RO],
    'cmi.student_data.max_time_allowed': [RO],
    'cmi.student_data.time_limit_action': [RO],
    'cmi.student_preference._children': [RO],
    'cmi.student_preference.audio': [RW, inRange(-1, 100)],
    'cmi.student_preference.language': [RW, maxLen(255)],
    'cmi.student_preference.speed': [RW, inRange(-100, 100)],
    'cmi.student_preference.text': [RW, inRange(-1, 1)],
    'cmi.interactions._children': [RO], 'cmi.interactions._count': [RO],
    'cmi.interactions.n.id': [WO, maxLen(255)],
    'cmi.interactions.n.objectives._count': [RO],
    'cmi.interactions.n.objectives.n.id': [WO, maxLen(255)],
    'cmi.interactions.n.time': [WO, function (v) { return /^\d{2}:\d{2}:\d{2}(\.\d{1,2})?$/.test(v); }],
    'cmi.interactions.n.type': [WO, vocab(['true-false', 'choice', 'fill-in', 'matching', 'performance', 'sequencing', 'likert', 'numeric'])],
    'cmi.interactions.n.correct_responses._count': [RO],
    'cmi.interactions.n.correct_responses.n.pattern': [WO, maxLen(255)],
    'cmi.interactions.n.weighting': [WO, isDecimal],
    'cmi.interactions.n.student_response': [WO, maxLen(255)],
    'cmi.interactions.n.result': [WO, function (v) { return ['correct', 'wrong', 'unanticipated', 'neutral'].indexOf(v) >= 0 || isDecimal(v); }],
    'cmi.interactions.n.latency': [WO, timespan12]
  };

  var CHILDREN_12 = {
    'cmi._children': 'core,suspend_data,launch_data,comments,objectives,student_data,student_preference,interactions',
    'cmi.core._children': 'student_id,student_name,lesson_location,credit,lesson_status,entry,score,total_time,lesson_mode,exit,session_time',
    'cmi.core.score._children': 'raw,min,max',
    'cmi.objectives._children': 'id,score,status',
    'cmi.objectives.n.score._children': 'raw,min,max',
    'cmi.student_data._children': 'mastery_score,max_time_allowed,time_limit_action',
    'cmi.student_preference._children': 'audio,language,speed,text',
    'cmi.interactions._children': 'id,objectives,time,type,correct_responses,weighting,student_response,result,latency'
  };

  var score04 = function (v) { return isDecimal(v); };
  var MODEL_2004 = {
    'cmi._version': [RO],
    'cmi.comments_from_learner._children': [RO], 'cmi.comments_from_learner._count': [RO],
    'cmi.comments_from_learner.n.comment': [RW, maxLen(4000)],
    'cmi.comments_from_learner.n.location': [RW, maxLen(250)],
    'cmi.comments_from_learner.n.timestamp': [RW, time04],
    'cmi.comments_from_lms._children': [RO], 'cmi.comments_from_lms._count': [RO],
    'cmi.comments_from_lms.n.comment': [RO], 'cmi.comments_from_lms.n.location': [RO], 'cmi.comments_from_lms.n.timestamp': [RO],
    'cmi.completion_status': [RW, vocab(['completed', 'incomplete', 'not attempted', 'unknown'])],
    'cmi.completion_threshold': [RO],
    'cmi.credit': [RO], 'cmi.entry': [RO],
    'cmi.exit': [WO, vocab(['time-out', 'suspend', 'logout', 'normal', ''])],
    'cmi.interactions._children': [RO], 'cmi.interactions._count': [RO],
    'cmi.interactions.n.id': [RW, maxLen(4000)],
    'cmi.interactions.n.type': [RW, vocab(['true-false', 'choice', 'fill-in', 'long-fill-in', 'matching', 'performance', 'sequencing', 'likert', 'numeric', 'other'])],
    'cmi.interactions.n.objectives._count': [RO],
    'cmi.interactions.n.objectives.n.id': [RW, maxLen(4000)],
    'cmi.interactions.n.timestamp': [RW, time04],
    'cmi.interactions.n.correct_responses._count': [RO],
    'cmi.interactions.n.correct_responses.n.pattern': [RW, any],
    'cmi.interactions.n.weighting': [RW, isDecimal],
    'cmi.interactions.n.learner_response': [RW, any],
    'cmi.interactions.n.result': [RW, function (v) { return ['correct', 'incorrect', 'unanticipated', 'neutral'].indexOf(v) >= 0 || isDecimal(v); }],
    'cmi.interactions.n.latency': [RW, duration04],
    'cmi.interactions.n.description': [RW, maxLen(250)],
    'cmi.launch_data': [RO], 'cmi.learner_id': [RO], 'cmi.learner_name': [RO],
    'cmi.learner_preference._children': [RO],
    'cmi.learner_preference.audio_level': [RW, function (v) { return isDecimal(v) && +v >= 0; }],
    'cmi.learner_preference.language': [RW, maxLen(250)],
    'cmi.learner_preference.delivery_speed': [RW, function (v) { return isDecimal(v) && +v >= 0; }],
    'cmi.learner_preference.audio_captioning': [RW, vocab(['-1', '0', '1'])],
    'cmi.location': [RW, maxLen(1000)],
    'cmi.max_time_allowed': [RO], 'cmi.mode': [RO],
    'cmi.objectives._children': [RO], 'cmi.objectives._count': [RO],
    'cmi.objectives.n.id': [RW, maxLen(4000)],
    'cmi.objectives.n.score._children': [RO],
    'cmi.objectives.n.score.scaled': [RW, inRange(-1, 1)],
    'cmi.objectives.n.score.raw': [RW, score04], 'cmi.objectives.n.score.min': [RW, score04], 'cmi.objectives.n.score.max': [RW, score04],
    'cmi.objectives.n.success_status': [RW, vocab(['passed', 'failed', 'unknown'])],
    'cmi.objectives.n.completion_status': [RW, vocab(['completed', 'incomplete', 'not attempted', 'unknown'])],
    'cmi.objectives.n.progress_measure': [RW, inRange(0, 1)],
    'cmi.objectives.n.description': [RW, maxLen(250)],
    'cmi.progress_measure': [RW, inRange(0, 1)],
    'cmi.scaled_passing_score': [RO],
    'cmi.score._children': [RO],
    'cmi.score.scaled': [RW, inRange(-1, 1)],
    'cmi.score.raw': [RW, score04], 'cmi.score.min': [RW, score04], 'cmi.score.max': [RW, score04],
    'cmi.session_time': [WO, duration04],
    'cmi.success_status': [RW, vocab(['passed', 'failed', 'unknown'])],
    'cmi.suspend_data': [RW, maxLen(64000)],
    'cmi.time_limit_action': [RO], 'cmi.total_time': [RO],
    'adl.nav.request': [RW, function (v) { return /^(continue|previous|exit|exitAll|abandon|abandonAll|suspendAll|_none_|\{target=[^}]+\}(choice|jump))$/.test(v); }],
    'adl.nav.request_valid.continue': [RO],
    'adl.nav.request_valid.previous': [RO]
  };

  var CHILDREN_2004 = {
    'cmi.comments_from_learner._children': 'comment,location,timestamp',
    'cmi.comments_from_lms._children': 'comment,location,timestamp',
    'cmi.interactions._children': 'id,type,objectives,timestamp,correct_responses,weighting,learner_response,result,latency,description',
    'cmi.learner_preference._children': 'audio_level,language,delivery_speed,audio_captioning',
    'cmi.objectives._children': 'id,score,success_status,completion_status,progress_measure,description',
    'cmi.objectives.n.score._children': 'scaled,raw,min,max',
    'cmi.score._children': 'scaled,raw,min,max'
  };

  var ERRORS_12 = {
    0: 'No error', 101: 'General exception', 201: 'Invalid argument error', 202: 'Element cannot have children',
    203: 'Element not an array - cannot have count', 301: 'Not initialized', 401: 'Not implemented error',
    402: 'Invalid set value, element is a keyword', 403: 'Element is read only', 404: 'Element is write only',
    405: 'Incorrect data type'
  };
  var ERRORS_2004 = {
    0: 'No Error', 101: 'General Exception', 102: 'General Initialization Failure', 103: 'Already Initialized',
    104: 'Content Instance Terminated', 111: 'General Termination Failure', 112: 'Termination Before Initialization',
    113: 'Termination After Termination', 122: 'Retrieve Data Before Initialization', 123: 'Retrieve Data After Termination',
    132: 'Store Data Before Initialization', 133: 'Store Data After Termination', 142: 'Commit Before Initialization',
    143: 'Commit After Termination', 201: 'General Argument Error', 301: 'General Get Failure', 351: 'General Set Failure',
    391: 'General Commit Failure', 401: 'Undefined Data Model Element', 402: 'Unimplemented Data Model Element',
    403: 'Data Model Element Value Not Initialized', 404: 'Data Model Element Is Read Only',
    405: 'Data Model Element Is Write Only', 406: 'Data Model Element Type Mismatch',
    407: 'Data Model Element Value Out Of Range', 408: 'Data Model Dependency Not Established'
  };

  // 2004 elements that always have a value (others answer 403 until set)
  var ALWAYS_INIT_2004 = ['cmi.completion_status', 'cmi.success_status', 'cmi.credit', 'cmi.entry', 'cmi.mode',
    'cmi.learner_id', 'cmi.learner_name', 'cmi.total_time', 'cmi._version', 'cmi.time_limit_action',
    'adl.nav.request', 'adl.nav.request_valid.continue', 'adl.nav.request_valid.previous'];

  // ------------------------------------------------------------- runtime
  /**
   * options: {
   *   version: '1.2' | '2004', cmi: {flat map}, learner: {id, name},
   *   masteryScore, scaledPassingScore, completionThreshold, launchData,
   *   onProgress(cmi, reason), onNav(request), onLog(level, message)
   * }
   */
  function ScormRuntime(options) {
    this.o = options;
    this.is12 = options.version === '1.2';
    this.model = this.is12 ? MODEL_12 : MODEL_2004;
    this.children = this.is12 ? CHILDREN_12 : CHILDREN_2004;
    this.errors = this.is12 ? ERRORS_12 : ERRORS_2004;
    this.state = 'new'; // new | running | terminated
    this.lastError = 0;
    this.lastDiag = '';
    this.data = {};
    this.startedAt = 0;
    this._prepareData(options.cmi || {});
  }

  ScormRuntime.prototype._prepareData = function (saved) {
    var d = this.data, k;
    for (k in saved) { if (Object.prototype.hasOwnProperty.call(saved, k)) { d[k] = String(saved[k]); } }
    var o = this.o, hadData = Object.keys(saved).length > 0;

    if (this.is12) {
      var exit = d['cmi.core.exit'] || '';
      d['cmi.core.entry'] = !hadData ? 'ab-initio' : (exit === 'suspend' ? 'resume' : '');
      d['cmi.core.student_id'] = o.learner.id;
      d['cmi.core.student_name'] = o.learner.name;
      d['cmi.core.credit'] = 'credit';
      d['cmi.core.lesson_mode'] = 'normal';
      if (!d['cmi.core.lesson_status']) { d['cmi.core.lesson_status'] = 'not attempted'; }
      if (!d['cmi.core.total_time']) { d['cmi.core.total_time'] = '0000:00:00.00'; }
      d['cmi.launch_data'] = o.launchData || '';
      d['cmi.student_data.mastery_score'] = o.masteryScore !== undefined && o.masteryScore !== null ? String(o.masteryScore) : '';
      d['cmi.student_data.max_time_allowed'] = '';
      d['cmi.student_data.time_limit_action'] = 'continue,no message';
      d['cmi.comments_from_lms'] = d['cmi.comments_from_lms'] || '';
      delete d['cmi.core.exit']; delete d['cmi.core.session_time'];
    } else {
      var exit04 = d['cmi.exit'] || '';
      d['cmi.entry'] = !hadData ? 'ab-initio' : (exit04 === 'suspend' || exit04 === 'logout' ? 'resume' : '');
      d['cmi._version'] = '1.0';
      d['cmi.learner_id'] = o.learner.id;
      d['cmi.learner_name'] = o.learner.name;
      d['cmi.credit'] = 'credit';
      d['cmi.mode'] = 'normal';
      if (!d['cmi.completion_status']) { d['cmi.completion_status'] = 'unknown'; }
      if (!d['cmi.success_status']) { d['cmi.success_status'] = 'unknown'; }
      if (!d['cmi.total_time']) { d['cmi.total_time'] = 'PT0H0M0S'; }
      d['cmi.time_limit_action'] = 'continue,no message';
      if (o.launchData) { d['cmi.launch_data'] = o.launchData; }
      if (o.scaledPassingScore !== undefined && o.scaledPassingScore !== null) { d['cmi.scaled_passing_score'] = String(o.scaledPassingScore); }
      if (o.completionThreshold !== undefined && o.completionThreshold !== null) { d['cmi.completion_threshold'] = String(o.completionThreshold); }
      d['adl.nav.request'] = '_none_';
      d['adl.nav.request_valid.continue'] = 'unknown';
      d['adl.nav.request_valid.previous'] = 'unknown';
      delete d['cmi.exit']; delete d['cmi.session_time'];
    }
  };

  ScormRuntime.prototype._err = function (code, diag) {
    this.lastError = code; this.lastDiag = diag || '';
    if (code) { this.o.onLog('warn', 'Error ' + code + ' (' + this.errors[code] + ')' + (diag ? ': ' + diag : '')); }
    return code;
  };

  ScormRuntime.prototype.initialize = function (arg) {
    if (arg !== '' && arg !== undefined && arg !== null) { this._err(201, 'Initialize argument must be ""'); return 'false'; }
    if (this.state === 'running') { this._err(this.is12 ? 101 : 103); return 'false'; }
    if (this.state === 'terminated') { this._err(this.is12 ? 101 : 104); return 'false'; }
    this.state = 'running';
    this.startedAt = Date.now();
    this._err(0);
    return 'true';
  };

  ScormRuntime.prototype.terminate = function (arg) {
    if (arg !== '' && arg !== undefined && arg !== null) { this._err(201); return 'false'; }
    if (this.state === 'new') { this._err(this.is12 ? 301 : 112); return 'false'; }
    if (this.state === 'terminated') { this._err(this.is12 ? 101 : 113); return 'false'; }
    this._finalize();
    this.state = 'terminated';
    this.o.onProgress(this.snapshot(), 'terminate');
    var nav = this.is12 ? 'finish' : (this.data['adl.nav.request'] || '_none_');
    this.o.onNav(nav);
    this._err(0);
    return 'true';
  };

  /** Called by the host when the SCO is unloaded without calling Finish/Terminate. */
  ScormRuntime.prototype.abandon = function () {
    if (this.state !== 'running') { return; }
    this._finalize();
    this.state = 'terminated';
    this.o.onProgress(this.snapshot(), 'unload');
  };

  ScormRuntime.prototype._finalize = function () {
    var d = this.data;
    if (this.is12) {
      var session = d['cmi.core.session_time'] ? parse12(d['cmi.core.session_time']) : (Date.now() - this.startedAt) / 1000;
      d['cmi.core.total_time'] = format12(parse12(d['cmi.core.total_time']) + session);
      if (d['cmi.core.lesson_status'] === 'not attempted') { d['cmi.core.lesson_status'] = 'incomplete'; }
      this._applyMastery12();
      // Keep exit so that the next launch knows whether to resume.
      d['cmi.core.exit'] = d['cmi.core.exit'] || '';
      delete d['cmi.core.session_time'];
    } else {
      var s04 = d['cmi.session_time'] ? parse04(d['cmi.session_time']) : (Date.now() - this.startedAt) / 1000;
      d['cmi.total_time'] = format04(parse04(d['cmi.total_time']) + s04);
      d['cmi.completion_status'] = this._completion04();
      d['cmi.success_status'] = this._success04();
      d['cmi.exit'] = d['cmi.exit'] || '';
      delete d['cmi.session_time'];
    }
  };

  ScormRuntime.prototype._applyMastery12 = function () {
    var d = this.data, ms = d['cmi.student_data.mastery_score'], raw = d['cmi.core.score.raw'];
    if (ms !== '' && ms !== undefined && raw !== undefined && raw !== '' && d['cmi.core.credit'] === 'credit') {
      d['cmi.core.lesson_status'] = +raw >= +ms ? 'passed' : 'failed';
    }
  };

  ScormRuntime.prototype._completion04 = function () {
    var d = this.data, ct = d['cmi.completion_threshold'], pm = d['cmi.progress_measure'];
    if (ct !== undefined && pm !== undefined && pm !== '') { return +pm >= +ct ? 'completed' : 'incomplete'; }
    return d['cmi.completion_status'] || 'unknown';
  };

  ScormRuntime.prototype._success04 = function () {
    var d = this.data, sps = d['cmi.scaled_passing_score'], sc = d['cmi.score.scaled'];
    if (sps !== undefined && sc !== undefined && sc !== '') { return +sc >= +sps ? 'passed' : 'failed'; }
    return d['cmi.success_status'] || 'unknown';
  };

  ScormRuntime.prototype.getValue = function (el) {
    el = el === undefined || el === null ? '' : String(el);
    if (this.state === 'new') { this._err(this.is12 ? 301 : 122); return ''; }
    if (this.state === 'terminated') { this._err(this.is12 ? 301 : 123); return ''; }
    if (!el) { this._err(this.is12 ? 201 : 301, 'Empty element name'); return ''; }

    var t = template(el);
    if (/\._children$/.test(el)) {
      if (this.children[t] === undefined) { this._err(this.is12 ? 202 : 301, el + ' has no children'); return ''; }
      this._err(0); return this.children[t];
    }
    if (/\._count$/.test(el)) {
      if (!this.model[t]) { this._err(this.is12 ? 203 : 301, el + ' is not a collection'); return ''; }
      this._err(0); return String(this._count(el.replace(/\._count$/, '')));
    }
    var spec = this.model[t];
    if (!spec) {
      if (/^adl\.nav\.request_valid\.choice\./.test(el)) { this._err(0); return 'unknown'; }
      this._err(this.is12 ? 201 : 401, el); return '';
    }
    if (spec[0] === WO) { this._err(this.is12 ? 404 : 405, el); return ''; }

    if (!this.is12) {
      if (el === 'cmi.completion_status') { this._err(0); return this._completion04(); }
      if (el === 'cmi.success_status') { this._err(0); return this._success04(); }
    }
    var v = this.data[el];
    if (v === undefined) {
      if (!this.is12 && ALWAYS_INIT_2004.indexOf(el) < 0) { this._err(403, el); return ''; }
      this._err(0); return '';
    }
    this._err(0);
    return v;
  };

  ScormRuntime.prototype.setValue = function (el, value) {
    el = el === undefined || el === null ? '' : String(el);
    value = value === undefined || value === null ? '' : String(value);
    if (this.state === 'new') { this._err(this.is12 ? 301 : 132); return 'false'; }
    if (this.state === 'terminated') { this._err(this.is12 ? 301 : 133); return 'false'; }
    if (!el) { this._err(this.is12 ? 201 : 351, 'Empty element name'); return 'false'; }

    if (/\._(children|count|version)$/.test(el)) { this._err(this.is12 ? 402 : 404, el); return 'false'; }
    var t = template(el), spec = this.model[t];
    if (!spec) { this._err(this.is12 ? 201 : 401, el); return 'false'; }
    if (spec[0] === RO) { this._err(this.is12 ? 403 : 404, el); return 'false'; }

    // Collection index checks: index must be <= current count (appending is allowed)
    var segs = el.split('.');
    for (var i = 1; i < segs.length; i++) {
      if (/^\d+$/.test(segs[i])) {
        var coll = segs.slice(0, i).join('.'), idx = +segs[i], count = this._count(coll);
        if (idx > count) { this._err(this.is12 ? 201 : 351, 'Index ' + idx + ' out of order for ' + coll); return 'false'; }
      }
    }

    if (spec[1] && !spec[1](value)) {
      var range = /score|progress_measure|audio|speed|text/.test(t) && isDecimal(value);
      this._err(this.is12 ? 405 : (range ? 407 : 406), el + ' = "' + value + '"');
      return 'false';
    }
    if (this.is12 && t === 'cmi.core.lesson_status' && value === 'not attempted') { this._err(405); return 'false'; }

    this.data[el] = value;
    this._err(0);
    return 'true';
  };

  ScormRuntime.prototype.commit = function (arg) {
    if (arg !== '' && arg !== undefined && arg !== null) { this._err(201); return 'false'; }
    if (this.state === 'new') { this._err(this.is12 ? 301 : 142); return 'false'; }
    if (this.state === 'terminated') { this._err(this.is12 ? 301 : 143); return 'false'; }
    this.o.onProgress(this.snapshot(), 'commit');
    this._err(0);
    return 'true';
  };

  ScormRuntime.prototype._count = function (collection) {
    var prefix = collection + '.', max = -1, k;
    for (k in this.data) {
      if (k.indexOf(prefix) === 0) {
        var n = parseInt(k.substring(prefix.length), 10);
        if (!isNaN(n) && n > max) { max = n; }
      }
    }
    return max + 1;
  };

  /** Copy of the CMI data to persist (computed values resolved, session values excluded). */
  ScormRuntime.prototype.snapshot = function () {
    var out = {}, k;
    for (k in this.data) { if (Object.prototype.hasOwnProperty.call(this.data, k)) { out[k] = this.data[k]; } }
    if (!this.is12) {
      out['cmi.completion_status'] = this._completion04();
      out['cmi.success_status'] = this._success04();
    } else if (this.state === 'running') {
      // Report mastery-based status on commit too
      var ms = out['cmi.student_data.mastery_score'], raw = out['cmi.core.score.raw'];
      if (ms && raw) { out['cmi.core.lesson_status'] = +raw >= +ms ? 'passed' : 'failed'; }
    }
    return out;
  };

  /** Builds the object the SCO talks to. */
  ScormRuntime.prototype.createApi = function () {
    var rt = this, log = function (call, args, result) {
      rt.o.onLog('api', call + '(' + args.map(function (a) { return JSON.stringify(a); }).join(', ') + ') → ' + JSON.stringify(result));
      return result;
    };
    if (this.is12) {
      return {
        LMSInitialize: function (a) { return log('LMSInitialize', [a], rt.initialize(a)); },
        LMSFinish: function (a) { return log('LMSFinish', [a], rt.terminate(a)); },
        LMSGetValue: function (e) { return log('LMSGetValue', [e], rt.getValue(e)); },
        LMSSetValue: function (e, v) { return log('LMSSetValue', [e, v], rt.setValue(e, v)); },
        LMSCommit: function (a) { return log('LMSCommit', [a], rt.commit(a)); },
        LMSGetLastError: function () { return String(rt.lastError); },
        LMSGetErrorString: function (c) { return ERRORS_12[c] || ''; },
        LMSGetDiagnostic: function (c) { return c === undefined || c === '' || +c === rt.lastError ? rt.lastDiag : (ERRORS_12[c] || ''); }
      };
    }
    return {
      version: '1.0',
      Initialize: function (a) { return log('Initialize', [a], rt.initialize(a)); },
      Terminate: function (a) { return log('Terminate', [a], rt.terminate(a)); },
      GetValue: function (e) { return log('GetValue', [e], rt.getValue(e)); },
      SetValue: function (e, v) { return log('SetValue', [e, v], rt.setValue(e, v)); },
      Commit: function (a) { return log('Commit', [a], rt.commit(a)); },
      GetLastError: function () { return String(rt.lastError); },
      GetErrorString: function (c) { return ERRORS_2004[c] || ''; },
      GetDiagnostic: function (c) { return c === undefined || c === '' || +c === rt.lastError ? rt.lastDiag : (ERRORS_2004[c] || ''); }
    };
  };

  global.ScormRuntime = ScormRuntime;
})(typeof window !== 'undefined' ? window : globalThis);
