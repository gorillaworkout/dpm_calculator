#!/usr/bin/env python3
"""Browser chunk-retry logic in templates/tool.html, run under node with fakes.

The core between '--- chunk-retry core' and '--- end chunk-retry core' has no DOM,
so it can be exercised here: backoff with jitter, waiting for the network,
resuming from /upload/status after a lost ack, login/redirect detection, and
the diagnostics shown in the final error. Skips when node is not installed.

python3 test_upload_retry.py
"""
import base64
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NODE = shutil.which("node") or shutil.which("nodejs")

HARNESS = r"""
var assert = require('assert');
function resp(status, body, opts) {
  opts = opts || {};
  return {status: status, ok: status >= 200 && status < 300, type: opts.type || 'basic',
          redirected: !!opts.redirected, url: opts.url || 'https://dw.example/upload/chunk',
          headers: {get: function () { return opts.ct || 'application/json'; }},
          json: function () { return body === undefined ? Promise.reject(new Error('no json')) : Promise.resolve(body); }};
}
function netErr() { var e = new TypeError('Failed to fetch'); return e; }
function env(script, o) {
  o = o || {};
  var t = 1000000, calls = [], delays = [], retries = [], waits = 0, statusCalls = 0;
  var onlineSeq = (o.online || []).slice();
  var e = {
    calls: calls, delays: delays, retries: retries,
    get waits() { return waits; }, get statusCalls() { return statusCalls; },
    fetch: function (url, opts) {
      calls.push({url: url, opts: opts});
      assert.strictEqual(opts.redirect, 'manual');
      var step = script[calls.length - 1];
      if (!step) step = script[script.length - 1];
      t += 1000;
      return typeof step === 'function' ? step() : (step instanceof Error ? Promise.reject(step) : Promise.resolve(step));
    },
    sleep: function (ms, tick) { delays.push(ms); tick(ms); t += ms; return Promise.resolve(); },
    waitOnline: function () { waits++; return Promise.resolve(); },
    online: function () { return onlineSeq.length ? onlineSeq.shift() : true; },
    now: function () { return t; },
    random: function () { return o.rand === undefined ? 0.5 : o.rand; },
    status: function () { statusCalls++; return o.status ? Promise.resolve(o.status()) : Promise.resolve(null); },
    onRetry: function (info) { retries.push(info); },
    origin: 'https://dw.example'
  };
  return e;
}
var REQ = {id: 'abcdef0123456789abcdef0123456789', index: 0, seq: 0, size: 2 * 1048576,
           makeBody: function () { return 'body'; }};
var tests = [];
function test(name, fn) { tests.push([name, fn]); }

test('delay: exponential, jitter 50-100%, capped at 30 s', function () {
  assert.strictEqual(retryDelay(1, function () { return 0; }), 500);
  assert.strictEqual(retryDelay(1, function () { return 1; }), 1000);
  assert.strictEqual(retryDelay(4, function () { return 1; }), 8000);
  assert.strictEqual(retryDelay(6, function () { return 1; }), 30000);
  assert.strictEqual(retryDelay(8, function () { return 0; }), 15000);
  assert.strictEqual(RETRY.max, 8);
});

test('two network errors then success', function () {
  var e = env([netErr(), netErr(), resp(200, {received: 10})]);
  return makeChunkSender(e)(REQ).then(function (r) {
    assert.strictEqual(r.received, 10);
    assert.strictEqual(r.diag.attempts, 3);
    assert.strictEqual(e.calls.length, 3);
    assert.deepStrictEqual(e.delays, [750, 1500]);
    assert.deepStrictEqual(e.retries.map(function (x) { return x.n + '/' + x.max; }), ['1/8', '2/8']);
    assert.strictEqual(e.statusCalls, 2, 'asks the server before each re-send');
  });
});

test('lost ack: the server already has the chunk, it is not sent again', function () {
  var e = env([netErr()], {status: function () { return {files: [{index: 0, chunks: 1, bytes: 5}]}; }});
  return makeChunkSender(e)(REQ).then(function (r) {
    assert.strictEqual(r.resumed, true);
    assert.strictEqual(e.calls.length, 1);
  });
});

test('gives up after 8 retries with diagnostics', function () {
  var e = env([netErr()], {rand: 1});
  return makeChunkSender(e)(REQ).then(function () { assert.fail('should fail'); }, function (err) {
    assert.strictEqual(err.diag.attempts, 9);
    assert.strictEqual(e.calls.length, 9);
    assert.deepStrictEqual(e.delays, [1000, 2000, 4000, 8000, 16000, 30000, 30000, 30000]);
    var text = diagText(err, 0, true);
    ['error=TypeError: Failed to fetch', 'online=yes', 'attempts=9 of 9', 'elapsed=2m', 'chunk=2.00 MB',
     'response=none', 'upload=abcdef01'].forEach(function (s) {
      assert.ok(text.indexOf(s) >= 0, s + ' missing in: ' + text);
    });
  });
});

test('offline: waits for the online event before retrying', function () {
  var e = env([netErr(), resp(200, {received: 1})], {online: [false]});
  return makeChunkSender(e)(REQ).then(function () {
    assert.strictEqual(e.waits, 1);
    assert.ok(e.retries[0].offline);
  });
});

test('redirect (Cloudflare Access / login): session expired, no retries', function () {
  var e = env([resp(0, undefined, {type: 'opaqueredirect'})]);
  return makeChunkSender(e)(REQ).then(function () { assert.fail(); }, function (err) {
    assert.ok(err.sessionExpired);
    assert.strictEqual(e.calls.length, 1);
    assert.ok(diagText(err, 0, true).indexOf('type=opaqueredirect redirected=no opaque') >= 0, diagText(err, 0, true));
  });
});

test('401 and HTML-instead-of-JSON are session problems', function () {
  var a = env([resp(401, {})]);
  var b = env([resp(200, undefined, {ct: 'text/html; charset=utf-8'})]);
  var c = env([resp(200, undefined, {redirected: true, url: 'https://team.cloudflareaccess.com/cdn-cgi/access/login?x=1', ct: 'text/html'})]);
  return Promise.all([a, b, c].map(function (e) {
    return makeChunkSender(e)(REQ).then(function () { assert.fail(); }, function (err) {
      assert.ok(err.sessionExpired, err.message);
      assert.strictEqual(e.calls.length, 1);
    });
  }));
});

test('server errors and 408 retry; 413 stops at once with the server text', function () {
  var e = env([resp(502, undefined), resp(408, {retry: true, error: 'incomplete'}), resp(200, {received: 3})]);
  var f = env([resp(413, {error: 'Too big.'})]);
  return makeChunkSender(e)(REQ).then(function (r) {
    assert.strictEqual(r.diag.attempts, 3);
    assert.strictEqual(e.retries[0].err.name, 'HTTP 502');
    return makeChunkSender(f)(REQ);
  }).then(function () { assert.fail(); }, function (err) {
    assert.strictEqual(err.message, 'Too big.');
    assert.ok(err.noRetry);
    assert.strictEqual(f.calls.length, 1);
    assert.ok(diagText(err, 0, true).indexOf('response=HTTP 413 type=basic redirected=no') >= 0);
  });
});

test('409 resolves from server state, or stops', function () {
  var ok = env([resp(409, {expected: 1})], {status: function () { return {files: [{index: 0, chunks: 1}]}; }});
  var bad = env([resp(409, {error: 'Chunk arrived out of order.', expected: 0})]);
  return makeChunkSender(ok)(REQ).then(function (r) {
    assert.ok(r.resumed);
    return makeChunkSender(bad)(REQ);
  }).then(function () { assert.fail(); }, function (err) {
    assert.strictEqual(err.message, 'Chunk arrived out of order.');
  });
});

test('a failing status check does not block the retry', function () {
  var e = env([netErr(), resp(200, {received: 1})], {status: function () { throw new Error('x'); }});
  return makeChunkSender(e)(REQ).then(function (r) { assert.strictEqual(r.diag.attempts, 2); });
});

(function run(i) {
  if (i >= tests.length) { console.log('node retry OK ' + tests.length); return; }
  Promise.resolve().then(tests[i][1]).then(function () { run(i + 1); }, function (err) {
    console.error('FAIL ' + tests[i][0] + ': ' + (err && err.stack || err));
    process.exit(1);
  });
})(0);
"""


def core():
    html = (ROOT / "templates" / "tool.html").read_text(encoding="utf-8")
    m = re.search(r"// --- chunk-retry core.*?\n(.*?)// --- end chunk-retry core ---", html, re.S)
    assert m, "chunk-retry core markers missing in tool.html"
    assert "{{" not in m.group(1) and "{%" not in m.group(1), "core must not use Jinja"
    return m.group(1)


def run_node(code):
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "t.js"
        p.write_text(code, encoding="utf-8")
        r = subprocess.run([NODE, str(p)], capture_output=True, text=True, timeout=60)
        assert r.returncode == 0, r.stdout + r.stderr
        return r.stdout


def test_retry_core():
    if not NODE:
        print("SKIP: node not installed")
        return
    print(run_node(core() + HARNESS).strip())


def test_rendered_page_script_parses():
    """The whole rendered <script> (Jinja filled in) must be valid JavaScript."""
    if not NODE:
        print("SKIP: node not installed")
        return
    tmp = Path(tempfile.mkdtemp())
    os.environ.setdefault("DW_AUTH_DB", str(tmp / "t.db"))
    os.environ.setdefault("DW_JOB_DIR", str(tmp / "jobs"))
    os.environ.setdefault("DW_ADMIN_USER", "tester")
    os.environ.setdefault("DW_ADMIN_PASS", "tester-password-1")
    from app import app
    app.config["TESTING"] = True
    c = app.test_client()
    c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(
        f"{os.environ['DW_ADMIN_USER']}:{os.environ['DW_ADMIN_PASS']}".encode()).decode()
    for path in ("/kvb/tool/segregate", "/tool/segregate", "/tool/dw"):
        html = c.get(path, headers={"Accept": "text/html"}).data.decode()
        scripts = re.findall(r"<script>(.*?)</script>", html, re.S)
        assert scripts, path
        with tempfile.TemporaryDirectory() as d:
            for i, s in enumerate(scripts):
                p = Path(d) / f"s{i}.js"
                p.write_text(s, encoding="utf-8")
                r = subprocess.run([NODE, "--check", str(p)], capture_output=True, text=True)
                assert r.returncode == 0, (path, i, r.stderr)
    print("OK rendered scripts parse")


if __name__ == "__main__":
    test_retry_core()
    test_rendered_page_script_parses()
    print("OK upload retry")
