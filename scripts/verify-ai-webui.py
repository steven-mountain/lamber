#!/usr/bin/env python3
"""Run alpha.5's complete WebUI in an isolated Lamber native probe window.

Default uses a deterministic local SSE model. --real --app-config PATH reads
only the existing AI settings; credentials enter the child environment only.
No production workspace, config, session history, bridge or business IPC is
opened. The runner owns and reaps both processes. Ctrl-C or closing the native
window stops the Host. This is stage-0 evidence, never full migration sign-off.
"""
import argparse
import hashlib
import http.server
import http.cookiejar
import json
import os
import plistlib
from pathlib import Path
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request

ROOT = Path(__file__).resolve().parent.parent
BRIDGE = ROOT / 'agent-bridge'
VERSION = '0.1.2-alpha.5'


def runtime_inventory():
    lock = json.loads((BRIDGE / 'package-lock.json').read_text())['packages']
    packages = ['dsh', 'dsh-web-app', 'dsh-web-frontend', 'dsh-client-modules',
                'dsh-client-connection', 'dsh-api-session-controller',
                'dsh-client-ui-chat', 'dsh-client-ui-approval', 'dsh-client-ui-theme']
    versions = {}
    dsh_count = 0
    for key, locked in lock.items():
        if '/@deepseek-ai/dsh' not in key:
            continue
        metadata = json.loads((BRIDGE / key / 'package.json').read_text())
        if metadata['version'] != locked['version']:
            raise RuntimeError(f'Installed dependency differs from lock: {metadata["name"]}')
        if metadata['name'].startswith('@deepseek-ai/dsh'):
            if metadata['version'] != VERSION:
                raise RuntimeError(f'Mixed dsh runtime: {metadata["name"]}')
            dsh_count += 1
    for name in packages:
        key = f'node_modules/@deepseek-ai/{name}'
        actual = json.loads((BRIDGE / key / 'package.json').read_text())['version']
        if lock[key]['version'] != VERSION or actual != VERSION:
            raise RuntimeError(f'WebUI runtime version mismatch: {name}')
        versions[name] = actual
    entry = BRIDGE / 'node_modules/@deepseek-ai/dsh-web-frontend/dist/index.html'
    return {'packages': versions, 'lockedDshPackageCount': dsh_count,
            'frontendIndexSha256': hashlib.sha256(entry.read_bytes()).hexdigest()}


def verify_transport(url):
    """Exercise the real Host boundary without exporting its launch credential."""
    origin = url.split('/?')[0]
    def status(opener, request):
        try:
            with opener.open(request, timeout=10) as response:
                return response.status, response.read().decode()
        except urllib.error.HTTPError as error:
            return error.code, ''
    plain = urllib.request.build_opener()
    unauth, _ = status(plain, origin + '/api/session/list')
    assert unauth == 401, f'unauthenticated API must reject, got {unauth}'
    jar = http.cookiejar.CookieJar()
    authenticated = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    code, html = status(authenticated, url)
    assert code == 200 and '__DSH_BOOT__' in html, 'full official boot graph missing'
    for package in ['dsh-client-ui-chat', 'dsh-client-ui-session', 'dsh-client-ui-settings',
                    'dsh-client-ui-approval', 'dsh-client-ui-attachment', 'dsh-client-ui-lamber-brand']:
        assert package in html, f'official client package missing: {package}'
    cross_origin, _ = status(authenticated, urllib.request.Request(
        origin + '/api/session/list', headers={'Origin': 'https://invalid.example'}))
    assert cross_origin == 403, f'cross-origin request must reject, got {cross_origin}'
    return {'unauthenticatedApi': unauth, 'crossOriginApi': cross_origin,
            'authenticatedBootGraph': True}


class Model(http.server.BaseHTTPRequestHandler):
    def log_message(self, *_args):
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
        # A response is deterministic but still crosses the real agent loop,
        # Session log, HTTP/WebSocket transport and upstream UI projection.
        prompt = json.dumps(body.get('messages', []), ensure_ascii=False)
        slow = '长回复' in prompt
        parts = ['官方 WebUI ', '中文🙂流式验证。\n\n', '| 项目 | 金额 |\n|---|---:|\n',
                 '| 合成样例 | 1,234.56 元 |\n\n', '本轮不执行任何业务写入。']
        if slow:
            parts = [f'第 {i} 条：仅用于验证停止按钮与流式显示。\n' for i in range(1, 81)]
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.end_headers()
        try:
            for part in parts:
                value = {'id': 'webui-probe', 'object': 'chat.completion.chunk', 'created': 1,
                         'model': 'deepseek-v4-flash', 'choices': [{'index': 0,
                         'delta': {'content': part}, 'finish_reason': None}]}
                self.wfile.write(('data: ' + json.dumps(value, ensure_ascii=False) + '\n\n').encode())
                self.wfile.flush()
                time.sleep(.25)
            value['choices'] = [{'index': 0, 'delta': {}, 'finish_reason': 'stop'}]
            self.wfile.write(('data: ' + json.dumps(value) + '\n\ndata: [DONE]\n\n').encode())
            self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError):
            pass


def stop(process):
    if process is None or process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=8)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=8)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--app', type=Path, default=ROOT / 'src-tauri/target/debug/benefit-calculator')
    parser.add_argument('--real', action='store_true')
    parser.add_argument('--app-config', type=Path)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    inventory = runtime_inventory()
    settings = {}
    if args.real:
        if not args.app_config:
            parser.error('--real requires --app-config; no implicit credential search')
        settings = json.loads(args.app_config.read_text()).get('aiAgent', {})
        if not settings.get('apiKey'):
            parser.error('Configured real-model credential is missing; refusing to skip')
    if not args.app.is_file():
        parser.error('Build the debug application with cargo build first')
    server = http.server.ThreadingHTTPServer(('127.0.0.1', 0), Model)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    host = app = None
    with tempfile.TemporaryDirectory(prefix='lamber-webui-probe-') as temp:
        root = Path(temp)
        home, cwd = root / 'home', root / '合成工作区 With Spaces'
        cwd.mkdir()
        profile = home / 'profiles/web'
        modules = profile / 'node_modules'
        modules.mkdir(parents=True)
        for source, name in [(BRIDGE / 'webui/lamber-brand', 'dsh-client-ui-lamber-brand'),
                             (ROOT / 'scripts/fixtures/webui-probe', 'lamber-webui-probe-policy')]:
            shutil.copytree(source, modules / name)
        (profile / 'package.json').write_text(json.dumps({'name': 'lamber-webui-probe', 'private': True,
            'dsh': {'profile': {'bundles': ['@deepseek-ai/dsh-base', '@deepseek-ai/dsh-web-app'],
                                'patchReload': 'startup'}}}))
        (profile / 'cordis.yml').write_text('[]\n')
        (profile / 'cordis.patch.yml').write_text('[]\n')
        preset = root / 'presets/lamber-probe'
        preset.mkdir(parents=True)
        (preset / 'preset.yml').write_text('name: Lamber 界面验证\ndescription: 隔离的只读聊天承载验证\n')
        (preset / 'agent.cordis.yml').write_text('''- id: persona
  name: '@deepseek-ai/dsh-persona'
  config:
    text: 你是 Lamber 的中文助手。本次仅验证官方聊天界面，不调用任何工具。
    complete: true
    includeRuntimeContext: false
''')
        endpoint = settings.get('baseUrl', 'https://api.deepseek.com') if args.real else f'http://127.0.0.1:{server.server_port}'
        patch = root / 'probe.patch.yml'
        patch.write_text(f'''- id: llm-deepseek
  config:
    baseURL: {json.dumps(endpoint)}
- id: agent-presets
  config:
    default: lamber-probe
    includeShippedRoot: false
    includeUserRoot: false
    roots:
      - path: {json.dumps(str(root / 'presets'))}
        trust: system
- id: ui-brand-official
  disabled: true
- id: directory-picker
  disabled: true
- id: agent-default-model
  config:
    provider: deepseek-official
    model: {json.dumps(settings.get('model', 'deepseek-v4-flash'))}
- insert:
    - id: lamber-brand
      name: dsh-client-ui-lamber-brand
    - id: lamber-probe-policy
      name: lamber-webui-probe-policy
''')
        env = {k: v for k, v in os.environ.items() if not k.startswith(('DSH_', 'LAMBER_', 'DEEPSEEK_'))}
        env.update(DSH_HOME=str(home), DSH_TELEMETRY_MODE='DISABLED',
                   DEEPSEEK_API_KEY=settings.get('apiKey', 'local-webui-probe'))
        report = {**inventory, 'mode': 'real' if args.real else 'controlled-sse',
                  'businessBridge': False, 'nativeBusinessIpc': False, 'nativeExitedCleanly': False}
        log_path = root / 'host.log'
        try:
            with log_path.open('w') as log:
                host = subprocess.Popen(['node', str(BRIDGE / 'node_modules/@deepseek-ai/dsh/lib/bin.js'),
                    '--profile', 'web', '--patch', str(patch), '--no-open', '--host', '127.0.0.1', '--port', '0'],
                    cwd=cwd, env=env, stdout=log, stderr=log)
            deadline = time.monotonic() + 45
            url = None
            while time.monotonic() < deadline and host.poll() is None:
                match = re.search(r'dsh web: (http://127\.0\.0\.1:\d+/\?token=\S+)', log_path.read_text())
                if match:
                    url = match.group(1)
                    break
                time.sleep(.1)
            if not url:
                diagnostic = log_path.read_text()[:1800]
                diagnostic = diagnostic.replace(env['DEEPSEEK_API_KEY'], '[REDACTED]')
                diagnostic = re.sub(r'token=[^\s&]+', 'token=[REDACTED]', diagnostic)
                diagnostic = diagnostic.replace(str(root), '<probe>')
                diagnostic = diagnostic.replace(str(ROOT), '<repo>')
                report['startupDiagnostic'] = diagnostic
                raise RuntimeError('Official WebUI startup failed: ' + diagnostic)
            report['transportChecks'] = verify_transport(url)
            env.pop('DEEPSEEK_API_KEY', None)
            env['LAMBER_WEBUI_PROBE_URL'] = url
            executable = args.app.resolve()
            if sys.platform == 'darwin':
                bundle = root / 'Lamber WebUI Probe.app'
                macos = bundle / 'Contents/MacOS'
                macos.mkdir(parents=True)
                shutil.copy2(executable, macos / 'lamber-webui-probe')
                (bundle / 'Contents/Info.plist').write_bytes(plistlib.dumps({
                    'CFBundleIdentifier': 'com.lamber.webui-probe',
                    'CFBundleName': 'Lamber WebUI Probe',
                    'CFBundleDisplayName': 'Lamber WebUI Probe',
                    'CFBundleExecutable': 'lamber-webui-probe',
                    'CFBundlePackageType': 'APPL',
                    'CFBundleVersion': '1',
                    'NSHighResolutionCapable': True,
                }))
                executable = macos / 'lamber-webui-probe'
            with (root / 'native.log').open('w') as log:
                app = subprocess.Popen([str(executable)], env=env, stdout=log, stderr=log)
            print(json.dumps({**report, 'status': 'native-window-started',
                              'nativeApp': str(bundle) if sys.platform == 'darwin' else str(executable),
                              'origin': url.split('/?')[0], 'workspace': str(cwd)}, ensure_ascii=False), flush=True)
            while app.poll() is None:
                if host.poll() is not None:
                    raise RuntimeError('WebUI Host exited while the native window was open')
                time.sleep(.2)
            report['nativeExitedCleanly'] = app.returncode == 0
        finally:
            stop(app)
            stop(host)
            server.shutdown()
            report['hostReaped'] = host is None or host.poll() is not None
            report['nativeReaped'] = app is None or app.poll() is not None
            if args.output:
                args.output.parent.mkdir(parents=True, exist_ok=True)
                args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
            print(json.dumps(report, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    signal.signal(signal.SIGTERM, lambda *_: (_ for _ in ()).throw(KeyboardInterrupt()))
    try:
        main()
    except KeyboardInterrupt:
        pass
