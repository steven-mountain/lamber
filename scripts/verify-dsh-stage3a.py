#!/usr/bin/env python3
"""Fail-loud alpha.5 subscription contract through the actual provisioned plugin.

Default: deterministic local SSE, no key/network. --real: official DeepSeek with
DEEPSEEK_API_KEY or --app-config PATH. Temporary home/workspace; no business writes.
Output contains only this synthetic prompt's display events, never credentials.
"""
import argparse
import http.server
import json
import os
from pathlib import Path
import queue
import shutil
import subprocess
import tempfile
import threading
import time
import uuid


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--real', action='store_true')
    parser.add_argument('--app-config', type=Path)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    base = Path(__file__).resolve().parent.parent / 'agent-bridge'
    key = os.environ.get('DEEPSEEK_API_KEY', '')
    if args.real and args.app_config:
        key = json.loads(args.app_config.read_text())['aiAgent'].get('apiKey', '')
    if args.real and not key:
        raise SystemExit('Real verification requires a key; refusing to skip')
    events, timings, commits, issues = [], [], [], []
    token = uuid.uuid4().hex
    binding = {'lamberSessionId': 'stage3a-contract', 'requestId': uuid.uuid4().hex}
    active_session = [None]
    start = time.monotonic()

    class Server(http.server.BaseHTTPRequestHandler):
        def log_message(self, *unused):
            pass

        def do_POST(self):
            body = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
            if self.path == '/lamber-bridge/handshake':
                assert self.headers.get('x-lamber-bridge-token') == token
                contract = json.loads((base / 'bridge-contract.json').read_text())
                assert body == contract
                data = json.dumps(contract).encode()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            if self.path == '/lamber-bridge/stream':
                assert self.headers.get('x-lamber-bridge-token') == token
                assert body['sessionId'] == active_session[0]
                if body['kind'] == 'bind':
                    result = binding
                else:
                    assert body['requestId'] == binding['requestId']
                    events.append({'method': 'session/stream', **binding, 'params': body})
                    if body['kind'] == 'delta':
                        timings.append(time.monotonic() - start)
                    if body['kind'] == 'error':
                        issues.append('Plugin reported stream failure')
                    result = {}
                data = json.dumps(result).encode()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            assert not args.real, 'Unexpected bridge route'
            self.send_response(200)
            self.send_header('Content-Type', 'text/event-stream')
            self.end_headers()
            # Deliberately split the SSE wire inside Chinese/emoji UTF-8 too.
            deltas = [{'reasoning_content': '逐步思考🙂'}, *({'content': text} for text in ['中文', '🙂', '金额', '1,234.56', '元。']), {}]
            for delta in deltas:
                data = {'id': 'stage3a', 'object': 'chat.completion.chunk', 'created': 1,
                        'model': 'deepseek-v4-flash', 'choices': [{'index': 0, 'delta': delta,
                        'finish_reason': None if delta else 'stop'}]}
                wire = ('data: ' + json.dumps(data, ensure_ascii=False) + '\n\n').encode()
                for byte in wire:
                    self.wfile.write(bytes([byte]))
                self.wfile.flush()
                time.sleep(.12)
            self.wfile.write(b'data: [DONE]\n\n')
            self.wfile.flush()

    server = http.server.ThreadingHTTPServer(('127.0.0.1', 0), Server)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        with tempfile.TemporaryDirectory(prefix='lamber-stage3a-') as temp:
            root = Path(temp)
            home = root / 'home'
            shutil.copytree(base / 'dsh-home-template', home)
            plugin = home / 'profiles/acp/node_modules/dsh-tool-lamber'
            plugin.mkdir(parents=True)
            shutil.copytree(base / 'dsh-tool-lamber/lib', plugin / 'lib')
            shutil.copy(base / 'dsh-tool-lamber/package.json', plugin / 'package.json')
            (plugin / 'node_modules').symlink_to(base / 'dsh-tool-lamber/node_modules', target_is_directory=True)
            patch = root / 'patch.yml'
            endpoint = 'https://api.deepseek.com' if args.real else f'http://127.0.0.1:{server.server_port}'
            patch.write_text((base / 'patch.yml').read_text() + f'\n- id: llm-deepseek\n  config:\n    baseURL: {endpoint}\n')
            env = dict(os.environ, DSH_HOME=str(home), DSH_TELEMETRY_MODE='DISABLED',
                       DEEPSEEK_API_KEY=key if args.real else 'local-contract', LAMBER_STREAM_DISPLAY='1',
                       LAMBER_BRIDGE_URL=f'http://127.0.0.1:{server.server_port}', LAMBER_BRIDGE_TOKEN=token)
            proc = subprocess.Popen(['node', str(base / 'node_modules/@deepseek-ai/dsh/lib/bin.js'),
                '--profile', 'acp', '--patch', str(patch)], stdin=subprocess.PIPE,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=env)
            messages = queue.Queue()

            def read_output():
                for line in proc.stdout:
                    try:
                        message = json.loads(line)
                    except ValueError:
                        continue
                    if message.get('method') == 'session/update':
                        params = message['params']
                        if params['update']['sessionUpdate'] in ['agent_message_chunk', 'agent_thought_chunk']:
                            commits.append(time.monotonic() - start)
                        events.append({'method': 'session/update', **binding, 'params': params})
                    messages.put(message)

            def read_errors():
                for line in proc.stderr:
                    if 'lamber: incremental' in line:
                        issues.append('Plugin incremental transport warning')

            threading.Thread(target=read_output, daemon=True).start()
            threading.Thread(target=read_errors, daemon=True).start()

            def rpc(number, method, params):
                proc.stdin.write(json.dumps({'jsonrpc': '2.0', 'id': number, 'method': method, 'params': params}) + '\n')
                proc.stdin.flush()
                deadline = time.monotonic() + 180
                while True:
                    try:
                        message = messages.get(timeout=max(.01, deadline - time.monotonic()))
                    except queue.Empty:
                        raise AssertionError(f'{method} timed out (no silent skip)')
                    if message.get('method') == 'session/requestPermission':
                        raise AssertionError('Synthetic stream prompt unexpectedly requested a tool')
                    if message.get('id') == number:
                        assert 'error' not in message, f'{method} failed; inspect locally without printing credentials'
                        return message['result']
            try:
                rpc(1, 'initialize', {'protocolVersion': 1, 'clientCapabilities': {}})
                session = rpc(2, 'session/new', {'cwd': str(root), 'mcpServers': []})['sessionId']
                active_session[0] = session
                result = rpc(3, 'session/prompt', {'sessionId': session, 'prompt': [{'type': 'text', 'text':
                    '不要使用工具。请用中文写二十条关于整理桌面文件的建议，每条至少二十字，结尾写“中文🙂金额1,234.56元”。'}]})
                events.append({'method': 'session/turn-ended', **binding, 'params': {'sessionId': session, **result}})
                assert not issues, issues
                assert len(timings) >= 3, 'assistant/chunk subscription disappeared: fewer than 3 deltas'
                assert commits and sum(t < commits[0] for t in timings) >= 3, 'Deltas did not arrive before ACP commit'
                kinds = sorted({e['params']['chunk']['type'] for e in events if e['method'] == 'session/stream' and e['params']['kind'] == 'delta'})
                output = args.output or Path(tempfile.gettempdir()) / ('lamber-stage3a-real-events.json' if args.real else 'lamber-stage3a-mock-events.json')
                output.write_text(json.dumps(events, ensure_ascii=False))
                subprocess.run(['node', str(base.parent / 'src-ui/scripts/test_dsh_runtime.cjs'), str(output)], check=True)
                print(json.dumps({'mode': 'real' if args.real else 'mock', 'passed': True,
                    'deltas': len(timings), 'kinds': kinds, 'firstDeltaSeconds': round(timings[0], 3),
                    'firstCommitSeconds': round(commits[0], 3), 'output': str(output)}, ensure_ascii=False))
            finally:
                proc.terminate()
                proc.wait(timeout=10)
    finally:
        server.shutdown()


if __name__ == '__main__':
    main()
