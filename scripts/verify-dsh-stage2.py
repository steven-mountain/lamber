#!/usr/bin/env python3
"""Run real ACP cancellation after observing an actual upstream text delta.

The loopback relay forwards unmodified requests only to DeepSeek's official API.
Credentials stay in process environments/headers. No request bodies are logged.
Usage: DEEPSEEK_API_KEY=<in environment> python3 scripts/verify-dsh-stage2.py
"""
import http.server
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import urllib.request
import urllib.error


def main():
    if '--relay' not in sys.argv and not os.environ.get('DEEPSEEK_API_KEY'):
        raise SystemExit('DEEPSEEK_API_KEY is required; this verification cannot be skipped')
    repo = Path(__file__).resolve().parent.parent
    with tempfile.TemporaryDirectory(prefix='lamber-dsh-stage2-') as temp:
        root = Path(temp)
        marker = root / 'upstream-delta'

        class Relay(http.server.BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def do_POST(self):
                body = self.rfile.read(int(self.headers['Content-Length']))
                request = urllib.request.Request('https://api.deepseek.com' + self.path,
                    data=body, headers={'Content-Type': 'application/json',
                    'Authorization': self.headers.get('Authorization', '')})
                try:
                    response = urllib.request.urlopen(request, timeout=120)
                except urllib.error.HTTPError as error:
                    response = error
                try:
                    with response:
                        self.send_response(response.status)
                        self.send_header('Content-Type', response.headers.get('Content-Type', 'text/event-stream'))
                        self.end_headers()
                        pending = b''
                        while chunk := response.read1(8192):
                            pending += chunk
                            while b'\n' in pending:
                                line, pending = pending.split(b'\n', 1)
                                if line.startswith(b'data:'):
                                    try:
                                        data = json.loads(line[5:])
                                        if any(choice.get('delta', {}).get('content') or
                                               choice.get('delta', {}).get('reasoning_content')
                                               for choice in data.get('choices', [])):
                                            marker.touch()
                                    except (ValueError, TypeError):
                                        pass
                            self.wfile.write(chunk)
                            self.wfile.flush()
                except (BrokenPipeError, ConnectionResetError):
                    pass  # Expected when ACP cancellation closes the upstream stream.

        server = http.server.ThreadingHTTPServer(('127.0.0.1', 0), Relay)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        patch = root / 'relay.yml'
        patch.write_text(f'- id: llm-deepseek\n  config:\n    baseURL: http://127.0.0.1:{server.server_port}\n')
        env = os.environ.copy()
        env['LAMBER_DSH_TRACE_PATH'] = str(Path(tempfile.gettempdir()) / 'lamber-stage2-real-stream.json')
        if '--relay' in sys.argv:
            try:
                print(json.dumps({'patch': str(patch), 'marker': str(marker)}), flush=True)
                sys.stdin.readline()
                return 0
            finally:
                server.shutdown()
        try:
            result = subprocess.run(['cargo', 'test', 'stage2_real_stream', '--', '--ignored', '--nocapture'],
                cwd=repo / 'src-tauri', env=env, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            output = result.stdout.replace(env['DEEPSEEK_API_KEY'], '[REDACTED]')
            (Path(tempfile.gettempdir()) / 'lamber-stage2-real-test.log').write_text(output)
            print(output[output.find('running 1 test'):] if 'running 1 test' in output else output[-5000:])
            return result.returncode
        finally:
            server.shutdown()


if __name__ == '__main__':
    sys.exit(main())
