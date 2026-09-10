"""Controlled model for the isolated product app only; never calls a real provider."""
import argparse
import http.server
import json
import re
import uuid
from pathlib import Path

class Model(http.server.BaseHTTPRequestHandler):
    def log_message(self, *_): pass
    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
        messages = body.get('messages', [])
        last_user = max((i for i, m in enumerate(messages) if m['role'] == 'user' and not str(m.get('content', '')).startswith('LAMBER_BUSINESS_CONTEXT')), default=0)
        prompt = str(messages[last_user].get('content', ''))
        turns = messages[last_user:]
        previous = [m for m in turns if m['role'] == 'tool']
        tools = [t.get('function', {}).get('name') for t in body.get('tools', [])]
        # Persist only the synthetic transcript; Authorization headers are never read or logged.
        evidence = Path(self.server.evidence)
        snapshot = evidence.with_name(evidence.stem + '-' + uuid.uuid4().hex + '.json')
        snapshot.write_text(json.dumps({'model': body.get('model'), 'tools': tools, 'messages': messages}, ensure_ascii=False, indent=2))
        # Title generation and the business request can run concurrently.
        # Keep every complete request; atomically publish the latest one.
        temporary = snapshot.with_suffix('.latest')
        temporary.write_bytes(snapshot.read_bytes())
        temporary.replace(evidence)
        function = None
        if '验证改文审批' in prompt:
            if not previous:
                function = {'name': 'read_template_fields', 'arguments': json.dumps({'templateId': 'ICT项目需求导入表模板.docx'})}
            elif len(previous) == 1:
                raw = previous[0].get('content', '')
                try:
                    saved = json.loads(raw)
                    if isinstance(saved, list): saved = json.loads(saved[0]['text'])
                    function = {'name': 'fill_template_fields', 'arguments': json.dumps({'projectId': saved['projectId'], 'templateId': saved['templateId'], 'fields': {'gen_demand_env_require': '受控模型建议：部署在客户机房，保留现有网络。'}})}
                except (ValueError, KeyError, TypeError): pass
        if function and function['name'] not in tools:
            text = '受控验证：模型工具目录缺少 ' + function['name'] + '；可见工具为 ' + ', '.join(str(x) for x in tools)
            function = None
        else:
            text = '受控模型已收到当前会话。仅用于验证官方界面与业务链路。'
            if previous: text = '工具实际返回：\n' + str(previous[-1].get('content', ''))
        self.send_response(200); self.send_header('Content-Type', 'text/event-stream'); self.end_headers()
        delta = {'tool_calls': [{'index': 0, 'id': 'fixture-call-' + uuid.uuid4().hex, 'type': 'function', 'function': function}]} if function else {'content': text}
        base = {'id': 'lamber-product-fixture', 'object': 'chat.completion.chunk', 'created': 1, 'model': body.get('model'), 'choices': [{'index': 0, 'delta': delta, 'finish_reason': None}]}
        try:
            self.wfile.write(('data: ' + json.dumps(base, ensure_ascii=False) + '\n\n').encode())
            base['choices'] = [{'index': 0, 'delta': {}, 'finish_reason': 'tool_calls' if function else 'stop'}]
            self.wfile.write(('data: ' + json.dumps(base) + '\n\ndata: [DONE]\n\n').encode()); self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError): pass

if __name__ == '__main__':
    parser = argparse.ArgumentParser(); parser.add_argument('--config', required=True); parser.add_argument('--evidence', required=True); parser.add_argument('--port', type=int, default=0)
    args = parser.parse_args(); path = Path(args.config)
    if path.parent.name != 'com.cmcc.benefitcalc.webui-test': parser.error('Only the dedicated synthetic app configuration is supported')
    server = http.server.ThreadingHTTPServer(('127.0.0.1', args.port), Model); server.evidence = args.evidence
    config = json.loads(path.read_text()); config['aiAgent']['baseUrl'] = f'http://127.0.0.1:{server.server_port}'; config['aiAgent']['apiKey'] = 'controlled-fixture-only'
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2))
    print('Isolated controlled model ready', flush=True); server.serve_forever()
