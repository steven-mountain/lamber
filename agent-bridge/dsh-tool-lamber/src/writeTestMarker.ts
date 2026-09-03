/**
 * `write_test_marker` — a deliberately harmless write used to exercise the
 * human-approval channel end to end.
 *
 * It exists only to prove that a write-capable tool is gated: it touches the
 * OS temp directory and nothing else. It must never be pointed at lamber's
 * workspace, database, or project files. Real write tools stay out of this
 * package until the approval channel is verified — see `AGENTS.md`, which
 * forbids the AI from modifying project data without user confirmation.
 */
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { defineTool } from '@deepseek-ai/dsh-tools';

/** Tool name, shared with the approval guard. */
export const WRITE_TEST_MARKER = 'write_test_marker';

export const writeTestMarker = defineTool({
  name: WRITE_TEST_MARKER,
  description: [
    '【测试用】在系统临时目录写一个带时间戳的标记文件，用于验证人工审批通道是否连通。',
    '这是一个需要用户确认的写操作：调用前会弹出审批对话框，用户拒绝则不会执行。',
    '它不会读写 lamber 的任何项目数据、数据库或文档，仅用于联调。',
  ].join(' '),
  parameters: {
    note: {
      type: 'string',
      description: '可选备注，会写进标记文件内容，便于人工核对是不是本次调用。',
    },
  },
  output: {
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        path: { type: 'string', required: true, description: '实际写入的标记文件绝对路径。' },
        writtenAt: { type: 'string', required: true, description: '写入时间（ISO 8601）。' },
        bytes: { type: 'integer', required: true, description: '写入的字节数。' },
      },
    },
    render(_args, value) {
      return [
        {
          type: 'text',
          text: `已写入测试标记文件：${value.path}（${value.bytes} 字节，${value.writtenAt}）`,
        },
      ];
    },
  },
  timeoutMs: 15_000,
  isConcurrencySafe: () => false,
  async execute(args, exec) {
    exec.signal.throwIfAborted();
    const writtenAt = new Date().toISOString();
    // A fresh directory per call keeps concurrent runs from clobbering each other
    // and keeps every artifact under the OS temp root.
    const dir = await mkdtemp(join(tmpdir(), 'lamber-agent-marker-'));
    const path = join(dir, `marker-${writtenAt.replace(/[:.]/g, '-')}.txt`);
    const body = [
      'lamber agent-bridge 审批通道测试标记',
      `写入时间: ${writtenAt}`,
      `备注: ${args.note ?? '(无)'}`,
      '',
      '此文件由 write_test_marker 工具生成，仅用于验证人工审批流程，可随时删除。',
    ].join('\n');
    await writeFile(path, body, 'utf8');
    return { path, writtenAt, bytes: Buffer.byteLength(body, 'utf8') };
  },
});
