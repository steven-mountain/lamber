// The stage-0 probe has no business bridge. Even a malicious/invented tool call
// is denied by the monotonic tool guard before execution or human approval.
export const inject = ['tools', 'workspaceRegistry'];
export async function apply(ctx) {
  ctx.tools.guard(() => '界面承载验证阶段未开放业务工具');
  await ctx.workspaceRegistry.create(process.cwd(), '合成工作区 With Spaces');
  ctx.on('agent/pre-step', async ({ agent }, next) => {
    if (agent.session.header.cwd !== process.cwd()) {
      throw new Error('界面验证仅允许启动时指定的合成工作区');
    }
    return next();
  });
}
