# 智算公式 Token 光标定位

- **Status:** Done
- **Objective:** 为折叠式 DIY 计算器增加 Token 间插入点，使参数、计算结果、固定值和运算符可以插入公式任意位置。

## Progress

1. [x] 将光标建模为 `0..tokens.length` 的本地位置索引。
2. [x] 在所有 Token 前后渲染可点击插入槽，当前槽显示主色竖线。
3. [x] 点击 Token 将插入点移动到该 Token 之后。
4. [x] 所有插入操作改为在当前光标位置写入，插入后自动前移。
5. [x] Token 删除后按相对位置修正光标。
6. [x] “回退”改为删除光标前一个 Token；清空后光标归零。
7. [x] 新增纯函数测试覆盖中间插入、删除和光标边界。
8. [x] Token 删除按钮默认隐藏，仅在 Token 悬停或删除按钮键盘聚焦时显示。
9. [x] 将公式预览与计算结果移动到展开计算区域顶部，优先展示当前计算状态。
10. [x] 公式预览不再使用花括号包裹引用，改为 Markdown 行内代码风格的语义强调块。

## Validation

- `npm run test:ai-compute-quote`
- `npx tsc --noEmit`
- `npx eslint src/features/ai-compute-quote`
- `npm run build`

## Scope Boundary

- 光标和展开状态均为组件本地 UI 状态，不写入蓝图或项目持久化数据。
- 未修改公式语义、ICT 联动或其他智算页面区域。
