# html-to-pptx

一个 Claude Code / Agent **Skill**:把一张自包含的 **HTML 页面**(架构图、深色玻璃拟态大屏/仪表盘、信息图、海报等)转换成**高保真、可编辑的 PowerPoint `.pptx`**,观感与源图一致,并能在 **PowerPoint、Keynote、LibreOffice、Google Slides** 里正常打开。

## 核心思路:保真优先的混合策略

PowerPoint 画不出浏览器的某些效果(磨砂玻璃、外发光、复杂渐变、任意 SVG),用原生形状硬凑会显得很平。所以本 skill 对每个元素二分处理:

- **原生(可编辑)**:凡是 PowerPoint 能 1:1 画的——文字、纯色/线性渐变形状、线条、虚线边框、圆点、徽标。
- **透明 PNG 抠图**:凡是它做不出的——磨砂玻璃、辉光、径向/锥形渐变、复杂 SVG 图标、网格/拓扑图、纹理底。
- **文字永远是原生、永远在最上层**,所以即使框体是图片,整页文字依然可编辑。

这个二分法也让产物天生跨软件兼容:其它软件会丢失的效果,恰好都被烤进了 PNG(图片到处渲染都一致)。

## 安装

克隆到你的 Claude skills 目录:

```bash
git clone https://github.com/KDevSec/html-to-pptx.git ~/.claude/skills/html-to-pptx
```

之后当你要求"把某个 HTML 转成可编辑 PPT"时会自动触发。(任何加载 `SKILL.md` 的 agent 环境都通用。)

## 前置依赖

`python-pptx`、`Pillow`、`playwright`(Chromium),以及 `LibreOffice`(`soffice`)+ `pdftoppm`。一行安装见 `references/playbook.md` §0。

## 用法

指向一个自包含的 `.html`,然后说:

> 把这个 html 转成可编辑 PPT(完全复刻)

完整管线(也可手动逐步执行):

```bash
SK=~/.claude/skills/html-to-pptx
# 1. 采集几何 + 参考截图
python $SK/scripts/capture.py geom INPUT.html WORK/ --region .slide
# 2a. 保真优先:把"文字+图标"隐藏,截一张完整视觉板
python $SK/scripts/capture.py board INPUT.html WORK/ --hide "<所有文字 + .ic,.ci 选择器>"
# 2b. 或逐元素抠图(让部分框保持可编辑形状)
python $SK/scripts/capture.py slice INPUT.html WORK/ --selectors ".mesh,.glasspanel" --isolate
# 3. 用 scripts/pptx_helpers.py 写 WORK/build.py(背景 → 抠图 → 原生形状 → 原生文字)
# 4. 渲染自检
python $SK/scripts/verify.py OUT.pptx --source INPUT.html --region .slide
```

## 目录说明

| 路径 | 作用 |
|---|---|
| `SKILL.md` | skill 本体:触发条件、决策表(原生 vs PNG)、坐标系、工作流、跨平台说明 |
| `scripts/capture.py` | 无头 Chromium 采集:几何 `geom` / 保真板 `board` / 透明抠图 `slice` |
| `scripts/pptx_helpers.py` | python-pptx 复用库(`E`/`P`/`rrect`/`textbox`/`chip`/`oval`/`picture`/`save`/`target_cjk_font` 等) |
| `scripts/verify.py` | LibreOffice 渲染 + 源图并排对比 + 可编辑性报告 |
| `references/playbook.md` | 完整方法论:坐标换算、采集技巧、组装配方、踩坑清单 |

## 字体与跨平台

产物永远是标准 `.pptx`(不生成 Keynote 的 `.key`)。默认中文字体为**微软雅黑**;按受众系统切换用 `target_cjk_font('mac'|'linux')`,跨机分发可在 PowerPoint 里嵌入字体以彻底锁定观感。渲染自检在任意系统都用 LibreOffice;最终观感请在你实际演示用的软件里再确认一次。

## 许可证

[Apache-2.0](LICENSE) © 2026 KDevSec
