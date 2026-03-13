# datalinker（数据处理）

基于 `Tauri + Vue 3 + Element Plus` 的桌面工具，按原 Java 工具迁移实现。  
当前目标是功能可用与迁移一致性，不做业务扩展。

## 功能概览

当前包含两个页面：
- `数据处理`
- `分组`

## 数据处理页

支持能力：
- 选择 `标准样本路径`、`样本路径`、`解析结果路径`
- 设置 `阈值`
- 添加任务到列表
- 右键菜单：复制选中行、删除选中行、删除所有行
- 列表路径通过“选择”按钮更新，支持“查看”打开路径
- 单行执行对比并导出结果

### 导出结果（5 Sheet）

文件名规则：
- `解析结果_<timestamp>.xlsx`

导出包含 5 张 Sheet：
- `亲子鉴定`（汇总）
- `真实性鉴定`（汇总）
- `相同位点`（明细）
- `差异位点`（明细）
- `缺失位点`（明细）

### 相同/差异/缺失判定规则

亲子鉴定（阈值参与）：
- 分别对样本 `A/B/C` 与标样 `A/B/C` 做阈值区间命中判断
- 三项结果求和
- `0` 判定为 `缺失`
- `>1` 判定为 `相同`
- `=1` 判定为 `差异`

真实性鉴定（阈值不参与）：
- 若标样 `A/B/C` 全为 `0`，判定为 `缺失`
- 否则将标样与样本各自 `A/B/C` 排序后比较
- 完全相等判定为 `相同`
- 不相等判定为 `差异`

## 分组页

支持能力：
- 选择原始 Excel 并解析
- 勾选/全选数据
- 执行分组
- 保存/加载分组配置

## 环境要求

- Windows 10/11（建议）
- Node.js 18+（建议 20 LTS）
- pnpm 9+
- Rust stable（通过 rustup 安装）
- Visual Studio Build Tools（Windows 构建 Tauri 必需）

## Rust 安装教程（Windows）

### 0. 重要：先把 Rust 目录改到非 C 盘（强烈建议）

为避免占用 C 盘空间，请先设置这两个环境变量，再安装 Rust：
- `RUSTUP_HOME`：rustup 工具与工具链目录
- `CARGO_HOME`：cargo 缓存与二进制目录

推荐目录：
- `D:\Rust\rustup`
- `D:\Rust\cargo`

在 PowerShell 执行（当前用户）：

```powershell
[Environment]::SetEnvironmentVariable("RUSTUP_HOME", "D:\Rust\rustup", "User")
[Environment]::SetEnvironmentVariable("CARGO_HOME", "D:\Rust\cargo", "User")

$targetBin = "D:\Rust\cargo\bin"
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if (-not $userPath) { $userPath = "" }
if (-not (($userPath -split ";") -contains $targetBin)) {
  $newPath = ($userPath.TrimEnd(";") + ";" + $targetBin).TrimStart(";")
  [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
}
```

执行后请重开终端，再继续安装 rustup。

如果你之前已经装在 C 盘，建议：
1. 先备份需要的数据
2. 设置好上面的环境变量
3. 重新安装 Rust（必要时先卸载旧安装）

### 1. 安装 Rust（rustup）

- 打开官网：`https://www.rust-lang.org/tools/install`
- 下载并运行 `rustup-init.exe`
- 安装选项可用默认值

### 2. 安装 C++ 构建工具（Tauri Windows 必需）

1. 打开官方页面下载 Build Tools：  
   `https://visualstudio.microsoft.com/visual-cpp-build-tools/`
2. 点击 `Download Build Tools`，下载并运行安装器 `vs_BuildTools.exe`
3. 在安装器的 `工作负载` 页签，勾选：`使用 C++ 的桌面开发`
4. 在右侧安装明细中，确认以下组件已勾选（默认一般会带上）：
   - `MSVC v143 - VS 2022 C++ x64/x86 build tools`
   - `Windows 10 SDK` 或 `Windows 11 SDK`
   - `C++ CMake tools for Windows`
5. 点击 `安装`，安装完成后重开终端

可选命令行安装（管理员 PowerShell）：

```powershell
winget install --id Microsoft.VisualStudio.2022.BuildTools -e
```

如果你已安装 `Visual Studio Community/Professional`，并且已勾选同样的 C++ 工作负载，也可以不单独安装 Build Tools。

### 3. 验证安装

```bash
rustc -V
cargo -V
rustup -V
```

### 4. 更新工具链（可选）

```bash
rustup update
```

## Node 与 pnpm 安装建议

```bash
node -v
corepack enable
pnpm -v
```

若 `pnpm` 不存在，可执行：

```bash
corepack prepare pnpm@9 --activate
```

## 本地开发

```bash
pnpm install
pnpm tauri dev
```

## 本地构建

前端构建：

```bash
pnpm build
```

Rust 检查：

```bash
cd src-tauri
cargo check
cd ..
```

构建 Windows 安装包（NSIS）：

```bash
pnpm tauri build --bundles nsis
```

构建产物通常位于：
- `src-tauri/target/release/bundle/nsis/*.exe`

构建 Windows 免安装便携版 EXE：

```bash
pnpm tauri build --no-bundle
```

便携版 EXE 通常位于：
- `src-tauri/target/release/datalinker.exe`

## GitHub Actions（已配置）

仓库内已包含两个工作流：
- `.github/workflows/build-windows-exe.yml`
  - push `main` 或手动触发
  - push `main` 时自动递增 tag（`vX.Y.Z`）并发布到 GitHub Releases
  - 发布文件名自动带版本号：
    - `datalinker-portable-vX.Y.Z.exe`
    - `datalinker-setup-vX.Y.Z.exe`
- `.github/workflows/release-windows-exe.yml`
  - push tag（`v*`）触发
  - 手动打 tag 的兜底发布流程

### 发布版本（推荐流程：只提交代码）

提交并推送到 `main` 即可自动发布：

```bash
git push origin main
```

### 手动发布（可选）

如果你希望手动控制版本号，也可以继续手动打 tag：

```bash
git tag v0.1.0
git push origin v0.1.0
```

发布后可在仓库 `Releases` 页面下载两类文件：
- `datalinker-portable-vX.Y.Z.exe`：免安装，直接双击运行
- `datalinker-setup-vX.Y.Z.exe`：NSIS 安装包，适合普通分发

## 配置存储

前端配置通过 `tauri-plugin-store` 持久化：
- 文件名：`datalinker.store.json`
- 位置：Tauri `app_data_dir`（随系统与应用标识而定）

## 常见问题

1. `icons/icon.ico not found`
- 原因：`src-tauri/icons/icon.ico` 缺失
- 处理：补齐该文件并提交到仓库

2. GitHub Actions 提示 `Unable to locate executable file: pnpm`
- 原因：`setup-node` 中启用了 `cache: pnpm`，但 `pnpm` 尚未安装
- 处理：先执行 `pnpm/action-setup`，再执行 `actions/setup-node`

3. 运行成功但只在 Actions 找到安装包
- 说明：这是 Artifact 模式
- 若希望出现在 Releases，请使用 tag 触发 `release-windows-exe.yml`

4. 想要“免安装 EXE”而不是安装包
- 本地：使用 `pnpm tauri build --no-bundle`
- CI/Releases：已同时产出 `datalinker-portable-vX.Y.Z.exe` 和 NSIS 安装包

## 目录结构

```text
datalinker/
  .github/workflows/  # CI/CD 工作流
  src/                # Vue 前端
  src-tauri/          # Tauri + Rust 后端
  public/             # 静态资源
```
