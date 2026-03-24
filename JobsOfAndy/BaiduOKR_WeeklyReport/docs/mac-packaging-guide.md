# macOS 打包指南

> 版本: 1.1.0 | 更新日期: 2026-03-22

---

## 一键打包

```bash
cd JobsOfAndy/BaiduOKR_WeeklyReport
chmod +x scripts/build_mac.sh
./scripts/build_mac.sh
```

脚本自动完成：环境检查 → 安装依赖 → Playwright 浏览器 → PyInstaller 打包 → ad-hoc 签名 → ZIP 压缩。

产物位于 `release/` 目录：

| 文件 | 说明 |
|------|------|
| `WeeklyReportTool_v{VER}_pyinstaller/` | 可直接运行的目录 |
| `WeeklyReportTool_v{VER}_macos.zip` | 分发用压缩包 |

---

## 前置条件

| 要求 | 安装方式 |
|------|---------|
| Python 3.8+ | `brew install python@3.12` |
| Xcode CLI Tools | `xcode-select --install` |
| Chrome（运行时需要） | `brew install --cask google-chrome` |

脚本会自动检查并提示缺失项。

---

## 手动打包（备用）

如果不想用一键脚本：

```bash
cd JobsOfAndy/BaiduOKR_WeeklyReport

# 1. 安装依赖
pip3 install -r requirements.txt
pip3 install pyinstaller pyarmor

# 2. 安装 Playwright 浏览器
python3 -m playwright install chromium

# 3. 打包（build.py 已自动处理签名和权限）
python3 scripts/build.py --method pyinstaller

# 4. 运行
./release/WeeklyReportTool_v*_pyinstaller/WeeklyReportTool
```

---

## 常见问题

### Gatekeeper 拦截

如果从其他机器拷贝过来的构建产物被拦截：

```bash
xattr -cr release/WeeklyReportTool_v*_pyinstaller/
```

或在 **系统设置 > 隐私与安全** 中点击"仍然允许"。

本机构建 + 一键脚本已自动处理签名，通常不会遇到此问题。

### Apple Silicon vs Intel

- M1/M2/M3/M4 构建的产物只能在 ARM Mac 上运行
- Intel 构建的产物只能在 Intel Mac 上运行
- 接收方可用 Rosetta 2 运行 Intel 版本：`arch -x86_64 ./WeeklyReportTool`

### PyArmor 警告

PyArmor 试用版对 >32KB 文件有限制，构建脚本会自动回退到 `.pyc` 编译。警告可安全忽略。

### CI 自动构建

项目配置了 GitHub Actions（`.github/workflows/build.yml`），支持 Windows/macOS/Linux 三平台同时构建：

```bash
# 推送标签自动构建 + 创建 Release
git tag v1.2.0
git push origin v1.2.0
```

也可在 GitHub Actions 页面手动触发。
