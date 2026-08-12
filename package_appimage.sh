#!/usr/bin/env bash
# 中文文档智能排版工具 —— 国产系统版 AppImage 封装脚本
# 产出：WordFormatterPro-<arch>.AppImage（单个文件，免安装、双击即可运行）
#
# 前置：build_linux.sh 已产出 dist/WordFormatterPro（本机架构单文件）。
# 本脚本用 appimagetool 将其打包成 AppImage（含 AppRun / .desktop / 图标）。
#
# 用法（在目标架构的国产机上执行，例如 x86_64 / ARM64 / LoongArch 机器）：
#   chmod +x package_appimage.sh
#   ./package_appimage.sh
#
# 说明：
#   - PyInstaller 单文件已自带 Python 与 tkinter，AppImage 内不再依赖系统 Python。
#   - 运行 AppImage 需要 FUSE；若目标机无 FUSE，可执行
#       ./WordFormatterPro-<arch>.AppImage --appimage-extract
#     解包后直接运行 squashfs-root/AppRun。
set -e

cd "$(dirname "$0")"
SCRIPT_DIR="$(pwd)"

ARCH="$(uname -m)"
APP="WordFormatterPro"
BUILD="$SCRIPT_DIR/dist/$APP"

# 1) 若尚未构建，先运行 build_linux.sh（PyInstaller 不支持交叉编译，需本机构建）
if [ ! -f "$BUILD" ]; then
  echo "未找到 $BUILD，先运行 build_linux.sh 构建国产系统版单文件…"
  ./build_linux.sh
fi
if [ ! -f "$BUILD" ]; then
  echo "构建失败，无法继续打包 AppImage。" >&2
  exit 1
fi

# 2) 准备 AppDir（AppImage 约定目录结构）
APPDIR="$SCRIPT_DIR/AppDir"
rm -rf "$APPDIR"
mkdir -p "$APPDIR/usr/bin"
cp "$BUILD" "$APPDIR/usr/bin/$APP"
chmod +x "$APPDIR/usr/bin/$APP"

# 3) AppRun（启动器）
cat > "$APPDIR/AppRun" <<'EOF'
#!/bin/sh
HERE="$(dirname "$(readlink -f "${0}")")"
exec "$HERE/usr/bin/WordFormatterPro" "$@"
EOF
chmod +x "$APPDIR/AppRun"

# 4) 桌面入口（.desktop）
cat > "$APPDIR/$APP.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=中文文档智能排版工具
Comment=一键式中文文档智能排版（国产系统版）
Exec=$APP
Icon=$APP
Categories=Office;Utility;
Terminal=false
EOF

# 5) 图标（内嵌 SVG 占位；如需正式图标替换为此文件名即可）
cat > "$APPDIR/$APP.svg" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">
  <rect width="256" height="256" rx="48" fill="#0096E6"/>
  <text x="128" y="168" font-size="140" text-anchor="middle" fill="#FFFFFF" font-family="sans-serif">文</text>
</svg>
EOF

# 6) 下载 appimagetool（按架构缓存，避免重复下载）
DL_DIR="$SCRIPT_DIR/.appimage_tools"
mkdir -p "$DL_DIR"
APPIMAGETOOL="$DL_DIR/appimagetool-$ARCH.AppImage"
if [ ! -f "$APPIMAGETOOL" ]; then
  echo "下载 appimagetool（$ARCH）…"
  if command -v curl >/dev/null 2>&1; then
    curl -L -o "$APPIMAGETOOL" "https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-$ARCH.AppImage"
  else
    wget -O "$APPIMAGETOOL" "https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-$ARCH.AppImage"
  fi
  chmod +x "$APPIMAGETOOL"
fi

# 7) 生成 AppImage
export ARCH
"$APPIMAGETOOL" "$APPDIR" "$SCRIPT_DIR/$APP-$ARCH.AppImage"
echo "AppImage 封装完成：$SCRIPT_DIR/$APP-$ARCH.AppImage"
