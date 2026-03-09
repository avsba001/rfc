# DOCX 批量转 PDF，再转 JPG（Windows x86 EXE）

## 功能
- 批量将 `.docx` 转换为 `.pdf`
- 将 PDF 每页转换为 JPG
- 支持递归目录扫描
- 设计为通过 GitHub Actions 打包为 **Windows x86 exe**，并将 LibreOffice 与 ImageMagick 一并打包

## 本地运行（Windows）
```bash
python src/main.py <输入目录> <输出目录> --recursive
```

可选参数：
- `--office-path`：手动指定 `soffice` 路径
- `--magick-path`：手动指定 `magick.exe` 路径
- `--dpi`：输出图片分辨率，默认 `200`
- `--quality`：JPG 质量，默认 `92`

示例：
```bash
python src/main.py docs out --recursive --dpi 220 --quality 90
```

输出结构：
- `out/pdf/*.pdf`
- `out/jpg/<文档名>/<文档名>-000.jpg`

## GitHub Actions 打包
工作流文件：`.github/workflows/build-windows-x86.yml`

流程：
1. 使用 `actions/setup-python` 安装 x86 Python。
2. 用 Chocolatey 安装 x86 LibreOffice 与 x86 ImageMagick。
3. 复制两者到仓库 `third_party`。
4. 使用 PyInstaller 打包 exe，并将 `third_party` 一起打进发布目录。
5. 上传构建产物到 `Actions Artifacts`。

> 注意：首次构建下载依赖时间较长，且体积较大。
