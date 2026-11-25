# VIN 批量生成器（Windows）

这是对 macOS「VIN 批量生成器」的 Windows 版本重做，包含相同的主要工作流程：逐格填写前 8 位主体和第 10–11 位、设置起始流水号和数量，自动计算第 9 位校验码，生成、复制或导出 Excel 工作簿（.xlsx）。界面以 `?` 显示第 9 位位置，实际生成时会自动计算正确校验码。

## 在 GitHub 上构建

1. 将本目录中的文件上传到你自己的 GitHub 仓库。
2. 打开仓库的 **Actions** 页面，选择 **Build Windows app**。
3. 点击 **Run workflow** 并确认运行。
4. 运行完成后，在该次 workflow 的 **Artifacts** 区域下载 `VIN批量生成器-Windows`，解压即可得到 `VIN批量生成器.exe`。

也可以推送形如 `v1.0.0` 的 tag 自动触发构建。GitHub Actions 使用 Windows runner 生成 Windows 可执行文件；无需在本机安装 Python。

## 本机运行（可选）

安装 Python 3.10 或更高版本，并安装 Excel 文件库后运行：

```powershell
python -m pip install openpyxl
python main.py
```

## 规则

- VIN 为 17 位：前 8 位 + 自动校验位 + 第 10–11 位 + 6 位流水号。
- 使用标准 VIN 校验位计算（字符换算值与权重求和后对 11 取余，余数 10 使用 `X`）。
- 禁止字符 `I`、`O`、`Q`；起始流水号范围 0–999999，每批最多 5,000 条。
- Excel 导出为 `.xlsx` 工作簿，包含 VIN 列、筛选和冻结表头。
