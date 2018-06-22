# pdf-album-factory

根据图片及内容，批量生成漂亮的 PDF 格式相册文件。

效果图（点击可查看大图）：

<p float="left">
  <a href="http://zxjsdp1.qiniudn.com/pdf-album-factory-001-squashed.jpg" title="左页面-正文 3 图片示例" target='_blank'>
    <img src="http://zxjsdp1.qiniudn.com/pdf-album-factory-001-squashed.jpg"
         alt="左页面-正文 3 图片示例"
         width="24%" />
  </a>
  <a href="http://zxjsdp1.qiniudn.com/pdf-album-factory-002-squashed.jpg" title="右页面-正文 3 图片示例" target='_blank'>
    <img src="http://zxjsdp1.qiniudn.com/pdf-album-factory-002-squashed.jpg"
         alt="右页面-正文 3 图片示例"
         width="24%" />
  </a>
  <a href="http://zxjsdp1.qiniudn.com/pdf-album-factory-003-squashed.jpg" title="左页面-正文 2 图片示例" target='_blank'>
    <img src="http://zxjsdp1.qiniudn.com/pdf-album-factory-003-squashed.jpg"
         alt="左页面-正文 2 图片示例"
         width="24%" />
  </a>
  <a href="http://zxjsdp1.qiniudn.com/pdf-album-factory-004-squashed.jpg" title="左页面-正文 2 图片示例" target='_blank'>
    <img src="http://zxjsdp1.qiniudn.com/pdf-album-factory-004-squashed.jpg"
         alt="左页面-正文 2 图片示例"
         width="24%" />
  </a>
</p>

# 使用方法

## 第一步：安装 Python3

若您使用的是 Windows，则可能并没有安装过 Python3。请前往 <https://www.python.org/downloads/windows/> 下载并安装，建议使用 3.6 以上的版本。

若您使用的是 Linux / macOS，则应该已有预置的 Python 环境。请检查是否为 Python3 版本。

## 第二步：安装 wkhtmltopdf

Wkhtmltopdf 是可用于将 HTML 转换为 PDF 的开源工具，GitHub 链接为：[GitHub - wkhtmltopdf/wkhtmltopdf](https://github.com/wkhtmltopdf/wkhtmltopdf/)。

可前往 <https://wkhtmltopdf.org/downloads.html> 下载您所在平台对应的安装包，并进行安装。

## 第三步：准备项目及安装 Python 包依赖

1. 克隆项目至本地

    - 使用 git 将项目克隆至本地：`git@github.com:zxjsdp/pdf-album-factory.git`。
    - 若不熟悉 git 或仅想快速使用，也可直接下载最新 master 分支版本文件：<https://github.com/zxjsdp/pdf-album-factory/archive/master.zip>。
    
2. 切换至工作目录

    可通过文件管理器进入下载的 `pdf-album-factory` 文件夹后，按住 shift 同时点击右键，选择 “在此处打开命令行窗口”。
    
3. 安装 Python 包依赖

    执行 pip 命令，安装 requirements.txt 中的所有 Python package（建议使用 virtualenv）

        pip install -r requirements.txt

### 其中需要特别关注的两个包分别为：

1. openpyxl

    读取及写入 Excel 2010 xlsx/xlsm/xltx/xltm 文件的包。文档可参考：[OpenPyXL](https://openpyxl.readthedocs.io/en/stable/);

2. pdfkit

    wkhtmltopdf 的 Python 封装包，用户将 HTML 转换为 PDF。文档可参考：[GitHub - JazzCore/python-pdfkit](https://github.com/JazzCore/python-pdfkit)。
    
## 第四步：修改模板文件及准备图片

默认的模板文件名为 `相册数据模板.xlsx`，请按照需求修改模板文件。

## 第五步：

执行 PDF 生成逻辑（建议在 virtualenv 中执行）：

    python generate_pdf_album.py

### 注意

在使用此工具的流程中，请不要随意修改任何文件的相对路径，避免找不到对应文件的情况。
