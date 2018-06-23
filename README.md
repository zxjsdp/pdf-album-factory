# pdf-album-factory

## 介绍

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

## 使用方法

### 第一步：安装 Python3

若您使用的是 Windows，则可能并没有安装过 Python3。请前往 <https://www.python.org/downloads/windows/> 下载并安装，建议使用 3.6 以上的版本。

若您使用的是 Linux / macOS，则应该已有预置的 Python 环境。请检查是否为 Python3 版本，注意，本脚本仅支持 Python3！

### 第二步：安装 wkhtmltopdf

Wkhtmltopdf 是用于将 HTML 转换为 PDF 的开源工具，可前往 <https://wkhtmltopdf.org/downloads.html> 下载您所在平台对应的安装包，并进行安装。

若您对此工具感兴趣，可前往 GitHub 查看：[GitHub - wkhtmltopdf/wkhtmltopdf](https://github.com/wkhtmltopdf/wkhtmltopdf/)。

### 第三步：准备项目及安装 Python 包依赖

1. 克隆项目至本地

    - 使用 git 将项目克隆至本地：`git@github.com:zxjsdp/pdf-album-factory.git`。
    - 若不熟悉 git 或仅想快速使用，也可直接下载最新 master 分支版本文件：<https://github.com/zxjsdp/pdf-album-factory/archive/master.zip>。
    
2. 切换至工作目录

    可通过文件管理器进入下载的 `pdf-album-factory` 文件夹后，按住 shift 同时点击右键，选择 “在此处打开命令行窗口”。
    
3. 安装 Python 包依赖

    执行 pip 命令，安装 requirements.txt 中的所有 Python package（建议使用 virtualenv）

        pip install -r requirements.txt

#### 其中需要特别关注的两个包分别为：

1. openpyxl

    读取及写入 Excel 2010 xlsx/xlsm/xltx/xltm 文件的包。文档可参考：[OpenPyXL](https://openpyxl.readthedocs.io/en/stable/);

2. pdfkit

    wkhtmltopdf 的 Python 封装包，用于将 HTML 转换为 PDF。文档可参考：[GitHub - JazzCore/python-pdfkit](https://github.com/JazzCore/python-pdfkit)。
    
### 第四步：修改模板文件及准备图片

默认的模板文件名为 `相册数据模板.xlsx`，请按照需求修改模板文件。

### 第五步：执行生成命令

执行 PDF 生成逻辑（建议在 virtualenv 中执行）：

    python generate_pdf_album.py

## 常见错误类型

1. > "No wkhtmltopdf executable found"

    出现此错误，说明在环境变量中没有找到对应的 wkhtmltopdf.exe 可执行文件。解决方案：
    
    - Windows, macOS, Linux 用户可检查是否执行了第二步：安装 wkhtmltopdf。
    - Windows 用户若确认已安装但仍报此错误，则可尝试重启。
    - Windows 下若重启仍无效，可找到对应的 wkhtmltopdf.exe 可执行文件的路径，并更新 generate_pdf_album.py 中的常量 `WKHTMLTOPDF_PATH`。Windows 下 x64 系统的默认配置是 `WKHTMLTOPDF_PATH = 'c:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'`。

2. > IOError: No such file or directory

    - 请检查当前所使用的 Python 版本是否为 Python2.x。处于维护成本考虑，本脚本仅支持 Python3.x。
    - 请检查是否无意中移动或重命名了模板文件 "相册数据模板.xlsx" 或其他文件。

3. > ModuleNotFoundError: No module named 'pdfkit'

    - 请检查是否执行了第三步：安装 Python 包依赖。

## 注意

在使用此工具的流程中，请不要随意修改任何文件的相对路径，避免找不到对应文件的情况。
