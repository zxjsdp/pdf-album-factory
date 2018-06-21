# -*- coding: utf-8 -*-

import os
import sys
import pdfkit
import openpyxl

# =========================================================================================
# 这些选项可修改
# -----------------------------------------------------------------------------------------
# 输出结果路径及文件名前缀
RESULT_FOLDER = '_PDF输出结果'
RESULT_PDF_NAME_TEMPLATE = '相册结果_%03d.pdf'

# 模板文件名
TEMPLATE_XLSX_FILE = '相册数据模板.xlsx'
# =========================================================================================


# =========================================================================================
# 这些选项若您不知道代表什么含义，请不要随意修改
# -----------------------------------------------------------------------------------------
# 默认标题栏，处理时文件的标题栏需要完全对应
DEFAULT_HEADER_CONTENT = [
    '名字', '学号', 'QQ 号', '微信号', '图片路径（资料图）', '图片路径（正文图片一）', '图片路径（正文图片二）',
    '图片路径（正文图片三）', '页面类型（左右页面）', '资料背景色（Hex）', '内容背景色（Hex）', '正文内容'
]

# pdfkit 选项
OPTIONS = {
    'page-size': 'A3',
    'orientation': 'Landscape',
    'dpi': '100',
    'image-dpi': '1000',
    'image-quality': '100',
    'margin-top': '0',
    'margin-right': '0',
    'margin-bottom': '0',
    'margin-left': '0',
    'encoding': "UTF-8",
    'no-outline': None
}

# 临时 web 文件名
TEMP_HTML_FILE = '_temp.html'
TEMP_CSS_FILE = '_temp.css'

# Web 模板文件夹
TEMPLATE_FOLDER = 'templates'

# Web 模板文件名
TEMPLATE_HTML_LEFT_2_IMAGES = 'album-left-2-images.html'
TEMPLATE_HTML_LEFT_3_IMAGES = 'album-left-3-images.html'
TEMPLATE_HTML_RIGHT_2_IMAGES = 'album-right-2-images.html'
TEMPLATE_HTML_RIGHT_3_IMAGES = 'album-right-3-images.html'
TEMPLATE_CSS_2_IMAGES = 'album-2-images.css'
TEMPLATE_CSS_3_IMAGES = 'album-3-images.css'

PLACEHOLDER_CSS_FILE = '__CSS_FILE_PATH__'  # css path
PLACEHOLDER_NAME = '__PERSON_NAME__'  # 名字
PLACEHOLDER_STUDENT_ID = '__STUDENT_ID__'  # 学号
PLACEHOLDER_QQ = '__QQ__'  # QQ 号
PLACEHOLDER_WECHAT = '__WECHAT__'  # 微信号
PLACEHOLDER_INTRO_IMAGE_PATH = '__INTRO_IMAGE_PATH__'  # 图片路径（资料图）
PLACEHOLDER_CONTENT_IMAGE_A_PATH = '__CONTENT_IMAGE_A_PATH__'  # 图片路径（正文图片一）
PLACEHOLDER_CONTENT_IMAGE_B_PATH = '__CONTENT_IMAGE_B_PATH__'  # 图片路径（正文图片二）sssss
PLACEHOLDER_CONTENT_IMAGE_C_PATH = '__CONTENT_IMAGE_C_PATH__'  # 图片路径（正文图片三）
PLACEHOLDER_PAGE_ORIENTATION = '__PAGE_ORIENTATION__'  # 页面类型（左右页面）
PLACEHOLDER_INTRO_BACKGROUND_COLOR_LEFT = '__INTRO_BACKGROUND_COLOR_LEFT__'  # 资料背景色（Hex）
PLACEHOLDER_CONTENT_BACKGROUND_COLOR_LEFT = '__CONTENT_BACKGROUND_COLOR_LEFT__'  # 内容背景色（Hex）
PLACEHOLDER_INTRO_BACKGROUND_COLOR_RIGHT = '__INTRO_BACKGROUND_COLOR_RIGHT__'  # 资料背景色（Hex）
PLACEHOLDER_CONTENT_BACKGROUND_COLOR_RIGHT = '__CONTENT_BACKGROUND_COLOR_RIGHT__'  # 内容背景色（Hex）
PLACEHOLDER_CONTENT_TEXT = '__CONTENT_TEXT__'  # 正文内容

template_cache = dict()  # 模板信息缓存 dict
# =========================================================================================


def prepare_template_cache():
    """准备模板信息缓存"""
    def join_path(folder, name):
        return os.path.join(folder, name)

    global template_cache

    with open(join_path(TEMPLATE_FOLDER, TEMPLATE_HTML_LEFT_2_IMAGES)) as template_left_2_html:
        template_cache[TEMPLATE_HTML_LEFT_2_IMAGES] = template_left_2_html.read()
    with open(join_path(TEMPLATE_FOLDER, TEMPLATE_HTML_LEFT_3_IMAGES)) as template_left_3_html:
        template_cache[TEMPLATE_HTML_LEFT_3_IMAGES] = template_left_3_html.read()
    with open(join_path(TEMPLATE_FOLDER, TEMPLATE_HTML_RIGHT_2_IMAGES)) as template_right_2_html:
        template_cache[TEMPLATE_HTML_RIGHT_2_IMAGES] = template_right_2_html.read()
    with open(join_path(TEMPLATE_FOLDER, TEMPLATE_HTML_RIGHT_3_IMAGES)) as template_right_3_html:
        template_cache[TEMPLATE_HTML_RIGHT_3_IMAGES] = template_right_3_html.read()
    with open(join_path(TEMPLATE_FOLDER, TEMPLATE_CSS_2_IMAGES)) as template_2_css:
        template_cache[TEMPLATE_CSS_2_IMAGES] = template_2_css.read()
    with open(join_path(TEMPLATE_FOLDER, TEMPLATE_CSS_3_IMAGES)) as template_3_css:
        template_cache[TEMPLATE_CSS_3_IMAGES] = template_3_css.read()


class XlsxFile(object):
    """处理 xlsx 文件并获取内容"""
    def __init__(self, excel_file):
        try:
            self.wb = openpyxl.load_workbook(excel_file)
        # Invalid xlsx format
        except openpyxl.utils.exceptions.InvalidFileException as e:
            raise ValueError('错误！无效的 xlsx 文件: {}！{}'.format(excel_file, e))
        except openpyxl.utils.exceptions.ReadOnlyWorkbookException as e:
            raise ValueError('错误！此 xlsx 文件当前无法修改: {}！{}'.format(excel_file, e))
        except IOError as e:
            raise IOError('错误！文件异常: {}！{}'.format(excel_file, e))
        except Exception as e:
            raise ValueError('错误！{}'.format(e))

        self.ws = self.wb.active
        if not self.ws:
            raise ValueError("错误！无法获取 xlsx 文件中的 active sheet：{}！".format(excel_file))
        self.ws_title = self.ws.title
        self.headers = list()
        self.rows = list()
        self._parse_xlsx()

    def _parse_xlsx(self):
        """Get a two dimensional matrix from the xlsx file."""
        self.headers = list()
        self.rows = list()
        for i, row in enumerate(self.ws.rows):
            if i == 0:
                for cell in row:
                    self.headers.append(cell.value)
            else:
                row_container = list()
                for cell in row:
                    row_container.append(cell.value)
                self.rows.append(tuple(row_container))


def validate_data(xlsx):
    """校验模板文件中的数据是否合法"""
    if not xlsx:
        sys.exit("错误！读取 xlsx 文件失败！")

    if not xlsx.headers:
        sys.exit("错误！未读取到有效的标题内容（请不要修改模板）！")

    if not xlsx.rows:
        sys.exit("错误！未读取到有效的数据内容（请不要修改模板）！")

    if xlsx.headers != DEFAULT_HEADER_CONTENT:
        sys.exit("错误！请不要修改模板标题行！")

    for i, row in enumerate(xlsx.rows):
        if not row or len(row) != len(xlsx.headers):
            sys.exit("错误！第 {} 行 cell 数目 {} 与模板标题 cell 数目 {} 不一致！".format(
                i + 1, len(row), len(xlsx.headers)))
        for j, cell in enumerate(row):
            if j not in [1, 2, 3, 7] and not cell:
                sys.exit("错误！第 {} 行第 {} 列 cell 不能为空！".format(i + 1, j + 1))
            if j == 8:
                if cell not in ['左', '右']:
                    sys.exit("错误！第 {} 行第 {} 列 cell 必须为左 / 右！".format(i + 1, j + 1))
            if j in [9, 10]:
                if not cell.startswith('#') and len(cell) == 7:
                    sys.exit("错误！第 {} 行第 {} 列 cell 必须为类似 #B2EBF2 的 hex color 格式！".format(i + 1, j + 1))

    print('共需处理 {} 页内容。'.format(len(xlsx.rows)))


def choose_target_template(row):
    """基于数据选择应当使用的模板"""
    if not row:
        return None, None
    if row[8] == '左':
        if not row[7]:
            return TEMPLATE_HTML_LEFT_2_IMAGES, TEMPLATE_CSS_2_IMAGES
        else:
            return TEMPLATE_HTML_LEFT_3_IMAGES, TEMPLATE_CSS_3_IMAGES
    elif row[8] == '右':
        if not row[7]:
            return TEMPLATE_HTML_RIGHT_2_IMAGES, TEMPLATE_CSS_2_IMAGES
        else:
            return TEMPLATE_HTML_RIGHT_3_IMAGES, TEMPLATE_CSS_3_IMAGES


def generate_html_by_templates(xlsx):
    """基于模板生成 PDF"""
    if not os.path.isdir(RESULT_FOLDER):
        os.mkdir(RESULT_FOLDER)

    for i, row in enumerate(xlsx.rows):
        print('>> 开始处理第 {} 个 PDF 页面...'.format(i + 1))

        name = row[0]
        student_id = row[1]
        qq = row[2]
        wechat = row[3]
        intro_image_path = row[4]
        content_image_a_path = row[5]
        content_image_b_path = row[6]
        content_image_c_path = row[7]
        orientation_type = row[8]
        intro_background_color = row[9]
        content_background_color = row[10]
        content_text = row[11]

        html_template, css_template = choose_target_template(row)
        html_content = template_cache.get(html_template)
        css_content = template_cache.get(css_template)
        html_content = html_content \
            .replace(PLACEHOLDER_CSS_FILE, TEMP_CSS_FILE) \
            .replace(PLACEHOLDER_NAME, name) \
            .replace(PLACEHOLDER_STUDENT_ID, '{}'.format(student_id)) \
            .replace(PLACEHOLDER_QQ, '{}'.format(qq)) \
            .replace(PLACEHOLDER_WECHAT, '{}'.format(wechat)) \
            .replace(PLACEHOLDER_INTRO_IMAGE_PATH, intro_image_path) \
            .replace(PLACEHOLDER_CONTENT_IMAGE_A_PATH, content_image_a_path) \
            .replace(PLACEHOLDER_CONTENT_IMAGE_B_PATH, content_image_b_path) \
            .replace(PLACEHOLDER_CONTENT_IMAGE_C_PATH, content_image_c_path if content_image_c_path else '') \
            .replace(PLACEHOLDER_PAGE_ORIENTATION, orientation_type) \
            .replace(PLACEHOLDER_CONTENT_TEXT, content_text)
        css_content = css_content \
            .replace(PLACEHOLDER_INTRO_BACKGROUND_COLOR_LEFT, intro_background_color) \
            .replace(PLACEHOLDER_CONTENT_BACKGROUND_COLOR_LEFT, content_background_color) \
            .replace(PLACEHOLDER_INTRO_BACKGROUND_COLOR_RIGHT, intro_background_color) \
            .replace(PLACEHOLDER_CONTENT_BACKGROUND_COLOR_RIGHT, content_background_color)

        with open(TEMP_HTML_FILE, 'w') as f:
            f.write(html_content)
        with open(TEMP_CSS_FILE, 'w') as f:
            f.write(css_content)

        result_file_path = os.path.join(RESULT_FOLDER, RESULT_PDF_NAME_TEMPLATE % (i + 1))
        pdfkit.from_file(TEMP_HTML_FILE, result_file_path, options=OPTIONS)

        print('<< 已生成 PDF 文件：' + result_file_path)


def do_cleaning():
    """执行清理操作，删除临时文件"""
    if os.path.isfile(TEMP_HTML_FILE):
        os.remove(TEMP_HTML_FILE)

    if os.path.isfile(TEMP_CSS_FILE):
        os.remove(TEMP_CSS_FILE)


def main():
    """执行主逻辑流程"""
    xlsx = XlsxFile(TEMPLATE_XLSX_FILE)
    validate_data(xlsx)
    prepare_template_cache()
    generate_html_by_templates(xlsx)
    do_cleaning()


if __name__ == '__main__':
    main()
