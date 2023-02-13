import xlwt
from xmindparser import xmind_to_dict
import openpyxl.styles
import sys
import time
import os

dataStyle = xlwt.XFStyle()
bgColor = xlwt.Pattern()
bgColor.pattern = xlwt.Pattern.SOLID_PATTERN
bgColor.pattern_fore_colour = 26  # 背景颜色
dataStyle.pattern = bgColor


class XMIND:

    def font(self, colour_index):
        titleStyle = xlwt.XFStyle()  # 初始化样式
        titleFont = xlwt.Font()
        titleFont.name = "宋体"
        titleFont.bold = True  # 加粗
        titleFont.height = 11 * 20  # 字号
        titleFont.colour_index = colour_index  # 字体颜色
        titleStyle.font = titleFont
        return titleStyle

    def xmind_file(self, filename):
        # filename = "C:\\Users\\firerock\\Desktop\\项目整理\\认证\\认证测试用例1.0.xmind"
            a = xmind_to_dict(filename)
            b = a[0]['topic']['topics']
            f = xlwt.Workbook()
            sheet = f.add_sheet(a[0]['topic']['title'], cell_overwrite_ok=True)
            # print('a：{}'.format(a))
            # print('b：{}'.format(b))
            # print('b[0]：{}'.format(b[0]))
            # print('b长度：{}'.format(len(b)))
            # print("b[0]['topics']：{}".format(len(b[0]['topics'])))
            return a, b, sheet, f

    def xmind_to_excel(self, filname):
        index = 0
        success = 0
        failed = 0
        no_executed = 0
        invalid = 0
        a, b, sheet, f = self.xmind_file(filname)
        """设置首行"""
        heard = ['用例编号', '测试项目', '测试模块', '页面/功能', '用例标题', '测试等级', '前置条件', '步骤', '预期结果', '成功/失败/未执行/无效', '测试数据', '统计']
        for s in range(len(heard)):
            sheet.write(0, s, heard[s], dataStyle)
        """写入数据"""
        for module in range(len(b)):  # 模块数量
            for page in range(len(b[module]['topics'])):  # 页面
                case_split1 = b[module]['topics'][page]['title'].split('-')
                preposition1 = b[module]['topics'][page]['title'].split('：')
                # if b[module]['topics'][page].__contains__('topics') is False:
                #     sheet.write(index + 1, 3, b[module]['topics'][page]['title'])
                #     sheet.write(index + 1, 1, a[0]['topic']['title'])
                #     sheet.write(index + 1, 2, b[module]['title'])
                #     sheet.write(index + 1, 0, 'cs{}'.format(index + 1))
                #     index += 1
                if case_split1[0] == 'tc':  # 判断模块下一级为测试用例
                    case_split2 = case_split1[1].split('：')
                    sheet.write(index + 1, 4, case_split2[1])
                    sheet.write(index + 1, 1, a[0]['topic']['title'])
                    sheet.write(index + 1, 2, b[module]['title'])
                    sheet.write(index + 1, 0, 'cs{}'.format(index + 1))
                    # if b[module]['topics'][page].__contains__('title'):
                    if b[module]['topics'][page].__contains__('topics') is False:   # 判断测试用例后面是否有数据
                        if b[module]['topics'][page].__contains__('makers') is False:
                            sheet.write(index + 1, 9, '未执行', self.font(colour_index=18))
                            no_executed += 1
                        elif b[module]['topics'][page]['makers'][0] == 'task-done':
                            sheet.write(index + 1, 9, '成功', self.font(colour_index=20))
                            success += 1
                        elif b[module]['topics'][page]['makers'][0] == 'tag-grey':
                            sheet.write(index + 1, 9, '无效', self.font(colour_index=39))
                            invalid += 1
                        else:
                            sheet.write(index + 1, 9, '失败', self.font(colour_index=30))
                            failed += 1
                        if case_split2[0] == 'p1':
                            sheet.write(index + 1, 5, 'height', self.font(colour_index=18))
                        elif case_split2[0] == 'p2':
                            sheet.write(index + 1, 5, 'mid', self.font(colour_index=20))
                        else:
                            sheet.write(index + 1, 5, 'low', self.font(colour_index=30))
                        #     sheet.write(index + 1, 8, )
                    else:
                        if b[module]['topics'][page]['topics'][0].__contains__('topics') and \
                                b[module]['topics'][page]['topics'][0].__contains__('title'):
                            sheet.write(index + 1, 7, b[module]['topics'][page]['topics'][0]['title'])  # 步骤
                            sheet.write(index + 1, 8, b[module]['topics'][page]['topics'][0]['topics'][0]['title'])  # 预期结果
                        if b[module]['topics'][page]['topics'][0].__contains__('topics') is False:
                            # b[module]['topics'][page]['topics'][0].__contains__('title') is False:
                            sheet.write(index + 1, 8, b[module]['topics'][page]['topics'][0]['title'])  # 没有步骤写预期结果
                        if b[module]['topics'][page].__contains__('makers') is False:
                            sheet.write(index + 1, 9, '未执行', self.font(colour_index=18))
                            no_executed += 1
                        elif b[module]['topics'][page]['makers'][0] == 'task-done':
                            sheet.write(index + 1, 9, '成功', self.font(colour_index=20))
                            success += 1
                        elif b[module]['topics'][page]['makers'][0] == 'tag-grey':
                            sheet.write(index + 1, 9, '无效', self.font(colour_index=39))
                            invalid += 1
                        else:
                            sheet.write(index + 1, 9, '失败', self.font(colour_index=30))
                            failed += 1
                        if case_split2[0] == 'p1':
                            sheet.write(index + 1, 5, 'height', self.font(colour_index=18))
                        elif case_split2[0] == 'p2':
                            sheet.write(index + 1, 5, 'mid', self.font(colour_index=20))
                        else:
                            sheet.write(index + 1, 5, 'low', self.font(colour_index=30))
                    index += 1
                elif preposition1[0] == 'pc':
                    sheet.write(index + 1, 6, preposition1[1])
                else:
                    for case in range(len(b[module]['topics'][page]['topics'])):   # 用例数量
                        case_split3 = b[module]['topics'][page]['topics'][case]['title'].split('-')
                        preposition2 = b[module]['topics'][page]['topics'][case]['title'].split('：')
                        # """用例"""
                        # sheet.write(index + 1, 4, b[module]['topics'][page]['topics'][case]['title'])
                        # sheet.write(index + 1, 3, b[module]['topics'][page]['title'])
                        # sheet.write(index + 1, 1, a[0]['topic']['title'])
                        # sheet.write(index + 1, 2, b[module]['title'])
                        # sheet.write(index + 1, 0, 'cs{}'.format(index + 1))
                        # if b[module]['topics'][page]['topics'][case].__contains__('topics'):
                        #     """预期结果"""
                        #     sheet.write(index + 1, 5, b[module]['topics'][page]['topics'][case]['topics'][0]['title'])
                        # index += 1
                        if case_split3[0] == 'tc':
                            # print(case_split3)
                            case_split4 = case_split3[1].split('：')
                            # print(case_split4)
                            sheet.write(index + 1, 4, case_split4[1])
                            sheet.write(index + 1, 1, a[0]['topic']['title'])
                            sheet.write(index + 1, 3, b[module]['topics'][page]['title'])
                            sheet.write(index + 1, 2, b[module]['title'])
                            sheet.write(index + 1, 0, 'cs{}'.format(index + 1))
                            if b[module]['topics'][page]['topics'][case].__contains__('topics'):
                                if b[module]['topics'][page]['topics'][case]['topics'][0].__contains__('topics') \
                                        is False:
                                    sheet.write(index + 1, 8,
                                                b[module]['topics'][page]['topics']
                                                [case]['topics'][0]['title'])  # 没有步骤写预期结果
                                else:
                                    sheet.write(index + 1, 7,
                                                b[module]['topics'][page]['topics'][case]['topics'][0]['title'])  # 步骤
                                    sheet.write(index + 1, 8,
                                                b[module]['topics'][page]['topics']
                                                [case]['topics'][0]['topics'][0]['title'])  # 预期结果
                            if b[module]['topics'][page]['topics'][case].__contains__('makers') is False:
                                sheet.write(index + 1, 9, '未执行', self.font(colour_index=18))
                                no_executed += 1
                            elif b[module]['topics'][page]['topics'][case]['makers'][0] == 'task-done':
                                sheet.write(index + 1, 9, '成功', self.font(colour_index=20))
                                success += 1
                            elif b[module]['topics'][page]['topics'][case]['makers'][0] == 'tag-grey':
                                sheet.write(index + 1, 9, '无效', self.font(colour_index=39))
                                invalid += 1
                            else:
                                sheet.write(index + 1, 9, '失败', self.font(colour_index=30))
                                failed += 1
                            if case_split4[0] == 'p1':
                                sheet.write(index + 1, 5, 'height', self.font(colour_index=18))
                            elif case_split4[0] == 'p2':
                                sheet.write(index + 1, 5, 'mid', self.font(colour_index=20))
                            else:
                                sheet.write(index + 1, 5, 'low', self.font(colour_index=30))
                            index += 1
                        elif preposition2[0] == 'pc':
                            sheet.write(index + 1, 6, preposition2[1])
                        else:
                            sheet.write(index + 1, 4, b[module]['topics'][page]['topics'][case]['title'])
                            sheet.write(index + 1, 1, a[0]['topic']['title'])
                            sheet.write(index + 1, 3, b[module]['topics'][page]['title'])
                            sheet.write(index + 1, 2, b[module]['title'])
                            sheet.write(index + 1, 0, 'cs{}'.format(index + 1))
                            index += 1
        status1 = ['成功', '失败', '未执行', '无效']
        status2 = success, failed, no_executed, invalid
        row = 11
        for i in range(len(status1)):
            sheet.write(i + 1, row, status1[i])
            sheet.write(i + 1, row + 1, status2[i])
        excel_name = filename.split('/')[0].split('xmind')[0] + 'xls'
        try:
            f.save(excel_name)
        except:
            print("请关闭excel文件再进行转换！")

    # def statistics(self, success, failed, no_executed, filname):
    #     sheet = self.xmind_file(filname)[2]


if __name__ == '__main__':
    XMIND = XMIND()
    filename = input('xmind路径：')
    # filename = 'C:\\Users\\firerock\\Desktop\\项目整理\\认证\\认证测试用例1.0.xmind'
    # filename = sys.argv[1]
    if filename.split(".") == "xmind":
        xlsname = filename.split("xmind")[0] + "xls"
        # print(xlsname)
        XMIND.xmind_to_excel(filename)
        print('运行完成')
        os.system(xlsname)
        time.sleep(2)
    else:
        print("文件格式错误")
