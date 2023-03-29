# -*- encoding: utf-8 -*-
"""
@Author: Aloha
@Time: 2023/3/19 23:46
@ProjectName: Practice
@FileName: xiaohongshu.py
@Software: PyCharm
"""
import os
import threading
import xlwt
import xlrd
import time
import json
import re
import execjs
from xlutils.copy import copy
from jsonpath import jsonpath
from requests_html import HTMLSession
# 构造请求对象
session = HTMLSession()


class XiaoHongShu(object):

    def __init__(self):
        self.url = "https://edith.xiaohongshu.com/api/sns/web/v1/homefeed"
        self.headers = {
            'authority': 'www.xiaohongshu.com',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'en,zh-CN;q=0.9,zh;q=0.8,en-US;q=0.7',
            'cache-control': 'max-age=0',
            'cookie': 'gid.sig=jUUJ8j30BuyEnFsH8HOflhuFzL4UK3vIVo5hdlDLcrQ; gid.ss=gSMQ9UOnDuZwH2oRGJG6BW6e4grs67TaYpnrW+8Wmd2A+uY3PnUwtQYlPTc3z0qq; timestamp2=1665557283121cdd4df5921b188df3a88c52607b1a312d980f5fd8894c12830; timestamp2.sig=ctqCrcFaLcEs_OyHzrAVRLSMIvs-GjnrpgsfHx-YcZ0; xhsTrackerId=186fb338-3314-432c-93d2-4e1bc1eedc75; xhsTrackerId.sig=r00ZEygp13x21wTCT8xZG2k6RvT8e965rofnTOxDP9s; a1=186d68913621madqypfrpkd042a74lm1xaifugql350000405317; webId=76cf297506a6621491186e860c52e71a; gid=yYKfKYjyj0C0yYKfKYjyq0kAKJy60fAvUiFU31qTf84J0i28W416y7888482qyW8Yjj0YfYW; gid.sign=UMjVO9YIj4Xyuc6xUg99+4lnvwg=; web_session=040069b3a673305d985646be38364bbb2fd224; xhsTracker=url=root&searchengine=baidu; xhsTracker.sig=kH7g0qqCGD4yamBpYpTbRDtYdtYDZi-5XRxuhBF924w; smidV2=20230314102946dfdd891cadc0d1ff4431b719384d3f7600fc44e604be3aa80; extra_exp_ids=h5_230301_origin,h5_1208_exp3,h5_1130_exp1,ios_wx_launch_open_app_exp,h5_video_ui_exp3,wx_launch_open_app_duration_origin,ques_clt2; extra_exp_ids.sig=7a-V7A-aDw2Yj_jOtwzrUZVeC2VvKiOIvgtgDlY9l64; websectiga=10f9a40ba454a07755a08f27ef8194c53637eba4551cf9751c009d9afb564467; sec_poison_id=c43b17d6-a2c8-4e05-bbc7-e48d04c43d4e; webBuild=1.2.6; xsecappid=xhs-pc-web; extra_exp_ids=h5_230301_origin,h5_1208_exp3,h5_1130_exp1,ios_wx_launch_open_app_exp,h5_video_ui_exp3,wx_launch_open_app_duration_origin,ques_clt2; extra_exp_ids.sig=7a-V7A-aDw2Yj_jOtwzrUZVeC2VvKiOIvgtgDlY9l64',
            'referer': 'https://www.xiaohongshu.com/website-login/captcha?redirectPath=https%3A%2F%2Fwww.xiaohongshu.com%2Fuser%2Fprofile%2F5f095dee000000000101cbd2',
            'sec-ch-ua': '"Google Chrome";v="105", "Not)A;Brand";v="8", "Chromium";v="105"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-user': '?1',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
        }

    def decrypt(self, e, t):
        # e, t为逆向所需参数
        params = execjs.compile(open('../xiaohongshu.js', 'r', encoding='utf-8').read()).call('sign', e, t)
        return params

    def firstPage(self):
        payload = "{\"cursor_score\":\"89119.28099999724\",\"num\":40,\"refresh_type\":1,\"note_index\":0,\"unread_begin_note_id\":\"63ea07e70000000007038aea\",\"unread_end_note_id\":\"640eeb7d0000000027002fe0\",\"unread_note_count\":50,\"category\":\"homefeed_recommend\"}"
        headers = {
            'authority': 'edith.xiaohongshu.com',
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en,zh-CN;q=0.9,zh;q=0.8,en-US;q=0.7',
            'content-type': 'application/json;charset=UTF-8',
            'cookie': 'gid.sig=jUUJ8j30BuyEnFsH8HOflhuFzL4UK3vIVo5hdlDLcrQ; gid.ss=gSMQ9UOnDuZwH2oRGJG6BW6e4grs67TaYpnrW+8Wmd2A+uY3PnUwtQYlPTc3z0qq; timestamp2=1665557283121cdd4df5921b188df3a88c52607b1a312d980f5fd8894c12830; timestamp2.sig=ctqCrcFaLcEs_OyHzrAVRLSMIvs-GjnrpgsfHx-YcZ0; xhsTrackerId=186fb338-3314-432c-93d2-4e1bc1eedc75; xhsTrackerId.sig=r00ZEygp13x21wTCT8xZG2k6RvT8e965rofnTOxDP9s; a1=186d68913621madqypfrpkd042a74lm1xaifugql350000405317; webId=76cf297506a6621491186e860c52e71a; gid=yYKfKYjyj0C0yYKfKYjyq0kAKJy60fAvUiFU31qTf84J0i28W416y7888482qyW8Yjj0YfYW; gid.sign=UMjVO9YIj4Xyuc6xUg99+4lnvwg=; web_session=040069b3a673305d985646be38364bbb2fd224; extra_exp_ids=h5_230301_origin,h5_1208_exp3,h5_1130_exp1,ios_wx_launch_open_app_exp,h5_video_ui_exp3,wx_launch_open_app_duration_origin,ques_clt2; extra_exp_ids.sig=7a-V7A-aDw2Yj_jOtwzrUZVeC2VvKiOIvgtgDlY9l64; xhsTracker=url=root&searchengine=baidu; xhsTracker.sig=kH7g0qqCGD4yamBpYpTbRDtYdtYDZi-5XRxuhBF924w; webBuild=1.2.3; websectiga=29098a4cf41f76ee3f8db19051aaa60c0fc7c5e305572fec762da32d457d76ae; sec_poison_id=ad58d408-8eff-45c4-b670-d5232b1ee47a; acw_tc=4ebf985f59088837a9ce37803c5a28de3e8b2bfbdb87c514f1d2fb0ab1a54591; smidV2=20230314102946dfdd891cadc0d1ff4431b719384d3f7600fc44e604be3aa80; xsecappid=xhs-pc-web; acw_tc=5229f571b912e93a1c6cffe1905aa2beda70ae65329f40e95c550bb6b4903771',
            'origin': 'https://www.xiaohongshu.com',
            'referer': 'https://www.xiaohongshu.com/',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36',
            'x-b3-traceid': '7561792b6cf73873',
            'x-s': 'slvGZjv+OgclsYT+Z2OvOgU60gUBs2wvO2Mlsl9bOiF3',
            'x-t': '1678761351062'
        }
        response = session.post(self.url, headers=headers, data=payload).json()
        # 获取下一页的关键字
        resp = response['data']
        next_key = resp['cursor_score']
        res = response['data']['items']
        for data in res:
            bid = data['id']  # 作品id
            title = data['note_card']['display_title']
            nickname = data['note_card']['user']['nickname']  # 昵称
            face = data['note_card']['user']['avatar']  # 头像
            uid = data['note_card']['user']['user_id']  # uid
            self.userInfo(uid)
            self.userWork(bid)
            data = {
                '小红书主页': [bid, title, nickname, face, uid]
            }
            # self.SaveExcels(data)
        return next_key

    def nextPage(self):
        next_key = self.firstPage()
        note_index = 10
        while True:
            # e, t为逆向所需参数
            e = '/api/sns/web/v1/homefeed'
            t = {
                "cursor_score": f"{next_key}",
                "num": "12",
                "refresh_type": "3",
                "note_index": f"{note_index}",
                "unread_begin_note_id": "",
                "unread_end_note_id": "",
                "unread_note_count": "0",
                "category": "homefeed_recommend"
            }
            params = self.decrypt(e, t)
            payload = "{\"cursor_score\":\"" + next_key + "\",\"num\":\"12\",\"refresh_type\":\"3\",\"note_index\":\"" + str(note_index) + "\",\"unread_begin_note_id\":\"\",\"unread_end_note_id\":\"\",\"unread_note_count\":\"0\",\"category\":\"homefeed_recommend\"}"
            headers = {
                'authority': 'edith.xiaohongshu.com',
                'accept': 'application/json, text/plain, */*',
                'accept-language': 'en,zh-CN;q=0.9,zh;q=0.8,en-US;q=0.7',
                'content-type': 'application/json;charset=UTF-8',
                'cookie': 'gid.sig=jUUJ8j30BuyEnFsH8HOflhuFzL4UK3vIVo5hdlDLcrQ; gid.ss=gSMQ9UOnDuZwH2oRGJG6BW6e4grs67TaYpnrW+8Wmd2A+uY3PnUwtQYlPTc3z0qq; timestamp2=1665557283121cdd4df5921b188df3a88c52607b1a312d980f5fd8894c12830; timestamp2.sig=ctqCrcFaLcEs_OyHzrAVRLSMIvs-GjnrpgsfHx-YcZ0; xhsTrackerId=186fb338-3314-432c-93d2-4e1bc1eedc75; xhsTrackerId.sig=r00ZEygp13x21wTCT8xZG2k6RvT8e965rofnTOxDP9s; xsecappid=xhs-pc-web; a1=186d68913621madqypfrpkd042a74lm1xaifugql350000405317; webId=76cf297506a6621491186e860c52e71a; gid=yYKfKYjyj0C0yYKfKYjyq0kAKJy60fAvUiFU31qTf84J0i28W416y7888482qyW8Yjj0YfYW; gid.sign=UMjVO9YIj4Xyuc6xUg99+4lnvwg=; web_session=040069b3a673305d985646be38364bbb2fd224; webBuild=1.2.3; xhsTracker=url=explore&searchengine=baidu; xhsTracker.sig=u1cFYHAwm89lKbFLL1Y8vp9JcskioXWTa56RKaAB2ys; websectiga=a9bdcaed0af874f3a1431e94fbea410e8f738542fbb02df1e8e30c29ef3d91ac; sec_poison_id=f9e89c49-24cb-4153-a301-d7f2c870f4f0; extra_exp_ids=h5_230301_origin,h5_1208_exp3,h5_1130_exp1,ios_wx_launch_open_app_exp,h5_video_ui_exp3,wx_launch_open_app_duration_origin,ques_clt2; extra_exp_ids.sig=7a-V7A-aDw2Yj_jOtwzrUZVeC2VvKiOIvgtgDlY9l64',
                'origin': 'https://www.xiaohongshu.com',
                'referer': 'https://www.xiaohongshu.com/',
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36',
                'x-b3-traceid': '5421b8be1ce96532',
                'x-s': params['X-s'],
                'x-t': str(params['X-t'])
            }
            response = session.post(self.url, headers=headers, data=payload).json()
            # 获取下一页的关键字
            resp = response['data']
            next_key = resp['cursor_score']
            res = response['data']['items']
            for data in res:
                bid = data['id']  # 作品id
                title = data['note_card']['display_title']
                nickname = data['note_card']['user']['nickname']  # 昵称
                face = data['note_card']['user']['avatar']  # 头像
                uid = data['note_card']['user']['user_id']  # uid
                self.userInfo(uid)
                self.userWork(bid)
                data = {
                    '小红书主页': [bid, title, nickname, face, uid]
                }
                # self.SaveExcels(data)
            next_key = next_key
            note_index += 10
            if note_index == 220:
                break

    def userInfo(self, uid):
        try:
            url = f'https://www.xiaohongshu.com/user/profile/{uid}'
            response = session.get(url, headers=self.headers).text
            resp = str(re.findall('window.__INITIAL_STATE__=(.*?)</script><script>', response)[0]).replace('undefined', 'false')
            res = json.loads(resp)
            if res['user']['userPageData'] != {}:
                red_id = res['user']['userPageData']['basicInfo']['redId']  # 小红书号
                nickname = res['user']['userPageData']['basicInfo']['nickname']  # 昵称
                face = res['user']['userPageData']['basicInfo']['images']  # 头像
                gender = res['user']['userPageData']['basicInfo']['gender']  # 性别
                location = res['user']['userPageData']['basicInfo']['ipLocation']  # ip地址
                interactions = jsonpath(res, '$...count')
                follows = interactions[0]  # 关注数
                fans = interactions[1]  # 粉丝数
                interaction = interactions[2]  # 获赞与收藏
                if res['user']['userPageData']['basicInfo']['desc'] == '':
                    signature = '无'
                else:
                    signature = str(res['user']['userPageData']['basicInfo']['desc']).replace('\n', ', ')  # 简介
                data = {
                    '用户信息': [red_id, nickname, face, gender, location, follows, fans, interaction, signature,
                             url]
                }
                # await self.SaveUser(data)
                print(f'用户 {nickname} 信息采集完毕--------------------------')
            else:
                red_id = '无'
                nickname = '无'
                face = '无'
                gender = '无'
                location = '无'
                follows = '无'
                fans = '无'
                interaction = '无'
                signature = '无'
                data = {
                    '用户信息': [red_id, nickname, face, gender, location, follows, fans, interaction, signature,
                             url]
                }
                # self.SaveUser(data)
                print(f'用户 {nickname} 信息采集完毕--------------------------')
        except Exception as e:
            print(e)

    def userWork(self, aid):
        try:
            # e, t为逆向所需参数
            e = '/api/sns/web/v1/feed'
            t = {"source_note_id": aid}
            params = self.decrypt(e, t)
            url = "https://edith.xiaohongshu.com/api/sns/web/v1/feed"
            payload = "{\"source_note_id\":\"" + aid + "\"}"
            headers = {
                'authority': 'edith.xiaohongshu.com',
                'accept': 'application/json, text/plain, */*',
                'accept-language': 'en,zh-CN;q=0.9,zh;q=0.8,en-US;q=0.7',
                'content-type': 'application/json;charset=UTF-8',
                'cookie': 'gid.sig=jUUJ8j30BuyEnFsH8HOflhuFzL4UK3vIVo5hdlDLcrQ; gid.ss=gSMQ9UOnDuZwH2oRGJG6BW6e4grs67TaYpnrW+8Wmd2A+uY3PnUwtQYlPTc3z0qq; timestamp2=1665557283121cdd4df5921b188df3a88c52607b1a312d980f5fd8894c12830; timestamp2.sig=ctqCrcFaLcEs_OyHzrAVRLSMIvs-GjnrpgsfHx-YcZ0; xhsTrackerId=186fb338-3314-432c-93d2-4e1bc1eedc75; xhsTrackerId.sig=r00ZEygp13x21wTCT8xZG2k6RvT8e965rofnTOxDP9s; a1=186d68913621madqypfrpkd042a74lm1xaifugql350000405317; webId=76cf297506a6621491186e860c52e71a; gid=yYKfKYjyj0C0yYKfKYjyq0kAKJy60fAvUiFU31qTf84J0i28W416y7888482qyW8Yjj0YfYW; gid.sign=UMjVO9YIj4Xyuc6xUg99+4lnvwg=; web_session=040069b3a673305d985646be38364bbb2fd224; xhsTracker=url=root&searchengine=baidu; xhsTracker.sig=kH7g0qqCGD4yamBpYpTbRDtYdtYDZi-5XRxuhBF924w; webBuild=1.2.3; smidV2=20230314102946dfdd891cadc0d1ff4431b719384d3f7600fc44e604be3aa80; xsecappid=xhs-pc-web; acw_tc=36554e5fe34db32dea09632fe9e1051d87bc200707e20ba5f0a5343287fbb017; extra_exp_ids=h5_230301_origin,h5_1208_exp3,h5_1130_exp1,ios_wx_launch_open_app_exp,h5_video_ui_exp3,wx_launch_open_app_duration_origin,ques_clt2; extra_exp_ids.sig=7a-V7A-aDw2Yj_jOtwzrUZVeC2VvKiOIvgtgDlY9l64; websectiga=9730ffafd96f2d09dc024760e253af6ab1feb0002827740b95a255ddf6847fc8; sec_poison_id=d92ed8a9-78b2-45cc-b4e1-ed0eb7bc23fa',
                'origin': 'https://www.xiaohongshu.com',
                'referer': 'https://www.xiaohongshu.com/',
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36',
                'x-b3-traceid': 'c63d082130b0bf14',
                'x-s': params['X-s'],
                'x-t': str(params['X-t'])
            }
            response = session.post(url, headers=headers, data=payload).json()['data']['items']
            for res in response:
                user_id = res['note_card']['user']['user_id']  # uid
                nickname = res['note_card']['user']['nickname']  # 昵称
                face = res['note_card']['user']['avatar']  # 头像
                title = res['note_card']['title']  # 标题
                content = str(res['note_card']['desc']).replace('\n', ', ').lstrip()  # 内容
                """ 格式化时间戳 """
                j = res['note_card']['time']  # 时间
                n = time.localtime(j / 1000)  # 将时间戳转换成时间元祖tuple
                j_time = time.strftime("%Y-%m-%d %H:%M:%S", n)  # 格式化输出时间
                liked = res['note_card']['interact_info']['liked_count']  # 点赞数
                share = res['note_card']['interact_info']['share_count']  # 分享数
                comment = res['note_card']['interact_info']['comment_count']  # 评论数
                collected = res['note_card']['interact_info']['collected_count']  # 收藏数
                tag_list = jsonpath(res, '$..name')
                image_list = jsonpath(res, '$..url')
                link = f'https://www.xiaohongshu.com/explore/{aid}'
                for img in image_list:
                    image = img + '?imageView2/2/h/1200/format/webp'
                data = {
                    '作品数据': [user_id, nickname, face, title, content, j_time, liked, share, comment, collected, tag_list, image, link]
                }
                # self.SaveWork(data)
                print(f'作品 {title} 采集完毕-------------------')
        except Exception as e:
            print(e)

    def SaveExcels(self, data):
        """
        使用前，请先阅读代码
        :param data: 需要保存的data字典(有格式要求)
        :return:
        格式要求:
            data = {
            '基本详情': ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']
        }
        """
        # 获取表的名称
        sheet_name = [i for i in data.keys()][0]
        # 创建保存excel表格的文件夹
        # os.getcwd() 获取当前文件路径
        os_mkdir_path = os.getcwd() + '/小红书数据/'
        # 判断这个路径是否存在，不存在就创建
        if not os.path.exists(os_mkdir_path):
            os.mkdir(os_mkdir_path)
        # 判断excel表格是否存在           工作簿文件名称
        os_excel_path = os_mkdir_path + '数据.xls'
        if not os.path.exists(os_excel_path):
            # 不存在，创建工作簿(也就是创建excel表格)
            workbook = xlwt.Workbook(encoding='utf-8')
            """工作簿中创建新的sheet表"""  # 设置表名
            worksheet1 = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            """设置sheet表的表头"""
            sheet1_headers = ('作品id', '作品标题', '昵称', '头像', 'uid')
            # 将表头写入工作簿
            for header_num in range(0, len(sheet1_headers)):
                # 设置表格长度
                worksheet1.col(header_num).width = 2560 * 3
                # 写入表头        行,    列,           内容
                worksheet1.write(0, header_num, sheet1_headers[header_num])
            # 循环结束，代表表头写入完成，保存工作簿
            workbook.save(os_excel_path)
        """=============================已有工作簿添加新表==============================================="""
        # 打开工作薄
        workbook = xlrd.open_workbook(os_excel_path)
        # 获取工作薄中所有表的名称
        sheets_list = workbook.sheet_names()
        # 如果表名称：字典的key值不在工作簿的表名列表中
        if sheet_name not in sheets_list:
            # 复制先有工作簿对象
            work = copy(workbook)
            # 通过复制过来的工作簿对象，创建新表  -- 保留原有表结构
            sh = work.add_sheet(sheet_name)
            # 给新表设置表头
            excel_headers_tuple = ('作品id', '作品标题', '昵称', '头像', 'uid')
            for head_num in range(0, len(excel_headers_tuple)):
                sh.col(head_num).width = 2560 * 3
                #               行，列，  内容，            样式
                sh.write(0, head_num, excel_headers_tuple[head_num])
            work.save(os_excel_path)
        """========================================================================================="""
        # 判断工作簿是否存在
        if os.path.exists(os_excel_path):
            # 打开工作簿
            workbook = xlrd.open_workbook(os_excel_path)
            # 获取工作薄中所有表的个数
            sheets = workbook.sheet_names()
            for i in range(len(sheets)):
                for name in data.keys():
                    worksheet = workbook.sheet_by_name(sheets[i])
                    # 获取工作薄中所有表中的表名与数据名对比
                    if worksheet.name == name:
                        # 获取表中已存在的行数
                        rows_old = worksheet.nrows
                        # 将xlrd对象拷贝转化为xlwt对象
                        new_workbook = copy(workbook)
                        # 获取转化后的工作薄中的第i张表
                        new_worksheet = new_workbook.get_sheet(i)
                        for num in range(0, len(data[name])):
                            new_worksheet.write(rows_old, num, data[name][num])
                        new_workbook.save(os_excel_path)

    def SaveUser(self, data):
        """
        使用前，请先阅读代码
        :param data: 需要保存的data字典(有格式要求)
        :return:
        格式要求:
            data = {
            '基本详情': ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']
        }
        """
        # 获取表的名称
        sheet_name = [i for i in data.keys()][0]
        # 创建保存excel表格的文件夹
        # os.getcwd() 获取当前文件路径
        os_mkdir_path = os.getcwd() + '/小红书数据/'
        # 判断这个路径是否存在，不存在就创建
        if not os.path.exists(os_mkdir_path):
            os.mkdir(os_mkdir_path)
        # 判断excel表格是否存在           工作簿文件名称
        # os_excel_path = os_mkdir_path + '数据.xls'
        os_excel_path = os_mkdir_path + '数据.xls'
        if not os.path.exists(os_excel_path):
            # 不存在，创建工作簿(也就是创建excel表格)
            workbook = xlwt.Workbook(encoding='utf-8')
            """工作簿中创建新的sheet表"""  # 设置表名
            worksheet1 = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            """设置sheet表的表头"""
            sheet1_headers = ('小红书号', '昵称', '头像', '性别(0--男, 1--女)', 'IP地址', '关注数', '粉丝数', '获赞与收藏', '简介', '用户主页')
            # 将表头写入工作簿
            for header_num in range(0, len(sheet1_headers)):
                # 设置表格长度
                worksheet1.col(header_num).width = 2560 * 3
                # 写入表头        行,    列,           内容
                worksheet1.write(0, header_num, sheet1_headers[header_num])
            # 循环结束，代表表头写入完成，保存工作簿
            workbook.save(os_excel_path)
        """=============================已有工作簿添加新表==============================================="""
        # 打开工作薄
        workbook = xlrd.open_workbook(os_excel_path)
        # 获取工作薄中所有表的名称
        sheets_list = workbook.sheet_names()
        # 如果表名称：字典的key值不在工作簿的表名列表中
        if sheet_name not in sheets_list:
            # 复制先有工作簿对象
            work = copy(workbook)
            # 通过复制过来的工作簿对象，创建新表  -- 保留原有表结构
            sh = work.add_sheet(sheet_name)
            # 给新表设置表头
            excel_headers_tuple = ('小红书号', '昵称', '头像', '性别(0--男, 1--女)', 'IP地址', '关注数', '粉丝数', '获赞与收藏', '简介', '用户主页')
            for head_num in range(0, len(excel_headers_tuple)):
                sh.col(head_num).width = 2560 * 3
                #               行，列，  内容，            样式
                sh.write(0, head_num, excel_headers_tuple[head_num])
            work.save(os_excel_path)
        """========================================================================================="""
        # 判断工作簿是否存在
        if os.path.exists(os_excel_path):
            # 打开工作簿
            workbook = xlrd.open_workbook(os_excel_path)
            # 获取工作薄中所有表的个数
            sheets = workbook.sheet_names()
            for i in range(len(sheets)):
                for name in data.keys():
                    worksheet = workbook.sheet_by_name(sheets[i])
                    # 获取工作薄中所有表中的表名与数据名对比
                    if worksheet.name == name:
                        # 获取表中已存在的行数
                        rows_old = worksheet.nrows
                        # 将xlrd对象拷贝转化为xlwt对象
                        new_workbook = copy(workbook)
                        # 获取转化后的工作薄中的第i张表
                        new_worksheet = new_workbook.get_sheet(i)
                        for num in range(0, len(data[name])):
                            new_worksheet.write(rows_old, num, data[name][num])
                        new_workbook.save(os_excel_path)

    def SaveWork(self, data):
        """
        使用前，请先阅读代码
        :param data: 需要保存的data字典(有格式要求)
        :return:
        格式要求:
            data = {
            '基本详情': ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']
        }
        """
        # 获取表的名称
        sheet_name = [i for i in data.keys()][0]
        # 创建保存excel表格的文件夹
        # os.getcwd() 获取当前文件路径
        os_mkdir_path = os.getcwd() + '/小红书数据/'
        # 判断这个路径是否存在，不存在就创建
        if not os.path.exists(os_mkdir_path):
            os.mkdir(os_mkdir_path)
        # 判断excel表格是否存在           工作簿文件名称
        # os_excel_path = os_mkdir_path + '数据.xls'
        os_excel_path = os_mkdir_path + '数据.xls'
        if not os.path.exists(os_excel_path):
            # 不存在，创建工作簿(也就是创建excel表格)
            workbook = xlwt.Workbook(encoding='utf-8')
            """工作簿中创建新的sheet表"""  # 设置表名
            worksheet1 = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            """设置sheet表的表头"""
            sheet1_headers = ('uid', '昵称', '头像', '标题', '内容', '发布时间', '点赞数', '分享数', '评论数', '收藏数', '标签', '作品图片', '详情页面')
            # 将表头写入工作簿
            for header_num in range(0, len(sheet1_headers)):
                # 设置表格长度
                worksheet1.col(header_num).width = 2560 * 3
                # 写入表头        行,    列,           内容
                worksheet1.write(0, header_num, sheet1_headers[header_num])
            # 循环结束，代表表头写入完成，保存工作簿
            workbook.save(os_excel_path)
        """=============================已有工作簿添加新表==============================================="""
        # 打开工作薄
        workbook = xlrd.open_workbook(os_excel_path)
        # 获取工作薄中所有表的名称
        sheets_list = workbook.sheet_names()
        # 如果表名称：字典的key值不在工作簿的表名列表中
        if sheet_name not in sheets_list:
            # 复制先有工作簿对象
            work = copy(workbook)
            # 通过复制过来的工作簿对象，创建新表  -- 保留原有表结构
            sh = work.add_sheet(sheet_name)
            # 给新表设置表头
            excel_headers_tuple = (
                'uid', '昵称', '头像', '标题', '内容', '发布时间', '点赞数', '分享数', '评论数', '收藏数', '标签', '作品图片', '详情页面')
            for head_num in range(0, len(excel_headers_tuple)):
                sh.col(head_num).width = 2560 * 3
                #               行，列，  内容，            样式
                sh.write(0, head_num, excel_headers_tuple[head_num])
            work.save(os_excel_path)
        """========================================================================================="""
        # 判断工作簿是否存在
        if os.path.exists(os_excel_path):
            # 打开工作簿
            workbook = xlrd.open_workbook(os_excel_path)
            # 获取工作薄中所有表的个数
            sheets = workbook.sheet_names()
            for i in range(len(sheets)):
                for name in data.keys():
                    worksheet = workbook.sheet_by_name(sheets[i])
                    # 获取工作薄中所有表中的表名与数据名对比
                    if worksheet.name == name:
                        # 获取表中已存在的行数
                        rows_old = worksheet.nrows
                        # 将xlrd对象拷贝转化为xlwt对象
                        new_workbook = copy(workbook)
                        # 获取转化后的工作薄中的第i张表
                        new_worksheet = new_workbook.get_sheet(i)
                        for num in range(0, len(data[name])):
                            new_worksheet.write(rows_old, num, data[name][num])
                        new_workbook.save(os_excel_path)

    def run(self):
        thread = []
        t = threading.Thread(target=self.nextPage)
        t.start()
        for i in thread:
            i.join()


if __name__ == '__main__':
    x = XiaoHongShu()
    x.run()
