# -*- coding: utf-8 -*-
'''
@ 目标： 计算单店效益，并后续便于每月更新
即计算出各个营业区的 收 - 支 - 利润
收： 主要来自新单保费带来的创费
支： 三个部分
    1. 人力成本
    2. 职场成本
    3. 促销费用

@ 涉及数据库及数据表
数据库： database_name = "险种清单2021.db"
数据表：
    促销费用： ('佣金发放',),     开始时间 2019.01
    外勤人力数据：'营销回单_个人')   开始时间 2019.01  其中2019年3月底基本法时间由21号开始切换到1号开始
    创费清单： 从险种清单计算出来的创费情况

@ Ray 2021 11 30 V2
'''

import sqlite3
import pandas as pd
import numpy as np
from time import sleep
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import psutil
import os

# 下载险种清单
class Downloader:
    '''
    使用selenium方法，从外挂系统下载险种清单
    '''

    def close_the_wps(self):
        print(">>> 清除后台 wps/et/chromedriver 等进程")
        pids = psutil.pids()
        for pid in pids:
            try:
                p = psutil.Process(pid)
                # print('pid=%s,pname=%s' % (pid, p.name()))
                if p.name() == 'et.exe':  # 关闭excel进程
                    cmd = 'taskkill /F /IM et.exe'
                    os.system(cmd)
                elif p.name() == 'wpsoffice.exe':  # 关闭wps 进程
                    cmd = 'taskkill /F /IM wpsoffice.exe'
                    os.system(cmd)
                elif p.name() == 'chromedriver.exe':  # 关闭已经在运行的 chromedriver 进程
                    cmd = 'taskkill /F /IM chromedriver.exe'
                    os.system(cmd)
            except Exception as e:
                print(e)

    def init_browser(self, download_path):

        print("  > 启动启动浏览器，开始下载险种清单")

        global browser
        chrome_path = r'D:\Users\lilei10\PycharmProjects\pythonProject\venv\Scripts\chromedriver.exe'
        chrome_options = Options()
        prefs = {"download.default_directory": download_path}
        chrome_options.add_experimental_option("prefs", prefs)
        #chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        desired_capabilities = DesiredCapabilities.CHROME  # get直接返回，不再等待界面加载完成
        desired_capabilities["pageLoadStrategy"] = "none"
        browser = webdriver.Chrome(options=chrome_options, executable_path=chrome_path)

        return browser

    def log_in(self, browser, url, username, password):
        '''
        登录外挂系统
        :param browser:
        :param url:
        :param username:
        :param password:
        :return:
        '''

        browser.get(url)

        print('  > 输入用户名、密码，并登录')
        browser.find_element_by_id('LoginName').clear()
        browser.find_element_by_id('LoginName').send_keys(username)
        browser.find_element_by_id('LoginPass').clear()
        browser.find_element_by_id('LoginPass').send_keys(password)
        # 定位登录按钮，点击
        browser.find_element_by_xpath('//*[@id="Button1"]').click()

        sleep(2)
        browser.implicitly_wait(20)

        return browser

    def locate_insurlist(self, browser, start_date, end_date):
        '''
        定位险种清单的地址，并下载
        :param browser:
        :param start_date:
        :param end_date:
        :return:
        '''

        print('  > 定位表格的xpath位置')

        # content = browser.page_source                 # 检查页面有没有获取到。
        # soup = BeautifulSoup(content, 'html.parser')
        # print(soup)

        iframe = browser.find_element_by_xpath('//*[@id="frmTitle"]/iframe')  # 切换到 iframe
        browser.switch_to.frame(iframe)
        browser.find_element_by_xpath('/html/body/table[1]/tbody/tr/td[2]').click()  # 个人系列
        browser.implicitly_wait(20)
        browser.find_element_by_xpath('//*[@id="LeftMenu__ctl0_LeftMenu_Sub__ctl0_Hyperlink1"]').click()  # 承保业绩统计
        browser.implicitly_wait(20)
        browser.switch_to.default_content()

        iframe = browser.find_element_by_xpath('//*[@id="mainFrame"]')  # 切换到 main iframe
        browser.switch_to.frame(iframe)
        browser.find_element_by_xpath('//*[@id="ddllx"]').click()  # 选择类型
        browser.implicitly_wait(20)
        browser.find_element_by_xpath('//*[@id="ddllx"]/option[7]').click()  # 险种清单

        # print('Sending dates...')
        browser.find_element_by_id('tcbdate_q').clear()
        browser.find_element_by_id('tcbdate_q').send_keys(start_date)  # 承保 起始时间

        browser.find_element_by_id('tcbdate_z').clear()
        browser.find_element_by_id('tcbdate_z').send_keys(end_date)  # 承保 结束时间

        print('  > Clicked and Downloading...Please waiting for a few minutes patiently!')
        browser.implicitly_wait(20)

        state = True
        time_start = time.time()
        try:
            while state == True:
                print('计时：', round(time.time() - time_start, 2), '秒', end="\r")
                browser.find_element_by_xpath('//*[@id="bdownload"]').click()  # 下载到指定文件夹。
                state = False
        except KeyboardInterrupt:
            print('end of time.')
        sleep(2)

        browser.implicitly_wait(20)
        browser.switch_to.default_content()
        browser.refresh()

        print('  > Downloaded yet!')

    def excecute_download(self, start_date, periods):
        '''
        根据时间批量下载险种清单
        :param start_date: 开始月份，str格式， '2021-01-01'
        :param periods: 处理几个月， int 格式, 1
        :return: None， 把险种清单的xls文件下载到了本地
        '''
        # 外挂基础设置，注意需要先登录一次，把验证码输入一次，后续就可以直接使用了。
        url = 'http://172.19.3.42/Login.aspx'
        username = 'lilei10'
        password = 'Jiangjin9_'

        # 险种清单下载地址
        dir_insurlist = r'D:\2019 市场企划\0 企划数据库\【险种清单】\原始险种清单'

        # 生成月初月末的时间列表
        data_range_begin = pd.date_range(start=start_date, periods=periods, freq='MS')  # 月初
        data_range_end = pd.date_range(start=start_date, periods=periods, freq='M')  # 月末

        # 清空后台进程，尤其是wps, et, chromedriver等
        self.close_the_wps()

        # 初始化浏览器
        #driver = Downloader.log_in(self, Downloader.init_browser(self, dir_insurlist), url, username, password)
        driver = self.log_in(self.init_browser(dir_insurlist), url, username, password)

        for begin, end in zip(data_range_begin.date, data_range_end.date):
            print("  > 正在下载 ：{} to {}".format(begin, end))
            # 按月下载险种清单
            self.locate_insurlist(driver, str(begin), str(end))

        # 等待5分钟后再退出，因为险种清单下载后生成需要一定时间
        sleep(300)
        browser.implicitly_wait(20)
        browser.quit()

        print('  > 下载完毕')


# 数据库相关的操作
class ManageDatabase:
    '''
    实现数据库的导入、查询操作
    '''
    def __init__(self, db_name, db_table_name):
        self.db_name = db_name
        self.db_table_name = db_table_name

    # 创建数据库、数据表，数据来源是DataFrame
    def create_table(self, df):  # , data_type

        print('  > 创建、上传至数据库和数据表')

        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 创建表
        df.to_sql(self.db_table_name, con=connect, if_exists= 'append', )  # if_exists {'fail'，'replace'，'append'}
        # 提交事务
        connect.commit()
        # 断开连接
        connect.close()

    def destroy_table(self):

        print('  > 删除数据库中的对应数据表')

        # 连接数据库
        connect = sqlite3.connect(self.db_name, isolation_level=None)
        # 删除表
        print(">>> DROP TABLE '{}'".format(self.db_table_name))
        connect.execute("DELETE FROM '{}'".format(self.db_table_name))
        connect.execute("DROP TABLE '{}'".format(self.db_table_name))
        connect.execute("VACUUM")
        # 提交事务
        connect.commit()
        # 断开连接
        connect.close()

    def delete_data(self, conditions):

        print('  > 删除{}表重点{}数据'.format(self.db_table_name, conditions))

        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 插入多条数据
        connect.execute("DELETE FROM {} WHERE {}".format(self.db_table_name, conditions))
        # 提交事务
        connect.commit()
        # 断开连接
        connect.close()


    def search_data(self, conditions):
        '''
        function: 查找特定的数据
        '''
        print("  > 按条件查找所有列: {}".format(conditions))
        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 创建游标
        cursor = connect.cursor()
        # 查找数据
        cursor.execute("SELECT * FROM {} WHERE {}".format(self.db_table_name, conditions))
        data = cursor.fetchall()
        # 关闭游标
        cursor.close()
        # 断开数据库连接
        connect.close()
        return data

    def search_by_columns(self, conditions, query_columns="all"):
        '''
            function: 使用pd.read_sql 直接读取数据按照关键词、给定的列名进行查询
        '''
        print("  > 按条件查找部分列: {}".format(conditions))
        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 创建游标
        cursor = connect.cursor()
        # 创建查询字符串
        if query_columns == 'all':
            # 查找数据
            sql = ("SELECT * FROM {} WHERE {}".format(self.db_table_name, conditions))
        else:
            query_columns_join = ','.join(str(i) for i in query_columns)
            sql = ("SELECT {} FROM {} WHERE {}".format(query_columns_join, self.db_table_name, conditions))
        print('  > 查询语句：{}'.format(sql))

        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 直接用pd.read_sql进行查询
        df = pd.read_sql(sql, con=connect, index_col=None, coerce_float=False)
        # 断开数据库连接
        connect.close()
        return df

    def read_table_names(self):

        print("  > 读取数据库中的所有数据表名称")

        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 创建游标
        cursor = connect.cursor()
        # 插入多条数据
        cursor.execute('select name from sqlite_master where type="table"')
        tables = cursor.fetchall()
        # 关闭游标
        cursor.close()
        # 断开数据库连接
        connect.close()
        # 输出表明的列表
        print("  > 数据库中的列表名称： ")
        print(tables)
        return tables

    def search_cross_tables(self, table1, table2, key1, key2, query_column1="all", query_column2="all", ):
        '''
            function: 使用pd.read_sql 直接读取数据按照关键词、给定的列名进行查询
            跨表查询， 得到 table.column, 满足 table1.key =table2.key
            talb1e: 要得到结果的主表
            table : 需要满足条件的辅助表
            key: 两个表都有关键列名
        '''
        print("  > 跨表查询")

        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 创建游标
        cursor = connect.cursor()
        # 创建查询字符串
        if query_column1 == 'all' and query_column2 == 'all':
            # 查找数据
            sql = ('select {0} from {1}, {2} where {3}.{4} = {5}.{6}'.format(
                '*', table1, table2, table1, key1, table2, key2))
        else:
            string1 = ['{0}.'.format(table1) + i for i in query_column1]
            string2 = ['{0}.'.format(table2) + j for j in query_column2]
            strings = string1 + string2
            query_columns_join = ','.join(str(i) for i in strings)  # 用 , 相连
            sql = ('select {0} from {1}, {2} where {3}.{4} = {5}.{6}'.format(
                query_columns_join, table1, table2, table1, key1, table2, key2))

        print('  > 正在查询：{}, {}, {}, on {} & {}'.format(self.db_name, table1, table2, key1, key2))
        print('  > SQL语句：{}'.format(sql))

        # 连接数据库
        connect = sqlite3.connect(self.db_name)
        # 直接用pd.read_sql进行查询
        df = pd.read_sql(sql, con=connect, index_col=None, coerce_float=False)
        # 断开数据库连接
        connect.close()

        return df

    def upload_to_db(self, start_date, periods):
        '''
        导入险种清单，进入数据库
        start_date:从哪个月开始导入，示例 '2021.01.01'
        periods: 导入几个月的数据, 一个月的就是 1
        input_columns: 输入的列是哪些，列表格式
        columns_types: 输入列的类型是什么，列表格式
        add_to_mode: 追加还是覆盖,'fail'，'replace'，'append'
        return: None
        '''



        dir_path = r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单'

        # 构造时间序列
        data_range_begin = pd.date_range(start=start_date, periods=periods, freq='MS')  # 月初
        data_range_end = pd.date_range(start=start_date, periods=periods, freq='M')  # 月末

        data_types = {'机构代码': str, '区代码': str, '部代码': str, '组代码': str, '营销员代码': str,
                     '入司时间': str, '推荐人代码': str, '预收时间': str, '承保时间': str, '客户签收时间': str,
                     '回单时间': str, '保单号': str, '险种': str, '投保书号': str, '被保人号': str,
                     '统计年月': str, }

        for begin, end in zip(data_range_begin.date, data_range_end.date):
            name = "承保_创费_险种清单" + str(begin) + '--' + str(end) + ".xlsx"
            full_name = os.path.join(dir_path, name)
            print(">>> 上传 {} 到数据库中: ".format(name))
            df = pd.read_excel(full_name, sheet_name='Sheet1', index_col=None, dtype=data_types)
            print('  > Shape of df', df.shape)

            self.create_table(df)

        print('  > 导入成功！')


# 运用 pandas 进行分区匹配及各类指标的计算
class DataMining:
    '''
    各种数据的匹配与计算
    revenue：创费，符合权责发生制条件下的营业收入， 通常也指财政收入。
    expense：促销费用。主要是“花费”、“开支”之意，如current expenses“日常开支”，selling expenses“销售费用”，travelling expenses“旅费”等
    cost：   人力、职场成本。其本义为“成本”、“原价”，常常用来表示对已取得的货物或劳务所支付的费用。
    manpower_KPI: 外勤人力指标，绩优、大绩优、有效人数等。
    '''


    def generate_revenue_list(self, start_date, periods):
        '''
        根据每月的险种清单，匹配创费率表，生成创费xlsx表格
        :param start_date: 开始月份，str格式， '2021-01-01'
        :param periods: 处理几个月， int 格式, 1
        :return: df
        '''
        print('>>> 计算创费、匹配分区，按月生成 创费_承保_险种清单')

        # 读取 险种_年期_汇总表.xlsx 表
        dir_income = r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单\产品费用率'
        file_income = '险种_年期_汇总表.xlsx'
        data_type_income = {'险种': str, '辅助列': str, }
        df_income = pd.read_excel(os.path.join(dir_income, file_income), sheet_name='外挂口径',
                                  dtype=data_type_income, usecols='A:F', skiprows=0, index_col=None, )
        # 读取分区表
        df_district = pd.read_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单\产品费用率', '营销员与营业区.xlsx'),
                                    sheet_name='Sheet1',
                                    dtype={'营销员代码': str},
                                    index_col=None,
                                    usecols='C:D')


        # 读取 每月的险种清单表，并匹配创费率，险种名称，计算首创值
        dir_insurlist = r'D:\2019 市场企划\0 企划数据库\【险种清单】\原始险种清单'
        data_type_insurlist = {'机构代码': str, '区代码': str, '部代码': str, '组代码': str, '营销员代码': str,
                               '入司时间': str, '推荐人代码': str, '预收时间': str, '承保时间': str, '客户签收时间': str,
                               '回单时间': str, '保单号': str, '险种': str, '投保书号': str, '被保人号': str, }
        # 生成月初月末的时间列表
        data_range_begin = pd.date_range(start=start_date, periods=periods, freq='MS')  # 月初
        data_range_end = pd.date_range(start=start_date, periods=periods, freq='M')  # 月末

        for begin, end in zip(data_range_begin.date, data_range_end.date):
            file_insurlist = os.path.join(dir_insurlist, '承保_全省险种清单{}--{}.xls'.format(str(begin), str(end)))
            print("  > Working on ：{}".format(file_insurlist))
            df_insurlist = pd.read_excel(os.path.join(dir_insurlist, file_insurlist), sheet_name='Sheet0',
                                         dtype=data_type_insurlist, usecols='A:AM', skiprows=0, index_col=None, )
            # 构造匹配的辅助列
            df_insurlist.loc[:, '辅助列'] = df_insurlist.loc[:, '险种'] + df_insurlist.loc[:, '交费年期'].astype('str')

            # 匹配创费率、产品大类
            df = pd.merge(df_insurlist, df_income, how='inner', on='辅助列',) #inner outer
            # 首创费用 = 规模保费 * 创费率
            df['首创'] = df.loc[:, '规保'] * df.loc[:, '首创率']
            # 时间标记
            df['统计年月'] = begin.strftime('%Y-%m').split('-')[0] + begin.strftime('%Y-%m').split('-')[1]

            # 匹配分区
            df1 = pd.merge(df, df_district, how='inner', on='营销员代码')

            try:
                # 去掉所有的空格
                df1.replace('\s+', '', regex=True, inplace=True)
                # 去掉重复项目
                # df = df.drop_duplicates(ignore_index=True)
            except Exception as e:
                print(e)
            # 删除重复的列
            df1 = df1.drop(['险种_y', '交费年期_y', ], 1)
            df1 = df1.rename(columns = {'险种_x': '险种', '交费年期_x': '交费年期', '营业区_x': '营业区','营业区_y': '分区'})
            # 排序
            df1 = df1.sort_values(by = '承保时间', ascending= True)
            dir_income1 = r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单'
            df1.to_excel(os.path.join(dir_income1, '承保_创费_险种清单{}--{}.xlsx'.format(str(begin), str(end))), index= None)

        print('  > 创费计算完毕')

    def calculate_district(self):
        print('>>> 计算营销员的分区数据')

        my_db = ManageDatabase('险种清单2021.db', '营销承保_个人')

        conditions = '统计年月>="202101"'
        query_columns = [ '姓名', '营销员代码',  '营业区', '机构']
        df =  my_db.search_by_columns(conditions, query_columns)


        df = df.drop_duplicates(ignore_index=True)
        #df= df.dropna(subset=['营业区'])
        #print(df)
        df.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单\产品费用率', '营销员与营业区.xlsx'))


    def calculate_revenue(self):

        print(">>> 生成创费清单汇总数据表")

        # 初始化 数据库类
        my_db = ManageDatabase('险种清单2021.db', '承保_创费_险种清单')


        conditions = '统计年月>="202101"'
        query_columns = ['机构名称', '分区','营销员代码', '规保', '价值', 'FYC', '期交保费', '首创', '统计年月', ]
        df = my_db.search_by_columns(conditions, query_columns)
        #print(df)

        # 按照营业区 数据分类汇总，以便呈现
        df_pivot = pd.pivot_table(df,
                                  index=['机构名称', '分区', ],
                                  columns=['统计年月'],
                                  values=['规保', '价值', 'FYC', '期交保费', '首创', ],
                                  aggfunc=np.sum,
                                  margins=1,
                                  )
        # 展平透视表，并去重，导出
        df_pivot = df_pivot.reset_index()
        df_pivot.drop_duplicates(ignore_index=True)
        df_pivot.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\单店效益',
                                       '【分区】创费_2021年1月起.xlsx'))

        # 按照中支 数据分类汇总，以便呈现
        df_pivot = pd.pivot_table(df,
                                  index=['机构名称', ],
                                  columns=['统计年月'],
                                  values=['规保', '价值', 'FYC', '期交保费', '首创', ],
                                  aggfunc=np.sum,
                                  margins=1,
                                  )
        # 展平透视表，并去重，导出
        df_pivot = df_pivot.reset_index()
        df_pivot.drop_duplicates(ignore_index=True)
        df_pivot.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\单店效益',
                                       '【中支】创费_2021年1月起.xlsx'))
        print('>>> 计算完成！')

    def calculate_promotional_expense(self):

        print(">>> 计算促销费用")

        # 初始化 数据库类
        my_db = ManageDatabase('险种清单2021.db',  '佣金发放')

        # 读取分区表
        df_district = pd.read_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单\产品费用率', '营销员与营业区.xlsx'),
                                    sheet_name='Sheet1',
                                    dtype={'营销员代码': str},
                                    index_col=None,
                                    usecols='C:E')

        conditions = '佣金所属年月>="202101"'
        query_columns = ['佣金所属年月', '姓名', '营销员编码', '初年度佣金', '业务推动费', '业务推动费_已发', '组织发展费',
            '组织发展费_已发', '新人责任津贴调节项', '应纳税收入额','实发' ]
        df_query = my_db.search_by_columns(conditions, query_columns)

        df = pd.merge(df_query, df_district, how='inner', left_on='营销员编码', right_on='营销员代码')

        # 按照 营业区分区 汇总，以便呈现
        df_pivot = pd.pivot_table(df,
                                  index=['机构', '营业区', ],
                                  columns=['佣金所属年月'],
                                  values=['初年度佣金', '业务推动费', '业务推动费_已发', '组织发展费', '组织发展费_已发',
                                          '新人责任津贴调节项', '应纳税收入额', '实发'],
                                  aggfunc=np.sum,
                                  margins=1,
                                  )
        # 展平透视表，并去重，导出
        df_pivot = df_pivot.reset_index()
        df_pivot.drop_duplicates(ignore_index=True)
        df_pivot.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\单店效益',
                                       '【分区】促销费用_2021年1月起.xlsx'))

        # 按照 机构 汇总，以便呈现
        df_pivot = pd.pivot_table(df,
                                  index=['机构', ],
                                  columns=['佣金所属年月'],
                                  values=['初年度佣金', '业务推动费', '业务推动费_已发', '组织发展费', '组织发展费_已发',
                                          '新人责任津贴调节项', '应纳税收入额', '实发'],
                                  aggfunc=np.sum,
                                  margins=1,
                                  )
        # 展平透视表，并去重，导出
        df_pivot = df_pivot.reset_index()
        df_pivot.drop_duplicates(ignore_index=True)
        df_pivot.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\单店效益',
                                       '【中支】促销费用_2021年1月起.xlsx'))
        print('>>> 计算完成！')

    def calculate_manpower_KPI(self):

        print(">>> 计算有效、绩优、大绩优人力等指标")

        # 初始化 数据库类
        my_db = ManageDatabase('险种清单2021.db', '营销回单_个人')

        conditions = '统计年月>="202001"'
        query_columns = ['机构', '营业区', '新单', '健康险', '价保', 'FYC',
                         '有效', '绩优', '大绩优', '统计年月']
        df = my_db.search_by_columns(conditions, query_columns)

        # 数据分类汇总，以便呈现
        df_pivot = pd.pivot_table(df,
                                  index=['机构', '营业区', ],
                                  columns=['统计年月'],
                                  values=['新单', '健康险', '价保', 'FYC', '有效', '绩优', '大绩优',],
                                  aggfunc=np.sum,
                                  margins=1,
                                  )
        # 展平透视表，并去重，导出
        df_pivot = df_pivot.reset_index()
        df_pivot.drop_duplicates(ignore_index=True)
        df_pivot.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\单店效益',
                                       '单店保费_人力_2020年1月起.xlsx'))
        print('>>> 计算完成！')


    def simple_query(self, conditions):
        '进行灵活的查询'
        print(">>> 生成创费清单汇总数据表")

        # 初始化 数据库类
        my_db = ManageDatabase('险种清单2021.db', '承保_创费_险种清单')

        # 查询相关
        query_columns = ['机构名称', '营业区', '分区', '营销员代码', '入司时间',
                         '承保时间','险种','交费年期','规保', '价值', 'FYC','件数',
                         '期交保费','险种名称','险种大类', '首创', '统计年月', ]
        df = my_db.search_by_columns(conditions, query_columns)
        #print(df)
        df.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】', '清单汇总.xlsx'))

        '''
        # 按照营业区 数据分类汇总，以便呈现
        df_pivot = pd.pivot_table(df,
                                  index=['机构名称', '分区', ],
                                  columns=['统计年月'],
                                  values=['规保', '价值', 'FYC', '期交保费', '首创', ],
                                  aggfunc=np.sum,
                                  margins=1,
                                  )
        # 展平透视表，并去重，导出
        df_pivot = df_pivot.reset_index()
        df_pivot.drop_duplicates(ignore_index=True)
        df_pivot.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】\单店效益',
                                       '【分区】创费_2021年1月起.xlsx'))
        '''

    def simple_query_raw(self, conditions):
        '进行灵活的查询'
        print(">>> 生成创费清单汇总数据表")

        # 初始化 数据库类
        my_db = ManageDatabase('承保清单2011.db', '险种清单')

        # 查询相关
        query_columns = ['险种', '交费年期',]
        df = my_db.search_by_columns(conditions, query_columns)
        df = df.drop_duplicates(ignore_index=True)

        df.to_excel(os.path.join(r'D:\2019 市场企划\0 企划数据库\【险种清单】', '险种年期清单.xlsx'))

if __name__ == "__main__":

    # 每次使用前，需要更新创费表, 数据来自 财务部
    my_path = r'D:\2019 市场企划\0 企划数据库\【险种清单】\创费清单\产品费用率\险种_年期_汇总表.xlsx'
    # 还要更新 营销承保_个人， 确保分区数据是完整的
    #DataMining().calculate_district()  # 先把分数数据计算好


    # 定义变量： 开始时间，计算月数
    start_year_month = '2021-01-01'
    month_numbers = 12+5

    # Step1: 初始化 下载（DownloadInsuranceLists）类，下载险种清单
    #my_nci = Downloader()
    #my_nci.excecute_download(start_date=start_year_month, periods=month_numbers)
    # 这里需要花几分钟的时间，去文件夹查看下是否下载完毕

    # Step2: 初始化 数据（DataMining）类，生成创费表, 2019.01起
    DataMining().generate_revenue_list(start_year_month, month_numbers)

    # Step3: 初始化 数据库（ManageDatabase）类，将创费表上传到数据库中的 承保_创费_险种清单表
    my_db = ManageDatabase('险种清单2021.db', '承保_创费_险种清单')
    my_db.upload_to_db(start_date= start_year_month, periods= month_numbers)

    # Step4: 计算分区<创费>情况： 由于分区的缘故，只计算2021.01起
    DataMining().calculate_revenue()

    # Step5: 计算分区<促销费用>情况：用到数据库 险种清单2021.db，佣金发放，由于分区的缘故，只计算2021.01起
    # DataMining().calculate_promotional_expense()

    # Step6: 计算分区<人力变化>情况：用到数据库 险种清单2021.db，营销回单_个人
    #DataMining().calculate_manpower_KPI()

    # Simple_query
    #conditions = '承保时间>="2021-01-01"'
    #DataMining().simple_query_raw(conditions)

    conditions = '统计年月>="202101"'
    DataMining().simple_query(conditions)

    # 注： 人力成本、职场固定成本等分别由人事、财务提供。