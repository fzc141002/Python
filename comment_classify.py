import pandas as pd
import re

# 读取Excel文件
ss = '08xx'         #输入评论对应的日期
file_path = "D:/商品评论" + ss + ".xlsx"        #找到文件所在的位置
comment_data = pd.read_excel(file_path, dtype={'订单ID': str, '商品ID': str})
comment_rule_b = pd.read_excel(r"D:\评价分类.xlsx", sheet_name="bad_comment")
comment_rule_g = pd.read_excel(r"D:\评价分类.xlsx", sheet_name="negative_commnet")
rules = [comment_rule_b, comment_rule_g]

#负评分类
keywords_b = {
    'bad_question': ['character'],
    '不好': ['nan!']
    # 继续添加更多类别和关键词
}
#正评分类
keywords_g = {
    'good_phenomenoon': ['character'],
    '好': ['nan!']
}
keywords = [keywords_b, keywords_g]
s = ['差评', '好评', '中评']
for i in range(len(s)):
    pattern = re.compile(s[i])  # 做好正则表达式
    filtered_data = comment_data[comment_data['评价等级'].apply(lambda x: bool(pattern.match(x)))].copy()

    # 假设我们要读取的列名为 '内容'
    content_column = '评价内容'
    date_column = '评价日期'
    # 将NaN值替换为空字符串
    filtered_data[content_column] = filtered_data[content_column].fillna('')
    if i == 2:          #对于中评进行123的分级评分
        filtered_data['first'] = 1
        filtered_data['second'] = 2
        filtered_data['third'] = 3
    if i < 2:           #对于负评和正评进行分类讨论
        # 初始化每个类别列为空字符串
        for category in keywords[i].keys():
            filtered_data[category] = ''

        # 根据关键词对内容进行分类
        def classify_content(content):
            categories = {category: '' for category in keywords[i].keys()}
            cout = 0
            for category, key_list in keywords[i].items():  # 这个方法可以遍历字典的键值对
                for keyword in key_list:
                    if keyword in content:
                        categories[category] = 'yes'  # 或者其他你想存储的信息
                        cout += 1
                        break
            if cout == 0 and i == 0:                #对于不属于特殊情况的评论进行一般化处理
                categories['不好'] = 'yes'
            elif cout == 0 and i == 1:
                categories['好'] = 'yes'
            return categories


        # 应用分类函数并填充对应的列
        classified_data = filtered_data[content_column].apply(classify_content)  # 这里的apply方法可以对每一行数据进行操作
        for category in keywords[i].keys():             #相当于把分类的结果返回去
            filtered_data.loc[:, category] = classified_data.apply(lambda x: x[category])

        # 确保商品编码和订单ID列保持为字符串格式
        filtered_data['订单ID'] = filtered_data['订单ID'].astype(str)
        filtered_data['商品ID'] = filtered_data['商品ID'].astype(str)

        # 转换日期格式
        filtered_data[date_column] = pd.to_datetime(filtered_data[date_column], format='%Y年%m月%d日 %H:%M:%S')
        filtered_data[date_column] = filtered_data[date_column].dt.strftime('%Y-%m-%d')

    # 已经获得了每一个是之后，怎么把一级二级放上去并且复制多个
    # 我觉得还是按照每一行进行复制
    # 把一个文件夹中所有的文件都合成到一块
    if i==2:        #如果是中评的话到这里就输出结束了
        output_file_path = 'D:/' + ss + s[i] + '.xlsx'
        filtered_data.to_excel(output_file_path, index=False)
    else:           #否则的话就需要进一步处理
        c_df = filtered_data
        #公司标准的评价分类情况

        def find_comment(sanji, judge):     #确定不同三级评论对应的二级评论情况
            id_c = rules[judge].loc[rules[judge]['third'] == sanji].index[0]
            yiji = rules[judge].loc[id_c, 'first']
            erji = rules[judge].loc[id_c, 'second']
            return [yiji, erji]

        name_c = c_df.columns[0:11]
        c_new = pd.DataFrame(columns=name_c)  # 新建一个空的DataFrame
        for i in range(len(c_df)):  # 复制每一行
            cout = 0
            for j in range(11, len(c_df.columns)):  # 相当于从第11开始需要计数
                if c_df.iloc[i, j] == 'yes':
                    cout += 1
                    name_comment = c_df.columns[j]  # 评论的数量
                    # 先找三级
                    r_c = find_comment(name_comment, 1)  # 注意这边不是写死的
                    yy = r_c[0]
                    ee = r_c[1]
                    # 再找一级二级
                    a = c_df.iloc[i, :11]
                    a['first'] = yy
                    a['second'] = ee
                    a['third'] = name_comment
                    d = pd.DataFrame(a).T
                    c_new = pd.concat([c_new, d], ignore_index=True)
        c_new.to_excel("D:/"+ss+s[i]+"结果.xlsx", index=False)
