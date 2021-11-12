# 简介
由于做研发测试每日需要从gerrit上筛选前一日合入的patch，并填写用于发送邮件的excel

每个链接需要单独点开，复制几段信息到表格中，工作较为繁琐耗时，所以写了这个脚本帮助我更快的完成表格
## 思路
1. 到网页上刷新并抓包，发现请求页面的数据的接口，以及返回的json数据

2. 到postman上验证接口，发现返回json为空

3. 应该是需要添加cookies，采用selenium仿真输入账户密码得到cookies

4. 使用request带cookies去访问接口，返回我们想要的数据。

5. 经过整理数据并裁剪数据得到需要的内容存为yaml

6. 使用openpyxl完成页面内容的填写，以及格式的设置
# 已完成内容
#### [2021.10.28]
  1. 抓包找到接口
  2. 完成selenium代码获取到cookies
  3. 完成openpyxl对页面内容的填写
  4. 主体代码跑通
#### [2021.10.29]
  1. username、password从代码中分离保存
  2. cookies从代码中分离保存
#### [2021.11.04]
  1. 完成数据整理裁剪存yaml部分
  2. 完成openpyxl对格式的设置
#### [2021.11.08]
  1. 完成log部分代码
#### [2021.11.12]
  1. 优化utils中的gen_fix_dict()中的删除操作

# 待完成内容
- 指定获取的开始和截止日期
- 整理utils中的gen_fix_dict()中对info_dict.yaml的操作
- 完成过程文件的分期储存，以备不时之需