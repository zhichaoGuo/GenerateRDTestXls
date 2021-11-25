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
#### [2021.11.15]
  1. 分离root_url和header到cfg.yaml中
  2. 设置xlsx文件储存位置
#### [2021.11.16]
  1. 整理utils中的gen_fix_dict()中对info_dict.yaml的操作
#### [2021.11.17]
  1. 修改项目git user.name user.email
  2. 增加log部分内容
  3. 删除过程数据文件，改为存储log
  4. 添加new_url，为指定开始截至日期提供条件
#### [2021.11.25]
  1. 优化main函数，为指定开始截至日期提供条件
  2. 指定获取的开始和截止日期

# 待完成内容
- 优化getcookies的操作
- 将每日生成的log和excelfile添加到.gitignore中
- 不整理指定库的合入如	vendor/htek/frameworks/apps/VoIP 库
- 优化解耦excel操作
- 增加日期标头的分类（在更新指定获取的开始和截至日期后有用，或周一获取周五周六周日等多日期时有用）