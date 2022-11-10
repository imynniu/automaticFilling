# automaticFillin
### 使用场景
- 微信小程序-小管家-的每日填报任务
- 大连海事大学的学生可以根据直接修改

### 使用流程（非海事大学的同学）
1. 使用fiddler等抓包工具对微信小程序进行抓包，并保存整个request文件。
2. 对data.xsl的数据定义进行修改。
3. 对test.py 的modify函数进行与第二步中的信息进行比对修改。
4. 将保存的request中的所有true替换成Ture，false替换成False。
5. request文件保存到data文件中。
6. 注意

###使用流程（海事大学同学）
1. 使用fiddler等抓包工具抓取request请求并保存。
2. 将保存的request中的所有true替换成Ture，false替换成False。
3. 修改data.xsl文件。