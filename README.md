!!!注意return.txt是各个物品的配置文件，放在同目录可以供程序识别各个物品，要不然只会显示物品ID

第一步打开https://www.wegame.com.cn/home/并按F12打开开发者模式启用网络抓包
第二步，刷新网页
第三步登陆你的QQ或微信账号
登陆成功后，按F12打开抓包界面，随便点击一个图片，选择标头选项
查看里面有没有叫openid=XXXXXXXXXXXXXXXXXXXXXXX; access_token=XXXXXXXXXXXXXXXXXXXXXXXXXX;的文字段
这就是程序需要的openid和token，接下来运行程序即可
我做的比较仓促，界面也不好看，有能力的可以改进一下
