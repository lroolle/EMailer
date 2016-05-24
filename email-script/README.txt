# EMailer, version：Beta_0.6.1

# Update: 5/23 2016
    * 修改 'mail_template'；
    * 修改登录机制；
    * 新增 'config.py', 用于配置文件目录信息；
    * 其他一些代码优化。

# 几点注意的：
#	1. 把 pdf 存为 txt 之后极有可能乱码什么的，如果不嫌麻烦，最好是用复制粘贴；
#	2. 复制到 'mailist.txt' 后记得 Ctrl+s 保存，不要另存；
# 	3. 你是否发过是根据 mailist_saved.txt 里的邮箱来判断的，
#	   每次发过的（仅限用这个脚本发过的）都会自动加到 _saved.txt文件里；
#	4. 发邮件的速度可能不稳定（看服务器）；


# 以下用法：
#
# 嗖嗖嗖先，双击 'config.txt' 改成你自己的信息
#
# 然后右击 'Start.py' 选择 Edit with IDLE 打开
#
# 其他不用管
#
# 按 'F5' 或者点击 run --> run module 运行脚本
#
# 有两个用处：
#*  1. 把邮件地址提取出来
#*  2. 把邮件地址提取出来并发送邮件
#
# 如果只想提取地址，按完 'F5' 出来
# 地址保存在'maillist_set.txt' 里。
# 如果要开始发邮件 按 'Enter' ...
#
#* Have fun!!!!!!

