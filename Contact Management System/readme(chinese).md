# 联系人管理系统

## 系统介绍

联系人管理系统是一个面向个人和小型企业用户的工具，旨在提供一个简单有效的方式来管理和维护联系人信息。该系统允许用户添加、搜索、修改、删除和筛选联系人，以及对联系人信息进行排序和统计分析。



## 安装说明

运行此程序需要以下库：

- `PyQt5`
- `openpyxl`
- `matplotlib`
- `smtplib`（通常作为Python标准库的一部分）
- `bypy`（可选，用于百度云同步）
- `playsound`

安装命令示例（使用pip）：

```bash
pip install PyQt5 openpyxl matplotlib playsound
```



## 使用指南

要运行程序，请执行以下步骤：

1. 打开命令行界面。
2. 导航到包含程序文件的目录。
3. 运行命令 `python ui.py` 来启动用户界面。
4. 或者直接运行ui.py以启动系统界面



## 功能概述

- **添加联系人**: 允许用户输入联系人的详细信息并将其保存到系统中。
- **搜索联系人**: 用户可以根据不同的信息来搜索联系人。
- **修改联系人**: 允许用户更新已存在联系人的信息。
- **删除联系人**: 用户可以删除一个或所有联系人。
- **排序联系人**: 根据不同的条件（如年龄、收入）对联系人进行排序。
- **筛选联系人**: 允许用户根据特定条件（如身高、体重）筛选联系人。

- **自动发送生日邮件**: 系统会自动检测即将过生日的联系人并发送定制的生日祝福邮件。
- **可视化分析**: 提供一个图表来展示不同类型联系人的比例分布。



## 程序入口

- 主程序入口文件为 `ui.py`。
- 启动程序后，用户界面将引导用户访问所有功能。



## 系统要求

- Python 3.x
- 适用于 Windows、macOS 和 Linux 操作系统。



## 额外信息

### 百度网盘的绑定

在开始之前，请确保Python和pip已经成功安装在您的电脑上。

1. 在命令行界面（cmd）中输入以下命令来安装`bypy`库：

```bash
pip install bypy
```

2. 接着，您需要授权`bypy`访问您的百度网盘。在命令行界面中输入以下命令：

```bash
bypy list
```

   3.执行此命令后，命令行界面会显示一个链接。请点击这个链接或将其复制到浏览器中打开。

   4.在浏览器中，您将被要求登录到百度网盘并授权应用程序。完成授权后，您会得到一个授权码。

   5.将这个授权码复制并粘贴回命令行界面，然后按回车键。这将完成`bypy`的绑定。

您可以在百度网盘上的“我的网盘>我的应用数据”中新建一个名为 `bypy` 的文件夹。程序将默认使用这个文件夹来存储和同步文件。



### 邮箱的绑定（设置SMTP服务及获取授权码）

#### QQ邮箱

##### 什么是授权码？

授权码是QQ邮箱用于登录第三方客户端的专用密码，适用于POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务。

#### 获取授权码步骤

1. 登录网页版QQ邮箱，进入 **设置** -> **帐户**。
2. 在“帐户”设置中找到 **服务设置项**。
3. 开启 **POP3/SMTP服务** 并验证密保。
4. 使用手机发送指定内容到指定号码，成功后获取授权码。
5. 在第三方客户端的密码框中输入16位授权码进行验证。

#### 注意事项

- 更改QQ密码或独立密码会导致授权码过期，需要重新获取。
- 授权码一般不会失效，除非上述情况发生。一旦获取，可保存备用。
- 登录时使用的是QQ邮箱地址，例如：`12345678@qq.com`。

### 网易邮箱

#### 获取授权码步骤

1. 打开网易邮箱网页版并登录。
2. 点击 **设置**，选择 **POP/SMTP/IMAP选项**。
3. 在客户端协议界面，选择开启相应的协议（如POP3/SMTP），并点击开启。
4. 在弹出窗口中选择发送短信方式。
5. 系统检测到短信发送成功后，会显示客户端授权码（16位字母组合）。

#### 注意事项

- 网易邮箱允许最多同时存在5个授权码。
- 授权码在开启时仅显示一次，但可以用于多个客户端。
- 推荐在需要时再次生成新的授权码。



**当设置自己的邮箱以及授权码时，在find_send_email.py中send_birthday_email的函数中修改密码以及授权码。**