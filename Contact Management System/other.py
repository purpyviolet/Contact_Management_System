from person import Person

import datetime

class Other:
    def __init__(self, username, full_name, email, birthdate, bio=None):
        # 初始化 Other 类的实例
        self.username = username  # 用户名
        self.full_name = full_name  # 用户全名
        self.email = email  # 电子邮件地址
        self.birthdate = birthdate  # 出生日期
        self.bio = bio if bio else ""  # 用户个人简介，默认为空字符串
        self.friends = []  # 存储用户的朋友列表
        self.posts = []  # 存储用户的帖子列表

    def __str__(self):
        # 返回 Other 类的字符串表示
        return f"Other(username='{self.username}', email='{self.email}')"

    def __repr__(self):
        # 返回 Other 类的正式表示
        return f"Other(username='{self.username}', email='{self.email}')"

    def add_friend(self, friend):
        # 添加朋友到用户的朋友列表
        if friend not in self.friends:
            self.friends.append(friend)
            friend.friends.append(self)

    def remove_friend(self, friend):
        # 从用户的朋友列表中移除朋友
        if friend in self.friends:
            self.friends.remove(friend)
            friend.friends.remove(self)

    def create_post(self, content):
        # 创建用户的帖子
        post = {
            "author": self,
            "content": content,
            "timestamp": datetime.datetime.now()
        }
        self.posts.append(post)

    def view_posts(self):
        # 查看用户的所有帖子
        return self.posts

    def update_bio(self, new_bio):
        # 更新用户的个人简介
        self.bio = new_bio

    def get_age(self):
        # 计算用户的年龄
        birth_year = int(self.birthdate.split("-")[0])
        current_year = datetime.datetime.now().year
        age = current_year - birth_year
        return age

# 示例用法
# user1 = Other("user123", "John Doe", "john@example.com", "1990-05-15", "A tech enthusiast.")
# user2 = Other("jane123", "Jane Smith", "jane@example.com", "1988-03-20", "Nature lover.")

# user1.add_friend(user2)

# user1.create_post("Excited about the new Python release!")
# user2.create_post("Enjoying a hike in the mountains.")

# print(user1.view_posts())
# print(user2.view_posts())

# user1.update_bio("Coding is my passion.")

# print(user1.bio)
# print(user2.get_age())

