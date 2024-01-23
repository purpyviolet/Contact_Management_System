
from person import Person

import datetime

class Friend(Person):
    def __init__(self, name, birthday, phone, email, height, weight, acquaintance_info=None):
        # 初始化 Friend 类的实例
        super().__init__(name, birthday, phone, email, height, weight, 'friend', acquaintance_info)
        self.acquaintance_info = acquaintance_info if acquaintance_info else {}  # 存储朋友的个人信息
        self.friend_requests_received = []  # 存储接收到的朋友请求
        self.friend_requests_sent = []  # 存储发送的朋友请求
        self.wall_posts = []  # 存储朋友的动态消息

    def __str__(self):
        # 返回 Friend 类的字符串表示
        return f"Friend(name='{self.name}', email='{self.email}')"

    def __repr__(self):
        # 返回 Friend 类的正式表示
        return f"Friend(name='{self.name}', email='{self.email}')"

    def add_acquaintance_info(self, key, value):
        # 添加朋友的个人信息
        self.acquaintance_info[key] = value

    def get_acquaintance_info(self, key):
        # 获取朋友的个人信息
        return self.acquaintance_info.get(key, None)

    def send_friend_request(self, friend):
        # 发送朋友请求
        if friend not in self.friend_requests_sent and friend != self:
            self.friend_requests_sent.append(friend)
            friend.friend_requests_received.append(self)

    def accept_friend_request(self, friend):
        # 接受朋友请求
        if friend in self.friend_requests_received:
            self.friend_requests_received.remove(friend)
            friend.friend_requests_sent.remove(self)
            self.add_friend(friend)

    def reject_friend_request(self, friend):
        # 拒绝朋友请求
        if friend in self.friend_requests_received:
            self.friend_requests_received.remove(friend)
            friend.friend_requests_sent.remove(self)

    def post_on_wall(self, author, content):
        # 在朋友的动态墙上发布消息
        if author in self.friends:
            post = {
                "author": author,
                "content": content,
                "timestamp": datetime.datetime.now()
            }
            self.wall_posts.append(post)

    def view_wall_posts(self):
        # 查看朋友的动态消息
        return self.wall_posts


# 示例用法
# friend1 = Friend("Alice", "1990-02-15", "555-123-4567", "alice@example.com", 170, 65)
# friend2 = Friend("Bob", "1988-07-10", "555-987-6543", "bob@example.com", 175, 70)

# friend1.add_acquaintance_info("Hobby", "Photography")
# friend2.add_acquaintance_info("Hobby", "Cooking")

# friend1.send_friend_request(friend2)
# friend2.accept_friend_request(friend1)

# friend1.post_on_wall(friend2, "Had a great time cooking together!")
# friend2.post_on_wall(friend1, "Thanks for the photography tips!")

# print(friend1.view_wall_posts())
# print(friend2.view_wall_posts())
