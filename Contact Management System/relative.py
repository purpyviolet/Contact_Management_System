
from person import Person

import datetime

class Relative(Person):
    def __init__(self, name, birthday, phone, email, height, weight, relationship):
        # 初始化 Relative 类的实例
        super().__init__(name, birthday, phone, email, height, weight, 'relative', relationship)
        self.relationship = relationship  # 亲属关系
        self.family_events = []  # 存储家庭事件的空列表

    def __str__(self):
        # 返回 Relative 类的字符串表示
        return f"Relative(name='{self.name}', email='{self.email}')"

    def __repr__(self):
        # 返回 Relative 类的正式表示
        return f"Relative(name='{self.name}', email='{self.email}')"

    def add_family_event(self, event_name, event_date):
        # 添加家庭事件
        event = {
            "event_name": event_name,  # 事件名称
            "event_date": event_date  # 事件日期
        }
        self.family_events.append(event)

    def view_family_events(self):
        # 查看家庭事件
        return self.family_events


# 示例用法
# relative1 = Relative("Alice", "1980-05-15", "555-123-4567", "alice@example.com", 165, 60, "sister")
# relative2 = Relative("Bob", "1975-03-20", "555-987-6543", "bob@example.com", 180, 80, "brother")

# relative1.add_family_event("Family Reunion", "2023-07-10")
# relative2.add_family_event("Birthday Party", "2023-04-15")

# print(relative1.view_family_events())
# print(relative2.view_family_events())
