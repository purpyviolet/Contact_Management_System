
from person import Person

import datetime

class Colleague(Person):
    def __init__(self, name, birthday, phone, email, height, weight, company, income):
        # 初始化 Colleague 类的实例
        super().__init__(name, birthday, phone, email, height, weight, 'colleague', [company, income])
        self.company = company  # 公司名称
        self.income = income  # 收入
        self.projects = []  # 存储项目信息的空列表

    def __str__(self):
        # 返回 Colleague 类的字符串表示
        return f"Colleague(name='{self.name}', email='{self.email}', company='{self.company}')"

    def __repr__(self):
        # 返回 Colleague 类的正式表示
        return f"Colleague(name='{self.name}', email='{self.email}', company='{self.company}')"

    def add_project(self, project_name, project_description):
        # 添加项目到 Colleague 的项目列表
        project = {
            "name": project_name,  # 项目名称
            "description": project_description  # 项目描述
        }
        self.projects.append(project)

    def view_projects(self):
        # 查看 Colleague 的所有项目
        return self.projects

    def promote(self, new_position):
        # 提升员工职位
        print(f"{self.name} has been promoted to {new_position} at {self.company}.")

    def update_income(self, new_income):
        # 更新员工的收入信息
        self.income = new_income

# 示例用法
# colleague1 = Colleague("Alice", "1985-05-15", "555-123-4567", "alice@example.com", 165, 60, "ABC Corporation", 75000)
# colleague2 = Colleague("Bob", "1980-03-20", "555-987-6543", "bob@example.com", 175, 70, "XYZ Inc.", 90000)

# colleague1.add_project("Project A", "Developing a new software product")
# colleague1.add_project("Project B", "Enhancing user experience")

# colleague2.add_project("Project X", "Launching a marketing campaign")

# print(colleague1.view_projects())
# print(colleague2.promote("Senior Developer"))

# colleague2.update_income(95000)
# print(f"New income for {colleague2.name}: ${colleague2.income}")

