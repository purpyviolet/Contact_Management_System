
from person import Person

import datetime

class Student(Person):
    def __init__(self, name, birthday, phone, email, height, weight, college, grade, major, gpa):
        # 初始化 Student 类的实例
        super().__init__(name, birthday, phone, email, height, weight, 'student', [college, grade, major, gpa])
        self.college = college  # 学院名称
        self.grade = grade  # 年级
        self.major = major  # 专业
        self.gpa = gpa  # GPA（平均绩点）
        self.courses = []  # 存储学生所选课程的空列表

    def __str__(self):
        # 返回 Student 类的字符串表示
        return f"Student(name='{self.name}', email='{self.email}', college='{self.college}')"

    def __repr__(self):
        # 返回 Student 类的正式表示
        return f"Student(name='{self.name}', email='{self.email}', college='{self.college}')"

    def add_course(self, course_name, course_code, credits):
        # 添加课程到学生的课程列表
        course = {
            "name": course_name,  # 课程名称
            "code": course_code,  # 课程代码
            "credits": credits,  # 学分
        }
        self.courses.append(course)

    def view_courses(self):
        # 查看学生所选的所有课程
        return self.courses

    def calculate_total_credits(self):
        # 计算学生所选课程的总学分
        total_credits = sum(course["credits"] for course in self.courses)
        return total_credits

    def enroll_in_course(self, course):
        # 将学生注册到课程中
        if course not in self.courses:
            self.courses.append(course)


# 示例用法
# student1 = Student("John Doe", "2000-01-15", "555-123-4567", "john@example.com", 175, 70, "ABC University", "Sophomore", "Computer Science", 3.5)
# student2 = Student("Jane Smith", "1999-03-20", "555-987-6543", "jane@example.com", 170, 65, "XYZ College", "Junior", "Biology", 3.2)

# student1.add_course("Introduction to Programming", "CS101", 4)
# student1.add_course("Data Structures", "CS201", 3)
# student2.add_course("Chemistry 101", "CHEM101", 4)

# print(student1.view_courses())
# print(student2.calculate_total_credits())

# new_course = {"name": "Calculus", "code": "MATH101", "credits": 5}
# student2.enroll_in_course(new_course)

# print(student2.view_courses())
