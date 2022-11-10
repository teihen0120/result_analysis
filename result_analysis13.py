import openpyxl as px

save_tyousa_path = r"D:\D_tokudome\D_desktop\test"

class wb(object):
    def __init__(self, name):
        self.name = name
    def create_workbook(self):
        wb = px.Workbook()
        wb["Sheet"].title = "Sheet1"
        return wb

        
workbook = wb("dsfa")
print(workbook.name)
workbook.save(save_tyousa_path + "//" + "test.xlsx")
class Person(object):
    def __init__(self):
        print("first")
    
    def say_something(self):
        print("hello")
        
person = Person()
person.say_something()
        