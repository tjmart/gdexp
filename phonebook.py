def print_menu():
    print()
    print("="*50)
    print("1. 조회, 2. 입력, 3. 수정, 4. 삭제, 5. 종료")
    print("="*50)

def phonebook_search():
    global phonebook
    print()
def phonebook_insert():
    pass

def phonebook_update():
    pass

def phonebook_delete():
    pass 



"""
[
['홍길동','010-1234-5678'],
['김철수','010-*5555-9966'],
['김대한','010-1111-2222']

]
"""
phonebook = []



print("파이썬 전화번호부 ver 1.0")
print_menu()

while True:
    m = input("기능을 선택하세요.(1~5 입력) ").strip()
    if m == "": continue

    if m == "1":
        pass
    if m == "2":
        pass
    if m == "3":
        pass
    if m == "4":
        pass
    if m == "5":
        break

