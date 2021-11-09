import pandas as pd
from pandas import DataFrame

# Global Path

print("Python 3.9.6")
print("후보추천서 자동화")
print()
print("\t\tMade by SW0000J")

all_student_path = input("전체 학생 xlsx 파일의 이름을 입력하세요.(ex: 소융대.xlsx): ")
recommand_student_path = input("비교할 후보추천서 xlsx 파일의 이름을 입력해주세요.(ex: 후보추천서.xlsx): ")

all_student_excel = pd.read_excel(all_student_path) # 소프트웨어융합대학.xlsx
recommand_student_excel = pd.read_excel(recommand_student_path) # 소융과_후보추천.xlsx

recommand_student_index_list = []
recommand_student_count = 0

def find_out_how_many():
    # have to initialize 0
    global recommand_student_count
    recommand_student_count = 0
    # have to make empty list
    recommand_student_index_list.clear()

    for i in range(len(recommand_student_excel.index)):
        buffer_list = all_student_excel.index[(all_student_excel["학번"] == recommand_student_excel["학번"][i])].tolist()
        
        if len(buffer_list) == 1:
            recommand_student_count += 1
            recommand_student_index_list.append(buffer_list[0])

    print("{}명의 사람이 추천하였습니다." .format(recommand_student_count))


def make_recommand_file(write_excel_file_name : str):
    new_recommand_dataframe = DataFrame(columns= ["이름", "전공", "학번", "전화번호", "서명"])

    for i in recommand_student_index_list:
        new_recommand_dataframe = new_recommand_dataframe.append(all_student_excel.loc[i, ["이름", "전공", "학번", "전화번호", "서명"]])
    
    # print(new_recommand_dataframe)
    excel_writer = pd.ExcelWriter(write_excel_file_name)
    new_recommand_dataframe.to_excel(excel_writer, sheet_name="Sheet1")
    excel_writer.close()


def check_write_file() -> str:
    print("\n")
    
    while True:
        str_input = input("추천서를 파일로 작성할 것인지 입력해주세요.[Y/N]: ")

        if str_input == "Y" or str_input == "y":
            write_excel_file_name = input("파일 이름은 무엇으로 할 것인지 입력해주세요.(ex: 후보추천서_결과.xlsx): ")
            make_recommand_file(write_excel_file_name)

            print("프로그램이 종료됩니다.")
            return "1"

        elif str_input == "N" or str_input == "n":
            
            print("프로그램이 종료됩니다.")
            return "1"

        else:
            print("다시 입력해주세요.")

if __name__ == "__main__":
    find_out_how_many()
    check_write_file()