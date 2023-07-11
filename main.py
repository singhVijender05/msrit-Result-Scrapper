import requests
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
from tabulate import tabulate
import json
from openpyxl import Workbook

studentsList = []
studentDictionary = dict()
gradeToMarks = {
    "O": 10,
    "A+": 9,
    "A": 8,
    "B+": 7,
    "B": 6,
    "C": 5,
    "P": 4,
    "F": 0,
    "NE": 0,
    "X": 0,
    "I": 0,
    "TAL": 0,
}


def payset(usn):
    payload = {
        "usn": usn,
        "osolCatchaTxt": "",
        "osolCatchaTxtInst": 0,
        "option": "com_examresult",
        "task": "getResult",
    }
    return payload


def constructData(response, usn):
    try:
        soup = BeautifulSoup(response.content, "html.parser")
        subjectCodeToGrade = dict()

        studentDictionary["usn"] = usn
        name = soup.find("div", class_="stu-data1").contents[1].text
        studentDictionary["name"] = name

        cgpa = soup.find("div", class_="credits-sec4").contents[-2].text
        studentDictionary["cgpa"] = cgpa

        sgpa = soup.find("div", class_="credits-sec3").contents[-2].text
        studentDictionary["sgpa"] = sgpa

        table = soup.find("table", class_="res-table")
        tbody = table.contents[-2]
        trs = tbody.find_all("tr")

        for tr in trs:
            subjectCode = tr.contents[1].text
            grade = tr.contents[-2].text
            subjectCodeToGrade[subjectCode] = grade

        studentDictionary["subjectCodeToGrade"] = subjectCodeToGrade
        studentsList.append(studentDictionary.copy())
        return studentDictionary
        # showGraph(studentDictionary)
    except:
        print("Error in constructing data for USN: " + usn)


def getResult(usn):
    try:
        url = "https://exam.msrit.edu/index.php"
        response = requests.post(url, data=payset(usn))
        studentDictionary = constructData(response, usn)
        return studentDictionary
    except:
        print("Error in getting result for USN: " + usn)


def getUsnList():
    startUSN = int(input("Enter the starting USN number: "))
    endUSN = int(input("Enter the ending USN number: "))
    branchCode = input("Enter the branch code: ")
    year = int(input("Enter the year: "))

    for i in range(startUSN, endUSN + 1):
        usn = "1MS" + str(year) + branchCode + str(i).zfill(3)
        getResult(usn)


def showGraph(studentDictionary):
    grades = ["F", "P", "C", "B", "B+", "A", "A+", "O"]
    subjectCode = list(studentDictionary["subjectCodeToGrade"].keys())
    studentGrades = list(studentDictionary["subjectCodeToGrade"].values())
    x = range(len(subjectCode))
    fig, ax = plt.subplots()

    ax.bar(
        x,
        [grades.index(grade) for grade in studentGrades],
        label=f"{studentDictionary['usn']}\n{studentDictionary['name']}\ncgpa:{studentDictionary['cgpa']}\nsgpa:{studentDictionary['sgpa']}",
        color="purple",
    )

    plt.subplots_adjust(top=0.9)

    ax.set_xticks(x)
    ax.set_xticklabels(subjectCode)

    ax.set_ylim([0, len(grades) - 1])

    ax.set_yticks(range(len(grades)))
    ax.set_yticklabels(grades)

    ax.legend()

    plt.xlabel("Subjects")
    plt.ylabel("Grades")
    plt.title(f"Grades of {studentDictionary['name']}")
    plt.show()


def createTable(studentsList):
    table_data = []
    for student in studentsList:
        usn = student.get("usn", "-")
        name = student.get("name", "-")
        sgpa = student.get("sgpa", "-")
        cgpa = student.get("cgpa", "-")
        table_data.append([usn, name, sgpa, cgpa])

    table = tabulate(
        table_data, headers=["USN", "Name", "SGPA", "CGPA"], tablefmt="pretty"
    )
    print(table)


def createExcel(studentsList, filename):
    workbook = Workbook()
    sheet = workbook.active

    # Write header row
    sheet.append(["USN", "Name", "SGPA", "CGPA"])

    # Write data rows
    for student in studentsList:
        usn = student.get("usn", "-")
        name = student.get("name", "-")
        sgpa = student.get("sgpa", "-")
        cgpa = student.get("cgpa", "-")
        sheet.append([usn, name, sgpa, cgpa])

    workbook.save(filename)


def calculateSubjectWiseAverage(studentsList):
    subjectCodeToGrade = dict()
    subjectCodeToCount = dict()

    valid_cgpa_values = []
    valid_sgpa_values = []

    for student in studentsList:
        for subjectCode, grade in student["subjectCodeToGrade"].items():
            if "AEC" in subjectCode:
                subjectCode = "AEC"
            elif "HS" in subjectCode:
                subjectCode = "HS"

            if subjectCode not in subjectCodeToGrade:
                subjectCodeToGrade[subjectCode] = [grade]
                subjectCodeToCount[subjectCode] = 1
            else:
                subjectCodeToGrade[subjectCode].append(grade)
                subjectCodeToCount[subjectCode] += 1

        try:
            cgpa = float(student["cgpa"])
            valid_cgpa_values.append(cgpa)
        except ValueError:
            pass

        try:
            sgpa = float(student["sgpa"])
            valid_sgpa_values.append(sgpa)
        except ValueError:
            pass

    subjectCodeToAverage = dict()
    for subjectCode, grades in subjectCodeToGrade.items():
        totalStudents = subjectCodeToCount[subjectCode]
        subjectCodeToAverage[subjectCode] = (
            sum([float(gradeToMarks[grade]) for grade in grades]) / totalStudents
        )

    totalStudents = len(studentsList)
    cgpa = sum(valid_cgpa_values) / totalStudents
    sgpa = sum(valid_sgpa_values) / totalStudents

    for subjectCode, grade in subjectCodeToAverage.items():
        if grade >= 10:
            subjectCodeToAverage[subjectCode] = "O"
        elif grade >= 9:
            subjectCodeToAverage[subjectCode] = "A+"
        elif grade >= 8:
            subjectCodeToAverage[subjectCode] = "A"
        elif grade >= 7:
            subjectCodeToAverage[subjectCode] = "B+"
        elif grade >= 6:
            subjectCodeToAverage[subjectCode] = "B"
        elif grade >= 5:
            subjectCodeToAverage[subjectCode] = "C"
        elif grade >= 4:
            subjectCodeToAverage[subjectCode] = "P"
        else:
            subjectCodeToAverage[subjectCode] = "F"

    averageData = {
        "name": "Average",
        "usn": "Average",
        "cgpa": round(cgpa, 2),
        "sgpa": round(sgpa, 2),
        "subjectCodeToGrade": subjectCodeToAverage,
    }

    showGraph(averageData)


while True:
    choice = int(
        input(
            "Enter\n1. For entire class data\n2. For individual student data\n3. Exit\n"
        )
    )
    match choice:
        case 1:
            getUsnList()
            with open("data.json", "w") as file:
                json.dump(studentsList, file)
            createTable(studentsList)
            createExcel(studentsList, "student_data.xlsx")
            calculateSubjectWiseAverage(studentsList)
            studentsList.clear()
            studentDictionary.clear()
        case 2:
            usn = input("Enter the USN number: ")
            branchCode = input("Enter the branch code: ")
            year = input("Enter the year: ")
            studentDictionary = getResult("1ms" + year + branchCode + str(usn).zfill(3))
            showGraph(studentDictionary)
        case 3:
            exit()
        case _:
            print("Invalid choice")
