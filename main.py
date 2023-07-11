import requests
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import json

studentsList = []
studentDictionary = dict()


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
        color = 'gray',
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
        case 2:
            usn = input("Enter the USN number: ")
            branchCode = input("Enter the branch code: ")
            year = input("Enter the year: ")
            studentDictionary = getResult("1ms" + year + branchCode + usn)
            showGraph(studentDictionary)
        case 3:
            exit()
        case _:
            print("Invalid choice")
