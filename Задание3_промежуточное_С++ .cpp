#include <Windows.h>
#include <string> //подключение объекта строк
#include "Class.h" //подключение заголовочного файла, где объявлены классы Student и University
#include <iostream>
#include <vector> //подключение контейнера vector

using namespace std;
//тело конструктора класса Student
Student::Student(int name, double* marks, int count) {
    Name = name;
    Marks = marks;
    Count = count;
}

//тело функции среднего балла класса Student
double Student::sredBall() {
    double summa = 0.0;
    for (int i = 0; i < Count; i++) {
        summa += Marks[i];
    }
    return summa / Count;
}
//тело функции индекса студента класса Student
int Student::nameSt() {
    return Name;
}
//тело функции наличия двойки у студента класса Student
bool Student::dva() {
    for (int i = 0; i < Count; i++) {
        if (Marks[i] == 2.0) {
            return true;
        }
    }
    return false;
}
//тело процедуры добавления студента класса University
void University::addStudent(Student& student) {
    students.push_back(student);
}

//тело процедуры анализа студента класса University
void University::processStudent() {
    for (Student& student : students) {
        if (student.dva()) {
            TheWorst.push_back(&student);
        }
        else {
            double avg = student.sredBall();
            if (avg >= 4.5) {
                TheBest.push_back(&student);
            }
            else if (avg > 3.5 && avg < 4.5) {
                good.push_back(&student);
            }
            else {
                middle.push_back(&student);
            }
        }
    }
}
//основная функция Dll
extern "C" __declspec(dllexport)
int __stdcall GetBestStudent(double* marks, int rows, int cols, int* TheBestCount)
{
    University u;
    for (int i = 0; i < rows; ++i) {
        Student student(i, &marks[i * cols], cols);
        u.addStudent(student);
    }

    u.processStudent();

    int bestIndex = -1;
    double bestAvg = -1.0;

    for (Student* s : u.TheBest) {
        double avg = s->sredBall();
        if (avg > bestAvg) {
            bestAvg = avg;
            bestIndex = s->nameSt();
        }
    }

    *TheBestCount = u.TheBest.size();
    return bestIndex;
}
