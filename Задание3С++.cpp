#include <Windows.h> 
#include <string> //подключение объекта строк
#include "Class.h" //подключение заголовочного файла, где объявлены классы Student и University
#include <fstream> //подключение файлов
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
//тело функции стипендии
long Student::stipend() {


    if (Student::dva()) {
        return 0;
    }

    double avg = Student::sredBall();

    if (avg >= 4.5) {
        return 10000;
    }
    else if (avg >= 4.0) {
        return 7000;
    }
    else {
        return 4000;
    }
}
//тело процедуры добавления студента класса University
void University::addStudent(Student& student) {
    students.push_back(student);
}
//функция для записи полей в файл
static void writeHeader(std::ofstream& f) {
    f << "Index\tAvg\tStipend\n";
}
//тело процедуры анализа студента класса University
void University::processStudent() {
    //создание фалов(их путей, хранения)
    ofstream fTheBest("C:\\1\\TheBest.txt");
    ofstream fgood("C:\\1\\good.txt");
    ofstream fmiddle("C:\\1\\middle.txt");
    ofstream fTheWorst("C:\\1\\TheWorst.txt");

    //создание файлов(их полей)
    writeHeader(fTheBest);
    writeHeader(fgood);
    writeHeader(fmiddle);
    writeHeader(fTheWorst);


    //цикл записи в файлы и добавление в контейнер студентов
    for (Student& student : students) {
        double avg = student.sredBall();
        bool hastwo = student.dva();
        long stp = student.stipend();

        if (hastwo) {
            TheWorst.push_back(&student);
            fTheWorst << student.nameSt() << "\t" << avg << "\t" << stp << "\t0\n";
        }
        else if (avg >= 4.5) {
            TheBest.push_back(&student);
            fTheBest << student.nameSt() << "\t" << avg << "\t" << (avg * 1000) << "\t" << stp << "\n";
        }
        else if (avg >= 3.5) {
            good.push_back(&student);
            fgood << student.nameSt() << "\t" << avg << "\t" << (avg * 700) << "\t" << stp << "\n";
        }
        else {
            middle.push_back(&student);
            fmiddle << student.nameSt() << "\t" << avg << "\t" << (avg * 300) << "\t" << stp << "\n";
        }
    }

    //закрытие файлов
    fTheBest.close();
    fgood.close();
    fmiddle.close();
    fTheWorst.close();

}
//основная функция Dll
extern "C" __declspec(dllexport)
int __stdcall GetBestStudent(double* marks, int rows, int cols, int* TheBestCount)
{
    University u; //создание объекта класса University

    //создание объекта класса Student 
    for (int i = 0; i < rows; ++i) {
        Student student(i, &marks[i * cols], cols);
        u.addStudent(student);
    }

    u.processStudent(); //вызов функции анализа

    *TheBestCount = static_cast<int>(u.TheBest.size());
    //лямбда функция для поиска наилучшего студента по категориям 
    auto find_best = [&](const std::vector<Student*>& grp) {
        double bestA = -1.0;
        int bestI = -1;
        for (auto* st : grp) {
            double a = st->sredBall();
            if (a > bestA) {
                bestA = a;
                bestI = st->nameSt();
            }
        }
        return std::make_pair(bestA, bestI);
        };

    auto [avg, idx] = find_best(u.TheBest);
    if (idx != -1) {
        return idx;
    }

    std::tie(avg, idx) = find_best(u.good);
    if (idx != -1) {
        return idx;
    }

    std::tie(avg, idx) = find_best(u.middle);
    if (idx != -1) {
        return idx;
    }

    std::tie(avg, idx) = find_best(u.TheWorst);
    return idx;
}
