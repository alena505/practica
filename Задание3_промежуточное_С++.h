#pragma once
#include <vector>

using namespace std;

//класс Студент:
class Student {
private:
    double* Marks;
    int Count;
    int Name;
public:
    Student(int name, double* marks, int count);

    double sredBall();

    int nameSt();

    bool dva();



};

//Класс универа:
class University {
private:
    vector<Student> students;
public:
    vector<Student*> TheBest;
    vector<Student*> good;
    vector<Student*> middle;
    vector<Student*> TheWorst;
    void addStudent(Student& student);
    void processStudent();
};
