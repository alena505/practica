#pragma once
#include <vector>
#include <string>

using namespace std;

//класс Студент:
class Student {
private:
    vector<double> marks_;
    int count_;
    string name_;
public:
    Student(string name, double* marks, int count);

    double GetAvg();

    string GetNameStudent();

    bool GetTwo();



    long Scholarship();
};

//Класс универа
class University {
private:
    vector<Student> students_;
public:

    vector<Student*> GetTheBest();
    vector<Student*> GetGood();
    vector<Student*> GetMiddle();
    vector<Student*> GetTheWorst();

    /* vector<Student*> the_best;
     vector<Student*> good;
     vector<Student*> middle;
     vector<Student*> the_worst;*/

    void AddStudent(Student& student);
    void ProcessStudent();
    const vector<Student>& GetStudents() const { return students_; }
};
