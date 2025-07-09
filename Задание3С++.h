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

    double GetAvgRounded();

    string GetNameStudent();

    bool GetTwo();

    long Scholarship();

    

};

//класс University
class University {
private:
    vector<Student> students_;

public:

    vector<Student*> GetTheBest();
    vector<Student*> GetGood();
    vector<Student*> GetMiddle();
    vector<Student*> GetTheWorst();

    void AddStudent(Student& student);

    void ProcessStudent(vector<int>& student_i);

    vector<Student>& GetStudents();

};


