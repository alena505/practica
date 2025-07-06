#include <Windows.h> 
#include <string> //����������� ������� �����
#include "Class.h" //����������� ������������� �����, ��� ��������� ������ Student � University
#include <fstream> //����������� ������
#include <iostream>
#include <vector> //����������� ���������� vector

using namespace std;
//���� ������������ ������ Student
Student::Student(int name, double* marks, int count) {
    Name = name;
    Marks = marks;
    Count = count;
}
//���� ������� �������� ����� ������ Student
double Student::sredBall() {
    double summa = 0.0;
    for (int i = 0; i < Count; i++) {
        summa += Marks[i];
    }
    return summa / Count;
}
//���� ������� ������� �������� ������ Student
int Student::nameSt() {
    return Name;
}
//���� ������� ������� ������ � �������� ������ Student
bool Student::dva() {
    for (int i = 0; i < Count; i++) {
        if (Marks[i] == 2.0) {
            return true;
        }
    }
    return false;
}
//���� ������� ���������
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
//���� ��������� ���������� �������� ������ University
void University::addStudent(Student& student) {
    students.push_back(student);
}
//������� ��� ������ ����� � ����
static void writeHeader(std::ofstream& f) {
    f << "Index\tAvg\tStipend\n";
}
//���� ��������� ������� �������� ������ University
void University::processStudent() {
    //�������� �����(�� �����, ��������)
    ofstream fTheBest("C:\\1\\TheBest.txt");
    ofstream fgood("C:\\1\\good.txt");
    ofstream fmiddle("C:\\1\\middle.txt");
    ofstream fTheWorst("C:\\1\\TheWorst.txt");

    //�������� ������(�� �����)
    writeHeader(fTheBest);
    writeHeader(fgood);
    writeHeader(fmiddle);
    writeHeader(fTheWorst);


    //���� ������ � ����� � ���������� � ��������� ���������
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

    //�������� ������
    fTheBest.close();
    fgood.close();
    fmiddle.close();
    fTheWorst.close();

}
//�������� ������� Dll
extern "C" __declspec(dllexport)
int __stdcall GetBestStudent(double* marks, int rows, int cols, int* TheBestCount)
{
    University u; //�������� ������� ������ University

    //�������� ������� ������ Student 
    for (int i = 0; i < rows; ++i) {
        Student student(i, &marks[i * cols], cols);
        u.addStudent(student);
    }

    u.processStudent(); //����� ������� �������

    *TheBestCount = static_cast<int>(u.TheBest.size());
    //������ ������� ��� ������ ���������� �������� �� ���������� 
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
