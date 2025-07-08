#include <Windows.h>
#include <string> //подключение объекта строк
#include "Class.h" //подключение заголовочного файла, где объявлены классы Student и University
#include <fstream> //подключение файлов
#include <iostream>
#include <vector> //подключение контейнера vector
#include <comutil.h>




using namespace std;
//тело конструктора класса Student
Student::Student(string name, double* marks, int count) : name_(name), marks_(marks, marks + count), count_(count) {}


//тело функции среднего балла класса Student
double Student::GetAvg() {
    double summa = 0.0;
    for (double mark : marks_) {
        summa += mark;
    }
    return summa / marks_.size();
}
//тело функции индекса студента класса Student
string Student::GetNameStudent() {
    return name_;
}

//тело функции наличия двойки у студента класса Student
bool Student::GetTwo() {
    for (double mark : marks_) {
        if (mark == 2.0) {
            return true;
        }
    }
    return false;
}

//тело функции стипендии
long Student::Scholarship() {


    if (Student::GetTwo()) {
        return 0;
    }

    double avg = Student::GetAvg();

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
void University::AddStudent(Student& student) {
    students_.push_back(student);
}
//функция для записи полей в файл
static void WriteHeader(ofstream& f) {
    f << "Name\tAvg\tScholarship\n";
}
vector<Student*> University::GetTheBest() {
    vector<Student*> result;
    for (Student& student : students_) {
        if (!student.GetTwo() && student.GetAvg() >= 4.5) {
            result.push_back(const_cast<Student*>(&student));
        }
    }
    return result;
}
vector<Student*> University::GetGood() {
    vector<Student*> result;
    for (Student& student : students_) {
        if (!student.GetTwo() && student.GetAvg() >= 4.0) {
            result.push_back(const_cast<Student*>(&student));
        }
    }
    return result;
}
vector<Student*> University::GetMiddle() {
    vector<Student*> result;
    for (Student& student : students_) {
        if (!student.GetTwo() && student.GetAvg() < 4.0) {
            result.push_back(const_cast<Student*>(&student));
        }
    }
    return result;
}
vector<Student*> University::GetTheWorst() {
    vector<Student*> result;
    for (Student& student : students_) {
        if (student.GetTwo()) {
            result.push_back(const_cast<Student*>(&student));
        }
    }
    return result;
}
void WriteStudentGroupToFile(const string& filename, const vector<Student*>& students) {
    ofstream file(filename); 

    if (!file.is_open()) {
        cerr << "Ошибка открытия файла: " << filename << endl;
        return;
    }

    // Записываем заголовок
    file << "Name\tAvg\tScholarship\n";

    // Записываем данные студентов
    for (Student* student : students) {
        if (student) { 
            double avg = student->GetAvg();
            long stp = student->Scholarship();
            file << student->GetNameStudent() << "\t" << avg << "\t" << stp << "\n";
        }
    }

    file.close();

}
//тело процедуры анализа студента класса University
void University::ProcessStudent() {
    vector<Student*> the_best = GetTheBest();
    vector<Student*> good = GetGood();
    vector<Student*> middle = GetMiddle();
    vector<Student*> the_worst = GetTheWorst();



    WriteStudentGroupToFile("TheBest.txt", the_best);
    WriteStudentGroupToFile("good.txt", good);
    WriteStudentGroupToFile("middle.txt", middle);
    WriteStudentGroupToFile("TheWorst.txt", the_worst);





}
//основная функция Dll
extern "C" __declspec(dllexport)
int __stdcall GetBestStudent(
    double* marks,
    VARIANT names_var,
    int rows,
    int cols,
    int* the_best_count
) {
    University university;

    if (!(names_var.vt & VT_ARRAY)) {
        *the_best_count = 0;
        return -1;
    }


    SAFEARRAY* psa = names_var.parray;
    long l_bound, u_bound;
    SafeArrayGetLBound(psa, 1, &l_bound);
    SafeArrayGetUBound(psa, 1, &u_bound);
    long cnt_names = u_bound - l_bound + 1;

    if (names_var.vt == (VT_ARRAY | VT_BSTR)) {
        BSTR* p_bstr = nullptr;
        SafeArrayAccessData(psa, reinterpret_cast<void**>(&p_bstr));
        for (long i = 0; i < rows && i < cnt_names; ++i) {
            _bstr_t bt(p_bstr[i]);
            string sname = static_cast<const char*>(bt);
            Student st(sname, &marks[i * cols], cols);
            university.AddStudent(st);
        }
        SafeArrayUnaccessData(psa);
    }
    else if (names_var.vt == (VT_ARRAY | VT_VARIANT)) {
        VARIANT* p_var = nullptr;
        SafeArrayAccessData(psa, reinterpret_cast<void**>(&p_var));
        for (LONG i = 0; i < rows && i < cnt_names; ++i) {
            if (p_var[i].vt == VT_BSTR) {
                _bstr_t bt(p_var[i].bstrVal);
                string sname = static_cast<const char*>(bt);
                Student st(sname, &marks[i * cols], cols);
                university.AddStudent(st);
            }
        }
        SafeArrayUnaccessData(psa);
    }
    else {
        for (int i = 0; i < rows; ++i) {
            Student st(string(), &marks[i * cols], cols);
            university.AddStudent(st);
        }
    }

    auto best_grp = university.GetTheBest();
    *the_best_count = static_cast<int>(best_grp.size());

    auto find_best = [&](const vector<Student*>& grp) {
        long best_s = -1;
        const Student* best_st = nullptr;
        for (auto* st : grp) {
            long sch = st->Scholarship();
            if (sch > best_s) {
                best_s = sch;
                best_st = st;
            }
        }
        if (!best_st) return -1;
        const auto& all = university.GetStudents();
        for (int i = 0; i < static_cast<int>(all.size()); ++i) {
            if (&all[i] == best_st) return i;
        }
        return -1;
        };

    std::vector<std::vector<Student*>> groups = {
        university.GetTheBest(),
        university.GetGood(),
        university.GetMiddle(),
        university.GetTheWorst()
    };

    for (auto& grp : groups) {
        if (grp.empty()) continue;
        int idx = find_best(grp);
        if (idx != -1) return idx;
    }

    return -1;
}
