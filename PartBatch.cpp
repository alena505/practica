#include "PartBatch.h"
#include <fstream>

//Конструктор базового класса
PartBatch::PartBatch(int norm_parts, int bad_parts, double porog, double fine)
    : norm_parts_(norm_parts), bad_parts_(bad_parts), porog_(porog), fine_(fine) {}

//Анализ процента бракованных деталей
double PartBatch::AnalyzeBraka() {
    if (norm_parts_ == 0) return 0.0;
    return (static_cast<double>(bad_parts_) / norm_parts_) * 100.0;
}
//Анализ является ли деталь бракованной
bool PartBatch::IsBrak() {
    return AnalyzeBraka() > porog_;
}
//Рассчёт штрафа
double PartBatch::CalculateFine() {
    return bad_parts_ * fine_;
}

//Класс потомка стандартные детали
StandardPartBatch::StandardPartBatch(int norm_parts, int bad_parts)
    : PartBatch(norm_parts, bad_parts, 5.0, 100000.0) {}


//Виртуальная переопределённая функция анализа является ли деталь бракованной
bool StandardPartBatch::IsBrak() {
    return PartBatch::IsBrak();
}
//Виртуальная переопределённая функция расчёта штрафа
double StandardPartBatch::CalculateFine() {
    return PartBatch::CalculateFine();
}

//Класс потомка критических деталей
CriticalPartBatch::CriticalPartBatch(int norm_parts, int bad_parts)
    : PartBatch(norm_parts, bad_parts, 2.0, 500000.0) {}
//Виртуальная переопределённая функция анализа является ли деталь бракованной
bool CriticalPartBatch::IsBrak() {
    return PartBatch::IsBrak();
}
//Виртуальная переопределённая функция расчёта штрафа
double CriticalPartBatch::CalculateFine() {
    return PartBatch::CalculateFine();
}
//Делаем функцию для того, чтобы приватное поле norm_parts было доступно с помощбю этой функции
int PartBatch::GetNormParts() {
    return norm_parts_;
}

int PartBatch::GetBadParts() {
    return bad_parts_;
}
