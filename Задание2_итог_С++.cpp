#include <Windows.h> // Подключение файла Windows API

//Подключение конвенции о связях stdcall для взаимодейстаия функции на С++ с VBA для поиска лучшего ученика
extern "C" __declspec(dllexport)
    int __stdcall GetBestStudent(double* marks, int rows, int cols)
{
    int best = -1;
    double maxAvg = -1.0;

    //Внешний цикл для прохода по строкам(ученикам)
    for (int i = 0; i < rows; ++i)
    {
        bool hasTwo = false; //флаг для двоек у учеников
        double sum = 0.0;    // Сумма оценок ученика по всем предметам

         //Внутренний цикл для прохода по столбцам(предметам)
        for (int j = 0; j < cols; ++j)
        {


            double m = marks[i * cols + j];

            // Проверяем наличие двойки
            if (m == 2.0) hasTwo = true;


            sum += m;
        }

        // Если у ученика есть хотя бы одна двойка - пропускаем его
        if (hasTwo) continue;

        // Вычисляем средний балл ученика
        double avg = sum / cols;

        // Проверяем, является ли текущий ученик лучшим
        if (avg > maxAvg)
        {
            maxAvg = avg;
            best = i;
        }
    }

    //возвращаем индекс лучшего ученика по среднему баллу
    return best;
}
