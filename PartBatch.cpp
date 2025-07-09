#pragma once
#include <vector>
#include <fstream>
//Класс партий(базовый)
class PartBatch {
private:
	int norm_parts_;
	int bad_parts_;
	double porog_;
	double fine_;
public:
	PartBatch(int norm_parts, int bad_parts, double porog, double fine);
	virtual double AnalyzeBraka();
	virtual bool IsBrak();
	virtual double CalculateFine();
	int GetNormParts();
	int GetBadParts();
};
//Класс потомок стандартных деталей
class StandardPartBatch : public PartBatch {
public:
	StandardPartBatch(int norm_parts, int bad_parts);
	bool IsBrak() override;
	double CalculateFine() override;
};
//класс потомок критических деталей
class CriticalPartBatch : public PartBatch {
public:
	CriticalPartBatch(int norm_parts, int bad_parts);
	bool IsBrak() override;
	double CalculateFine() override;
};