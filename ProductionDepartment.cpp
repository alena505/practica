#include "WorkshopMaster.h"
#include "PartBatch.h"
#pragma once


class ProductionDepartment {
private:
	vector<WorkshopMaster*> name_master_;
public:
	vector<WorkshopMaster*> AnalyzeMasters();
	double CalculateGeneralBrak();
	void AddMaster(WorkshopMaster* master);
	void Svodka(int& summ_details, int& summ_defect_details, double& summ_avg_deffect_procent);
};