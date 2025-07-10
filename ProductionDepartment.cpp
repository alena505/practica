#include "ProductionDepartment.h"
#include <vector>
#include <fstream>


using namespace std;

//Функция добавления мастера
void ProductionDepartment::AddMaster(WorkshopMaster* master) {
	name_master_.push_back(master);
}

//Анализ мастеров
vector<WorkshopMaster*> ProductionDepartment::AnalyzeMasters() {
	vector<WorkshopMaster*> bad_masters;
	for (auto* master : name_master_) {
		bool has_narush = false;
		for (PartBatch* part : master->details_) {
			double defect_procent = part->AnalyzeBraka();

			if (part->IsBrak()) {
				has_narush = true;
				break;
			}
			
	

		}
		if (has_narush) {
			bad_masters.push_back(master);
		}
		
	}


	return bad_masters;

}
//Общий расчёт
double ProductionDepartment::CalculateGeneralBrak() {
	double summa_fines = 0.0;
	for (WorkshopMaster* master : name_master_) {
		for (PartBatch* part : master->details_) {
			summa_fines += part->CalculateFine();
		}
	}
	return summa_fines;
}

//Сводка по деталям
void ProductionDepartment::Svodka(int& summ_details, int& summ_defect_details, double& summ_avg_deffect_procent) {
	summ_details = 0;
	summ_defect_details = 0;
	for (WorkshopMaster* master : name_master_) {
		for (PartBatch* part : master->details_) {
			summ_details += part->GetNormParts() + part->GetBadParts();
			summ_defect_details += part->GetBadParts();
		}
	}

	if (summ_details > 0) {
		summ_avg_deffect_procent = (static_cast<double>(summ_defect_details) / summ_details) * 100.0;
	}
	else {
		summ_avg_deffect_procent = 0.0;
	}

}
