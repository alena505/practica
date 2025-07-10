#include "WorkshopMaster.h"
#include <vector>

using namespace std;

WorkshopMaster::WorkshopMaster(string name): name_master_(name){}
//Деструктор
WorkshopMaster::~WorkshopMaster() {
	for (auto* detail:details_) {
		delete detail;
	}
}
//Функция добавления деталей
void WorkshopMaster::AddDetails(PartBatch* details) {
	details_.push_back(details);
}
//Функция для получения приватного имени
string WorkshopMaster::GetName() const {
	return name_master_;
}

