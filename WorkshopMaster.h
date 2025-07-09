#include <Windows.h>
#include "ProductionDepartment.h"
#include <iostream>
#include <fstream>
#include <comutil.h>
#include <map>
#include <set>
#include <vector>
#include <string>

using namespace std;
//Функция dll
extern "C" __declspec(dllexport)
int __stdcall AnalyzeProduction(int* data, int* brak_data, int* type_data, char* names, int w_n, int n) {

	int summa = 0;

	

	static vector<PartBatch*> batch;
	static set<string> workers;

	int ta = 0;
	int tb = 0;

	int j = 0;
	char d[256];
	while (j < w_n) {
		int t = 0;
		while (names[j] != '@') {
			d[t] = names[j];
			t++;
			j++;
		}
		j++;
		d[t] = 0;
		workers.insert(string (d));
	}

	map<string, WorkshopMaster*> wm;

	for (auto it = workers.begin(); it != workers.end(); ++it) {
		wm[*it] = new WorkshopMaster(*it);
	}

	 j = 0;

	

	for (int i = 0; i < n; i++) {
		PartBatch* p;

		ta += data[i];
		tb += brak_data[i];

		int t = j;
		while (names[t] != '@') {
			d[t - j] = names[t];
			++t;
		}

		d[t - j] = 0;
		string key = string(d);

		j = t + 1;

		if (type_data[i] == 1) {
			p = new StandardPartBatch(data[i], brak_data[i]);
		}
		else {
			p = new CriticalPartBatch(data[i], brak_data[i]);
		}

		wm[key]->AddDetails(p);

		batch.push_back(p);





	

	}


	ProductionDepartment depart;

	for (auto it = wm.begin(); it != wm.end(); ++it) {
		depart.AddMaster(it->second);
	}

	auto bad = depart.AnalyzeMasters();

	ofstream bad_file("OverDefected.txt");

	for (auto it = bad.begin(); it != bad.end(); ++it) {
		bad_file << (*it)->GetName() << endl;
	}

	bad_file.close();


	double fines = 0;

	for (int i = 0; i < batch.size(); i++) {
		fines += batch[i]->CalculateFine();
	}
	

	ofstream ps("ProductionSummary.txt");

	int sd = 0;
	int sdd = 0;
	double bd = 0.0;

	depart.Svodka(sd, sdd, bd);
	ps << sd << endl << sdd << endl << bd << endl;

	ps.close();



	return bad.size();


}



