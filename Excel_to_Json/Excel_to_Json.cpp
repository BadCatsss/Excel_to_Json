
#include "pch.h"
#include "BasicExcel.hpp"
#include <string>
#include <set>
#include <boost/property_tree/ptree.hpp>
#include <boost/property_tree/json_parser.hpp>
#include <boost/optional.hpp>

#define _SCL_SECURE_NO_WARNINGS
#define _CRT_SECURE_NO_WARNINGS
#pragma warning(disable: 4996)
using namespace YExcel;
using namespace std;
using boost::property_tree::ptree;
size_t maxRows;
size_t maxCols;
set<string> Names;
BasicExcelCell* cell;
multimap<string, int> Data;
BasicExcelWorksheet* sheet1;
vector<vector<int>> DataBlock;
string f_name;
void PrintNames()
{
	for (auto it : Names)
	{
		cout << "Found Names" << setw(10) << it << endl;
	}
}
void PrintData()
{

	for (auto i : Data)
	{
		cout << i.first << setw(4) << i.second << endl;
	}
}

void PrintBlocksToFile()
{



	auto it = Names.begin();


	// Short alias for this namespace
	namespace pt = boost::property_tree;

	// Create a root
	pt::ptree root;


	// Add a matrix

	vector<vector<int>> matrix(DataBlock.size());
	for (int i = 0; i < DataBlock.size(); i++)
	{
		matrix[i] = (vector<int>(1000));

	}
	for (int c = 0; c < DataBlock.size(); c++)
	{

		pt::ptree matrix_node;
		for (int r = 0; r < DataBlock[c].size(); r += 4)
		{
			pt::ptree row;
			// Create an unnamed value
			pt::ptree cell;

			// Add the value to our row
			for (size_t i = 0; i < 4; i++)
			{
				cell.put_value(DataBlock[c][r + i]);
				row.push_back(std::make_pair("", cell));
			}
			matrix_node.push_back(std::make_pair("", row));
		}
		// Add the row to our matrix

		const char* value1 = (*it).c_str();
		root.add_child(value1, matrix_node);
		it++;

	}


	// Add the node to the root

	string s= strtok(const_cast<char*>(f_name.c_str()), ".");
	write_json(s + ".json", root);



	cout << "file wrote" << endl;





}
void GetDigitalStandartName()
{

	for (size_t r = 1; r < maxRows; ++r)
	{
		BasicExcelCell* cell = sheet1->Cell(r, 0);
		Names.emplace(cell->GetString());
	}
	PrintNames();
};
void GetData()
{
	int r = 1;
	auto Names_it = Names.begin();
	vector<vector<int>> Datablock;


	for (; r < maxRows; r++)
	{
		for (int c = 1; c <= 4; c++)
		{
			Data.insert(pair< string, int >(sheet1->Cell(r, 0)->GetString(), sheet1->Cell(r, c)->GetInteger()));
		}
	}


	//PrintData();

};
void CreateBlockByName()
{

	DataBlock.resize((Names.size()));
	auto it = Names.begin();
	for (size_t i = 0; i < Names.size(); i++)
	{
		DataBlock[i].reserve(1000);
	}
	int i = 0;


	for (auto k : Data)
	{
		if (k.first != *it && i < Names.size() - 1)
		{
			DataBlock[i].shrink_to_fit();
			i++;
			it++;
			/*cout << k.first;*/
		}
		DataBlock[i].push_back(k.second);
	}


	PrintBlocksToFile();
}

int main()
{
	BasicExcel e;
	cout << "enter file name:" << endl;
	cout << "example: 'file.xls'" << endl;
	cin >> f_name;
	const char* f_ch_name = f_name.c_str();

	if (e.Load(f_ch_name))
	{
		cout << "file is open" << endl;
	}


	cout << "enter sheet name:" << endl;
	cout << "example: 'Sheet1'" << endl;
	string sheet_name;
	cin >> sheet_name;
	

	sheet1 = e.GetWorksheet(sheet_name.c_str());

	if (sheet1)
	{
		cout << "sheet is open" << endl;
		maxRows = sheet1->GetTotalRows();
		maxCols = sheet1->GetTotalCols();


		GetDigitalStandartName();
		GetData();
		CreateBlockByName();

	}
	cout << endl;



	return 0;
}

// Запуск программы: CTRL+F5 или меню "Отладка" > "Запуск без отладки"
// Отладка программы: F5 или меню "Отладка" > "Запустить отладку"

// Советы по началу работы 
//   1. В окне обозревателя решений можно добавлять файлы и управлять ими.
//   2. В окне Team Explorer можно подключиться к системе управления версиями.
//   3. В окне "Выходные данные" можно просматривать выходные данные сборки и другие сообщения.
//   4. В окне "Список ошибок" можно просматривать ошибки.
//   5. Последовательно выберите пункты меню "Проект" > "Добавить новый элемент", чтобы создать файлы кода, или "Проект" > "Добавить существующий элемент", чтобы добавить в проект существующие файлы кода.
//   6. Чтобы снова открыть этот проект позже, выберите пункты меню "Файл" > "Открыть" > "Проект" и выберите SLN-файл.
