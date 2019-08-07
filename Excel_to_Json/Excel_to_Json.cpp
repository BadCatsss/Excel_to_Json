
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


using boost::property_tree::ptree;
#include "Data_Interface.h"


void Convertor::PrintBlocksToFile()
{



	auto it = this->Names.begin();


	// Short alias for this namespace
	namespace pt = boost::property_tree;

	// Create a root
	pt::ptree root;


	// Add a matrix

	vector<vector<int>> matrix(this->DataBlock.size());
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

	string s = strtok(const_cast<char*>(this->f_str_name.c_str()), ".");
	write_json(s + ".json", root);



	cout << "file wrote" << endl;





}
void Convertor::GetDigitalStandartName()
{

	for (size_t r = 1; r < maxRows; ++r)
	{
		YExcel::BasicExcelCell* cell = sheet1->Cell(r, 0);
		Names.emplace(cell->GetString());
	}
	if (this->PrintNameFlag)
	{
		PrintNames();
	}

};
void Convertor::GetData()
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

	if (PrintDataFlag)
	{
		PrintData();
	}
	

};
void Convertor::CreateBlockByName()
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
void Convertor::PrintNames()
{
	for (auto it : Names)
	{
		cout << "Found Names" << setw(10) << it << endl;
	}
}
void Convertor::PrintData()
{

	for (auto i : Data)
	{
		cout << i.first << setw(4) << i.second << endl;
	}
}
void Convertor::OpenFile()
{
	if (e.Load(this->f_str_name.c_str()))
	{
		cout << "file is open" << endl;
	}
}
void Convertor::SetSheet(string sheet_name)
{
	this->sheet_name = sheet_name;
}
string Convertor::GetSheetNameStr()
{
	return this->sheet_name;
}
const char*  Convertor::GetSheetNameChar()
{
	return this->sheet_name.c_str();
}
void Convertor::OpenSheet(YExcel::BasicExcelWorksheet* p)
{
	this->sheet1 = p;
}
YExcel::BasicExcelWorksheet* Convertor::GetSheetObj()
{
	return this->sheet1;
}
string Convertor::GetFileStrName()
{
	return this->f_str_name;
}
const char* Convertor::GetFileCharName()
{
	return this->f_str_name.c_str();
}
void Convertor::SetMaxRows(rsize_t s)
{
	this->maxRows = s;
}
void Convertor::SetMaxColumns(rsize_t s)
{
	this->maxCols = s;
}
void Convertor::Generate()
{
	this->GetDigitalStandartName();
	this->GetData();
	this->CreateBlockByName();
}





int main()
{
	string f_name;
	string sh_name;
	cout << "enter file name:" << endl;
	cout << "example: 'file.xls'" << endl;
	cin >> f_name;

	cout << "enter sheet name:" << endl;
	cout << "example: 'Sheet1'" << endl;
	cin >> sh_name;
	Convertor cnv(f_name,false,true);
	cnv.OpenFile();
	cnv.SetSheet(sh_name.c_str());
		

	cnv.OpenSheet( cnv.e.GetWorksheet (cnv.GetSheetNameChar()));

	if (cnv.GetSheetObj())
	{
		cout << "sheet is open" << endl;
		cnv.SetMaxRows( cnv.GetSheetObj()->GetTotalRows());
		cnv.SetMaxColumns( cnv.GetSheetObj()->GetTotalCols());
		cnv.Generate();

		

	}
	cout << endl;



	return 0;
}

