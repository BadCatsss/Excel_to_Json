


#include <QCoreApplication>
//#include <boost/property_tree/ptree.hpp>
//#include <boost/property_tree/json_parser.hpp>
//#include <boost/optional.hpp>
#include <QCoreApplication>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
#include <QDebug>
using namespace QXlsx;
using namespace std;

//using boost::property_tree::ptree;
#include "Data_Interface.h"


void Convertor::PrintBlocksToFile()
{
    auto it_Names = Names.begin();
    auto it_Data_Block=DataBlock.begin();
    // Текущий json объект, с которым производится работа
    QJsonObject m_currentJsonObject;
    QJsonArray listArray;
    // Создаём объект текста
    int c=0;
    int r=0;
    int current_size;

    for (;it_Data_Block!=DataBlock.end();it_Data_Block++,c++) {
        r=0;
        QJsonArray textsArray = m_currentJsonObject[it_Names->toString()].toArray();
        for(vector<int>::iterator it2 = (*it_Data_Block).begin() ; it2 != (*it_Data_Block).end(),r+4<=DataBlock[c].size(); ++it2 ,r += 4){

            for (int var = 0; var < 4; ++var) {
                textsArray.push_back(DataBlock[c][r+var]);
            }
            listArray.push_back(textsArray);


            m_currentJsonObject[it_Names->toString()] = listArray;
            current_size =textsArray.size();
            for (int var = current_size; var >=0; --var) {
                textsArray.removeAt(var);
            }

            //
        }

        it_Names++;
        // Добавляем объект текста в массив
        // Сохраняем массив обратно в текущий объект

    }


    QString saveFileName="old.json";

    // Создаём объект файла и открываем его на запись
    QFile jsonFile(saveFileName);
    if (!jsonFile.open(QIODevice::WriteOnly))
    {
        cout<<"error"<<endl;
        return;
    }

    // Записываем текущий объект Json в файл
    jsonFile.write(QJsonDocument(m_currentJsonObject).toJson(QJsonDocument::Indented));
    jsonFile.close();   // Закрываем файл




}
void Convertor::GetDigitalStandartName()
{

    for (size_t r = 2; r < maxRows; r++)
    {
        Cell* cell = this->GetBook()->cellAt(r, 1);
        Names.emplace(cell->value());
    }
    if (this->PrintNameFlag)
    {
        PrintNames();
    }

};
void Convertor::GetData()
{
    int r = 2;
    auto Names_it = Names.begin();
    vector<vector<int>> Datablock;

    QTextStream qtout(stdout);
    for (; r < maxRows; r++)
    {

        for (int c = 2; c <maxCols; c++)
        {

// при 0 - падает
            Data.insert(this->xlsxR->cellAt(r, 1)->value().toString(), this->xlsxR->cellAt(r, c)->value().toInt());
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
    QMultiMap<QString, int > ::iterator Data_it=Data.begin();
    for (size_t i = 0; i < Names.size(); i++)
    {
        DataBlock[i].reserve(1000);
    }
    int i = 0;


    for (; Data_it!=Data.end();Data_it++)
    {

        if (Data_it.key() != *it && i < Names.size() - 1)
        {
            DataBlock[i].shrink_to_fit();
            i++;
            it++;
            /*cout << k.first;*/
        }
        DataBlock[i].push_back(Data_it.value());
    }


    PrintBlocksToFile();
}
void Convertor::PrintNames()
{
    QTextStream qtout(stdout);
    for (auto it : Names)
    {
        qtout   << it.toString() << endl;
    }
}
void Convertor::PrintData()
{ QTextStream qtout(stdout);
    QMultiMap<QString, int > ::iterator it=Data.begin();
    for (;it!=Data.end();it++)
    {
        qtout<<it.key();
        qtout<<it.value()<<endl;
    }
}
void Convertor::OpenBook()
{
    if (this->xlsxR->load())
    {
        cout << "file is open" << endl;
    }
    else {
        throw exception(" Cant open File\n");
    }
}
QXlsx::Document* Convertor:: GetBook() const
{
    return this->xlsxR;
}
QStringList  Convertor::OpenWorkSheet()
{


    cout<<"Sheet is found"<<endl;
    QStringList list_;
    QTextStream qtout(stdout);
    this->sheet_count;
    for (auto i:  this->xlsxR->sheetNames())
    {
        list_.push_back(i);
        cout<<sheet_count<<"  ";
        qtout<<i<<endl;
        sheet_count++;
    }

    return list_;
}
int Convertor:: GetSheetCount() const
{
    return this->sheet_count;
}

void Convertor::SetActivetWorkSheet(QString p)
{
    this->ActiveSheet= this->xlsxR->sheet(p);

    this->xlsxR->selectSheet(this->ActiveSheet->sheetName());
}
QXlsx::AbstractSheet* Convertor:: GetActivetWorkSheet() const
{
    QTextStream qtout(stdout);
    qtout<<this->ActiveSheet->sheetName()<<endl;
    return this->ActiveSheet;

}




void Convertor::SetMaxRows( )
{
    int row = 1;
    Cell* cell = this->xlsxR->cellAt(row, 1); // get cell pointer.
    while (cell!=NULL) {
        row++;
        cell = this->xlsxR->cellAt(row, 1);
    }
    this->maxRows=row;
    cout<<"row "<<this->maxRows<<endl;

}
void Convertor::SetMaxColumns( )
{
    int col = 1;
    Cell* cell = this->xlsxR->cellAt(1, col); // get cell pointer.
    while (cell!=NULL) {
        col++;
        cell = this->xlsxR->cellAt(1, col);
    }
    this->maxCols=col;
    cout<<"col "<<this->maxCols<<endl;

}
void Convertor::Generate()
{
    this->GetDigitalStandartName();
    this->GetData();
    this->CreateBlockByName();
}





int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    QTextStream qtin(stdin);

    QTextStream qtout(stdout);
    QString f_name;
    QString sh_name;

    bool PrintDataFlag=true,PrintNameFlag=true;
    int ch_1,ch_2;
    qtout<<"welcome to xlsx_to_json converter"<<endl;
    qtout << "enter  open file name:" << endl;
    qtout << "example: 'file.xlsx'" << endl;
    qtin >> f_name;
    qtout<<" want you print names?"<<endl;
    qtout<<"press 0 if not? else press any key"<<endl;
    qtin>>ch_1;
    PrintNameFlag=ch_1;
    qtout<<" want you  also print data to console?"<<endl;
    qtout<<"press 0 if not? else press any key"<<endl;
    qtin>>ch_2;
    PrintDataFlag=ch_2;
    Convertor cnv(f_name,PrintDataFlag,PrintNameFlag);
    try {
        cnv.OpenBook();

    } catch (exception& ex) {
        cout<<ex.what();
        return -1;
    }
    QStringList sh_list= cnv.OpenWorkSheet();
    int ch_number=0;
    cout<<"choose current worksheet. enter the number"<<endl;
    cin>>ch_number;
    if (ch_number>=0 && ch_number<= cnv.GetSheetCount() ) {
        cnv.SetActivetWorkSheet(sh_list[ch_number]);
    }
    cnv.GetActivetWorkSheet();
    cnv.SetMaxRows();
    cnv.SetMaxColumns();
    cnv.Generate();


    cout << endl;



    return a.exec();
}

