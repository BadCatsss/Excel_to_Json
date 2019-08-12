#include <QCoreApplication>
#include <QCoreApplication>
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
#include <QDebug>
#include "Convertor.h"
using namespace QXlsx;
using namespace std;

//initialize
Convertor::Convertor(const QString& f_str_path, bool PrintDataFlag, bool PrintNameFlag) :PrintDataFlag(PrintDataFlag), PrintNameFlag(PrintNameFlag),sheet_count(0),f_str_path(f_str_path)
{
    xlsxR = new QXlsx::Document(f_str_path);
    string tmp_s = f_str_path.toStdString();
    tmp_s.erase(tmp_s.begin() + tmp_s.find("."), tmp_s.end());
    this->f_str_name=QString::fromStdString(tmp_s);
}

//book and sheet setup
bool Convertor::openBook()
{
    if (this->xlsxR->load())
    {
        return true;
    }
    else {
        return false;
    }
}
void Convertor::calculateNotEmptyRowsCount( )
{
    int row = 1;
    Cell* cell = this->xlsxR->cellAt(row, 1); // get cell pointer.
    while (cell!=nullptr) {
        row++;
        cell = this->xlsxR->cellAt(row, 1);
    }
    this->maxRows=row;
    cout << "row "<< this->maxRows << endl;
}
void Convertor::calculateNotEmptyColumnsCount( )
{
    int col = 1;
    Cell* cell = this->xlsxR->cellAt(1, col); // get cell pointer.
    while (cell!=nullptr) {
        col++;
        cell = this->xlsxR->cellAt(1, col);
    }
    this->maxCols=col;
    cout<< "col " << this->maxCols << endl;

}
void Convertor::setActivetWorkSheet(QString p)
{
    this->ActiveSheet= this->xlsxR->sheet(p);
    this->xlsxR->selectSheet(this->ActiveSheet->sheetName());
}
QStringList  Convertor::getSheetsList()
{
    cout << "Sheet is found" << endl;
    QStringList list_;
    QTextStream qtout(stdout);
    this->sheet_count;
    for (auto i:  this->xlsxR->sheetNames())
    {
        list_.push_back(i);
        cout << sheet_count << "  ";
        qtout<<i<<endl;
        sheet_count++;
    }
    return list_;
}

//pharse
void Convertor::convert()
{
    QStringList sh_list = this->getSheetsList();
    int choose_number = 0;
    cout << "choose current worksheet. enter the number" << endl;
    cin >> choose_number;
    if (choose_number>=0 && choose_number<= this->sheet_count ) {
        this->setActivetWorkSheet(sh_list[choose_number]);
    }
    QTextStream qtout(stdout);
    qtout << "You chose " << this->ActiveSheet->sheetName() << endl;
    this->calculateNotEmptyRowsCount();
    this->calculateNotEmptyColumnsCount();
    this->GetDigitalStandartName();
    this->GetData();
    this->CreateBlockByName();
}
void Convertor::GetDigitalStandartName()
{
    for (size_t r = 2; r < maxRows; r++)
    {
        Cell* cell = this->xlsxR->cellAt(r, 1);
        Names.emplace(cell->value());
    }
    if (this->PrintNameFlag)
    {
        printNamesToConsole();
    }

}
void Convertor::GetData()
{
    int r = 2;
    auto Names_it = Names.begin();
    vector<vector<int>> Datablock;
    QTextStream qtout(stdout);
    for (; r < maxRows; r++)
    {

        for (int c = 2; c < maxCols; c++)
        {

            // при 0 - падает
            Data.insert(this->xlsxR->cellAt(r, 1)->value().toString(), this->xlsxR->cellAt(r, c)->value().toInt());
        }
    }

    if (PrintDataFlag)
    {
        printDataToConsole();
    }


}
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

//out
//to console
void Convertor::printNamesToConsole()
{
    QTextStream qtout(stdout);
    for (auto it : Names)
    {
        qtout << it.toString() << endl;
    }
}
void Convertor::printDataToConsole()
{ QTextStream qtout(stdout);
    QMultiMap<QString, int > ::iterator it = Data.begin();
    for (;it!=Data.end();it++)
    {
        qtout << it.key();
        qtout << it.value() << endl;
    }
}
//to file
void Convertor::PrintBlocksToFile()
{
    auto it_Names = Names.begin();
    auto it_Data_Block = DataBlock.begin();
    // Текущий json объект, с которым производится работа
    QJsonObject m_currentJsonObject;
    QJsonArray listArray;
    // Создаём объект текста
    int c = 0;
    int r = 0;
    int current_size;
    QJsonArray textsArray = m_currentJsonObject[it_Names->toString()].toArray();
    for (;it_Data_Block!=DataBlock.end();it_Data_Block++,c++) {
        r=DataBlock[c].size();

        for(vector<int>::iterator it2 = (*it_Data_Block).begin() ; it2 != (*it_Data_Block).end(),r-4 >= 0; ++it2 ,r -= 4){
            vector<int> tmp_v;
            int tmp = r-1;
            for (int var = 0; var < 4 ; var++) {
                if (tmp-var >= 0) {
                    textsArray.push_back(DataBlock[c][tmp-var]);
                };

            }

            listArray.push_back(textsArray);




            current_size =textsArray.size();
            for (int i = 0; i < current_size; i++) {

                for (int var = 0; var < current_size; var++) {
                    textsArray.removeAt(var);

                }
                m_currentJsonObject[it_Names->toString()] = listArray;

            }

            //
        }

        it_Names++;
        int size_=listArray.size();
        for (int i = 0; i < size_; i++) {
            listArray.pop_back();
        }
        // Добавляем объект текста в массив
        // Сохраняем массив обратно в текущий объект

    }


    QString saveFileName = QString::fromStdString(this->f_str_name.toStdString()+".json");

    // Создаём объект файла и открываем его на запись
    QFile jsonFile(saveFileName);
    if (!jsonFile.open(QIODevice::WriteOnly))
    {
        cout << "error" << endl;
        return;
    }

    // Записываем текущий объект Json в файл
    jsonFile.write(QJsonDocument(m_currentJsonObject).toJson(QJsonDocument::Indented));
    jsonFile.close();   // Закрываем файл
    cout << "file correct create" << endl;




}



