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
#include <fstream>
using namespace QXlsx;
using namespace std;


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
    while (cell != nullptr) {
        row++;
        cell = this->xlsxR->cellAt(row, 1);
    }
    this->maxRows = row;
    cout << "row " << this->maxRows << endl;
}
void Convertor::calculateNotEmptyColumnsCount( )
{
    int col = 1;
    Cell* cell = this->xlsxR->cellAt(1, col); // get cell pointer.
    while (cell != nullptr) {
        col++;
        cell = this->xlsxR->cellAt(1, col);
    }
    this->maxCols=col;
    cout << "col " << this->maxCols << endl;

}
void Convertor::setActivetWorkSheet(QString p)
{
    this->activeSheet = this->xlsxR->sheet(p);
    this->xlsxR->selectSheet(this->activeSheet->sheetName());
}
QStringList  Convertor::getSheetsList()
{
    cout << "Sheet is found" << endl;
    QStringList sheetList;
    QTextStream qtout(stdout);
    for (auto i:  this->xlsxR->sheetNames())
    {
        sheetList.push_back(i);
        cout << sheet_count << "  ";
        qtout << i << endl;
        sheet_count++;
    }
    return sheetList;
}

//pharse
Convertor::Convertor(const QString& param)
{
    QFileInfo info(param);
    filePath = info.absoluteFilePath();
    xlsxR = new QXlsx::Document(filePath);
    string tmpString = filePath.toStdString();
    tmpString.erase(tmpString.begin() + tmpString.find("."), tmpString.end());
    this->savePath = tmpString;
}
void Convertor:: convert()
{


    this->calculateNotEmptyRowsCount();
    this->calculateNotEmptyColumnsCount();
    this->readXlsxFile();
    this->createJsonObject();
    this->writeJsonFile();



}
void Convertor:: readXlsxFile()
{
    for (int r = 2; r < maxRows; ++r) {
        QJsonArray arr;
        for (int c = 2; c < maxCols; ++c) {
            if (this->xlsxR->cellAt(r,1)->readValue().toString() == "UMTS" && c == maxCols - 1 || c == maxCols - 2) {
                arr.append(30);
            }
            else {
                arr.append((this->xlsxR->cellAt(r,c)->readValue().toJsonValue()));
            }

        }
        valeMap[this->xlsxR->cellAt(r,1)->readValue().toString()].append(arr);
    }

}
void  Convertor ::createJsonObject()
{

    for (auto i : valeMap.keys()) {

        this->valueJsonObject[i] = valeMap.take(i);
    }

}
void Convertor::writeJsonFile()
{

    QString saveFileName = QString::fromStdString(this->savePath + ".json");

    // Создаём объект файла и открываем его на запись
    QFile jsonFile(saveFileName);
    if ( !jsonFile.open(QIODevice::WriteOnly) )
    {
        cout << "error" << endl;
        return;
    }

    // Записываем текущий объект Json в файл
    jsonFile.write(QJsonDocument(this->valueJsonObject).toJson(QJsonDocument::Indented));
    jsonFile.close();   // Закрываем файл
    cout << "file correct create" << endl;

}

