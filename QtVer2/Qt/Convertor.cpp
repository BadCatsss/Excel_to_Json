#include <QCoreApplication>
#include <QCoreApplication>
#include <QDebug>
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
#include "Convertor.h"

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
    this->maxCols = col;
    cout << "col " << this->maxCols << endl;

}
void Convertor::setActivetWorkSheet(QString chosenSheet)
{
    this->activeSheet = this->xlsxR->sheet(chosenSheet);
    this->xlsxR->selectSheet(this->activeSheet->sheetName());
}
QStringList  Convertor::getSheetsList()
{
    cout << "Sheet is found" << endl;
    QStringList sheetList;
    QTextStream qtout(stdout);
    for (auto sheet:  this->xlsxR->sheetNames())
    {
        sheetList.push_back(sheet);
        cout << sheetCount << "  ";
        qtout << sheet << endl;
        sheetCount++;
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
        QJsonArray arrayOfValue;
        for (int c = 2; c < maxCols; ++c) {
            if (this->xlsxR->cellAt(r,1)->readValue().toString() == "UMTS" && c == maxCols - 1 ) {
                arrayOfValue.append((this->xlsxR->cellAt(r,c)->readValue().toJsonValue()));
                arrayOfValue.append(30);
                arrayOfValue.append(30);
            }
            else {
                arrayOfValue.append((this->xlsxR->cellAt(r,c)->readValue().toJsonValue()));
            }

        }
        valeMap[this->xlsxR->cellAt(r,1)->readValue().toString()].append(arrayOfValue);
    }
}
void  Convertor ::createJsonObject()
{
    for (auto currentKey : valeMap.keys()) {
        this->valueJsonObject[currentKey] = valeMap.take(currentKey);
    }
}
void Convertor::writeJsonFile()
{
    QString saveFileName = QString::fromStdString(this->savePath + ".json");
    // Создаём объект файла и открываем его на запись
    QFile jsonFile(saveFileName);
    if ( !jsonFile.open(QIODevice::WriteOnly) )
    {
        cout << "write error" << endl;
        exit(-1);
    }
    // Записываем текущий объект Json в файл
    jsonFile.write(QJsonDocument(this->valueJsonObject).toJson(QJsonDocument::Indented));
    jsonFile.close();   // Закрываем файл
    cout << "file correct create" << endl;

}

