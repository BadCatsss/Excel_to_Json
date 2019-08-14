#include <QCoreApplication>
#include <QCoreApplication>
#include <QDebug>
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
#include "SettingsConverter.h"

using namespace QXlsx;
using namespace std;

void SettingsConverter:: setLastError(QString errorMessege)
{
    this->errorValue=errorMessege;
}
QString SettingsConverter::getLastError()
{
    return this->errorValue;
}
//book and sheet setup
bool SettingsConverter::openBook()
{
    string tmpString = filePath.toStdString();
    int startSearchPosition=tmpString.find(".");
    if (tmpString.find("xlsx",startSearchPosition)!=std::string::npos) {
        xlsxR = new QXlsx::Document(filePath);

        tmpString.erase(tmpString.begin() + tmpString.find("."), tmpString.end());
        this->savePath = tmpString;
    }
    else {
        setLastError("incorrect format");
        return false;
    }

    if (this->xlsxR->load())
    {
        return true;
    }
    else {
        setLastError("cant open file");
        return false;
    }
}
void SettingsConverter::calculateNotEmptyRowsCount( )
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
void SettingsConverter::calculateNotEmptyColumnsCount( )
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
void SettingsConverter::setActivetWorkSheet(QString chosenSheet)
{
    this->activeSheet = this->xlsxR->sheet(chosenSheet);
    this->xlsxR->selectSheet(this->activeSheet->sheetName());
}
QStringList  SettingsConverter::getSheetsList()
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
SettingsConverter::SettingsConverter(const QString& param)
{
    QFileInfo info(param);
    filePath = info.absoluteFilePath();
}
void SettingsConverter:: convert()
{
    this->calculateNotEmptyRowsCount();
    this->calculateNotEmptyColumnsCount();
    this->readXlsxFile();
    this->createJsonObject();
    this->writeJsonFile();
}
void SettingsConverter:: readXlsxFile()
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
void  SettingsConverter ::createJsonObject()
{
    for (auto currentKey : valeMap.keys()) {
        this->valueJsonObject[currentKey] = valeMap.take(currentKey);
    }
}
void SettingsConverter::writeJsonFile()
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

