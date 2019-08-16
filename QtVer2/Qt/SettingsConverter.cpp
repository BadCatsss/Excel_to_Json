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

void SettingsConverter:: addErrorToList(QString errorMessege)
{
    this->errorValues.push_back(errorMessege);
}
void SettingsConverter:: printErrorMesseges()
{
    for (auto errorMessege : this->errorValues) {
        cout << errorMessege.toStdString() << endl;
    }
}
//book and sheet setup
bool SettingsConverter::openBook()
{
    string tmpString = filePath.toStdString();
    int startSearchPosition = tmpString.find(".");
    if (tmpString.find("xlsx",startSearchPosition) != std::string::npos) {
        xlsxR = new QXlsx::Document(filePath);
        tmpString.erase(tmpString.begin() + tmpString.find("."), tmpString.end());
        this->savePath = tmpString;
    }
    else {
        addErrorToList("incorrect format");
        return false;
    }

    if (this->xlsxR->load())
    {
        return true;
    }
    else {
        addErrorToList("cant open file");
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
void SettingsConverter::setActivetWorkSheet(const QString& chosenSheet)
{
    this->activeSheet = this->xlsxR->sheet(chosenSheet);
    this->xlsxR->selectSheet(this->activeSheet->sheetName());
}
QStringList  SettingsConverter::getSheetsList()
{
    cout << "Sheet is found" << endl;
    QStringList sheetList;
    QTextStream qtout(stdout);
    for (auto sheet : this->xlsxR->sheetNames())
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
bool SettingsConverter:: convert()
{
    this->calculateNotEmptyRowsCount();
    this->calculateNotEmptyColumnsCount();
    this->createJsonObject();
    if ( !this->readXlsxFile() || !this->createJsonObject() || !this->writeJsonFile()) {
        this->addErrorToList("convert error");
        return false;
    }
    else {
        return true;
    }


}
bool SettingsConverter:: readXlsxFile()
{


    bool valueIsCorrect = true;
    for ( int r = 2; r < maxRows; ++r) {
        QJsonArray arrayOfValue;
        for (  int c = 2; c < maxCols; ++c) {
            if (this->xlsxR->cellAt(r,1)->readValue().toString() == "UMTS" && c == maxCols - 1 ) {
                arrayOfValue.append((this->xlsxR->cellAt(r,c)->readValue().toJsonValue()));
                arrayOfValue.append(30);
                arrayOfValue.append(30);
            }
            else {
                arrayOfValue.append((this->xlsxR->cellAt(r,c)->readValue().toJsonValue()));
            }
        }
        valuesMap[this->xlsxR->cellAt(r,1)->readValue().toString()].append(arrayOfValue);
    }

    for (auto currentArray : valuesMap) {
        for (auto currentValue : currentArray) {
            if (currentValue.isNull()) {
                valueIsCorrect = false;
            }
        }
    }
    if (valueIsCorrect && valuesMap.size() != 0) {
        return true;
    }
    else {
        this->addErrorToList("read error");
        return false;
    }
}
bool  SettingsConverter::createJsonObject()
{
    for (auto currentKey : valuesMap.keys()) {
        this->valueJsonObject[currentKey] = valuesMap.take(currentKey);
    }
    if (this->valueJsonObject.size() != 0) {
        return true;
    }
    else {
        this->addErrorToList("JsonObject create error");
        return false;
    }
}
bool SettingsConverter::writeJsonFile()
{
    QString saveFileName = QString::fromStdString(this->savePath + ".json");
    // Создаём объект файла и открываем его на запись
    QFile jsonFile(saveFileName);
    if ( !jsonFile.open(QIODevice::WriteOnly) )
    {
        addErrorToList("write error");
        return false;
    }
    // Записываем текущий объект Json в файл
    jsonFile.write(QJsonDocument(this->valueJsonObject).toJson(QJsonDocument::Indented));
    jsonFile.close();   // Закрываем файл
    cout << "file correct create" << endl;
    return true;
}

