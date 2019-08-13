#pragma once
#include <string>
#include <iostream>
#include <iomanip>
#include <map>
#include <QtCore>
#include <QJsonArray>
#include "xlsxdocument.h"

using namespace std;



class Convertor
{
    size_t maxRows = 0;
    size_t maxCols = 0;
    QXlsx::Document* xlsxR = nullptr;
    QXlsx::AbstractSheet* activeSheet = nullptr;
    QString filePath=nullptr;
    string savePath=nullptr;
    int sheet_count = 0;
    QMap<QString ,QJsonArray> valeMap;
    QJsonObject valueJsonObject;

    QStringList  getSheetsList();
    bool openBook();
    void setActivetWorkSheet(QString);
    void calculateNotEmptyRowsCount();
    void calculateNotEmptyColumnsCount();
    void readXlsxFile();
    void createJsonObject();
    void writeJsonFile();






public:

    void convert();
    Convertor(const char* p);


};



