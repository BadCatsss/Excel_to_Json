#pragma once
#include <string>
#include <set>
#include <iostream>
#include <iomanip>
#include <vector>
#include <map>
#include <QtCore>
#include "xlsxdocument.h"

using namespace std;

class Convertor
{
private:
    /////parsing/////
    void GetDigitalStandartName();
    void GetData();
    void CreateBlockByName();
    void PrintBlocksToFile();
    ///////////////////////////

    QString f_str_name = nullptr;
    QString f_str_path = nullptr;
    size_t maxRows = 0;
    size_t maxCols = 0;
    bool PrintDataFlag = false;
    bool PrintNameFlag = false;
    set<QVariant> Names;
    QMultiMap<QString, int > Data;
    vector<vector<int>> DataBlock;
    QXlsx::Document* xlsxR = nullptr;
    QXlsx::AbstractSheet* ActiveSheet = nullptr;
    int sheet_count = 0;

    void calculateNotEmptyRowsCount();
    void calculateNotEmptyColumnsCount();
    void printNamesToConsole();
    void printDataToConsole();
    QStringList  getSheetsList();
    void setActivetWorkSheet(QString);
//    int getSheetCount() const;

public:
    bool openBook();
    void convert();
    Convertor(const QString& f_str_path, bool PrintDataFlag, bool PrintNameFlag);
};
