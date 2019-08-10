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


    QString f_str_name;
    QString f_str_path;
    size_t maxRows;
    size_t maxCols;
    bool PrintDataFlag,PrintNameFlag;
    set<QVariant> Names;
    QMultiMap<QString, int >  Data;
    vector<vector<int>> DataBlock;
    QXlsx::Document* xlsxR;
    QXlsx::AbstractSheet* ActiveSheet;

    int sheet_count;
    void PrintNames();
    void PrintData();


public:
    static QString  ParsePath(QString path);
    void OpenBook();
    QXlsx::Document* GetBook() const;
    QStringList  OpenWorkSheet();
    void SetActivetWorkSheet(QString);
    QXlsx::AbstractSheet* GetActivetWorkSheet() const;
    int GetSheetCount() const;



    void SetMaxRows();
    void SetMaxColumns();
    void Generate();






    Convertor(QString f_str_path, bool f1, bool f2) ;





};
