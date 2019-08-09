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

class Convertor:exception
{
private:
    /////parsing/////
    void GetDigitalStandartName();
    void GetData();
    void CreateBlockByName();
    void PrintBlocksToFile();
    ///////////////////////////


    QString f_str_name;
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
    void OpenBook();
    QXlsx::Document* GetBook() const;
    QStringList  OpenWorkSheet();
    void SetActivetWorkSheet(QString);
     QXlsx::AbstractSheet* GetActivetWorkSheet() const;
    int GetSheetCount() const;


    void SetMaxRows();
    void SetMaxColumns();
    void Generate();






    Convertor(QString f_str_name, bool f1, bool f2) :f_str_name(f_str_name), PrintDataFlag(f1), PrintNameFlag(f2),sheet_count(0)
    {
        xlsxR= new  QXlsx::Document(f_str_name);
    }




};
