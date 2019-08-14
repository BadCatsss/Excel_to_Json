#pragma once
#include <string>
#include <iostream>
#include <iomanip>
#include <map>
#include <QtCore>
#include <QJsonArray>
#include "xlsxdocument.h"

using namespace std;



class SettingsConverter
{
    size_t maxRows = 0;
    size_t maxCols = 0;
    QXlsx::Document* xlsxR = nullptr;
    QXlsx::AbstractSheet* activeSheet = nullptr;
    QString filePath = nullptr;
    string savePath = "";
    QString errorValue=nullptr;
    int sheetCount = 0;
    QMap<QString ,QJsonArray> valeMap;
    QJsonObject valueJsonObject;
    void calculateNotEmptyRowsCount();
    void calculateNotEmptyColumnsCount();
    void readXlsxFile();
    void createJsonObject();
    void writeJsonFile();
    void setLastError(QString error);


public:
    void convert();
    bool openBook();
    QStringList  getSheetsList();
    void setActivetWorkSheet(QString);
    SettingsConverter(const QString& p);
    QString getLastError();
};



